#!/usr/bin/env python3
"""
法答网精选答问 爬虫脚本 (v3.0)
================================
数据源：人民法院报 (rmfyb.com) — 覆盖全部37批
功能：
1. 从 rmfyb.com 搜索并提取所有"法答网精选答问"批次链接
2. 逐个抓取每批次的详细内容
3. 生成格式化的 Markdown 文件，每批次一个文件
4. 支持增量更新：仅抓取新增批次（自动检测网站是否有新批次）

用法：
  python fetch_fadawang.py              # 增量模式：仅抓取新批次
  python fetch_fadawang.py --force      # 强制模式：重新抓取所有批次
  python fetch_fadawang.py --list       # 仅列出发现的批次
  python fetch_fadawang.py --batch 38   # 仅抓取指定批次

输出目录：脚本所在目录下的 md/ 子目录
格式：法答网精选答问（第XX批）.md

依赖：Python 3.8+, requests
"""

import re
import sys
import html as html_mod
import json
import time
import argparse
from pathlib import Path
from datetime import datetime

import requests

# 修复 Windows 终端编码问题
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ============================================================================
# 配置
# ============================================================================

BASE_URL = "https://www.rmfyb.com"
SEARCH_URL = f"{BASE_URL}/search.html"
SEARCH_KEYWORD = "法答网"
OUTPUT_DIR = Path(__file__).parent / "md"
META_FILE = Path(__file__).parent / ".fadawang_meta.json"
DELAY = 1.5  # 请求间隔（秒）
TIMEOUT = 30  # 请求超时（秒）
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# ============================================================================
# 工具函数
# ============================================================================

def fetch_page(url, params=None, encoding="utf-8"):
    """通用的页面抓取函数，含错误处理和重试"""
    for attempt in range(3):
        try:
            resp = requests.get(
                url, params=params, headers=HEADERS, timeout=TIMEOUT
            )
            resp.raise_for_status()
            if resp.apparent_encoding and "gb" in resp.apparent_encoding.lower():
                resp.encoding = "gbk"
            else:
                resp.encoding = encoding
            return resp.text
        except requests.RequestException as e:
            print(f"  [重试 {attempt+1}/3] 请求失败: {e}")
            if attempt < 2:
                time.sleep(DELAY * 2)
            else:
                raise
    return None


def safe_strip(text):
    """去除首尾空白，并压缩中间多余空白"""
    if not text:
        return ""
    text = html_mod.unescape(text)
    # 替换 &nbsp; 等
    text = text.replace(" ", " ").replace("　", " ")
    text = re.sub(r"&nbsp;", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_batch_number(title):
    """从标题中提取批次号。返回 int，失败返回 None"""
    clean_title = re.sub(r"<[^>]+>", "", title)
    cn_num_map = {
        "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
        "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
    }
    m = re.search(r"第([一二三四五六七八九十百零\d]+)批", clean_title)
    if not m:
        return None
    cn_num = m.group(1)
    if cn_num.isdigit():
        return int(cn_num)
    result = 0
    if "百" in cn_num:
        parts = cn_num.split("百")
        result += (cn_num_map.get(parts[0], 0) * 100) if parts[0] else 100
        cn_num = parts[1] if len(parts) > 1 else ""
    if "十" in cn_num:
        parts = cn_num.split("十")
        result += (cn_num_map.get(parts[0], 0) * 10) if parts[0] else 10
        if len(parts) > 1 and parts[1]:
            result += cn_num_map.get(parts[1], 0)
    elif cn_num in cn_num_map:
        result += cn_num_map[cn_num]
    return result if result > 0 else None


# ============================================================================
# 第一步：搜索所有法答网精选答问批次链接 (rmfyb.com)
# ============================================================================

def search_all_batches():
    """搜索 rmfyb.com 所有法答网精选答问批次，返回 {batch_num: {title, url, date}}"""
    batches = {}
    page = 1

    while True:
        print(f"  搜索第 {page} 页...")
        try:
            html_content = fetch_page(
                SEARCH_URL,
                params={
                    "wd": SEARCH_KEYWORD,
                    "search_range": "1",  # 标题搜索
                    "order": "2",         # 时间倒序
                    "begin_time": "",
                    "end_time": "",
                    "p": str(page),
                },
            )
        except Exception as e:
            print(f"  搜索页 {page} 请求出错: {e}")
            break

        # 提取总分页数（仅第一页需要）
        if page == 1:
            total_pages_m = re.search(r'total_page_num\s*=\s*"(\d+)"', html_content)
            total_pages = int(total_pages_m.group(1)) if total_pages_m else 1
            print(f"  共 {total_pages} 页搜索结果")
        else:
            total_pages_m = re.search(r'total_page_num\s*=\s*"(\d+)"', html_content)
            total_pages = int(total_pages_m.group(1)) if total_pages_m else page

        # 提取每个搜索结果
        # 结构: <a href="/content/..."> <div class='content_item'> <div class='item_title'>...</div> ... </div> </a>
        result_blocks = re.findall(
            r'<a\s+href="(/content/[^"]+)"[^>]*>(.*?)</a>',
            html_content, re.DOTALL
        )

        for url_path, block in result_blocks:
            # 提取标题
            title_m = re.search(r"<div class='item_title'>(.*?)</div>", block, re.DOTALL)
            if not title_m:
                continue
            title_raw = title_m.group(1).strip()
            title_clean = re.sub(r"<[^>]+>", "", title_raw).strip()

            # 只处理"法答网精选答问"标题
            if "法答网精选答问" not in title_clean:
                continue

            batch_num = extract_batch_number(title_clean)
            if batch_num is None:
                continue

            # 避免重复（同一批次可能出现在多个搜索结果中）
            if batch_num in batches:
                continue

            # 提取日期/版面信息
            page_m = re.search(r"<span class='page_num'>([^<]+)</span>", block)
            date_str = page_m.group(1).strip() if page_m else ""
            # 格式: "2025-11-06 07版" -> 提取日期部分
            publish_date = date_str.split(" ")[0] if date_str else ""

            batches[batch_num] = {
                "title": title_clean,
                "url": BASE_URL + url_path,
                "date": publish_date,
            }

        # 翻页判断
        if page >= total_pages:
            break
        page += 1
        time.sleep(DELAY)

    print(f"  共发现 {len(batches)} 批法答网精选答问")
    return batches


# ============================================================================
# 第二步：抓取单个批次详情 (rmfyb.com)
# ============================================================================

def fetch_batch_detail(batch_num, batch_info):
    """抓取单个批次详情，返回结构化字典"""
    url = batch_info["url"]
    print(f"  抓取第 {batch_num} 批: {url}")

    try:
        html_content = fetch_page(url)
    except Exception as e:
        print(f"    [FAIL] 请求失败: {e}")
        return None

    # 提取标题: <div class="title_nav_text">法答网精选答问（第XX批）</div>
    title_m = re.findall(
        r'<div class="title_nav_text">(.*?)</div>', html_content
    )
    page_title = ""
    for t in title_m:
        t_clean = re.sub(r"<[^>]+>", "", t).strip()
        if "法答网精选答问" in t_clean:
            page_title = t_clean
            break
    if not page_title:
        page_title = batch_info["title"]

    # 提取正文: <div class="conten_text_box"> ... <div align="left"> ... </div> ... </div>
    content_m = re.search(
        r'<div class="conten_text_box">(.*?)</div>\s*<!--文章结束-->',
        html_content, re.DOTALL
    )
    if not content_m:
        # 尝试更宽松匹配
        content_m = re.search(
            r'<div class="conten_text_box">(.*?)</div>\s*</div>\s*<!--\s*放大字体',
            html_content, re.DOTALL
        )
    if not content_m:
        content_m = re.search(
            r'<div class="conten_text_box">(.*?)</div>',
            html_content, re.DOTALL
        )

    if not content_m:
        print(f"    [WARN] 无法提取正文内容")
        return None

    raw_html = content_m.group(1)

    # 提取发布日期（从搜索结果的 date 字段）
    publish_date = batch_info.get("date", "")

    return {
        "batch_num": batch_num,
        "title": page_title,
        "url": url,
        "publish_date": publish_date,
        "raw_html": raw_html,
    }


# ============================================================================
# 第三步：HTML -> Markdown 转换
# ============================================================================

def html_to_markdown(detail):
    """将批次详情中的 HTML 内容转换为 Markdown"""
    raw_html = detail["raw_html"]
    lines = []

    # 文件头
    lines.append(f"# {detail['title']}")
    lines.append("")
    lines.append(f"> **来源：** 人民法院报 (rmfyb.com) - 法答网精选答问")
    lines.append(f"> **链接：** {detail['url']}")
    if detail["publish_date"]:
        lines.append(f"> **发布日期：** {detail['publish_date']}")
    lines.append(f"> **抓取时间：** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")

    # 清理 HTML：
    # 1. 移除图片标签
    text = re.sub(r'<img[^>]*/?>', '', raw_html)
    # 2. 移除 font 标签但保留内容
    text = re.sub(r'</?font[^>]*>', '', text)
    # 3. 移除 div 标签但保留内容
    text = re.sub(r'</?div[^>]*>', '', text)
    # 4. 将 <br> 和 <br/> 替换为换行
    text = re.sub(r'<br\s*/?>', '\n', text)
    # 5. 移除 <b> 和 </b> 标签（文本中用 ** 标记）
    text = re.sub(r'</?b>', '', text)
    # 6. 解码 HTML 实体
    text = html_mod.unescape(text)
    # 7. 处理 &nbsp; 和全角空格
    text = text.replace(' ', ' ').replace('　', ' ')
    text = re.sub(r'&nbsp;', ' ', text, flags=re.IGNORECASE)

    # 按行分割
    raw_lines = text.split('\n')

    first_question_seen = False
    in_opening = True

    for line in raw_lines:
        line = safe_strip(line)
        if not line:
            continue

        # 检测问题标题 (问题N：...)
        q_match = re.match(r'问题(\d+)[：:]\s*(.*)', line)
        if q_match:
            q_num = int(q_match.group(1))
            q_text = q_match.group(2) if q_match.group(2) else ""
            first_question_seen = True
            in_opening = False
            lines.append("")
            lines.append("---")
            lines.append("")
            if q_text:
                lines.append(f"## 问题{q_num}：{q_text}")
            else:
                # 问题标题可能跨行，先只放标题
                lines.append(f"## 问题{q_num}")
            lines.append("")
            continue

        # 检测开栏的话（在第一个问题之前的内容视为开栏语）
        if "开栏的话" in line:
            in_opening = True
            lines.append("> **开栏的话**")
            lines.append(">")
            continue

        # 检测答疑意见
        if line.startswith("答疑意见：") or line.startswith("答疑意见:"):
            in_opening = False
            answer_text = line[len("答疑意见："):].strip()
            if answer_text:
                lines.append(f"**答疑意见：** {answer_text}")
            else:
                lines.append(f"**答疑意见：**")
            lines.append("")
            continue

        # 检测答疑专家
        if line.startswith("答疑专家：") or line.startswith("答疑专家:"):
            expert_text = line[len("答疑专家："):].strip()
            lines.append(f"*答疑专家：{expert_text}*")
            lines.append("")
            continue

        # 检测咨询人
        if line.startswith("咨询人：") or line.startswith("咨询人:"):
            advisor_text = line[len("咨询人："):].strip()
            lines.append(f"*咨询人：{advisor_text}*")
            lines.append("")
            continue

        # 检测点评意见
        if line.startswith("点评意见：") or line.startswith("点评意见:"):
            review_text = line[len("点评意见："):].strip()
            lines.append(f"**点评意见：** {review_text}")
            lines.append("")
            continue

        # 检测点评专家
        if line.startswith("点评专家：") or line.startswith("点评专家:"):
            reviewer_text = line[len("点评专家："):].strip()
            lines.append(f"*点评专家：{reviewer_text}*")
            lines.append("")
            continue

        # 开栏部分的内容
        if in_opening and not first_question_seen:
            lines.append(f"> {line}")
            lines.append(">")
            continue

        # 普通段落
        lines.append(line)
        lines.append("")

    # 后处理
    markdown = "\n".join(lines)
    # 合并连续空行
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    # 清理空引用行 (">" 单独成行)
    markdown = re.sub(r"\n>\s*\n>", "\n>", markdown)

    return markdown


# ============================================================================
# 第四步：保存 Markdown 文件
# ============================================================================

def save_markdown(batch_num, detail, force=False):
    """保存 Markdown 文件"""
    output_dir = OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    filename = f"法答网精选答问（第{batch_num:02d}批）.md"
    filepath = output_dir / filename

    if filepath.exists() and not force:
        print(f"    [SKIP] 文件已存在: {filename}")
        return False

    markdown = html_to_markdown(detail)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(markdown)

    print(f"    [OK] 已保存: {filename} ({len(markdown)} 字符)")
    return True


# ============================================================================
# 元数据管理
# ============================================================================

def load_meta():
    """加载元数据"""
    if META_FILE.exists():
        try:
            with open(META_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return {
        "version": 3,
        "source": "rmfyb.com",
        "last_run": None,
        "processed_batches": {},
    }


def save_meta(meta):
    """保存元数据"""
    meta["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta["processed_batches"] = dict(
        sorted(meta["processed_batches"].items(), key=lambda x: int(x[0]))
    )
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def get_new_batches(batches, meta, force=False):
    """获取需要抓取的新批次"""
    if force:
        return set(batches.keys())
    processed = {int(k) for k in meta.get("processed_batches", {}).keys()}
    return set(batches.keys()) - processed


# ============================================================================
# 主流程
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="法答网精选答问爬虫 - 从人民法院报(rmfyb.com)抓取并生成 Markdown"
    )
    parser.add_argument("--force", action="store_true",
                        help="强制重新抓取所有批次（覆盖已有文件）")
    parser.add_argument("--batch", type=int, default=None,
                        help="仅抓取指定批次（如 --batch 38）")
    parser.add_argument("--list", action="store_true",
                        help="仅列出已发现/已抓取的批次，不抓取")
    args = parser.parse_args()

    print("=" * 60)
    print("  法答网精选答问 爬虫脚本 v3.0")
    print("  数据来源：人民法院报 (rmfyb.com)")
    print(f"  运行时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    print()

    # 加载元数据
    meta = load_meta()
    if meta.get("last_run"):
        print(f"[*] 上次运行时间：{meta['last_run']}")
        print(f"[*] 已抓取批次数量：{len(meta.get('processed_batches', {}))}")
    else:
        print("[*] 首次运行，将抓取所有批次")
    print()

    # ===== 第一步：搜索所有批次 =====
    print("[1/4] 搜索所有法答网精选答问批次...")
    batches = search_all_batches()

    if not batches:
        print("[FAIL] 未找到任何批次，请检查网络连接或网站是否可访问")
        sys.exit(1)

    if args.list:
        print("\n[*] 发现的批次列表：")
        all_nums = sorted(batches.keys())
        for num in all_nums:
            info = batches[num]
            processed = str(num) in meta.get("processed_batches", {})
            status = "[OK] 已抓取" if processed else "[NEW] 新批次"
            print(f"  {status}  第 {num:02d} 批: {info['title']}")

        # 检测批次编号缺口
        if len(all_nums) > 1:
            gaps = []
            for i in range(all_nums[0], all_nums[-1] + 1):
                if i not in batches:
                    gaps.append(i)
            if gaps:
                gap_str = ", ".join(f"第{g}批" for g in gaps)
                print(f"\n[WARN] 检测到批次编号缺口（可能网站未发布或已删除）：")
                print(f"  缺失批次: {gap_str}")
            else:
                print(f"\n[OK] 批次编号连续，无缺口。")

        print(f"\n共 {len(batches)} 批")
        return

    # ===== 第二步：筛选需抓取的批次 =====
    if args.batch is not None:
        if args.batch not in batches:
            print(f"[FAIL] 批次 {args.batch} 未找到")
            sys.exit(1)
        to_fetch = {args.batch}
    else:
        to_fetch = get_new_batches(batches, meta, force=args.force)

    if not to_fetch:
        print("[OK] 没有新批次需要抓取，知识库已是最新状态。")
        print(f"  (当前共 {len(batches)} 批，已全部抓取)")
        return

    print(f"\n[2/4] 开始抓取 {len(to_fetch)} 个批次...")
    print()

    success_count = 0
    fail_count = 0

    # 按批次号从大到小排序
    for batch_num in sorted(to_fetch, reverse=True):
        batch_info = batches[batch_num]
        print(f"  [{batch_num:02d}] {batch_info['title']}")

        detail = fetch_batch_detail(batch_num, batch_info)
        if detail is None:
            fail_count += 1
            continue

        try:
            saved = save_markdown(batch_num, detail, force=args.force)
            if saved or args.force:
                meta["processed_batches"][str(batch_num)] = {
                    "title": detail["title"],
                    "url": detail["url"],
                    "publish_date": detail["publish_date"],
                    "md_file": f"法答网精选答问（第{batch_num:02d}批）.md",
                    "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            success_count += 1
        except Exception as e:
            print(f"    [FAIL] 保存失败: {e}")
            fail_count += 1

        time.sleep(DELAY)

    # ===== 保存元数据 =====
    save_meta(meta)

    # ===== 汇总 =====
    print()
    print("=" * 60)
    print(f"  完成！")
    print(f"  总批次数：{len(batches)}")
    print(f"  本次抓取：{len(to_fetch)} 批")
    print(f"  成功：{success_count} | 失败：{fail_count}")
    print(f"  MD 文件目录：{OUTPUT_DIR.absolute()}")
    print(f"  元数据文件：{META_FILE.absolute()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
