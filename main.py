# -*- coding: utf-8 -*-
"""
AI 文件整理工具 —— 第一步：分析文件，生成 Excel 总表

用法：python main.py

流程：扫描文件夹 → 提取文字内容 → DeepSeek 逐文件分析（关键词/总结）
     → DeepSeek 通盘全局分类（参考整个文件库）→ 生成 Excel 总表
特点：
  - 已分析过的文件会写入 cache.json，重复运行不再消耗 API 费用
  - 分类由 AI 根据关键词、总结和整个文件库通盘给出，不照抄原有文件夹名
  - 总表为矩阵布局：每个分类一列，每个文件一个格子，点击格子直接打开文件
  - 重跑时保留你在总表里手动改过的分类（按文件链接或 文件名+总结 匹配）
"""

import hashlib
import json
import math
import os
import sys
import time
from pathlib import Path
from urllib.parse import unquote

from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from markitdown import MarkItDown

# ---------------- 路径与常量 ----------------
# 数据目录：源码运行时用脚本所在目录；打包成 exe 后自动改用 exe 所在目录，
# 这样用户把 exe 拷到哪里，配置/缓存/日志就存在哪里（可整体分发）。
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
CACHE_FILE = BASE_DIR / "cache.json"
GLOBAL_KEY = "__global__"  # cache.json 里存全局分类结果的固定键

SUPPORTED_EXTS = {".doc", ".docx", ".pdf", ".ppt", ".pptx", ".txt"}
MAX_TEXT_CHARS = 4000  # 发给 AI 的内容上限，控制 token 消耗
# 失败类结果不进缓存（或视为需要重试）：否则 API Key 失效一次，
# 失败结果会被永久缓存，之后每次运行都"缓存命中"直接跳过，永远无法恢复
FAIL_CATEGORIES = {"分析失败", "解析失败", "无法识别"}


# ---------------- 配置与缓存 ----------------
def load_config():
    if not CONFIG_FILE.exists():
        return {}  # 首次使用：还没有配置文件，由图形界面引导填写
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        sys.exit(f"config.json 格式有误：{e}\n注意路径里的反斜杠要写成两个，例如 C:\\\\Users\\\\test")


def load_cache():
    if CACHE_FILE.exists():
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    return {}


def save_cache(cache):
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def cache_key(path):
    """用文件内容的 MD5 作为缓存标识：重命名、移动都不影响缓存，相同内容只分析一次。"""
    try:
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return legacy_key(path)  # 文件被占用读不了时退化为旧标识，后续解析会标为失败


def legacy_key(path):
    """旧版缓存标识：文件名+大小+修改时间（仅用于迁移旧缓存）。"""
    st = path.stat()
    return f"{path.name}|{st.st_size}|{int(st.st_mtime)}"


# ---------------- AI 分析 ----------------
def build_prompt(filename, text, existing_categories):
    cats = "、".join(sorted(existing_categories)) if existing_categories else "（暂无）"
    return f"""文件名：{filename}

文件内容（节选）：
{text}

已有分类：{cats}

请完成三件事：
1. 提取 3 个关键词；
2. 用一句话（30 字以内）总结这个文件的内容；
3. 给出一个分类名：2-6 个字的简短中文词，从内容领域或使用场景命名（如"AI学习""工作文档""烹饪食谱""医疗""合同"）。已有分类里有合适的就直接复用，都不合适再新建。

严格按以下 JSON 格式返回，不要输出其他内容：
{{"keywords": ["关键词1", "关键词2", "关键词3"], "summary": "一句话总结", "category": "分类名"}}"""


def normalize_result(data):
    keywords = data.get("keywords") or []
    if isinstance(keywords, str):
        keywords = [keywords]
    return {
        "keywords": [str(k) for k in keywords][:3],
        "summary": str(data.get("summary", "")).strip(),
        "category": str(data.get("category", "未分类")).strip() or "未分类",
    }


def analyze(client, converter, path, model, existing_categories):
    """提取文件文字并调用 DeepSeek 分析。任何一步失败都不会中断整体流程。"""
    try:
        text = (converter.convert(str(path)).text_content or "").strip()
    except Exception as e:
        return {"keywords": [], "summary": f"文件解析失败：{e}", "category": "解析失败"}
    if not text:
        return {"keywords": [], "summary": "未能提取到文字内容", "category": "无法识别"}

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "你是专业的文件整理助手，擅长快速理解文档内容并归类。始终输出合法 JSON。"},
                {"role": "user", "content": build_prompt(path.name, text[:MAX_TEXT_CHARS], existing_categories)},
            ],
            response_format={"type": "json_object"},
            temperature=0.3,
        )
        return normalize_result(json.loads(resp.choices[0].message.content))
    except Exception as e:
        return {"keywords": [], "summary": f"AI 分析失败：{e}", "category": "分析失败"}


# ---------------- Excel 生成 ----------------
def make_file_url(path):
    """文件的 file:/// 链接：单元格挂上它，点击即可用默认程序打开文件。"""
    return "file:///" + str(path).replace("\\", "/")


def norm_path(p):
    """路径标准化，用于匹配（Windows 不区分大小写）。"""
    return os.path.normcase(os.path.normpath(str(p)))


def path_from_cell(cell):
    """从单元格的超链接还原文件路径（兼容 Excel 重新保存后的百分号编码）。"""
    if not cell.hyperlink or not cell.hyperlink.target:
        return None
    t = cell.hyperlink.target
    if t.lower().startswith("file:///"):
        t = t[8:]
    return unquote(t)


def write_excel(excel_path, rows):
    """矩阵布局：每个分类一列，每个文件一个格子；格子本身是可点击的文件链接。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "文件总表"

    groups = {}
    for r in rows:
        groups.setdefault(r["category"], []).append(r)

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    wrap_top = Alignment(wrap_text=True, vertical="top")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 文件多的分类排前面，方便查看
    categories = sorted(groups, key=lambda c: (-len(groups[c]), c))
    row_lines = {}
    for col, category in enumerate(categories, start=1):
        h = ws.cell(row=1, column=col, value=category)
        h.fill = header_fill
        h.font = header_font
        h.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = 46

        for row_i, r in enumerate(sorted(groups[category], key=lambda x: x["name"].lower()), start=2):
            text = (f"{r['name']}\n"
                    f"关键词：{'、'.join(r['keywords'])}\n"
                    f"总结：{r['summary']}\n"
                    f"修改日期：{r['mtime_str']}")
            c = ws.cell(row=row_i, column=col, value=text)
            c.alignment = wrap_top
            c.border = border
            c.hyperlink = make_file_url(r["path"])  # 点击格子直接打开文件
            # 估算这个格子需要的行数（文件名和总结会折行）
            lines = math.ceil(len(r["name"]) / 20) + math.ceil(len(r["summary"]) / 20) + 2
            row_lines[row_i] = max(row_lines.get(row_i, 0), lines)

    for row_i, lines in row_lines.items():
        ws.row_dimensions[row_i].height = lines * 14 + 8
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    wb.save(excel_path)


# ---------------- 全局分类 ----------------
def fingerprint(keys):
    """文件库指纹：文件有增删改时自动重新做一次全局分类。"""
    return hashlib.md5("\n".join(sorted(keys)).encode("utf-8")).hexdigest()


def classify_globally(client, model, infos):
    """把全部文件的关键词+总结一次性发给 AI，通盘设计分类体系并逐文件分配。"""
    lines = [
        f"{i}. {r['name']} | 关键词：{'、'.join(r['keywords']) or '无'} | 总结：{r['summary']}"
        for i, r in enumerate(infos, 1)
    ]
    prompt = f"""这是一个文件夹里全部 {len(infos)} 个文件的信息：

{chr(10).join(lines)}

请你通盘考虑整个文件库，完成两件事：
1. 设计一套统一的分类体系：5-12 个分类，互斥、粒度适中；分类名为 2-8 个字的简短中文词，从内容领域或使用场景命名（如"AI学习""简历求职""考公资料""烹饪食谱"），要概括内容主题，不要照抄文件名或文件夹名；
2. 把每个文件归入最合适的一类。内容相似的文件必须归入同一个分类，分类名写法完全一致。实在无法归类的才用"未分类"。

严格按以下 JSON 格式返回，键是文件编号（每个文件都要有）：
{{"assignments": {{"1": "分类名", "2": "分类名"}}}}"""
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "你是专业的文件整理专家，擅长为整个文件库设计统一、清晰的分类体系。始终输出合法 JSON。"},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.3,
        )
        data = json.loads(resp.choices[0].message.content)
        raw = data.get("assignments", data)
        return {str(k): str(v).strip() for k, v in raw.items()}
    except Exception as e:
        print(f"全局分类失败，改用每个文件各自的建议分类。原因：{e}")
        return {}


def load_manual_categories(excel_path):
    """读取已有总表里人工调整过的分类，重跑时保留。
    返回两个字典：按路径匹配（新版矩阵布局）、按 文件名+总结 匹配（旧版行式布局）。"""
    by_path, by_namesum = {}, {}
    if not excel_path.exists():
        return by_path, by_namesum
    try:
        ws = load_workbook(excel_path).active
        headers = [c.value for c in ws[1]]
        if "分类" in headers and "文件名" in headers:
            # 旧版行式布局
            h = {v: i for i, v in enumerate(headers)}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[h["文件名"]] and row[h["分类"]]:
                    key = (str(row[h["文件名"]]), str(row[h["一句话总结"]] or ""))
                    by_namesum[key] = str(row[h["分类"]])
        else:
            # 新版矩阵布局：表头即分类，带超链接的格子即文件
            for col_cells in ws.iter_cols(min_row=1):
                category = col_cells[0].value
                if not category:
                    continue
                for c in col_cells[1:]:
                    p = path_from_cell(c)
                    if p:
                        by_path[norm_path(p)] = str(category)
    except Exception as e:
        # 读不出来就中止，而不是重新生成——否则你手动改过的分类会被静默覆盖
        sys.exit(f"无法读取已有总表：{e}\n请关闭 Excel/WPS 里打开的总表后重新运行。（为保护你手动改过的分类，本次未重新生成）")
    return by_path, by_namesum


# ---------------- 主流程 ----------------


def main():
    config = load_config()
    if not config.get("source_folder"):
        sys.exit("还没有配置要整理的文件夹。\n请在图形界面「基本设置」里填写文件夹路径并保存，或直接编辑 config.json 的 source_folder 字段。")
    folder = Path(config["source_folder"])
    if not folder.is_dir():
        sys.exit(f"找不到文件夹：{folder}\n请打开 config.json，把 source_folder 改成你要整理的文件夹路径。")

    files = sorted(
        p for p in folder.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS and not p.name.startswith("~$")
    )
    if not files:
        sys.exit("该文件夹里没有找到 Word / PDF / PPT / TXT 文件。")
    print(f"共找到 {len(files)} 个文件\n")

    cache = load_cache()
    keys = {}
    migrated = 0
    for fp in files:
        k = cache_key(fp)
        if k not in cache:
            old = legacy_key(fp)
            if old in cache:
                cache[k] = cache[old]  # 迁移旧缓存到新标识，避免重复调 API
                migrated += 1
        keys[fp] = k
    if migrated:
        save_cache(cache)
        print(f"已迁移旧版缓存 {migrated} 条\n")
    need_ai = [fp for fp in files
               if keys[fp] not in cache or cache[keys[fp]].get("category") in FAIL_CATEGORIES]

    fp_hash = fingerprint(list(keys.values()))
    global_entry = cache.get(GLOBAL_KEY, {})
    need_global = global_entry.get("fingerprint") != fp_hash

    client = converter = None
    if need_ai or need_global:
        if "在这里" in config.get("api_key", ""):
            sys.exit("请先在 config.json 里填入你的 DeepSeek API 密钥（api_key）。")
        client = OpenAI(api_key=config["api_key"], base_url=config.get("base_url", "https://api.deepseek.com"))
        if need_ai:
            converter = MarkItDown()

    model = config.get("model", "deepseek-chat")
    existing_categories = {v.get("category", "") for v in cache.values()} - {""}
    for i, fp in enumerate(files, 1):
        if keys[fp] in cache and cache[keys[fp]].get("category") not in FAIL_CATEGORIES:
            print(f"[{i}/{len(files)}] 缓存命中：{fp.name}")
            continue
        print(f"[{i}/{len(files)}] AI 分析中：{fp.name}")
        result = analyze(client, converter, fp, model, existing_categories)
        if result["category"] in FAIL_CATEGORIES:
            # 失败结果不落缓存：下次运行自动重试（失败条目不写盘）
            cache.pop(keys[fp], None)
            cache[keys[fp]] = result  # 仅内存中保留，供本次生成总表使用
            save_cache({k: v for k, v in cache.items()
                        if k == GLOBAL_KEY or v.get("category") not in FAIL_CATEGORIES})
        else:
            cache[keys[fp]] = result
            save_cache(cache)
        existing_categories.add(result["category"])

    # 第二步：全局分类——把全部文件的关键词和总结发给 AI，通盘设计分类体系
    if need_global:
        print("\n正在根据整个文件库做全局分类...")
        infos = [{"name": fp.name, **cache[keys[fp]]} for fp in files]
        mapping = classify_globally(client, model, infos)
        if mapping:
            cache[GLOBAL_KEY] = {"fingerprint": fp_hash, "mapping": mapping}
            save_cache(cache)
    else:
        mapping = global_entry.get("mapping", {})

    excel_path = folder / config.get("excel_name", "文件整理总表.xlsx")
    manual_path, manual_namesum = load_manual_categories(excel_path)

    rows = []
    n_manual = 0
    for i, fp in enumerate(files, 1):
        info = cache[keys[fp]]
        is_manual = norm_path(fp) in manual_path or (fp.name, info["summary"]) in manual_namesum
        if is_manual:
            n_manual += 1
        category = (manual_path.get(norm_path(fp))
                    or manual_namesum.get((fp.name, info["summary"]))
                    or mapping.get(str(i))
                    or info["category"])
        rows.append({
            "category": category,
            "name": fp.name,
            "keywords": info["keywords"],
            "summary": info["summary"],
            "mtime_str": time.strftime("%Y-%m-%d %H:%M", time.localtime(fp.stat().st_mtime)),
            "path": fp,
        })

    try:
        write_excel(excel_path, rows)
    except PermissionError:
        sys.exit("无法写入 Excel：总表正被打开，请先关闭它再重新运行。")

    print(f"\n完成！总表已生成：{excel_path}")
    print(f"分类来源：AI 全局分类 {len(files) - n_manual} 个，保留人工修改 {n_manual} 个")
    print("接下来：打开总表检查并修改「分类」列，然后运行  python organize.py  进行归档。")


if __name__ == "__main__":
    try:
        main()
    finally:
        # 仅交互式终端运行时提示按回车（网页界面调用时会设置 FO_SILENT 自动跳过）
        if not os.environ.get("FO_SILENT") and sys.stdin.isatty():
            input("\n按回车键退出...")
