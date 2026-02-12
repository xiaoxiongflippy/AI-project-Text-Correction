import argparse
import re
import unicodedata
from dataclasses import dataclass

from export_utils import export_to_pdf, export_to_word


ZERO_WIDTH_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")
EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001F6FF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA70-\U0001FAFF"
    "\U00002600-\U000026FF"
    "\U00002700-\U000027BF"
    "]+",
    flags=re.UNICODE,
)


@dataclass
class CleanOptions:
    remove_markdown: bool = True
    normalize_punctuation: bool = True
    normalize_whitespace: bool = True
    merge_wrapped_lines: bool = True
    remove_emoji: bool = False
    indent_paragraph: bool = True
    keep_tables: bool = True


PARAGRAPH_INDENT = "　　"


PUNCT_MAP = {
    "“": '"',
    "”": '"',
    "‘": "'",
    "’": "'",
    "—": "-",
    "–": "-",
    "…": "...",
    "【": "[",
    "】": "]",
    "（": "(",
    "）": ")",
    "，": ",",
    "。": ".",
    "；": ";",
    "：": ":",
    "！": "!",
    "？": "?",
    "、": ",",
}


LIST_MARKER_PATTERN = re.compile(r"^\s*(?:[-*+•]|\d+[.)])\s+")
CODE_STRONG_PATTERNS = (
    re.compile(r"^(?:from\s+[A-Za-z_][\w.]*\s+import\b|import\s+[A-Za-z_][\w.]*)"),
    re.compile(r"^(?:class|def)\s+[A-Za-z_]\w*"),
    re.compile(r"^(?:if|elif|else|for|while|try|except|finally|with)\b.*:\s*$"),
    re.compile(r"^(?:return|yield|break|continue|pass)\b"),
    re.compile(r"^[A-Za-z_][A-Za-z0-9_]*\s*=\s*.+"),
    re.compile(r"^[A-Za-z_][\w.]*\([^)]*\)\s*$"),
    re.compile(r"^@[A-Za-z_][\w.]*"),
)
CODE_SYMBOL_PATTERN = re.compile(r"[{}\[\]=;]")
CN_PUNCT_HINT_PATTERN = re.compile(r"[，。！？；：、“”‘’（）【】《》]")
FENCED_CODE_PATTERN = re.compile(r"^\s*```")


def is_code_line_candidate(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if any(pattern.match(stripped) for pattern in CODE_STRONG_PATTERNS):
        return True
    if CODE_SYMBOL_PATTERN.search(stripped) and not CN_PUNCT_HINT_PATTERN.search(stripped):
        return True
    return False


def detect_code_line_indexes(lines: list[str]) -> set[int]:
    code_indexes: set[int] = set()
    weak_indexes: set[int] = set()
    in_fenced_code = False

    for index, raw in enumerate(lines):
        stripped = raw.strip()

        if FENCED_CODE_PATTERN.match(stripped):
            in_fenced_code = not in_fenced_code
            continue

        if in_fenced_code:
            code_indexes.add(index)
            continue

        if not stripped:
            continue

        if is_code_line_candidate(raw):
            code_indexes.add(index)
            continue

        if raw.startswith(("    ", "\t")):
            weak_indexes.add(index)
            continue

        if stripped.startswith("#"):
            weak_indexes.add(index)

    changed = True
    while changed:
        changed = False
        for index in list(weak_indexes):
            near_code = (index - 1 in code_indexes) or (index + 1 in code_indexes)
            skip_blank_before = (
                index - 2 in code_indexes
                and index - 1 >= 0
                and not lines[index - 1].strip()
            )
            skip_blank_after = (
                index + 2 in code_indexes
                and index + 1 < len(lines)
                and not lines[index + 1].strip()
            )
            if near_code or skip_blank_before or skip_blank_after:
                code_indexes.add(index)
                weak_indexes.remove(index)
                changed = True

    return code_indexes


def clean_text(raw_text: str, options: CleanOptions) -> str:
    text = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    text = ZERO_WIDTH_PATTERN.sub("", text)

    text = normalize_list_markers(text)
    text = remove_repeated_noise(text)

    if options.remove_markdown:
        text = strip_markdown(text, keep_tables=options.keep_tables)

    if options.keep_tables:
        text = normalize_table_blocks(text)
    else:
        text = convert_table_blocks_to_bullets(text)

    if options.normalize_punctuation:
        text = normalize_punctuation(text)

    if options.remove_emoji:
        text = EMOJI_PATTERN.sub("", text)

    if options.normalize_whitespace:
        text = normalize_whitespace(text)

    text = normalize_cjk_latin_spacing(text)

    if options.merge_wrapped_lines:
        text = merge_lines(text)

    if options.indent_paragraph:
        text = indent_paragraphs(text, options.merge_wrapped_lines)

    return text.strip("\n\r")


def strip_markdown(text: str, keep_tables: bool = False) -> str:
    lines = text.split("\n")
    table_lines: set[int] = set()
    code_lines = detect_code_line_indexes(lines)
    for i, line in enumerate(lines):
        if is_table_row_line(line) or is_table_separator_line(line):
            table_lines.add(i)

    non_table_parts: list[str] = []
    result_lines: list[str] = []
    for i, line in enumerate(lines):
        if i in table_lines or i in code_lines:
            if non_table_parts:
                result_lines.append(_strip_markdown_inline("\n".join(non_table_parts)))
                non_table_parts.clear()
            result_lines.append(line)
        else:
            non_table_parts.append(line)
    if non_table_parts:
        result_lines.append(_strip_markdown_inline("\n".join(non_table_parts)))

    return "\n".join(result_lines)


def _strip_markdown_inline(text: str) -> str:
    text = re.sub(r"^\s{0,3}#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*>\s?", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*([-*_])(?:\s*\1){2,}\s*$", "", text, flags=re.MULTILINE)
    text = re.sub(r"```(?:[\w+-]+)?\n?", "", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"__([^_]+)__", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = re.sub(r"_([^_]+)_", r"\1", text)
    text = re.sub(r"!\[[^\]]*\]\([^\)]*\)", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^\)]*\)", r"\1", text)
    text = re.sub(r"^\s*[-*+]\s+", "• ", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*(\d+)[.)]\s+", r"\1. ", text, flags=re.MULTILINE)
    return text


def strip_markdown_tables(text: str) -> str:
    lines = text.split("\n")
    cleaned_lines = []
    for line in lines:
        if re.match(r"^\s*\|?\s*[-:]+\s*(\|\s*[-:]+\s*)+\|?\s*$", line):
            continue
        if "|" in line and line.count("|") >= 2:
            cells = [cell.strip() for cell in line.strip("|").split("|")]
            cleaned_lines.append(" | ".join(cells))
        else:
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def normalize_table_blocks(text: str) -> str:
    lines = text.split("\n")
    normalized = []
    index = 0

    while index < len(lines):
        line = lines[index]
        if not is_table_row_line(line):
            normalized.append(line)
            index += 1
            continue

        block = []
        while index < len(lines) and is_table_row_line(lines[index]):
            block.append(lines[index])
            index += 1

        normalized.extend(normalize_table_block(block))

    return "\n".join(normalized)


def normalize_table_block(lines: list[str]) -> list[str]:
    rows, has_header = parse_table_block(lines)
    if not rows:
        return lines

    col_count = max(len(row) for row in rows)
    rows = [row + [""] * (col_count - len(row)) for row in rows]

    normalized = [format_table_row(row) for row in rows]

    if has_header or len(normalized) >= 2:
        normalized.insert(1, format_table_separator(col_count))

    return normalized


def parse_table_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    parts = [part.strip() for part in stripped.split("|")]
    return parts


def format_table_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def format_table_separator(col_count: int) -> str:
    return "| " + " | ".join(["---"] * col_count) + " |"


def is_table_separator_line(line: str) -> bool:
    stripped = line.strip()
    if "|" not in stripped:
        return False
    cells = [cell.strip() for cell in stripped.strip("|").split("|")]
    if not cells:
        return False
    for cell in cells:
        if not cell:
            return False
        if not re.fullmatch(r":?-{3,}:?", cell):
            return False
    return True


HEADER_HINTS = {
    "网站",
    "网址",
    "核心用途",
    "优势",
    "用途",
    "说明",
    "备注",
    "类型",
    "名称",
    "标题",
    "平台",
    "link",
    "url",
    "website",
    "purpose",
    "usage",
    "advantage",
    "benefit",
}

TITLE_HINTS = {"网站", "名称", "标题", "平台", "站点", "资源"}
URL_HINTS = {"网址", "链接", "url", "URL", "Link", "link"}


def parse_table_block(lines: list[str]) -> tuple[list[list[str]], bool]:
    rows = []
    has_header = False

    for raw in lines:
        if is_table_separator_line(raw):
            has_header = True
            continue
        rows.append(parse_table_row(raw))

    if not rows:
        return [], False

    if not has_header and is_probable_header(rows[0], rows[1:]):
        has_header = True

    return rows, has_header


def is_probable_header(header: list[str], data_rows: list[list[str]]) -> bool:
    header_text = " ".join(header).strip().lower()
    if any(hint in header_text for hint in HEADER_HINTS):
        return True
    if any("http" in cell.lower() for cell in header):
        return False
    if all(len(cell.strip()) <= 10 for cell in header if cell.strip()) and len(data_rows) >= 1:
        return True
    return False


def convert_table_blocks_to_bullets(text: str) -> str:
    lines = text.split("\n")
    converted = []
    index = 0

    while index < len(lines):
        line = lines[index]
        if not is_table_row_line(line):
            converted.append(line)
            index += 1
            continue

        block = []
        while index < len(lines) and is_table_row_line(lines[index]):
            block.append(lines[index])
            index += 1

        converted.extend(table_block_to_bullets(block))

    return "\n".join(converted)


def table_block_to_bullets(lines: list[str]) -> list[str]:
    rows, has_header = parse_table_block(lines)
    if not rows:
        return lines

    header = rows[0] if has_header else []
    data_rows = rows[1:] if has_header else rows
    if header:
        max_cols = max(len(header), max(len(row) for row in data_rows))
        header = header + [""] * (max_cols - len(header))
    else:
        max_cols = max(len(row) for row in data_rows)
        header = ["" for _ in range(max_cols)]

    bullets = []
    for row in data_rows:
        row = row + [""] * (max_cols - len(row))
        line = format_table_bullet(header, row)
        if line:
            bullets.append(line)

    return bullets


def format_table_bullet(header: list[str], row: list[str]) -> str:
    header_clean = [cell.strip() for cell in header]
    row_clean = [cell.strip() for cell in row]

    title_index = None
    for index, cell in enumerate(header_clean):
        if any(hint in cell for hint in TITLE_HINTS):
            title_index = index
            break

    url_index = None
    for index, cell in enumerate(header_clean):
        if any(hint.lower() in cell.lower() for hint in URL_HINTS):
            url_index = index
            break

    title = row_clean[title_index] if title_index is not None else ""
    url = row_clean[url_index] if url_index is not None else ""

    details = []
    for index, (label, value) in enumerate(zip(header_clean, row_clean)):
        if index == title_index or index == url_index:
            continue
        if not value:
            continue
        if label:
            details.append(f"{label}：{value}")
        else:
            details.append(value)

    if title:
        line = f"• {title}"
        if url:
            line += f"（{url}）"
        if details:
            line += " — " + "；".join(details)
        return line

    parts = []
    for label, value in zip(header_clean, row_clean):
        if not value:
            continue
        if label:
            parts.append(f"{label}：{value}")
        else:
            parts.append(value)
    if not parts:
        return ""
    return "• " + "；".join(parts)


def normalize_punctuation(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    normalized_lines = []
    for index, line in enumerate(lines):
        if index in code_lines:
            normalized_lines.append(line)
            continue
        normalized = unicodedata.normalize("NFKC", line)
        for source, target in PUNCT_MAP.items():
            normalized = normalized.replace(source, target)
        normalized_lines.append(normalized)
    return "\n".join(normalized_lines)


def normalize_whitespace(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    normalized_lines = []
    for index, raw in enumerate(lines):
        line = raw.replace("\u3000", " ").replace("\u00a0", " ")
        if index in code_lines:
            normalized_lines.append(line.rstrip())
            continue
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r" +([,.;:!?])", r"\1", line)
        normalized_lines.append(line)
    text = "\n".join(normalized_lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def normalize_list_markers(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    normalized = []

    for index, raw in enumerate(lines):
        if index in code_lines:
            normalized.append(raw.rstrip())
            continue

        line = raw.strip()
        if not line:
            normalized.append("")
            continue

        if re.match(r"^([-*_])(?:\s*\1){2,}$", line):
            normalized.append(line)
            continue

        bullet = re.match(r"^(?:[-*+•·●▪◦‣⁃])\s*(.+)$", line)
        if bullet:
            normalized.append(f"• {bullet.group(1).strip()}")
            continue

        ordered = re.match(r"^[（(]?(\d{1,3})[)）.、]\s*(.+)$", line)
        if ordered:
            normalized.append(f"{ordered.group(1)}. {ordered.group(2).strip()}")
            continue

        normalized.append(line)

    return "\n".join(normalized)


def remove_repeated_noise(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    cleaned_lines = []

    for index, raw in enumerate(lines):
        if index in code_lines:
            cleaned_lines.append(raw.rstrip())
            continue

        line = raw.strip()
        if not line:
            cleaned_lines.append("")
            continue

        if is_table_row_line(line) or is_table_separator_line(line):
            cleaned_lines.append(line)
            continue

        line = re.sub(r"([:：])(?:\s*[^:：\n]{1,20}\1){2,}", r"\1", line)
        line = re.sub(r"^((?:[^:：\n]{1,20})[:：])(?:\1){1,}", r"\1", line)
        line = re.sub(r"([。.!?！？；;，,、])(?:\s*\1)+", r"\1", line)
        line = re.sub(r"([\w\u4e00-\u9fff]{2,20})(?:\s*[，,;；]\s*\1){1,}", r"\1", line)

        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def normalize_cjk_latin_spacing(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    normalized_lines = []
    for index, line in enumerate(lines):
        if index in code_lines:
            normalized_lines.append(line)
            continue
        normalized = re.sub(r"([\u4e00-\u9fff])([A-Za-z0-9])", r"\1 \2", line)
        normalized = re.sub(r"([A-Za-z0-9])([\u4e00-\u9fff])", r"\1 \2", normalized)
        normalized = re.sub(r"([A-Za-z])(\d)", r"\1 \2", normalized)
        normalized = re.sub(r"(\d)([A-Za-z])", r"\1 \2", normalized)
        normalized_lines.append(normalized)
    return "\n".join(normalized_lines)


def merge_lines(text: str) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    merged = []
    buffer = ""

    for index, current in enumerate(lines):
        if index in code_lines:
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            merged.append(current.rstrip())
            continue

        line = current.strip()
        if not line:
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            continue

        if is_table_row_line(line):
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            merged.append(line)
            continue

        if LIST_MARKER_PATTERN.match(line):
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            merged.append(line)
            continue

        if not buffer:
            buffer = line
            continue

        if should_break_paragraph(buffer, line):
            merged.append(buffer.strip())
            buffer = line
        else:
            buffer = f"{buffer}{joiner(buffer, line)}{line}"

    if buffer:
        merged.append(buffer.strip())

    text = "\n".join(merged)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def should_break_paragraph(previous: str, current: str) -> bool:
    if is_table_row_line(previous) or is_table_row_line(current):
        return True
    if re.match(r"^(?:\d+[.)、]\s+|•\s+|-\s+|[（(]\d+[)）])", current):
        return True
    if is_heading_like(previous) or is_heading_like(current):
        return True
    return bool(re.search(r"[.!?;:。！？；：]$", previous))


def is_heading_like(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if LIST_MARKER_PATTERN.match(stripped):
        return False
    if is_table_row_line(stripped):
        return False

    if len(stripped) <= 30 and re.search(r"[:：]$", stripped):
        return True

    if len(stripped) <= 24 and re.search(r"[:：]", stripped) and not re.search(r"[，,。.!?！？]$", stripped):
        parts = [item for item in re.split(r"[:：]", stripped) if item]
        if 1 < len(parts) <= 3:
            return True

    return False


def joiner(previous: str, current: str) -> str:
    if is_cjk(previous[-1]) and is_cjk(current[0]):
        return ""
    return " "


def is_cjk(char: str) -> bool:
    return "\u4e00" <= char <= "\u9fff"


def indent_paragraphs(text: str, treat_line_breaks_as_paragraphs: bool = False) -> str:
    lines = text.split("\n")
    code_lines = detect_code_line_indexes(lines)
    formatted_lines = []
    paragraph_start = True

    for index, raw in enumerate(lines):
        if index in code_lines:
            formatted_lines.append(raw.rstrip())
            paragraph_start = True
            continue

        line = raw.strip()
        if not line:
            formatted_lines.append("")
            paragraph_start = True
            continue

        if treat_line_breaks_as_paragraphs:
            if should_indent(line):
                line = f"{PARAGRAPH_INDENT}{line}"
            formatted_lines.append(line)
            paragraph_start = True
            continue

        if paragraph_start and should_indent(line):
            line = f"{PARAGRAPH_INDENT}{line}"

        formatted_lines.append(line)
        paragraph_start = False

    return "\n".join(formatted_lines)


def should_indent(line: str) -> bool:
    if is_heading_like(line):
        return False
    if is_table_row_line(line):
        return False
    if is_code_line_candidate(line) or line.startswith(("    ", "\t")):
        return False
    return not bool(re.match(r"^(?:\d+[.)、]\s+|•\s+|-\s+|[（(]\d+[)）])", line))


def is_table_row_line(line: str) -> bool:
    if "|" not in line:
        return False
    parts = [part.strip() for part in line.split("|") if part.strip()]
    return len(parts) >= 2


CH_PUNCT = set("，。！？；：、（）【】“”‘’《》")
EN_PUNCT = set(",.!?;:()[]\"'")


def punctuation_consistency_warnings(text: str) -> list[str]:
    lines = [line for line in text.split("\n") if line.strip()]
    if not lines:
        return []

    mixed_lines = 0
    for line in lines:
        has_ch = any(ch in CH_PUNCT for ch in line)
        has_en = any(ch in EN_PUNCT for ch in line)
        if has_ch and has_en:
            mixed_lines += 1

    warnings = []
    if mixed_lines:
        warnings.append(f"标点混用 {mixed_lines} 处")
    return warnings


def cli() -> None:
    parser = argparse.ArgumentParser(description="清理大模型回复文本中的符号和排版")
    parser.add_argument("--text", help="直接传入要清理的文本")
    parser.add_argument("--input", help="输入文件路径")
    parser.add_argument("--output", help="输出文件路径")
    parser.add_argument("--keep-markdown", action="store_true", help="保留 Markdown 标记")
    parser.add_argument("--keep-lines", action="store_true", help="保留原始换行")
    parser.add_argument("--remove-emoji", action="store_true", help="移除 Emoji")
    parser.add_argument("--no-indent", action="store_true", help="关闭段首空两格")
    parser.add_argument("--export-word", help="导出为 Word 文件（.docx）")
    parser.add_argument("--export-pdf", help="导出为 PDF 文件（.pdf）")
    args = parser.parse_args()

    source_text = ""
    if args.text:
        source_text = args.text
    elif args.input:
        with open(args.input, "r", encoding="utf-8") as file:
            source_text = file.read()
    else:
        source_text = input("粘贴待清理文本后回车：\n")

    options = CleanOptions(
        remove_markdown=not args.keep_markdown,
        merge_wrapped_lines=not args.keep_lines,
        remove_emoji=args.remove_emoji,
        indent_paragraph=not args.no_indent,
    )
    cleaned = clean_text(source_text, options)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as file:
            file.write(cleaned)
        print(f"清理完成，已写入: {args.output}")
    else:
        print(cleaned)

    if args.export_word:
        try:
            output_word = export_to_word(cleaned, args.export_word)
            print(f"Word 导出成功: {output_word}")
        except RuntimeError as error:
            print(str(error))

    if args.export_pdf:
        try:
            output_pdf = export_to_pdf(cleaned, args.export_pdf)
            print(f"PDF 导出成功: {output_pdf}")
        except RuntimeError as error:
            print(str(error))


if __name__ == "__main__":
    cli()
