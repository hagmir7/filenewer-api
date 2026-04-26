"""
Text / file comparison service functions.
"""

import logging

logger = logging.getLogger(__name__)


def compare_texts(
    text1: str,
    text2: str,
    compare_mode: str = "line",
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    ignore_blank_lines: bool = False,
    context_lines: int = 3,
    output_format: str = "unified",
) -> dict:
    """
    Compare two texts and return differences.

    Args:
        text1             : first text (original)
        text2             : second text (modified)
        compare_mode      : 'line'  → line by line diff
                            'word'  → word by word diff
                            'char'  → character by character diff
                            'sentence' → sentence by sentence diff
        ignore_case       : ignore case differences      (default: False)
        ignore_whitespace : ignore whitespace changes    (default: False)
        ignore_blank_lines: ignore blank line changes    (default: False)
        context_lines     : context lines around changes (default: 3)
        output_format     : 'unified'  → unified diff format
                            'html'     → HTML with highlights
                            'json'     → structured JSON diff
                            'side_by_side' → side by side comparison

    Returns:
        {
            'differences'   : list,
            'stats'         : dict,
            'output'        : str,
            'is_identical'  : bool,
            'similarity'    : float,
        }
    """
    import difflib
    import re

    if text1 is None or text2 is None:
        raise ValueError("Both text1 and text2 are required.")

    # ── Preprocess texts ──────────────────────────
    def preprocess(text: str) -> str:
        if ignore_case:
            text = text.lower()
        if ignore_whitespace:
            text = "\n".join(" ".join(line.split()) for line in text.splitlines())
        if ignore_blank_lines:
            text = "\n".join(line for line in text.splitlines() if line.strip())
        return text

    t1 = preprocess(text1)
    t2 = preprocess(text2)

    # ── Tokenize based on mode ─────────────────────
    def tokenize(text: str, mode: str) -> list:
        if mode == "line":
            return text.splitlines(keepends=True)
        elif mode == "word":
            return re.findall(r"\S+|\s+", text)
        elif mode == "char":
            return list(text)
        elif mode == "sentence":
            sentences = re.split(r"(?<=[.!?])\s+", text)
            return [s + "\n" for s in sentences if s.strip()]
        return text.splitlines(keepends=True)

    tokens1 = tokenize(t1, compare_mode)
    tokens2 = tokenize(t2, compare_mode)

    # ── Compute diff ──────────────────────────────
    matcher = difflib.SequenceMatcher(None, tokens1, tokens2)
    opcodes = matcher.get_opcodes()
    similarity = round(matcher.ratio() * 100, 2)
    is_identical = similarity == 100.0

    # ── Build differences list ─────────────────────
    differences = []
    added = 0
    removed = 0
    changed = 0
    unchanged = 0

    for tag, i1, i2, j1, j2 in opcodes:
        old_content = "".join(tokens1[i1:i2])
        new_content = "".join(tokens2[j1:j2])

        if tag == "equal":
            unchanged += len(tokens1[i1:i2])
            differences.append(
                {
                    "type": "equal",
                    "old_start": i1 + 1,
                    "old_end": i2,
                    "new_start": j1 + 1,
                    "new_end": j2,
                    "old_content": old_content,
                    "new_content": new_content,
                }
            )
        elif tag == "insert":
            added += len(tokens2[j1:j2])
            differences.append(
                {
                    "type": "insert",
                    "old_start": i1 + 1,
                    "old_end": i2,
                    "new_start": j1 + 1,
                    "new_end": j2,
                    "old_content": "",
                    "new_content": new_content,
                }
            )
        elif tag == "delete":
            removed += len(tokens1[i1:i2])
            differences.append(
                {
                    "type": "delete",
                    "old_start": i1 + 1,
                    "old_end": i2,
                    "new_start": j1 + 1,
                    "new_end": j2,
                    "old_content": old_content,
                    "new_content": "",
                }
            )
        elif tag == "replace":
            changed += 1
            differences.append(
                {
                    "type": "replace",
                    "old_start": i1 + 1,
                    "old_end": i2,
                    "new_start": j1 + 1,
                    "new_end": j2,
                    "old_content": old_content,
                    "new_content": new_content,
                }
            )

    # ── Stats ──────────────────────────────────────
    stats = {
        "added": added,
        "removed": removed,
        "changed": changed,
        "unchanged": unchanged,
        "total_changes": added + removed + changed,
        "text1_length": len(text1),
        "text2_length": len(text2),
        "text1_lines": len(text1.splitlines()),
        "text2_lines": len(text2.splitlines()),
        "text1_words": len(text1.split()),
        "text2_words": len(text2.split()),
        "text1_chars": len(text1),
        "text2_chars": len(text2),
        "similarity": similarity,
        "is_identical": is_identical,
        "compare_mode": compare_mode,
    }

    # ── Generate output ────────────────────────────
    output = ""

    if output_format == "unified":
        output = _generate_unified_diff(
            text1,
            text2,
            context_lines=context_lines,
            ignore_case=ignore_case,
        )

    elif output_format == "html":
        output = _generate_html_diff(text1, text2)

    elif output_format == "side_by_side":
        output = _generate_side_by_side(
            text1,
            text2,
            context_lines=context_lines,
        )

    elif output_format == "json":
        output = differences

    return {
        "differences": differences,
        "stats": stats,
        "output": output,
        "is_identical": is_identical,
        "similarity": similarity,
        "output_format": output_format,
        "compare_mode": compare_mode,
    }


def _generate_unified_diff(
    text1: str,
    text2: str,
    context_lines: int = 3,
    ignore_case: bool = False,
) -> str:
    """Generate unified diff format (like git diff)."""
    import difflib

    t1 = text1.lower() if ignore_case else text1
    t2 = text2.lower() if ignore_case else text2

    lines1 = t1.splitlines(keepends=True)
    lines2 = t2.splitlines(keepends=True)

    diff = difflib.unified_diff(
        lines1,
        lines2,
        fromfile="text1 (original)",
        tofile="text2 (modified)",
        n=context_lines,
        lineterm="",
    )
    return "\n".join(list(diff))


def _generate_html_diff(text1: str, text2: str) -> str:
    """Generate HTML diff with highlighted changes."""
    import difflib

    lines1 = text1.splitlines(keepends=True)
    lines2 = text2.splitlines(keepends=True)

    differ = difflib.HtmlDiff(
        tabsize=4,
        wrapcolumn=80,
    )
    html = differ.make_table(
        lines1,
        lines2,
        fromdesc="Original",
        todesc="Modified",
        context=True,
        numlines=3,
    )
    return html


def _generate_side_by_side(
    text1: str,
    text2: str,
    context_lines: int = 3,
    width: int = 60,
) -> list:
    """Generate side-by-side comparison."""
    import difflib

    lines1 = text1.splitlines()
    lines2 = text2.splitlines()

    matcher = difflib.SequenceMatcher(None, lines1, lines2)
    result = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for l1, l2 in zip(lines1[i1:i2], lines2[j1:j2]):
                result.append(
                    {
                        "type": "equal",
                        "left": l1,
                        "right": l2,
                    }
                )
        elif tag == "insert":
            for l2 in lines2[j1:j2]:
                result.append(
                    {
                        "type": "insert",
                        "left": "",
                        "right": l2,
                    }
                )
        elif tag == "delete":
            for l1 in lines1[i1:i2]:
                result.append(
                    {
                        "type": "delete",
                        "left": l1,
                        "right": "",
                    }
                )
        elif tag == "replace":
            max_len = max(i2 - i1, j2 - j1)
            l1_lines = lines1[i1:i2] + [""] * (max_len - (i2 - i1))
            l2_lines = lines2[j1:j2] + [""] * (max_len - (j2 - j1))
            for l1, l2 in zip(l1_lines, l2_lines):
                result.append(
                    {
                        "type": "replace",
                        "left": l1,
                        "right": l2,
                    }
                )

    return result


def compare_files(
    source1,
    source2,
    encoding: str = "utf-8",
    compare_mode: str = "line",
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    output_format: str = "unified",
    context_lines: int = 3,
) -> dict:
    """
    Compare two text files.

    Args:
        source1   : first file object OR bytes
        source2   : second file object OR bytes
        encoding  : file encoding           (default: utf-8)
        ...rest same as compare_texts()

    Returns:
        Same as compare_texts() + file metadata
    """

    # ── Read files ────────────────────────────────
    def read_source(src, name: str) -> tuple:
        if hasattr(src, "read"):
            raw = src.read()
            filename = getattr(src, "name", name)
        elif isinstance(src, bytes):
            raw = src
            filename = name
        else:
            raise ValueError(f"Invalid source: {name}")

        if isinstance(raw, bytes):
            text = raw.decode(encoding, errors="replace")
        else:
            text = raw

        return text, filename, len(raw)

    text1, fname1, size1 = read_source(source1, "file1.txt")
    text2, fname2, size2 = read_source(source2, "file2.txt")

    # ── Compare ───────────────────────────────────
    result = compare_texts(
        text1,
        text2,
        compare_mode=compare_mode,
        ignore_case=ignore_case,
        ignore_whitespace=ignore_whitespace,
        output_format=output_format,
        context_lines=context_lines,
    )

    # ── Add file metadata ──────────────────────────
    result["file1"] = {
        "name": fname1,
        "size_kb": round(size1 / 1024, 2),
        "lines": result["stats"]["text1_lines"],
        "words": result["stats"]["text1_words"],
        "chars": result["stats"]["text1_chars"],
    }
    result["file2"] = {
        "name": fname2,
        "size_kb": round(size2 / 1024, 2),
        "lines": result["stats"]["text2_lines"],
        "words": result["stats"]["text2_words"],
        "chars": result["stats"]["text2_chars"],
    }

    return result
