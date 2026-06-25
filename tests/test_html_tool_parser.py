import re
from pathlib import Path

def parse_bundle_content(text: str) -> list[dict]:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks = []
    current_path = ""
    current_lines = []

    def flush_block():
        nonlocal current_path, current_lines
        if current_path:
            content = "\n".join(current_lines).strip()
            if content:
                blocks.append({"path": current_path, "text": content})

    for line in lines:
        trim_line = line.strip()

        if trim_line.startswith("### FILE:"):
            flush_block()
            current_path = trim_line[9:].strip()
            current_lines = []
            continue

        okf_match = re.search(r"<!--\s*OKF:BEGIN\s+concept=[\"']?([^\s\"'>]+)", trim_line, re.IGNORECASE)
        if okf_match:
            flush_block()
            current_path = okf_match.group(1).strip()
            current_lines = []
            continue

        if trim_line.startswith("### END FILE") or trim_line.startswith("<!-- OKF:END"):
            flush_block()
            current_path = ""
            current_lines = []
            continue

        if not current_path:
            continue

        current_lines.append(line)

    flush_block()
    return blocks


def test_kanban_example_parsing() -> None:
    example_path = Path(__file__).parent.parent / "_meta" / "KanBan Example"
    assert example_path.exists()
    
    text = example_path.read_text(encoding="utf-8")
    blocks = parse_bundle_content(text)
    
    assert len(blocks) == 15
    
    board_block = [b for b in blocks if b["path"].endswith("board.md")]
    task_blocks = [b for b in blocks if "tasks/" in b["path"]]
    
    assert len(board_block) == 1
    assert len(task_blocks) == 14
    
    assert board_block[0]["path"] == "projects/air-cargo-ai/board.md"
    assert task_blocks[0]["path"] == "projects/air-cargo-ai/tasks/T-0001.md"
    
    # Check that T-0001 contains the expected frontmatter type
    assert "type: Task" in task_blocks[0]["text"]
    assert "id: T-0001" in task_blocks[0]["text"]
