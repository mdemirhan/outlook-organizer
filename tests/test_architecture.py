from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "src" / "outlook_organizer"


def imports_under(directory: Path) -> set[str]:
    imported: set[str] = set()
    for path in directory.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module:
                imported.add(node.module)
            elif isinstance(node, ast.Import):
                imported.update(alias.name for alias in node.names)
    return imported


def test_brief_does_not_depend_on_triage_audit_or_database() -> None:
    imports = imports_under(ROOT / "brief")
    forbidden = (
        "outlook_organizer.triage",
        "outlook_organizer.audit",
        "outlook_organizer.database",
    )
    assert not any(name.startswith(forbidden) for name in imports)


def test_triage_does_not_depend_on_brief() -> None:
    assert not any(
        name.startswith("outlook_organizer.brief") for name in imports_under(ROOT / "triage")
    )


def test_mcp_does_not_import_triage_audit_or_mail_writer() -> None:
    source = (ROOT / "mcp_server.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    imports = {
        node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom) and node.module
    }
    assert not any(name.startswith("outlook_organizer.triage") for name in imports)
    assert not any(name.startswith("outlook_organizer.audit") for name in imports)
    assert "MailWriter" not in source
    read_bootstrap = (ROOT / "read_bootstrap.py").read_text(encoding="utf-8")
    assert "outlook_organizer.triage" not in read_bootstrap
    assert "outlook_organizer.audit" not in read_bootstrap
    assert "OutlookAdapter" not in read_bootstrap


def test_removed_monoliths_do_not_exist() -> None:
    for name in ("service.py", "state.py", "threads.py", "models.py", "config.py"):
        assert not (ROOT / name).exists()
