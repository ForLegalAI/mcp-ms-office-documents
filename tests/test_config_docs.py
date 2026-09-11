"""Every environment variable config.py reads is documented.

``config.py`` is the only module that reads ``os.environ``. This test
extracts the variable names it reads and checks that each one appears in
``.env.example`` (as an assignment, possibly commented out) and in
``docs/configuration.md`` (in backticks). Documentation of environment
variables is what goes stale first; this keeps the three in step.
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

_ENV_READ_RE = re.compile(r"""os\.environ\.get\(\s*["']([A-Z][A-Z0-9_]*)["']""")


def _read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def _variables_read_by_config() -> set[str]:
    names = set(_ENV_READ_RE.findall(_read("config.py")))
    assert names, "no os.environ.get() calls found in config.py; the regex is stale"
    return names


def test_config_is_the_only_env_reader():
    offenders = []
    for path in ROOT.rglob("*.py"):
        rel = path.relative_to(ROOT).as_posix()
        if rel == "config.py" or rel.startswith("tests/") or "/.venv/" in f"/{rel}":
            continue
        text = path.read_text(encoding="utf-8")
        if "os.environ" in text or "os.getenv" in text:
            offenders.append(rel)
    assert not offenders, f"read the environment through get_config(), not directly: {offenders}"


def test_every_variable_is_in_env_example():
    example = _read(".env.example")
    present = set(re.findall(r"^\s*#?\s*([A-Z][A-Z0-9_]*)=", example, flags=re.M))
    missing = sorted(_variables_read_by_config() - present)
    assert not missing, f"add to .env.example: {missing}"


def test_every_variable_is_in_configuration_docs():
    doc = _read("docs/configuration.md")
    present = set(re.findall(r"`([A-Z][A-Z0-9_]*)`", doc))
    missing = sorted(_variables_read_by_config() - present)
    assert not missing, f"document in docs/configuration.md: {missing}"
