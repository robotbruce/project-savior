#!/usr/bin/env python3
"""Collect objective facts about a repository for a handover / teaching deck.

Stdlib only. Read-only: never executes project code, never prints secret values
(only env var *names* are collected; .env / key files are flagged, not read).

Usage:
    python3 scan_repo.py <repo_path> [--out <dir>] [--max-depth 3]

Writes <out>/facts.json and <out>/facts.md (default out: ./<repo-name>-facts).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

try:
    import tomllib  # py3.11+
except ImportError:  # pragma: no cover
    tomllib = None

IGNORE_DIRS = {
    ".git", "node_modules", ".venv", "venv", "env", "__pycache__", ".mypy_cache",
    ".pytest_cache", ".ruff_cache", "dist", "build", ".next", ".nuxt", "target",
    ".idea", ".vscode", ".gradle", "coverage", ".tox", ".eggs", "site-packages",
    ".terraform", ".cache", "vendor", ".DS_Store", ".ipynb_checkpoints",
}

LANG_BY_EXT = {
    ".py": "Python", ".ipynb": "Jupyter Notebook", ".js": "JavaScript", ".jsx": "JavaScript",
    ".ts": "TypeScript", ".tsx": "TypeScript", ".vue": "Vue", ".go": "Go", ".java": "Java",
    ".kt": "Kotlin", ".scala": "Scala", ".rs": "Rust", ".rb": "Ruby", ".php": "PHP",
    ".cs": "C#", ".cpp": "C++", ".cc": "C++", ".c": "C", ".h": "C/C++ Header",
    ".swift": "Swift", ".sh": "Shell", ".bash": "Shell", ".zsh": "Shell", ".sql": "SQL",
    ".html": "HTML", ".css": "CSS", ".scss": "SCSS", ".yaml": "YAML", ".yml": "YAML",
    ".json": "JSON", ".toml": "TOML", ".md": "Markdown", ".tf": "Terraform",
    ".proto": "Protobuf", ".r": "R", ".dart": "Dart", ".lua": "Lua",
}
CODE_LANGS = {
    "Python", "Jupyter Notebook", "JavaScript", "TypeScript", "Vue", "Go", "Java", "Kotlin",
    "Scala", "Rust", "Ruby", "PHP", "C#", "C++", "C", "Swift", "Shell", "SQL", "Terraform",
    "R", "Dart", "Lua", "HTML", "CSS", "SCSS",
}

KEY_FILE_PATTERNS = {
    "docs": [r"^README", r"^CLAUDE\.md$", r"^AGENTS\.md$", r"^CONTRIBUTING", r"^CHANGELOG",
             r"^docs?/", r"^ARCHITECTURE", r"\.md$"],
    "python_deps": [r"(^|/)pyproject\.toml$", r"(^|/)requirements[^/]*\.txt$", r"(^|/)setup\.py$",
                    r"(^|/)setup\.cfg$", r"(^|/)Pipfile$", r"(^|/)environment\.ya?ml$"],
    "lockfiles": [r"(^|/)(poetry|uv|Pipfile|package-lock|yarn|pnpm-lock|Cargo|Gemfile|composer)\.lock$",
                  r"(^|/)package-lock\.json$", r"(^|/)pnpm-lock\.yaml$", r"(^|/)go\.sum$"],
    "js_deps": [r"(^|/)package\.json$"],
    "other_deps": [r"(^|/)go\.mod$", r"(^|/)pom\.xml$", r"(^|/)build\.gradle(\.kts)?$",
                   r"(^|/)Cargo\.toml$", r"(^|/)Gemfile$", r"(^|/)composer\.json$"],
    "container": [r"(^|/)Dockerfile[^/]*$", r"(^|/)docker-compose[^/]*\.ya?ml$",
                  r"(^|/)compose\.ya?ml$", r"(^|/)\.dockerignore$"],
    "ci_cd": [r"^\.github/workflows/", r"(^|/)\.gitlab-ci\.yml$", r"(^|/)Jenkinsfile",
              r"(^|/)azure-pipelines\.yml$", r"(^|/)bitbucket-pipelines\.yml$",
              r"(^|/)\.circleci/", r"(^|/)cloudbuild\.ya?ml$", r"(^|/)buildspec\.ya?ml$"],
    "k8s_deploy": [r"(^|/)Chart\.yaml$", r"(^|/)values[^/]*\.ya?ml$", r"(^|/)kustomization\.ya?ml$",
                   r"(^|/)(k8s|kubernetes|helm|charts|manifests|deploy|deployment)/",
                   r"(^|/)bentofile\.ya?ml$", r"(^|/)serverless\.ya?ml$", r"(^|/)Procfile$",
                   r"\.tf$", r"(^|/)databricks\.ya?ml$"],
    "build_tasks": [r"(^|/)Makefile$", r"(^|/)justfile$", r"(^|/)Taskfile\.ya?ml$",
                    r"(^|/)tox\.ini$", r"(^|/)noxfile\.py$"],
    "config": [r"(^|/)\.env\.(example|sample|template)$", r"(^|/)config[^/]*\.(ya?ml|json|toml|ini|py)$",
               r"(^|/)settings[^/]*\.(py|ya?ml|json)$", r"(^|/)conf/", r"(^|/)configs?/"],
    "entrypoints": [r"(^|/)main\.(py|go|ts|js)$", r"(^|/)app\.(py|ts|js)$", r"(^|/)server\.(py|ts|js)$",
                    r"(^|/)index\.(ts|js)$", r"(^|/)manage\.py$", r"(^|/)__main__\.py$",
                    r"(^|/)wsgi\.py$", r"(^|/)asgi\.py$", r"(^|/)service\.py$", r"(^|/)cli\.py$",
                    r"^cmd/", r"(^|/)run[^/]*\.(py|sh)$", r"(^|/)handler\.py$", r"(^|/)lambda_function\.py$"],
    "tests": [r"(^|/)tests?/", r"(^|/)test_[^/]*\.py$", r"(^|/)[^/]*_test\.(py|go)$",
              r"\.(test|spec)\.(ts|tsx|js|jsx)$", r"(^|/)__tests__/", r"(^|/)pytest\.ini$",
              r"(^|/)conftest\.py$", r"(^|/)jest\.config", r"(^|/)vitest\.config"],
}

SENSITIVE_PATTERNS = [r"(^|/)\.env$", r"(^|/)\.env\.(?!example|sample|template)[^/]+$",
                      r"\.pem$", r"\.key$", r"(^|/)id_rsa", r"\.p12$", r"\.pfx$",
                      r"(^|/)credentials(\.json)?$", r"service[-_]account[^/]*\.json$",
                      r"(^|/)secrets?\.(ya?ml|json)$"]

ENV_PATTERNS = [
    re.compile(r"os\.environ(?:\.get)?\(\s*['\"]([A-Z][A-Z0-9_]{2,})['\"]"),
    re.compile(r"os\.environ\[\s*['\"]([A-Z][A-Z0-9_]{2,})['\"]\s*\]"),
    re.compile(r"os\.getenv\(\s*['\"]([A-Z][A-Z0-9_]{2,})['\"]"),
    re.compile(r"process\.env\.([A-Z][A-Z0-9_]{2,})"),
    re.compile(r"process\.env\[\s*['\"]([A-Z][A-Z0-9_]{2,})['\"]"),
    re.compile(r"import\.meta\.env\.([A-Z][A-Z0-9_]{2,})"),
    re.compile(r"os\.Getenv\(\s*\"([A-Z][A-Z0-9_]{2,})\""),
    re.compile(r"System\.getenv\(\s*\"([A-Z][A-Z0-9_]{2,})\""),
    re.compile(r"\$\{([A-Z][A-Z0-9_]{2,})(?::?-[^}]*)?\}"),
]
ENV_SCAN_EXT = {".py", ".js", ".ts", ".tsx", ".jsx", ".go", ".java", ".kt", ".sh", ".yaml",
                ".yml", ".toml", ".ini", ".cfg", ".env.example", ".vue", ".rb", ".php"}
TODO_RE = re.compile(r"\b(TODO|FIXME|HACK|XXX|BUG)\b[:\s]?(.*)")
MAX_FILE_BYTES = 512_000


def run_git(repo: Path, *args: str) -> str:
    try:
        out = subprocess.run(["git", "-C", str(repo), *args], capture_output=True,
                             text=True, timeout=60)
        return out.stdout.strip() if out.returncode == 0 else ""
    except Exception:
        return ""


def walk(repo: Path):
    for root, dirs, files in os.walk(repo):
        dirs[:] = sorted(d for d in dirs if d not in IGNORE_DIRS and not d.endswith(".egg-info"))
        for f in sorted(files):
            if f in IGNORE_DIRS:
                continue
            p = Path(root) / f
            yield p, p.relative_to(repo).as_posix()


def build_tree(repo: Path, max_depth: int, max_entries: int = 40) -> list[str]:
    lines: list[str] = [repo.name + "/"]

    def rec(d: Path, prefix: str, depth: int):
        try:
            entries = sorted(
                [e for e in d.iterdir() if e.name not in IGNORE_DIRS and not e.name.endswith(".egg-info")],
                key=lambda e: (not e.is_dir(), e.name.lower()))
        except PermissionError:
            return
        shown = entries[:max_entries]
        for i, e in enumerate(shown):
            last = i == len(shown) - 1 and len(entries) <= max_entries
            conn = "└── " if last else "├── "
            if e.is_dir():
                n = sum(1 for _ in e.rglob("*") if _.is_file() and not any(p in IGNORE_DIRS for p in _.parts))
                lines.append(f"{prefix}{conn}{e.name}/  ({n} files)")
                if depth < max_depth:
                    rec(e, prefix + ("    " if last else "│   "), depth + 1)
            else:
                lines.append(f"{prefix}{conn}{e.name}")
        if len(entries) > max_entries:
            lines.append(f"{prefix}└── … (+{len(entries) - max_entries} more)")

    rec(repo, "", 1)
    return lines


def read_text(p: Path, limit: int = MAX_FILE_BYTES) -> str:
    try:
        if p.stat().st_size > limit:
            return ""
        return p.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""


def parse_deps(repo: Path, rel: str) -> dict:
    p = repo / rel
    name = p.name
    txt = read_text(p)
    info: dict = {"file": rel}
    if name == "package.json":
        try:
            data = json.loads(txt)
            info["name"] = data.get("name")
            info["scripts"] = data.get("scripts", {})
            info["dependencies"] = sorted((data.get("dependencies") or {}).keys())
            info["devDependencies"] = sorted((data.get("devDependencies") or {}).keys())
        except Exception:
            info["error"] = "unparseable"
    elif name == "pyproject.toml" and tomllib:
        try:
            data = tomllib.loads(txt)
            proj = data.get("project", {})
            poetry = data.get("tool", {}).get("poetry", {})
            info["name"] = proj.get("name") or poetry.get("name")
            info["python"] = proj.get("requires-python") or (poetry.get("dependencies") or {}).get("python")
            deps = list(proj.get("dependencies", [])) or [k for k in (poetry.get("dependencies") or {}) if k != "python"]
            info["dependencies"] = deps
            opt = proj.get("optional-dependencies") or {}
            if opt:
                info["optional"] = {k: v for k, v in opt.items()}
            scripts = proj.get("scripts") or poetry.get("scripts")
            if scripts:
                info["scripts"] = scripts
            info["tools"] = sorted(data.get("tool", {}).keys())
        except Exception:
            info["error"] = "unparseable"
    elif name.startswith("requirements") and name.endswith(".txt"):
        info["dependencies"] = [l.strip() for l in txt.splitlines()
                                if l.strip() and not l.strip().startswith(("#", "-"))]
    elif name == "go.mod":
        m = re.search(r"^module\s+(\S+)", txt, re.M)
        info["name"] = m.group(1) if m else None
        info["dependencies"] = re.findall(r"^\s*([\w./-]+\.\w+/[\w./-]+)\s+v", txt, re.M)
    elif name == "Makefile":
        info["targets"] = sorted(set(re.findall(r"^([A-Za-z0-9_.-]+):(?!=)", txt, re.M)) - {".PHONY"})
    return info


def git_facts(repo: Path) -> dict:
    if not run_git(repo, "rev-parse", "--is-inside-work-tree"):
        return {"is_git": False}
    f: dict = {"is_git": True}
    f["remote"] = run_git(repo, "remote", "get-url", "origin")
    f["branch"] = run_git(repo, "rev-parse", "--abbrev-ref", "HEAD")
    f["branches"] = [b.strip() for b in run_git(repo, "branch", "-a", "--format=%(refname:short)").splitlines()][:30]
    f["commit_count"] = run_git(repo, "rev-list", "--count", "HEAD")
    f["first_commit"] = run_git(repo, "log", "--reverse", "--format=%ad", "--date=short", "--max-parents=0").splitlines()[:1]
    f["last_commit"] = run_git(repo, "log", "-1", "--format=%ad %h %s", "--date=short")
    f["recent_commits"] = run_git(repo, "log", "-20", "--format=%ad|%h|%an|%s", "--date=short").splitlines()
    f["contributors_all"] = run_git(repo, "shortlog", "-sn", "--no-merges", "HEAD").splitlines()[:12]
    f["contributors_12m"] = run_git(repo, "shortlog", "-sn", "--no-merges", "--since=12.months", "HEAD").splitlines()[:12]
    f["tags"] = run_git(repo, "tag", "--sort=-creatordate").splitlines()[:10]
    # Hotspots: most frequently changed files in the last 12 months
    churn = run_git(repo, "log", "--since=12.months", "--name-only", "--format=", "--no-merges")
    c = Counter(l for l in churn.splitlines() if l.strip())
    f["hotspots_12m"] = [f"{n:>4}  {p}" for p, n in c.most_common(15)]
    f["dirty"] = bool(run_git(repo, "status", "--porcelain"))
    return f


def scan(repo: Path, max_depth: int) -> dict:
    repo = repo.resolve()
    facts: dict = {"name": repo.name, "path": str(repo)}
    facts["git"] = git_facts(repo)
    facts["tree"] = build_tree(repo, max_depth)

    lang_files: Counter = Counter()
    lang_lines: Counter = Counter()
    key_files: dict[str, list[str]] = {k: [] for k in KEY_FILE_PATTERNS}
    sensitive: list[str] = []
    env_vars: dict[str, set] = {}
    todos: list[str] = []
    todo_count = 0
    k8s_kinds: Counter = Counter()
    total_files = 0
    compiled = {k: [re.compile(p, re.I) for p in v] for k, v in KEY_FILE_PATTERNS.items()}
    sens_re = [re.compile(p, re.I) for p in SENSITIVE_PATTERNS]

    for p, rel in walk(repo):
        total_files += 1
        ext = p.suffix.lower()
        lang = LANG_BY_EXT.get(ext)
        if any(r.search(rel) for r in sens_re):
            sensitive.append(rel)
            continue  # never read sensitive files
        for cat, pats in compiled.items():
            if any(r.search(rel) for r in pats):
                if cat == "docs" and rel.count("/") > 2 and not rel.lower().startswith("doc"):
                    continue
                key_files[cat].append(rel)
        if lang:
            lang_files[lang] += 1
        txt = ""
        if lang in CODE_LANGS or ext in ENV_SCAN_EXT or rel.endswith(".env.example"):
            txt = read_text(p)
        if not txt:
            continue
        if lang:
            lang_lines[lang] += txt.count("\n") + 1
        if ext in {".yaml", ".yml"}:
            for kind in re.findall(r"^kind:\s*(\w+)", txt, re.M):
                k8s_kinds[kind] += 1
        for rx in ENV_PATTERNS:
            for m in rx.findall(txt):
                env_vars.setdefault(m, set()).add(rel)
        if rel.endswith((".env.example", ".env.sample", ".env.template")):
            for m in re.findall(r"^\s*([A-Z][A-Z0-9_]{2,})\s*=", txt, re.M):
                env_vars.setdefault(m, set()).add(rel)
        if lang in CODE_LANGS:
            for i, line in enumerate(txt.splitlines(), 1):
                m = TODO_RE.search(line)
                if m and ("#" in line or "//" in line or "--" in line or "/*" in line):
                    todo_count += 1
                    if len(todos) < 40:
                        todos.append(f"{rel}:{i}  {m.group(1)} {m.group(2).strip()[:100]}")

    facts["total_files"] = total_files
    facts["languages"] = [
        {"language": l, "files": lang_files[l], "lines": lang_lines.get(l, 0)}
        for l, _ in lang_files.most_common()
    ][:15]
    facts["key_files"] = {k: v[:40] for k, v in key_files.items() if v}
    facts["dependencies"] = [parse_deps(repo, rel) for rel in
                             (key_files["python_deps"] + key_files["js_deps"] + key_files["other_deps"]
                              + key_files["build_tasks"])[:15]]
    facts["k8s_kinds"] = dict(k8s_kinds.most_common())
    facts["env_vars"] = {k: sorted(v)[:5] for k, v in sorted(env_vars.items())}
    facts["sensitive_files_present"] = sensitive[:30]
    facts["todo_count"] = todo_count
    facts["todos_sample"] = todos
    readme = next((r for r in key_files["docs"] if re.match(r"^README", r, re.I)), None)
    facts["readme"] = readme
    facts["readme_head"] = read_text(repo / readme)[:3000] if readme else ""
    return facts


def to_markdown(f: dict) -> str:
    g = f["git"]
    out = [f"# Repo facts: {f['name']}", "", f"- Path: `{f['path']}`",
           f"- Files (excluding vendored/build dirs): {f['total_files']}"]
    if g.get("is_git"):
        out += [f"- Remote: {g.get('remote') or '(none)'}", f"- Current branch: {g.get('branch')}",
                f"- Commits: {g.get('commit_count')}, first: {(g.get('first_commit') or ['?'])[0]}, last: {g.get('last_commit')}",
                f"- Uncommitted changes: {'yes' if g.get('dirty') else 'no'}"]
    else:
        out.append("- Not a git repository")
    out += ["", "## Languages", "| Language | Files | Lines |", "|---|---:|---:|"]
    out += [f"| {l['language']} | {l['files']} | {l['lines'] or '-'} |" for l in f["languages"]]
    out += ["", "## Directory tree", "```", *f["tree"], "```", "", "## Key files"]
    for cat, items in f["key_files"].items():
        out.append(f"- **{cat}**: " + ", ".join(f"`{i}`" for i in items[:20]) + (" …" if len(items) > 20 else ""))
    if f["dependencies"]:
        out += ["", "## Dependencies / scripts"]
        for d in f["dependencies"]:
            out.append(f"### `{d['file']}`")
            for k in ("name", "python", "scripts", "targets", "dependencies", "devDependencies", "optional", "tools"):
                if d.get(k):
                    v = d[k]
                    if isinstance(v, list):
                        v = ", ".join(map(str, v[:40])) + (" …" if len(v) > 40 else "")
                    elif isinstance(v, dict):
                        v = "; ".join(f"{a}: {b}" for a, b in list(v.items())[:25])
                    out.append(f"- {k}: {v}")
    if f["k8s_kinds"]:
        out += ["", "## Kubernetes / manifest kinds", ", ".join(f"{k}×{v}" for k, v in f["k8s_kinds"].items())]
    if f["env_vars"]:
        out += ["", "## Environment variables referenced (names only)"]
        out += [f"- `{k}` — {', '.join(v)}" for k, v in f["env_vars"].items()]
    if f["sensitive_files_present"]:
        out += ["", "## ⚠ Sensitive-looking files present (NOT read — do not put contents in the deck)"]
        out += [f"- `{s}`" for s in f["sensitive_files_present"]]
    if g.get("is_git"):
        out += ["", "## Contributors (last 12 months)", "```", *(g.get("contributors_12m") or ["(none)"]), "```",
                "## Contributors (all time)", "```", *g.get("contributors_all", []), "```",
                "## Hotspots (most-changed files, last 12 months)", "```", *(g.get("hotspots_12m") or ["(none)"]), "```",
                "## Recent commits", "```", *g.get("recent_commits", []), "```"]
        if g.get("tags"):
            out += ["## Tags (newest first)", ", ".join(g["tags"])]
    out += ["", f"## TODO/FIXME markers: {f['todo_count']}"]
    out += [f"- {t}" for t in f["todos_sample"]]
    if f["readme_head"]:
        out += ["", f"## README head (`{f['readme']}`)", "", f["readme_head"]]
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("repo")
    ap.add_argument("--out")
    ap.add_argument("--max-depth", type=int, default=3)
    a = ap.parse_args()
    repo = Path(a.repo).expanduser()
    if not repo.is_dir():
        sys.exit(f"not a directory: {repo}")
    facts = scan(repo, a.max_depth)
    out = Path(a.out or f"./{repo.resolve().name}-facts").expanduser()
    out.mkdir(parents=True, exist_ok=True)
    (out / "facts.json").write_text(json.dumps(facts, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "facts.md").write_text(to_markdown(facts), encoding="utf-8")
    print(f"wrote {out/'facts.json'} and {out/'facts.md'}")


if __name__ == "__main__":
    main()
