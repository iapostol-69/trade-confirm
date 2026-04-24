#!/usr/bin/env python3
"""
deploy.py — Tag a release and package the skill for deployment.

Usage:
    python deploy.py <version>          # e.g.  python deploy.py 1.1

What it does:
  1. Updates  version: <x.y>  in SKILL.md frontmatter
  2. Commits the change with message  "Release v<x.y>"
  3. Creates an annotated git tag  v<x.y>
  4. Pushes commits + tag to origin
  5. Builds  trade-confirm-skill-v<x.y>.zip  (SKILL.md + scripts/)
"""

import re
import subprocess
import sys
import zipfile
from pathlib import Path


# ── Helpers ────────────────────────────────────────────────────────────────────

def run(cmd: list[str], **kwargs) -> subprocess.CompletedProcess:
    result = subprocess.run(cmd, text=True, capture_output=True, **kwargs)
    if result.returncode != 0:
        print(f"ERROR running {' '.join(cmd)}")
        print(result.stderr or result.stdout)
        sys.exit(1)
    return result


def check_clean_working_tree():
    result = subprocess.run(
        ['git', 'status', '--porcelain'],
        text=True, capture_output=True, cwd=ROOT,
    )
    dirty = [l for l in result.stdout.splitlines()
             if not l.lstrip().startswith('??')]   # ignore untracked files
    if dirty:
        print("ERROR: working tree has uncommitted changes:")
        for line in dirty:
            print(f"  {line}")
        print("Commit or stash them before releasing.")
        sys.exit(1)


# ── Main ───────────────────────────────────────────────────────────────────────

ROOT = Path(__file__).parent


def main():
    if len(sys.argv) != 2:
        print("Usage: python deploy.py <version>   e.g. python deploy.py 1.1")
        sys.exit(1)

    version = sys.argv[1].lstrip('v')          # accept "1.1" or "v1.1"
    tag     = f'v{version}'
    releases_dir = ROOT / 'releases'
    releases_dir.mkdir(exist_ok=True)
    zip_out = releases_dir / f'trade-confirm-skill-{tag}.zip'

    # ── Guard: no uncommitted changes ──────────────────────────────────────────
    check_clean_working_tree()

    # ── Guard: tag must not already exist ──────────────────────────────────────
    existing = run(['git', 'tag'], cwd=ROOT).stdout.split()
    if tag in existing:
        print(f"ERROR: tag {tag} already exists. Choose a different version.")
        sys.exit(1)

    # ── 1. Update version in SKILL.md ─────────────────────────────────────────
    skill_path = ROOT / 'SKILL.md'
    text = skill_path.read_text(encoding='utf-8')
    updated, n = re.subn(r'^version:.*$', f'version: {version}', text,
                         flags=re.MULTILINE)
    if n == 0:
        print("ERROR: could not find 'version:' field in SKILL.md frontmatter.")
        sys.exit(1)
    skill_path.write_text(updated, encoding='utf-8')
    print(f"[1/5] Updated SKILL.md  version -> {version}")

    # ── 2. Commit ──────────────────────────────────────────────────────────────
    run(['git', 'add', 'SKILL.md'], cwd=ROOT)
    run(['git', 'commit', '-m', f'Release {tag}'], cwd=ROOT)
    print(f"[2/5] Committed  'Release {tag}'")

    # ── 3. Annotated tag ───────────────────────────────────────────────────────
    run(['git', 'tag', '-a', tag, '-m', f'Release {tag}'], cwd=ROOT)
    print(f"[3/5] Tagged     {tag}")

    # ── 4. Push commits + tag ─────────────────────────────────────────────────
    run(['git', 'push', 'origin', 'HEAD'], cwd=ROOT)
    run(['git', 'push', 'origin', tag], cwd=ROOT)
    print(f"[4/5] Pushed     commits + {tag} to origin")

    # ── 5. Build zip ──────────────────────────────────────────────────────────
    if zip_out.exists():
        zip_out.unlink()
    with zipfile.ZipFile(zip_out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(skill_path, 'SKILL.md')
        for f in sorted((ROOT / 'scripts').rglob('*')):
            if f.is_file() and '__pycache__' not in f.parts:
                zf.write(f, 'scripts/' + f.name)
    print(f"[5/5] Built      {zip_out.name}")
    print()
    print(f"Done. Release {tag} is live on GitHub and packaged as {zip_out.name}")


if __name__ == '__main__':
    main()