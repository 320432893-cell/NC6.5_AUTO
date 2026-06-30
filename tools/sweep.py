# 职责：每日大体检的机械引擎——跑机器闸/算扇入/数漂移/分类改动面,把事实写进 logs/sweep-report.md
# 不做什么：不删码、不改白名单、不动 git——只读 + 写报告;判断/动手交 rules/review/daily.md(AI/人)
# 允许依赖层：标准库 + 子进程调 ruff/vulture/radon/lint-imports(经 SWEEP_TOOLBIN 解析)
# 谁不应该 import：本文件是 CLI 入口,不被业务 import

import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOGS = ROOT / "logs"
DRIFT_LOG = LOGS / "sweep-drift.log"
WHITELIST = ROOT / ".vulture_whitelist.py"
LINE_RATCHET = 600


def tool(name, module=None):
    """解析工具命令:SWEEP_TOOLBIN → .venv/bin → PATH → python -m。缺则 None。"""
    tb = os.environ.get("SWEEP_TOOLBIN")
    for cand in ([Path(tb) / name] if tb else []) + [ROOT / ".venv" / "bin" / name]:
        if cand.exists():
            return [str(cand)]
    w = shutil.which(name)
    if w:
        return [w]
    if module:
        return [sys.executable, "-m", module]
    return None


def run(cmd, **kw):
    return subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, **kw)


def changed_py():
    """工作区相对 HEAD 改动的 .py(含新增未跟踪),排除 archive。"""
    out = run(["git", "status", "--porcelain"]).stdout.splitlines()
    files = []
    for line in out:
        p = line[3:].strip()
        if p.endswith(".py") and "archive/" not in p and (ROOT / p).exists():
            files.append(p)
    return files


def gate_ruff():
    cmd = tool("ruff")
    if not cmd:
        return "skip(无 ruff)", []
    r = run(cmd + ["check", "core", "tools"])
    bad = [ln for ln in r.stdout.splitlines() if "-->" in ln and "archive/" not in ln]
    return ("🟢 绿" if not bad else f"🔴 {len(bad)} 处正式码"), bad


def gate_importlinter():
    cmd = tool("lint-imports")
    if not cmd:
        return "skip(无 lint-imports)"
    r = run(cmd)
    return (
        "🟢 kept"
        if "broken" in r.stdout and " 0 broken" in r.stdout
        else ("🔴 broken" if "broken" in r.stdout else "skip(无契约)")
    )


def gate_vulture():
    cmd = tool("vulture")
    if not cmd:
        return "skip(无 vulture)", []
    args = ["core", "tools", "tests"]
    if WHITELIST.exists():
        args.append(str(WHITELIST.name))
    r = run(cmd + args + ["--exclude", "*/archive/*", "--min-confidence", "60"])
    found = [ln for ln in r.stdout.splitlines() if "unused" in ln]
    return ("🟢 绿" if not found else f"🔴 {len(found)} 名单外"), found


def drift():
    wl = 0
    if WHITELIST.exists():
        wl = sum(
            1
            for ln in WHITELIST.read_text(encoding="utf-8").splitlines()
            if "# unused" in ln
        )
    skips = run(
        [
            "git",
            "grep",
            "-rEc",
            r"@pytest.mark.skip|pytest.mark.xfail|skipif",
            "--",
            "tests/",
        ]
    )
    skip_n = sum(
        int(ln.split(":")[-1]) for ln in skips.stdout.splitlines() if ":" in ln
    )
    over = []
    for d in ("core", "tools"):
        for f in sorted((ROOT / d).glob("*.py")):
            n = sum(1 for _ in f.open(encoding="utf-8", errors="ignore"))
            if n > LINE_RATCHET:
                over.append((n, f.relative_to(ROOT)))
    over.sort(reverse=True)
    return wl, skip_n, over


def fan_in(modnames):
    rows = []
    for m in modnames:
        r = run(
            [
                "git",
                "grep",
                "-rl",
                "-E",
                rf"(from (core|tools)\.{m} import|import (core|tools)\.{m}\b)",
                "--",
                "core",
                "tools",
            ]
        )
        files = [
            x
            for x in r.stdout.splitlines()
            if "archive/" not in x and f"/{m}.py" not in x
        ]
        rows.append((len(files), m))
    rows.sort(reverse=True)
    return rows


def git_age(symbol, path):
    r = run(
        [
            "git",
            "log",
            "--diff-filter=A",
            f"-S def {symbol}",
            "--format=%as",
            "--",
            path,
        ]
    )
    dates = r.stdout.split()
    return dates[-1] if dates else "?"


def main():
    LOGS.mkdir(exist_ok=True)
    now = datetime.now()
    stamp = now.strftime("%Y-%m-%d-%H%M")
    ts = now.strftime("%Y-%m-%d %H:%M")
    report = LOGS / f"sweep-report-{stamp}.md"
    changed = changed_py()
    changed_mods = sorted({Path(f).stem for f in changed})
    ruff_s, ruff_bad = gate_ruff()
    il_s = gate_importlinter()
    vul_s, vul_found = gate_vulture()
    wl, skip_n, over = drift()
    fi = fan_in(changed_mods)

    L = []
    L.append(f"# 每日体检报告 · {ts}")
    L.append(
        "\n> 机械引擎产出(只读);判断/动手见 rules/review/daily.md。**不代表架构无恙,只证 L4/L5。**"
    )
    L.append("\n## 机器闸")
    L.append("| 闸 | 状态 |\n|---|---|")
    L.append(f"| ruff(正式码) | {ruff_s} |")
    L.append(f"| import-linter | {il_s} |")
    L.append(f"| vulture(白名单+排除archive) | {vul_s} |")
    if ruff_bad:
        L.append("\n正式码 ruff:\n```")
        L += ruff_bad
        L.append("```")
    L.append("\n## 漂移仪表")
    L.append(f"- 白名单条目: {wl}")
    L.append(f"- skip/xfail: {skip_n}  {'🔴' if skip_n else '🟢'}")
    L.append(
        f"- 超 {LINE_RATCHET} 行正式文件: {len(over)} 个  {'🔴' if over else '🟢'}"
    )
    for n, f in over[:15]:
        L.append(f"    - {n}  {f}")
    L.append("\n## 承重扇入(本次改动模块,fan-in≥1,top12)")
    shown = [(n, m) for n, m in fi if n >= 1][:12]
    if shown:
        L.append("| 模块 | fan-in |\n|---|---|")
        for n, m in shown:
            mark = "  ← 承重,动它带测试" if n >= 10 else ""
            L.append(f"| {m} | {n}{mark} |")
    else:
        L.append("(无被依赖的改动模块)")
    L.append("\n## 死码/孤儿(名单外)+ git 年龄分诊")
    if vul_found:
        for line in vul_found[:30]:
            path = line.split(":", 1)[0]
            sym = line.split("'")[1] if "'" in line else "?"
            age = git_age(sym, path) if sym != "?" else "?"
            L.append(f"- {line}  [出生 {age}]")
        L.append("\n→ 删前走 daily.md:grep 自验;老=rot 删 / 新=WIP 宽限,不自动删。")
    else:
        L.append("🟢 无(名单外零死码)")
    L.append("\n## 待办 / 待批(judgment·AI/人按 daily.md 填)")
    L.append("- [ ] 死码处置(grep验后删 / 宽限 / 入白名单待你批)")
    L.append("- [ ] 漂移红灯 → 是否排 stage")
    L.append("- [ ] 白名单/skip pending → 待你批")

    report.write_text("\n".join(L) + "\n", encoding="utf-8")
    with DRIFT_LOG.open("a", encoding="utf-8") as fh:
        fh.write(
            f"{stamp}\twhitelist={wl}\tskip={skip_n}\tover{LINE_RATCHET}={len(over)}\n"
        )
    print(f"报告已写 {report.relative_to(ROOT)}")
    print(f"闸: ruff {ruff_s} | import-linter {il_s} | vulture {vul_s}")
    print(
        f"漂移: 白名单{wl} skip{skip_n} 超{LINE_RATCHET}行{len(over)}个(趋势见 {DRIFT_LOG.relative_to(ROOT)})"
    )


if __name__ == "__main__":
    main()
