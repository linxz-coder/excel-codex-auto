#!/usr/bin/env python3
import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


YEARLY_DAILY_SHEET_FORMAT = "{year}年日报"


def ensure_utf8_locale():
    if sys.getfilesystemencoding().lower() != "ascii":
        return
    if os.environ.get("CODEX_UTF8_REEXEC") == "1":
        return
    env = os.environ.copy()
    env["LC_ALL"] = env.get("LC_ALL") or "zh_CN.utf8"
    env["LANG"] = env.get("LANG") or "zh_CN.utf8"
    env["PYTHONIOENCODING"] = env.get("PYTHONIOENCODING") or "utf-8"
    env["CODEX_UTF8_REEXEC"] = "1"
    os.execvpe(sys.executable, [sys.executable] + sys.argv, env)


def make_absolute_path(raw_path, base_dir):
    path = Path(raw_path).expanduser()
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def run_git(repo_root, args):
    result = subprocess.run(
        ["git"] + args,
        cwd=str(repo_root),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=True,
    )
    if result.stdout:
        print(result.stdout.rstrip())
    if result.returncode != 0:
        raise SystemExit(result.returncode)
    return result


def main():
    ensure_utf8_locale()
    parser = argparse.ArgumentParser(description="Stage, commit, and push the daily formal report files to GitHub.")
    parser.add_argument("--report-date", default=datetime.now().strftime("%Y-%m-%d"), help="Report date in YYYY-MM-DD format")
    parser.add_argument("--xlsx", default="日报及月报发送记录.xlsx", help="Path to the formal workbook")
    parser.add_argument("--out-dir", default="csv_exports", help="Directory containing exported CSV files")
    parser.add_argument("--backup-dir", default="csv_backups", help="Directory containing backup CSV files")
    parser.add_argument("--sheet", help="Sheet name. Defaults to '<year>年日报' based on --report-date.")
    parser.add_argument("--remote", default="origin", help="Git remote name")
    parser.add_argument("--branch", default="main", help="Git branch name")
    parser.add_argument("--commit-message", help="Commit message override")
    args = parser.parse_args()

    report_date = datetime.strptime(args.report_date, "%Y-%m-%d")
    sheet_name = args.sheet or YEARLY_DAILY_SHEET_FORMAT.format(year=report_date.year)

    repo_root = Path(__file__).resolve().parent
    xlsx_path = make_absolute_path(args.xlsx, repo_root)
    out_dir = make_absolute_path(args.out_dir, repo_root)
    backup_dir = make_absolute_path(args.backup_dir, repo_root)
    csv_path = out_dir / (sheet_name + ".csv")
    backup_path = backup_dir / ("%s_%s.csv" % (report_date.strftime("%Y-%m-%d"), sheet_name))

    target_files = [xlsx_path, csv_path, backup_path]
    missing = [str(path) for path in target_files if not path.exists()]
    if missing:
        raise SystemExit("缺少要提交的文件：%s" % ", ".join(missing))

    rel_paths = [str(path.relative_to(repo_root)) for path in target_files]
    run_git(repo_root, ["add"] + rel_paths)

    diff_result = subprocess.run(
        ["git", "diff", "--cached", "--quiet", "--"] + rel_paths,
        cwd=str(repo_root),
    )
    if diff_result.returncode == 0:
        print("没有新的日报改动可提交。")
        return
    if diff_result.returncode != 1:
        raise SystemExit(diff_result.returncode)

    commit_message = args.commit_message or "report: update %s daily status" % report_date.strftime("%Y-%m-%d")
    run_git(repo_root, ["commit", "-m", commit_message])
    run_git(repo_root, ["push", args.remote, args.branch])

    print("synced:")
    for rel_path in rel_paths:
        print(rel_path)


if __name__ == "__main__":
    main()
