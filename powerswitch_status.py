from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import subprocess


def parse_args() -> argparse.Namespace:
    today = datetime.now().strftime("%Y%m%d")
    parser = argparse.ArgumentParser(description="Show live progress for a PowerSwitch scrape run.")
    parser.add_argument(
        "--log-file",
        default=f"Output/{today}_powerswitch_run.log",
        help="Path to the scrape log file generated via tee.",
    )
    parser.add_argument(
        "--total-scenarios",
        type=int,
        default=392,
        help="Expected total scenario count for this run.",
    )
    parser.add_argument(
        "--tail-lines",
        type=int,
        default=10,
        help="How many recent non-empty lines to show.",
    )
    parser.add_argument(
        "--process-match",
        default="powerswitch_monthly_scrape.py",
        help="Substring to identify the active scraper process command line.",
    )
    return parser.parse_args()


def detect_windows_process_start(process_match: str) -> datetime | None:
    ps_script = (
        "Get-CimInstance Win32_Process -Filter \"Name = 'python.exe'\" | "
        f"Where-Object {{ $_.CommandLine -like '*{process_match}*' }} | "
        "Select-Object -ExpandProperty CreationDate | "
        "ForEach-Object { $_.ToString('yyyyMMddHHmmss') }"
    )
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception:
        return None

    if proc.returncode != 0:
        return None

    starts: list[datetime] = []
    for raw in proc.stdout.splitlines():
        value = raw.strip()
        if len(value) < 14:
            continue
        stamp = value[:14]
        if not stamp.isdigit():
            continue
        try:
            starts.append(datetime.strptime(stamp, "%Y%m%d%H%M%S"))
        except ValueError:
            continue

    if not starts:
        return None
    return min(starts)


def resolve_started_at(log_path: Path, process_match: str) -> tuple[datetime, str]:
    process_start = detect_windows_process_start(process_match)
    if process_start is not None:
        return process_start, "process"
    return datetime.fromtimestamp(log_path.stat().st_ctime), "log-ctime"


def main() -> None:
    args = parse_args()
    log_path = Path(args.log_file)

    if not log_path.exists():
        raise FileNotFoundError(f"Log file not found: {log_path}")
    if args.total_scenarios < 1:
        raise ValueError("--total-scenarios must be at least 1")

    lines = log_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    non_empty_lines = [line for line in lines if line.strip()]

    started = sum(1 for line in lines if line.startswith("[W"))
    completed = sum(1 for line in lines if line.startswith("    ->"))
    attempt_failures = sum(1 for line in lines if line.startswith("    !!"))

    now = datetime.now()
    started_at, started_source = resolve_started_at(log_path, args.process_match)
    elapsed_min = (now - started_at).total_seconds() / 60

    pct = (completed / args.total_scenarios) * 100
    rows_per_min = completed / elapsed_min if elapsed_min > 0 else 0.0

    eta_text = "unknown"
    if completed >= 3 and elapsed_min > 0:
        est_total_min = elapsed_min * args.total_scenarios / completed
        remain_min = max(est_total_min - elapsed_min, 0.0)
        eta_finish = now + (now - started_at) * ((args.total_scenarios - completed) / completed)
        eta_text = f"~{remain_min:.1f} min remaining (finish around {eta_finish.strftime('%H:%M')})"

    print(f"Log file: {log_path}")
    print(f"Run started: {started_at.strftime('%Y-%m-%d %H:%M:%S')} ({started_source})")
    print(f"Elapsed: {elapsed_min:.1f} min")
    print(f"Scenarios started: {started}/{args.total_scenarios}")
    print(f"Scenarios completed: {completed}/{args.total_scenarios} ({pct:.1f}%)")
    print(f"Retry/failure attempts logged: {attempt_failures}")
    print(f"Throughput: {rows_per_min:.2f} scenarios/min")
    print(f"ETA: {eta_text}")

    tail_count = max(args.tail_lines, 0)
    if tail_count > 0:
        print("Recent log lines:")
        for line in non_empty_lines[-tail_count:]:
            print(line)


if __name__ == "__main__":
    main()