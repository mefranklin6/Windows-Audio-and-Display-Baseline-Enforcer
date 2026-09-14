"""Concurrently inspect CTS deployments without putting monitoring logic in Flutter."""

import argparse
from concurrent import futures
from datetime import datetime
import json
import logging as log
from pathlib import Path
import re
import subprocess
import sys

PROJECT_ROOT = Path(__file__).parent
LOGS_DIR = PROJECT_ROOT / "logs"
LOGS_DIR.mkdir(exist_ok=True)
LOG_FILE = LOGS_DIR / f"monitor-{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
DEPLOYMENT_RECORDS_DIR = LOGS_DIR / "deployment_records"
log.basicConfig(
    level=log.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        log.FileHandler(LOG_FILE, encoding="utf-8"),
        log.StreamHandler(sys.stdout),
    ],
)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect CTS deployment health.")
    parser.add_argument("--target", action="append", dest="targets")
    parser.add_argument("--targets-file", type=Path)
    parser.add_argument("--max-workers", type=int, default=10)
    parser.add_argument("--result-file", type=Path, required=True)
    return parser.parse_args()


def load_targets(args: argparse.Namespace) -> list[str]:
    if args.targets:
        raw_targets = args.targets
    else:
        targets_file = args.targets_file or PROJECT_ROOT / "targets.txt"
        if not targets_file.is_absolute():
            targets_file = PROJECT_ROOT / targets_file
        if not targets_file.exists():
            raise FileNotFoundError(f"Targets file not found: {targets_file}")
        raw_targets = targets_file.read_text(encoding="utf-8").splitlines()

    targets = [
        target.strip()
        for target in raw_targets
        if target.strip() and not target.lstrip().startswith("#")
    ]
    if not targets:
        raise ValueError("No target PCs were provided.")
    return list(dict.fromkeys(targets))


def failed_result(pc: str, message: str) -> dict[str, object]:
    return {
        "pc": pc,
        "online": False,
        "winrm": False,
        "error": message,
        "cts_deployed": False,
        "audio_configured": False,
        "display_configured": False,
        "logout_shortcut": False,
        "reboot_shortcut": False,
        "bginfo_deployed": False,
        "bginfo_executable": False,
        "bginfo_profile": False,
        "bginfo_background": False,
        "bginfo_startup": False,
        "bginfo_startup_method": "",
        "audio_device_cmdlets_versions": [],
        "display_config_versions": [],
        "deployment_intent": load_deployment_intent(pc),
    }


def deployment_record_path(pc: str) -> Path:
    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", pc).strip("._") or "unknown"
    return DEPLOYMENT_RECORDS_DIR / f"{safe_name.lower()}.json"


def load_deployment_intent(pc: str) -> dict[str, object] | None:
    """Load only the latest persisted deployment request for a target."""

    record_path = deployment_record_path(pc)
    try:
        payload = json.loads(record_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict):
        return None
    return payload


def inspect_target(pc: str) -> dict[str, object]:
    log.info("%s: Monitoring started", pc)
    script = PROJECT_ROOT / "utility_scripts" / "MonitorTarget.ps1"
    try:
        completed = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-ComputerName",
                pc,
            ],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as error:
        log.error("%s: Could not start PowerShell: %s", pc, error)
        return failed_result(pc, f"Could not start PowerShell: {error}")

    output_lines = [
        line.strip() for line in completed.stdout.splitlines() if line.strip()
    ]
    if completed.returncode != 0 or not output_lines:
        message = (
            completed.stderr.strip() or "PowerShell returned no monitoring result."
        )
        log.error("%s: Monitoring failed: %s", pc, message)
        return failed_result(pc, message)
    try:
        result = json.loads(output_lines[-1])
    except json.JSONDecodeError as error:
        message = f"Invalid PowerShell monitoring result: {error}"
        log.error("%s: %s", pc, message)
        return failed_result(pc, message)

    result["deployment_intent"] = load_deployment_intent(pc)

    status = "offline" if not result.get("online") else "complete"
    if result.get("online") and not result.get("winrm"):
        status = "error"
    log.info("%s: Monitoring complete (%s)", pc, status)
    return result


def main() -> int:
    args = parse_arguments()
    if args.max_workers < 1:
        log.error("Maximum workers must be at least 1.")
        return 1
    try:
        targets = load_targets(args)
    except (FileNotFoundError, ValueError) as error:
        log.error(str(error))
        return 1

    results_by_pc: dict[str, dict[str, object]] = {}
    with futures.ThreadPoolExecutor(max_workers=args.max_workers) as executor:
        tasks = {executor.submit(inspect_target, pc): pc for pc in targets}
        for task in futures.as_completed(tasks):
            pc = tasks[task]
            try:
                results_by_pc[pc] = task.result()
            except Exception as error:  # Defensive boundary for one target.
                log.exception("%s: Unexpected monitoring failure", pc)
                results_by_pc[pc] = failed_result(pc, str(error))
            print(f"MONITOR_PROGRESS {pc}", flush=True)

    result_file = args.result_file
    if not result_file.is_absolute():
        result_file = PROJECT_ROOT / result_file
    result_file.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "log_file": str(LOG_FILE.resolve()),
        "pcs": [results_by_pc[pc] for pc in targets],
    }
    result_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    log.info("Monitoring complete for %d target(s)", len(targets))
    return 0


if __name__ == "__main__":
    sys.exit(main())
