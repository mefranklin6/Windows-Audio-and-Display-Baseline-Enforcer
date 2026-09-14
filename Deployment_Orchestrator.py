"""Concurrently deploy the selected configuration scripts."""

import argparse
from concurrent import futures
from dataclasses import asdict, dataclass, field
from datetime import datetime
import json
import logging as log
from pathlib import Path
import re
import subprocess
import sys
import threading
import os

import config

MAX_WORKERS = 10  # Concurrent PCs to install on. Default is 10.

now = datetime.now()
project_root = Path(__file__).parent
logs_dir = project_root / "logs"
logs_dir.mkdir(exist_ok=True)
log_file = logs_dir / f"{now.strftime('%Y-%m-%d_%H-%M-%S')}.log"
log.basicConfig(
    level=log.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        log.FileHandler(log_file, encoding="utf-8"),
        log.StreamHandler(sys.stdout),
    ],
)

# Do not modify. Set feature flags in config.py now. See config.py.example
# {Installer Path (str) : Enabled (bool)}
#
# Ordering is important here:
#     - First, we need the individual installers to run
#     - Second, we need Cleanup to consilidate artifacts left by the installers
#     - Finally, add_shortcuts relies on consolidation work of Cleanup, so it must be last
pwsh_scripts: dict[str, bool] = {}
bginfo_folder = config.BGINFO_FOLDER

SEVERITY_RANK = {"info": 0, "warning": 1, "error": 2, "fatal": 3}
DEPLOYMENT_RECORDS_DIR = logs_dir / "deployment_records"


@dataclass
class ScriptResult:
    """Result of one installer script on one PC."""

    name: str
    severity: str = "info"
    messages: list[str] = field(default_factory=list)


@dataclass
class PcResult:
    """Aggregated deployment result for one PC."""

    pc: str
    highest_severity: str = "info"
    issues: list[dict[str, str]] = field(default_factory=list)
    scripts: list[ScriptResult] = field(default_factory=list)


result_lock = threading.Lock()
run_results: dict[str, PcResult] = {}

LOG_LEVEL_PATTERN = re.compile(
    r"^\s*(DEBUG|INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\s*[:\-]?\s*(.*)$",
    re.IGNORECASE,
)


def normalize_severity(level_name: str) -> str:
    """Return one of the severity values used by the result report."""

    match level_name.upper():
        case "WARN" | "WARNING":
            return "warning"
        case "ERROR":
            return "error"
        case "FATAL" | "CRITICAL":
            return "fatal"
        case _:
            return "info"


def record_result_event(
    pc: str,
    severity: str,
    message: str,
    component: str,
    script_name: str | None = None,
) -> None:
    """Record a warning or failure for the final per-PC report."""

    normalized = normalize_severity(severity)
    if normalized == "info":
        return

    with result_lock:
        pc_result = run_results[pc]
        if SEVERITY_RANK[normalized] > SEVERITY_RANK[pc_result.highest_severity]:
            pc_result.highest_severity = normalized

        if script_name is None:
            pc_result.issues.append(
                {
                    "component": component,
                    "severity": normalized,
                    "message": message,
                }
            )
            return

        script_result = next(
            script
            for script in reversed(pc_result.scripts)
            if script.name == script_name
        )
        if SEVERITY_RANK[normalized] > SEVERITY_RANK[script_result.severity]:
            script_result.severity = normalized
        script_result.messages.append(message)


def record_script_start(pc: str, script_name: str) -> None:
    """Add a script to the ordered list of scripts run for a PC."""

    with result_lock:
        run_results[pc].scripts.append(ScriptResult(name=script_name))


def log_script_line(pc: str, script_name: str, line: str, default_level: str) -> None:
    """Write a single captured line at the most appropriate logging level."""

    parsed_line = LOG_LEVEL_PATTERN.match(line)

    match parsed_line:
        case None:
            level_name = default_level
            message = line
        case _:
            level_name = parsed_line.group(1).upper()
            message = parsed_line.group(2) or line

    formatted_message = f"{pc}: {script_name}: {message}"
    record_result_event(pc, level_name, message, script_name, script_name)
    match level_name:
        case "DEBUG":
            log.debug(formatted_message)
        case "INFO":
            log.info(formatted_message)
        case "WARN" | "WARNING":
            log.warning(formatted_message)
        case "ERROR":
            log.error(formatted_message)
        case "FATAL" | "CRITICAL":
            log.critical(formatted_message)
        case _:
            log.info(formatted_message)


def log_process_output(
    pc: str, script_name: str, result: subprocess.CompletedProcess[str]
) -> None:
    """Write captured PowerShell output streams to the deployment log."""

    for stream_name, output in (("stdout", result.stdout), ("stderr", result.stderr)):
        match stream_name:
            case "stdout":
                default_level = "INFO"
            case "stderr":
                default_level = "ERROR"
            case _:
                default_level = "INFO"

        if output:
            for line in output.splitlines():
                log_script_line(pc, script_name, line, default_level)


def ping(pc: str) -> bool:
    """Return whether the target host responds to a single ping request."""

    result = subprocess.run(
        ["powershell.exe", "ping", "-n", "1", pc],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0


def test_winRM(pc: str) -> bool:
    """Return whether a basic WinRM command succeeds on the target host."""

    result = subprocess.run(
        [
            "powershell.exe",
            "Invoke-Command",
            "-ComputerName",
            pc,
            "-ScriptBlock",
            "{1}",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0


def check_online(pc: str) -> bool:
    """Validate that the target host is reachable and accepts WinRM commands."""

    if not ping(pc):
        log.warning(f"{pc}: Ping test failed")
        record_result_event(pc, "warning", "Ping test failed", "Connectivity")
        return False
    if test_winRM(pc):
        return True
    # PC probably hung / lost domain trust
    log.error(f"{pc}: WinRM test failed")
    record_result_event(pc, "error", "WinRM test failed", "Connectivity")
    return False


def RunCommand(pc: str) -> None:
    """Run remote configuration scripts for a target PC when checks fail."""

    if not check_online(pc):
        log.info(f"{pc}: Deployment complete with connectivity issues")
        return

    pwsh_script_items = pwsh_scripts.items()

    for pwsh_script_tuple in pwsh_script_items:
        script = pwsh_script_tuple[0]
        enabled = pwsh_script_tuple[1]

        if not enabled:
            continue

        record_script_start(pc, script)
        log.info(f"{pc}: Running {script}")
        script_path = project_root / Path(script)
        try:
            if "bginfo" in script.lower():  # Handle extra param for this script
                result = subprocess.run(
                    [
                        "powershell.exe",
                        "-File",
                        str(script_path),
                        pc,
                        "false",
                        bginfo_folder,
                    ],
                    capture_output=True,
                    text=True,
                    check=False,
                )
            else:
                result = subprocess.run(
                    ["powershell.exe", "-File", str(script_path), pc, "false"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
        except OSError as error:
            message = f"Could not start script: {error}"
            record_result_event(pc, "fatal", message, script, script)
            log.critical(f"{pc}: {script}: {message}")
            continue
        log_process_output(pc, script, result)
        if result.returncode == 0:
            log.info(f"{pc}: Completed {script}")
        else:
            failure_message = f"Failed with exit code {result.returncode}"
            record_result_event(pc, "error", failure_message, script, script)
            log.error(f"{pc}: {script} {failure_message.lower()}")

    log.info(f"{pc}: Deployment complete")


def parse_arguments() -> argparse.Namespace:
    """Parse optional CLI overrides used by both scripts and the Flutter app."""

    parser = argparse.ArgumentParser(
        description="Deploy the configured audio and display baseline to Windows PCs."
    )
    parser.add_argument(
        "--target",
        action="append",
        dest="targets",
        help="Target hostname. Repeat this option to bypass targets.txt.",
    )
    parser.add_argument(
        "--targets-file",
        type=Path,
        default=project_root / "targets.txt",
        help="Target file to use when --target is not provided.",
    )
    parser.add_argument(
        "--audio-recall",
        action=argparse.BooleanOptionalAction,
        default=config.AUDIO_RECALL,
        help="Install the audio recall feature.",
    )
    parser.add_argument(
        "--display-recall",
        action=argparse.BooleanOptionalAction,
        default=config.DISPLAY_RECALL,
        help="Install the display recall feature.",
    )
    parser.add_argument(
        "--bginfo-install",
        action=argparse.BooleanOptionalAction,
        default=config.BGINFO_INSTALL,
        help="Install BGInfo.",
    )
    parser.add_argument(
        "--add-desktop-shortcuts",
        action=argparse.BooleanOptionalAction,
        default=config.ADD_DESKTOP_SHORTCUTS,
        help="Add the save/recall desktop shortcuts.",
    )
    parser.add_argument(
        "--bginfo-folder",
        default=config.BGINFO_FOLDER,
        help="Folder name beneath the repository BGInfo directory.",
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=MAX_WORKERS,
        help="Maximum number of target PCs processed concurrently.",
    )
    parser.add_argument(
        "--result-file",
        type=Path,
        help="Optional JSON file that receives a structured per-PC result report.",
    )
    args = parser.parse_args()
    if args.max_workers < 1:
        parser.error("--max-workers must be at least 1")
    if args.bginfo_install and not args.bginfo_folder.strip():
        parser.error("--bginfo-folder cannot be empty when BGInfo is enabled")
    return args


def load_targets(args: argparse.Namespace) -> list[str]:
    """Return direct targets or load targets from the selected text file."""

    if args.targets:
        raw_targets = args.targets
    else:
        targets_file = args.targets_file
        if not targets_file.is_absolute():
            targets_file = project_root / targets_file
        if not targets_file.exists():
            raise FileNotFoundError(
                f"File {targets_file} does not exist. Create it or pass --target."
            )
        raw_targets = targets_file.read_text(encoding="utf-8").splitlines()

    targets = [
        target.strip()
        for target in raw_targets
        if target.strip() and not target.lstrip().startswith("#")
    ]
    if not targets:
        raise ValueError("No target PCs were provided.")
    return list(dict.fromkeys(targets))


def configure_run(args: argparse.Namespace) -> None:
    """Apply feature overrides for the current deployment run."""

    global bginfo_folder  # noqa: PLW0603
    global pwsh_scripts  # noqa: PLW0603

    bginfo_folder = args.bginfo_folder
    pwsh_scripts = {
        "installer_scripts\\./InstallAudioDeviceCmdlets.ps1": args.audio_recall,
        "installer_scripts\\./InstallDisplayConfig.ps1": args.display_recall,
        "installer_scripts\\./InstallBGInfo.ps1": args.bginfo_install,
        "installer_scripts\\./Cleanup.ps1": True,
        "installer_scripts\\./AddShortcuts.ps1": args.add_desktop_shortcuts,
    }


def write_result_report(result_file: Path | None) -> None:
    """Write the structured GUI report when a result path was requested."""

    if result_file is None:
        return
    if not result_file.is_absolute():
        result_file = project_root / result_file

    with result_lock:
        pc_results = [asdict(result) for result in run_results.values()]
    payload = {
        "log_file": str(log_file.resolve()),
        "pcs": pc_results,
    }
    result_file.parent.mkdir(parents=True, exist_ok=True)
    result_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    log.info(f"Result report: {result_file.resolve()}")


def deployment_record_path(pc: str) -> Path:
    """Return a stable, filesystem-safe record path for one target."""

    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", pc).strip("._") or "unknown"
    return DEPLOYMENT_RECORDS_DIR / f"{safe_name.lower()}.json"


def write_deployment_records(targets: list[str], args: argparse.Namespace) -> None:
    """Persist the most recent requested feature set for every target."""

    DEPLOYMENT_RECORDS_DIR.mkdir(parents=True, exist_ok=True)
    recorded_at = datetime.now().astimezone().isoformat()
    for pc in targets:
        payload = {
            "pc": pc,
            "recorded_at": recorded_at,
            "audio_recall": args.audio_recall,
            "display_recall": args.display_recall,
            "bginfo_install": args.bginfo_install,
            "desktop_shortcuts": args.add_desktop_shortcuts,
        }
        record_path = deployment_record_path(pc)
        temporary_path = record_path.with_suffix(f".{os.getpid()}.tmp")
        temporary_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temporary_path.replace(record_path)


def main() -> int:
    """Run a complete deployment and return a process exit code."""

    args = parse_arguments()
    try:
        targets = load_targets(args)
    except (FileNotFoundError, ValueError) as error:
        log.error(str(error))
        return 1

    configure_run(args)
    try:
        write_deployment_records(targets, args)
    except OSError as error:
        log.warning("Could not save deployment records: %s", error)
    with result_lock:
        run_results.clear()
        run_results.update({pc: PcResult(pc=pc) for pc in targets})
    log.info("Starting remote deployment run")
    log.info(f"Targets: {len(targets)}; maximum concurrent targets: {args.max_workers}")
    with futures.ThreadPoolExecutor(max_workers=args.max_workers) as executor:
        submitted_tasks = {}
        for pc in targets:
            log.info(f"{pc}: Queuing configuration check")
            submitted_tasks[executor.submit(RunCommand, pc)] = pc

        for task in futures.as_completed(submitted_tasks):
            pc = submitted_tasks[task]
            try:
                task.result()
            except Exception as error:
                message = f"Unexpected deployment failure: {error}"
                record_result_event(pc, "fatal", message, "Orchestrator")
                log.exception(f"{pc}: {message}")

    log.info("Remote deployment run complete")
    log.info(f"Detailed log: {log_file}")
    try:
        write_result_report(args.result_file)
    except OSError as error:
        log.error(f"Could not write result report: {error}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
