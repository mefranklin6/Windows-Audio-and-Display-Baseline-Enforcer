"""This script will conurrently deploy all configuration scripts."""

from concurrent import futures
from datetime import datetime
import logging as log
from pathlib import Path
import re
import subprocess
import sys

import config

MAX_WORKERS = 10  # Concurrent PC's to install on. Default is 10

now = datetime.now()
project_root = Path(__file__).parent
logs_dir = project_root / "logs"
logs_dir.mkdir(exist_ok=True)
log_file = logs_dir / f"{now.strftime('%Y-%m-%d_%H-%M-%S')}.log"
log.basicConfig(
    filename=log_file,
    level=log.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)

file_path = project_root / "targets.txt"
if not file_path.exists():
    log.error(
        f"File {file_path} does not exist. Please create it with the list of target PCs."
    )
    sys.exit(1)

with file_path.open("r") as f:
    TargetPCList = f.read().splitlines()

# Do not modify. Set feature flags in config.py now. See config.py.example
# {Installer Path (str) : Enabled (bool)}
#
# Ordering is important here:
#     - First, we need the individual installers to run
#     - Second, we need Cleanup to consilidate artifacts left by the installers
#     - Finally, add_shortcuts relies on consolidation work of Cleanup, so it must be last
pwsh_scripts = {
    "installer_scripts\\./InstallAudioDeviceCmdlets.ps1": config.AUDIO_RECALL,
    "installer_scripts\\./InstallDisplayConfig.ps1": config.DISPLAY_RECALL,
    "installer_scripts\\./InstallBGInfo.ps1": config.BGINFO_INSTALL,
    "installer_scripts\\./Cleanup.ps1": True,
    "installer_scripts\\./AddShortcuts.ps1": config.ADD_DESKTOP_SHORTCUTS,
}


Executor = futures.ThreadPoolExecutor(max_workers=MAX_WORKERS)

LOG_LEVEL_PATTERN = re.compile(
    r"^\s*(DEBUG|INFO|WARNING|WARN|ERROR)\s*[:\-]?\s*(.*)$",
    re.IGNORECASE,
)


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
    print(formatted_message)  # noqa: T201

    match level_name:
        case "DEBUG":
            log.debug(formatted_message)
        case "INFO":
            log.info(formatted_message)
        case "WARN" | "WARNING":
            log.warning(formatted_message)
        case "ERROR":
            log.error(formatted_message)
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
        return False
    if test_winRM(pc):
        return True
    # PC probably hung / lost domain trust
    log.error(f"{pc}: WinRM test failed")
    return False


def RunCommand(pc: str) -> None:
    """Run remote configuration scripts for a target PC when checks fail."""

    if not check_online(pc):
        return

    pwsh_script_items = pwsh_scripts.items()

    for pwsh_script_tuple in pwsh_script_items:
        script = pwsh_script_tuple[0]
        enabled = pwsh_script_tuple[1]

        if not enabled:
            continue

        log.info(f"{pc}: Running {script}")
        script_path = project_root / Path(script)
        if "bginfo" in script.lower():  # Handle extra param for this script
            result = subprocess.run(
                [
                    "powershell.exe",
                    "-File",
                    str(script_path),
                    pc,
                    "false",
                    config.BGINFO_FOLDER,
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
        log_process_output(pc, script, result)
        if result.returncode == 0:
            log.info(f"{pc}: Completed {script}")
        else:
            log.error(f"{pc}: {script} failed with exit code {result.returncode}")


if __name__ == "__main__":
    log.info("Starting remote deployment run")
    for pc in TargetPCList:
        log.info(f"{pc}: Queuing configuration check")
        Executor.submit(RunCommand, pc)

    Executor.shutdown(wait=True)
    log.info("Remote deployment run complete")
