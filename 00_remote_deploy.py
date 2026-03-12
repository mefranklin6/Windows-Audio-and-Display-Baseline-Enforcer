"""This script will cocurrently deploy all configuration scripts."""

from concurrent import futures
import logging as log
from pathlib import Path
import subprocess
import sys

MAX_WORKERS = 10

file_path = Path("targets.txt")
if not file_path.exists():
    log.error(
        f"File {file_path} does not exist. Please create it with the list of target PCs."
    )
    sys.exit(1)

with file_path.open("r") as f:
    TargetPCList = f.read().splitlines()


pwsh_scripts = [
    "installer_scripts\\./InstallAudioDeviceCmdlets.ps1",
    "installer_scripts\\./InstallDisplayConfig.ps1",
    # "installer_scripts\\./InstallBGInfo.ps1",
    "installer_scripts\\./Cleanup.ps1",  # Cleanup must be last
]


Executor = futures.ThreadPoolExecutor(max_workers=MAX_WORKERS)


def ping(ip: str) -> bool:
    """Return whether the target host responds to a single ping request."""

    result = subprocess.run(
        ["powershell.exe", "ping", "-n", "1", ip],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0


def test_winRM(ip: str) -> bool:
    """Return whether a basic WinRM command succeeds on the target host."""

    result = subprocess.run(
        [
            "powershell.exe",
            "Invoke-Command",
            "-ComputerName",
            ip,
            "-ScriptBlock",
            "{1}",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0


def check_online(ip: str) -> bool:
    """Validate that the target host is reachable and accepts WinRM commands."""

    if not ping(ip):
        log.warning(f"{ip}: Ping test failed")
        return False
    if test_winRM(ip):
        return True
    # PC probably hung / lost domain trust
    log.error(f"{ip}: //////////////// WinRM test failed ////////////////")
    return False


def RunCommand(pc: str) -> None:
    """Run remote configuration scripts for a target PC when checks fail."""

    if not check_online(pc):
        return

    for pwsh_script in pwsh_scripts:
        subprocess.run(["powershell.exe", pwsh_script, pc, "false"], check=False)


if __name__ == "__main__":
    for pc in TargetPCList:
        log.info(f"{pc}: Queuing configuration check")
        Executor.submit(RunCommand, pc)

    Executor.shutdown(wait=True)
