"""
Status Collector - Postcheck + Compare

Collects operational state from network devices after a maintenance window,
creates a postcheck ZIP archive, and compares the latest postcheck against
the latest precheck for the same change ticket.

This sample uses placeholder device names. Replace the inventory with devices
from your environment before running.
"""

from netmiko import ConnectHandler
from datetime import datetime
import os
import zipfile
import getpass
import difflib
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

from rich.console import Console
from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn

console = Console()

change_ticket = input("Change Ticket (Jira/ServiceNow/etc.): ").strip().upper()
username = input("Username: ")
password = getpass.getpass("Password: ")

arista_devices = [
    {"device_type": "arista_eos", "host": "arista-core-1.example.com", "username": username, "password": password},
    {"device_type": "arista_eos", "host": "arista-core-2.example.com", "username": username, "password": password},
]

cisco_devices = [
    {"device_type": "cisco_ios", "host": "cisco-core-1.example.com", "username": username, "password": password},
    {"device_type": "cisco_ios", "host": "cisco-core-2.example.com", "username": username, "password": password},
]

palo_devices = [
    {"device_type": "paloalto_panos", "host": "firewall-east.example.com", "username": username, "password": password},
    {"device_type": "paloalto_panos", "host": "firewall-west.example.com", "username": username, "password": password},
]

arista_commands = [
    "show mlag config-sanity",
    "show ip interface brief",
    "show vlan brief",
    "show mac address-table",
    "show ip arp",
    "show interfaces status",
    "show interfaces trunk",
    "show port-channel summary",
    "show interfaces transceiver",
    "show lldp neighbors",
    "show ip ospf neighbor",
    "show ip ospf interface",
    "show ip route ospf",
    "show ip bgp summary",
    "show ip route",
    "show ip bgp",
    "show running-config",
    "show version",
]

cisco_commands = [
    "show ip interface brief",
    "show vlan brief",
    "show mac address-table",
    "show ip arp",
    "show interfaces status",
    "show interfaces trunk",
    "show etherchannel summary",
    "show cdp neighbors",
    "show ip ospf neighbor",
    "show ip ospf interface",
    "show ip route ospf",
    "show ip bgp summary",
    "show ip route",
    "show ip bgp",
    "show running-config",
    "show version",
]

palo_commands = [
    "show system info",
    "show high-availability state",
    "show routing route",
    "show routing protocol bgp peer",
    "show routing protocol ospf neighbor",
    "set cli config-output-format set",
    "show config running",
]

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")

reports_dir = "Reports"
maintenance_dir = os.path.join(reports_dir, change_ticket)
precheck_dir = os.path.join(maintenance_dir, "Precheck")
postcheck_dir = os.path.join(maintenance_dir, "Postcheck")
compare_dir = os.path.join(maintenance_dir, "Compare")

os.makedirs(postcheck_dir, exist_ok=True)
os.makedirs(compare_dir, exist_ok=True)

folder_name = os.path.join(postcheck_dir, f"postcheck_{timestamp}")
os.makedirs(folder_name, exist_ok=True)

zip_name = os.path.join(postcheck_dir, f"postcheck_{timestamp}.zip")


def get_arista_hostname(conn, fallback):
    output = conn.send_command("show hostname")

    for line in output.splitlines():
        if line.startswith("Hostname:"):
            return line.split(":")[1].strip()

    return fallback


def get_cisco_hostname(conn, fallback):
    output = conn.send_command("show running-config | include ^hostname")

    for line in output.splitlines():
        if line.startswith("hostname"):
            return line.split()[1].strip()

    return fallback


def get_palo_hostname(conn, fallback):
    output = conn.send_command("show system info | match hostname")

    for line in output.splitlines():
        if line.startswith("hostname:"):
            return line.split(":")[1].strip()

    return fallback


def collect_device(device, commands, hostname_function, progress, overall_task):
    try:
        progress.console.log(f"Connecting to {device['host']}...")
        conn = ConnectHandler(**device)

        hostname = hostname_function(conn, device["host"])
        progress.console.log(f"Connected to {hostname}")

        file_path = os.path.join(folder_name, f"{hostname}.txt")

        with open(file_path, "w", encoding="utf-8") as file:
            file.write(f"Hostname: {hostname}\n")
            file.write(f"Address: {device['host']}\n")
            file.write(f"Generated: {datetime.now()}\n")
            file.write("=" * 80 + "\n")

            for command in commands:
                progress.update(overall_task, description=f"{hostname}")

                try:
                    output = conn.send_command(command, read_timeout=180)
                except Exception as cmd_error:
                    output = f"COMMAND FAILED:\n{cmd_error}"

                file.write(f"\n\n### {command} ###\n")
                file.write("-" * 80 + "\n")
                file.write(output)
                file.write("\n")

                progress.advance(overall_task, 1)

        conn.disconnect()

    except Exception as error:
        progress.console.log(f"FAILED: {device['host']}")
        progress.console.log(str(error))

        failed_file = os.path.join(folder_name, f"{device['host']}_FAILED.txt")

        with open(failed_file, "w", encoding="utf-8") as file:
            file.write(f"FAILED TO CONNECT TO {device['host']}\n")
            file.write(str(error))

        progress.advance(overall_task, len(commands))


def create_postcheck_zip():
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_name in os.listdir(folder_name):
            file_path = os.path.join(folder_name, file_name)
            zip_file.write(file_path, arcname=file_name)


def find_latest_folder(parent_dir, prefix):
    if not os.path.exists(parent_dir):
        return None

    folders = sorted([
        folder for folder in os.listdir(parent_dir)
        if folder.startswith(prefix)
        and os.path.isdir(os.path.join(parent_dir, folder))
    ])

    if not folders:
        return None

    return os.path.join(parent_dir, folders[-1])


def normalize_line(command, line):
    line = line.rstrip("\n")

    if command == "show interfaces transceiver":
        return None

    if command in ["show running-config", "show config running"]:
        return line

    noisy_starts = [
        "Generated:",
        "Uptime:",
        "Free memory:",
        "Last table change time",
        "Number of table inserts",
        "Number of table deletes",
        "time:",
        "uptime:",
        "url-filtering-version:",
        "Last update age:",
        "Update messages:",
        "Total messages:",
        "Flap counts:",
        "lifetime remain:",
    ]

    if any(line.strip().startswith(item) for item in noisy_starts):
        return None

    line = re.sub(r"\s+\d+:\d+:\d+ ago$", "", line)
    line = re.sub(r"\s+\d+ days?,.*ago$", "", line)

    if command == "show ip bgp summary":
        parts = line.split()

        if "Estab" in parts:
            estab_index = parts.index("Estab")
            return " ".join(parts[0:3] + parts[estab_index:])

        if "Idle(Admin)" in parts:
            idle_index = parts.index("Idle(Admin)")
            return " ".join(parts[0:3] + parts[idle_index:])

        return line

    if command == "show ip ospf neighbor":
        parts = line.split()

        if len(parts) >= 8:
            return " ".join(parts[0:5] + parts[6:])

        return line

    if command == "show ip arp":
        parts = line.split()

        if len(parts) >= 4 and re.match(r"\d+:\d+:\d+", parts[1]):
            return " ".join([parts[0]] + parts[2:])

        return line

    if command == "show mac address-table":
        line = re.sub(r"\s+\d+:\d+:\d+ ago$", "", line)
        line = re.sub(r"\s+\d+ days?,.*ago$", "", line)
        return line

    if command == "show routing route":
        parts = line.split()

        if len(parts) >= 5:
            return " ".join([p for p in parts if not p.isdigit()])

        return line

    if command == "show routing protocol bgp peer":
        stripped = line.strip()

        bgp_noise = [
            "Peer status:",
            "Update messages:",
            "Total messages:",
            "Last update age:",
            "Flap counts:",
        ]

        if any(stripped.startswith(item) for item in bgp_noise):
            if stripped.startswith("Peer status:"):
                if "," in stripped:
                    return stripped.split(",")[0]
                return stripped

            return None

        return line

    if command == "show routing protocol ospf neighbor":
        if line.strip().startswith("lifetime remain:"):
            return None

        return line

    if command == "show system info":
        stripped = line.strip()

        system_noise = [
            "time:",
            "uptime:",
            "url-filtering-version:",
            "global-protect-client-package-version:",
            "global-protect-clientless-vpn-version:",
            "app-version:",
            "av-version:",
            "threat-version:",
            "wildfire-version:",
        ]

        if any(stripped.startswith(item) for item in system_noise):
            return None

        return line

    return line


def parse_sections(file_path):
    sections = {}
    current_command = "HEADER"
    sections[current_command] = []

    with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
        for line in file.readlines():
            clean_line = line.rstrip("\n")

            if clean_line.startswith("### ") and clean_line.endswith(" ###"):
                current_command = clean_line.replace("###", "").strip()
                sections[current_command] = []
            else:
                normalized = normalize_line(current_command, clean_line)

                if normalized is None:
                    continue

                sections[current_command].append(normalized)

    return sections


def run_compare():
    precheck_folder = find_latest_folder(precheck_dir, "precheck_")
    postcheck_folder = find_latest_folder(postcheck_dir, "postcheck_")

    if precheck_folder is None:
        console.print("[yellow]No precheck folder found. Skipping compare.[/yellow]")
        return

    if postcheck_folder is None:
        console.print("[yellow]No postcheck folder found. Skipping compare.[/yellow]")
        return

    compare_file = os.path.join(compare_dir, f"compare_{timestamp}.txt")

    pre_files = sorted(os.listdir(precheck_folder))
    post_files = sorted(os.listdir(postcheck_folder))

    common_files = sorted(set(pre_files) & set(post_files))
    missing_post = sorted(set(pre_files) - set(post_files))
    new_post = sorted(set(post_files) - set(pre_files))

    with open(compare_file, "w", encoding="utf-8") as report:
        report.write("Pre/Post Maintenance Comparison Report\n")
        report.write("=" * 80 + "\n\n")
        report.write(f"Change Ticket:    {change_ticket}\n")
        report.write(f"Precheck Folder:  {precheck_folder}\n")
        report.write(f"Postcheck Folder: {postcheck_folder}\n\n")

        report.write("File Summary\n")
        report.write("-" * 80 + "\n")
        report.write(f"Common files: {len(common_files)}\n")
        report.write(f"Missing in postcheck: {len(missing_post)}\n")
        report.write(f"New in postcheck: {len(new_post)}\n\n")

        if missing_post:
            report.write("Missing in Postcheck:\n")
            for file_name in missing_post:
                report.write(f"- {file_name}\n")
            report.write("\n")

        if new_post:
            report.write("New in Postcheck:\n")
            for file_name in new_post:
                report.write(f"+ {file_name}\n")
            report.write("\n")

        for file_name in common_files:
            pre_path = os.path.join(precheck_folder, file_name)
            post_path = os.path.join(postcheck_folder, file_name)

            pre_sections = parse_sections(pre_path)
            post_sections = parse_sections(post_path)

            all_commands = sorted(set(pre_sections.keys()) | set(post_sections.keys()))

            report.write("\n")
            report.write("=" * 80 + "\n")
            report.write(f"Device/File: {file_name}\n")
            report.write("=" * 80 + "\n")

            device_changed = False

            for command in all_commands:
                pre_lines = pre_sections.get(command, [])
                post_lines = post_sections.get(command, [])

                if pre_lines == post_lines:
                    continue

                device_changed = True

                report.write("\n")
                report.write("-" * 80 + "\n")
                report.write(f"Command: {command}\n")
                report.write("-" * 80 + "\n")
                report.write("Differences detected.\n\n")

                diff = difflib.ndiff(pre_lines, post_lines)

                for line in diff:
                    if line.startswith("- "):
                        report.write(f"- {line[2:]}\n")
                    elif line.startswith("+ "):
                        report.write(f"+ {line[2:]}\n")

            if not device_changed:
                report.write("\nNo meaningful changes detected.\n")

    console.print(f"Compare report created: {compare_file}")


all_jobs = []

for device in arista_devices:
    all_jobs.append((device, arista_commands, get_arista_hostname))

for device in cisco_devices:
    all_jobs.append((device, cisco_commands, get_cisco_hostname))

for device in palo_devices:
    all_jobs.append((device, palo_commands, get_palo_hostname))

total_commands = sum(len(commands) for _, commands, _ in all_jobs)
max_workers = 5

with Progress(
    TextColumn("[bold blue]{task.description}"),
    BarColumn(),
    TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
    TimeElapsedColumn(),
    TimeRemainingColumn(),
    console=console,
) as progress:

    overall_task = progress.add_task("Postcheck Progress", total=total_commands)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []

        for device, commands, hostname_function in all_jobs:
            futures.append(
                executor.submit(
                    collect_device,
                    device,
                    commands,
                    hostname_function,
                    progress,
                    overall_task,
                )
            )

        for future in as_completed(futures):
            try:
                future.result()
            except Exception as error:
                progress.console.log(f"Thread failed: {error}")

create_postcheck_zip()
run_compare()

console.print()
console.print("[bold green]SUCCESS[/bold green]")
console.print(f"Postcheck ZIP created: {zip_name}")