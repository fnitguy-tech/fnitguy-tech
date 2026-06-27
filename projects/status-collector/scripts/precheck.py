"""
Status Collector - Precheck

Collects operational state from network devices before a maintenance window.

This sample uses placeholder device names. Replace the inventory with devices
from your environment before running.
"""

from netmiko import ConnectHandler
from datetime import datetime
import os
import zipfile
import getpass
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

os.makedirs(precheck_dir, exist_ok=True)

folder_name = os.path.join(precheck_dir, f"precheck_{timestamp}")
os.makedirs(folder_name, exist_ok=True)

zip_name = os.path.join(precheck_dir, f"precheck_{timestamp}.zip")


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


def create_precheck_zip():
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_name in os.listdir(folder_name):
            file_path = os.path.join(folder_name, file_name)
            zip_file.write(file_path, arcname=file_name)


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

    overall_task = progress.add_task("Precheck Progress", total=total_commands)

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

create_precheck_zip()

console.print()
console.print("[bold green]SUCCESS[/bold green]")
console.print(f"Precheck ZIP created: {zip_name}")