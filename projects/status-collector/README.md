# Status Collector

> Automated pre-maintenance and post-maintenance validation for enterprise network environments.

## Overview

Status Collector automates one of the most repetitive tasks performed by network engineers during maintenance windows: collecting operational data before and after changes and validating that the network returned to its expected state.

The tool connects to multiple devices in parallel, executes standardized operational commands, archives the results, and automatically compares post-maintenance output against the latest pre-maintenance baseline.

Rather than manually opening dozens of text files, engineers receive a structured comparison report highlighting only meaningful changes.

---

## Features

- Multi-threaded device collection
- Pre-maintenance data collection
- Post-maintenance data collection
- Automatic comparison against the latest precheck
- ZIP archive creation
- Timestamped reports
- Noise filtering to suppress expected runtime changes
- Vendor-specific command sets
- Rich progress bars for real-time status

---

## Supported Platforms

### Currently Supported

- Arista EOS
- Palo Alto PAN-OS

### Planned

- Cisco IOS / IOS XE
- Cisco NX-OS
- Juniper Junos

---

## Repository Structure

```text
status-collector/
│
├── README.md
├── CHANGELOG.md
├── requirements.txt
│
├── docs/
│   ├── Installation.md
│   └── Usage.md
│
├── reports/
├── sample-data/
├── scripts/
│   ├── precheck.py
│   └── postcheck.py
│
└── tests/
```

---

## Workflow

```text
            Maintenance Window

                  │
                  ▼

        Run precheck.py

                  │

      Collect Operational State

                  │

          Save & ZIP Reports

                  │

        Perform Maintenance

                  │

        Run postcheck.py

                  │

      Collect Operational State

                  │

      Automatically Compare

                  │

      Generate Difference Report

                  ▼

      Validate Network Health
```

---

## Example Output

```text
Reports/

└── CHG000123
    ├── Precheck
    │   ├── precheck_2026-06-27_1900
    │   └── precheck_2026-06-27_1900.zip
    │
    ├── Postcheck
    │   ├── postcheck_2026-06-27_2130
    │   └── postcheck_2026-06-27_2130.zip
    │
    └── Compare
        └── compare_2026-06-27_2130.txt
```

---

## Why This Project Exists

Many network maintenance activities still rely on engineers manually collecting command output, saving text files, and visually comparing results after a change.

Status Collector standardizes that workflow, reducing manual effort while making post-maintenance validation faster, more repeatable, and easier to review.

---

## Roadmap

### Version 0.1

- Precheck collection
- Postcheck collection
- Automated comparison
- Report generation
- ZIP archive creation

### Future

- External inventory files
- YAML command profiles
- HTML reports
- Configuration files
- Additional vendor support
- Plugin architecture

---

## Requirements

- Python 3.10+
- Netmiko
- Rich

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Disclaimer

This project is intended for lab and operational automation.

Always validate command sets and outputs in a non-production environment before deploying in production.