# Hi, I'm Zack Rogers 👋

I'm a lead network engineer moving the WAN to infrastructure as code.

I build tools for the part of the job that happens at 2 a.m.: the change
is done, and you need to know what moved. Everything here is free to
use, and every example uses made-up devices and addresses.

**Writing, case studies, and free tools: [fnitguy.tech](https://fnitguy.tech)**

---

## Start here

### 🦀 [netshell](https://github.com/fnitguy-tech/netshell): `mw`, the maintenance check in one file

`mw` takes a snapshot of every device before a change and again after.
Then it shows you what moved.

```text
mw before NET-123 -r      capture before the change, secrets stripped
   (do the change)
mw after  NET-123 -r      capture again, quick text diff
mw report NET-123         the full HTML report
```

- **One file, nothing to install.** Copy it to a Windows, macOS, or
  Linux machine and run it.
- **It starts in a millisecond.** A 200-device report builds in 37 ms.
- **It only reads.** Every command it sends is a `show`.
- **It checks SSH host keys** before it sends your password.

It talks to Arista EOS, Cisco, Juniper Junos, and Palo Alto PAN-OS. I've
run it against real Arista and Palo Alto gear. Cisco and Juniper have
only met a test server so far, so try one device of each first.

No devices handy? `mw demo` replays a made-up change and writes both
reports. [Download v0.2.0](https://github.com/fnitguy-tech/netshell/releases/latest),
or read [how the rewrite went](https://fnitguy.tech/projects/mw-rust-port/).

### 🔍 [prepost-check](https://github.com/fnitguy-tech/prepost-check): the same check in Python

This is the original. Its report matches `mw` byte for byte, and the two
share a file format, so you can capture with one and report with the
other.

Pick it when your gear isn't on `mw`'s list. Its SSH library knows more
than a hundred device types. It's also the easier one to change if you
already live in Python.

`python3 scripts/demo.py` runs a made-up four-device change from start
to finish.

### 🧩 [project-template](https://github.com/fnitguy-tech/project-template): a starting point for automation projects

The skeleton I start every infrastructure project from:

- A dry run comes first, so you see the change before it happens.
- One bad host doesn't stop the rest.
- Secrets live in OpenBao, never in the repo.
- Tests run with no devices attached.

It's a GitHub template. Click **Use this template** to start from it.

---

## Free tools on fnitguy.tech

These run in your browser. Nothing to sign up for.

| Tool | What you get |
| --- | --- |
| [Design a network](https://fnitguy.tech/design/) | A diagram, an address plan, and a parts list from a few answers |
| [Is this change safe enough to run?](https://fnitguy.tech/change/) | A risk read on a change before the window |
| [What does the second box buy?](https://fnitguy.tech/resilience/) | Uptime math for redundancy, against your contract target |
| [What did the outage tell you?](https://fnitguy.tech/postmortem/) | A blameless post-mortem you can send |
| [Subnet it, then prove you can](https://fnitguy.tech/subnet/) | A subnet calculator with practice drills |

There's also a [MOP template](https://fnitguy.tech/projects/mop-template/)
with five worked examples, from an IS-IS migration to a switch upgrade.

---

## YouTube

The [`youtube/`](youtube/) folder holds the scripts and lab files for
videos on the **FN IT Guy** channel:

- Cisco switch STIG hardening
- Cisco Tcl scripts
- Juniper SRX upgrades
- Palo Alto upgrades
- Threat modeling

---

## What I work with

| Area | Tools |
| --- | --- |
| Networking | Cisco, Arista, Juniper, Palo Alto, Cradlepoint |
| Automation | Python, Rust, PowerShell, Netmiko, REST APIs, Git |
| Infrastructure | Linux (RHEL), KVM, Docker, Ceph, Pacemaker, Foreman/Katello |
| Observability and security operations | Prometheus, Grafana, Loki, Wazuh, Zabbix, NetBox |
| Identity and security | FreeIPA, FreeRADIUS, OpenBao, PKI, DISA STIG, OpenSCAP |

---

## Why I share this

Most network engineers solve the same problems alone, and the fix never
leaves their laptop. I write mine down so you can skip the part I
already got wrong.

Found a bug, or want a device type added? Open an issue on the repo.
Want to talk shop? [Reach out](https://fnitguy.tech/contact/).
