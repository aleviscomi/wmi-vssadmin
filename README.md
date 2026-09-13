# wmi-vssadmin

> Remote VSS Shadow Copy Manager via DCOM/WMI.

[![Python](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://python.org)
[![Impacket](https://img.shields.io/badge/impacket-%3E%3D%200.11.0-orange.svg)](https://github.com/fortra/impacket)

`wmi-vssadmin` is a Python command-line tool that manages **Volume Shadow Copy Service (VSS)** snapshots on remote Windows hosts over a DCOM/WMI channel, using `Win32_ShadowCopy` class methods dispatched via Impacket's native RPC stack.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Command Reference](#command-reference)
- [Disclaimer](#disclaimer)
- [License](#license)

---

## Features

| Feature | Details |
|---|---|
| **Create** | Create a `ClientAccessible` VSS snapshot on any volume |
| **Delete** | Remove a snapshot by GUID *or* by shadow copy number |
| **List** | Enumerate every snapshot on the remote host in a sortable table |
| **Query** | Retrieve the full `Win32_ShadowCopy` property set for one snapshot |

---

## Installation

```bash
git clone https://github.com/aleviscomi/wmi-vssadmin.git
cd wmi-vssadmin
pip install -r requirements.txt
```

---

## Quick Start

```bash
# Create a shadow copy of C:\
python wmi_vssadmin.py CORP/Administrator:Password@192.168.1.10 -create

# List all shadow copies
python wmi_vssadmin.py CORP/Administrator:Password@192.168.1.10 -list

# Query shadow copy #3
python wmi_vssadmin.py CORP/Administrator:Password@192.168.1.10 -query 3

# Delete by GUID
python wmi_vssadmin.py CORP/Administrator:Password@192.168.1.10 \
    -delete '{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}'
```

---

## Command Reference

```
python wmi_vssadmin.py [-h] (-create | -delete GUID|N | -list | -query GUID|N) [-volume VOLUME] [-debug] [-hashes LMHASH:NTHASH] [-no-pass] [-k] [-aesKey HEX] target

POSITIONAL ARGUMENTS:
  target                [[domain/]username[:password]@]<targetName|IP address>

OPTIONS:
  -h, --help            show this help message and exit
  -create               Create a new ClientAccessible VSS shadow copy
  -delete GUID|N        Delete a shadow copy by GUID ({xxxxxxxx-...}) or shadow number
  -list                 List all shadow copies present on the remote host
  -query GUID|N         Display all WMI properties of a specific shadow copy

CREATE OPTIONS:
  -volume VOLUME        Volume to snapshot — trailing backslash required (default: C:\)

GENERAL OPTIONS:
  -debug                Enable DEBUG-level logging

AUTHENTICATION:
  -hashes LMHASH:NTHASH
                        NTLM hashes, format is LMHASH:NTHASH (LM hash may be empty: :NTHASH)
  -no-pass              Don't ask for password (useful for -k)
  -k                    Use Kerberos authentication. Grabs credentials from ccache file (KRB5CCNAME) based on target parameters. If valid credentials cannot be found, it will use the ones specified in the command line
  -aesKey HEX           AES key to use for Kerberos Authentication (128 or 256 bits)
```

---

## Disclaimer

This tool is designed for **authorised security research, defensive security testing, and system administration purposes**.

The tool provides functionality for managing Windows Volume Shadow Copies through DCOM/WMI. While the functionality itself is legitimate, Shadow Copies may provide access to data that is otherwise protected or inaccessible during normal system operation. As a result, the tool may be incorporated into security testing or other activities that could have security-sensitive or offensive implications.

Use this software only on systems and environments for which you have explicit authorisation.

The authors do not condone unauthorised access, data extraction, credential theft, or any other unlawful or harmful use of this software. Users are solely responsible for ensuring that their use of the tool complies with applicable laws, regulations, policies, and authorisation requirements.

**The authors accept no responsibility for misuse of the software or for any damage, data loss, or other consequences resulting from its use**.

---

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
