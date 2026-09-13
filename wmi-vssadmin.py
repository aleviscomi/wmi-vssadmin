#!/usr/bin/env python3
r"""
wmi-vssadmin — Remote VSS Shadow Copy Manager via DCOM/WMI
===========================================================

Manage Volume Shadow Copy Service (VSS) snapshots on remote Windows hosts
using a native DCOM/WMI transport layer.

ACTIONS (mutually exclusive, one required)
------------------------------------------
  -create            Create a new ClientAccessible shadow copy
  -delete <ref>      Delete a shadow copy by GUID or shadow number
  -list              Enumerate every shadow copy on the remote host
  -query  <ref>      Display full WMI properties of one shadow copy

  <ref> accepts either a GUID  ({xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx})
  or an integer shadow number  (e.g. 5  ->  HarddiskVolumeShadowCopy5).

SYNOPSIS
--------
  python wmi_vssadmin.py [domain/]user[:pass]@<host>  -create [-volume VOL]
  python wmi_vssadmin.py [domain/]user[:pass]@<host>  -list
  python wmi_vssadmin.py [domain/]user[:pass]@<host>  -query  <ref>
  python wmi_vssadmin.py [domain/]user[:pass]@<host>  -delete <ref>
  python wmi_vssadmin.py -hashes :NTHASH user@<host>  -create

REQUIREMENTS
------------
  Python   >= 3.7
  impacket >= 0.11.0
"""

from __future__ import annotations

import re
import sys
import logging
import argparse
from datetime import datetime, timezone, timedelta
from typing import Any, Dict, Iterator, List, Optional, Tuple

from impacket.examples import logger as impacket_logger
from impacket.examples.utils import parse_target
from impacket import version as impacket_version
from impacket.dcerpc.v5 import dcomrt
from impacket.dcerpc.v5.dcom import wmi
from impacket.dcerpc.v5.dtypes import NULL


# =============================================================================
#  Package metadata
# =============================================================================

__version__ = "1.0"
__author__  = "@aleviscomi"
__license__ = "Apache-2.0"

# =============================================================================
#  ASCII Art
# =============================================================================

ascii_art = "\033[33m" + r"""
               .__                                          .___      .__        
__  _  _______ |__|         ___  ________ ___________     __| _/_____ |__| ____  
\ \/ \/ /     \|  |  ______ \  \/ /  ___//  ___/\__  \   / __ |/     \|  |/    \ 
 \     /  Y Y  \  | /_____/  \   /\___ \ \___ \  / __ \_/ /_/ |  Y Y  \  |   |  \
  \/\_/|__|_|  /__|           \_//____  >____  >(____  /\____ |__|_|  /__|___|  /
             \/                       \/     \/      \/      \/     \/        \/ 

Version: %s
Author:  %s
""" % (__version__, __author__) + "\033[0m"

# =============================================================================
#  WMI / VSS constants
# =============================================================================

_WMI_NAMESPACE = "//./root/cimv2"
_WMI_CLASS     = "Win32_ShadowCopy"

# Win32_ShadowCopy::Create() method return codes.
# Reference: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/vsswmi/create-method-in-class-win32-shadowcopy
_VSS_CREATE_CODES: Dict[int, str] = {
    0: "Success",
    1: "Access denied",
    2: "Invalid argument",
    3: "Specified volume not found",
    4: "Volume not supported",
    5: "Shadow copy context not supported",
    6: "Insufficient storage available for provider",
    7: "Volume is in use / could not be locked",
    8: "Maximum number of shadow copies reached (512 per volume)",
    9: "Another VSS operation is already in progress",
}

# Win32_ShadowCopy.State property value descriptions.
_VSS_STATES: Dict[int, str] = {
    0:  "Unknown",
    1:  "Preparing",
    2:  "ProcessingPrepare",
    3:  "Prepared",
    4:  "ProcessingPrecommit",
    5:  "Precommitted",
    6:  "ProcessingCommit",
    7:  "Committed",
    8:  "ProcessingPostcommit",
    9:  "Created",
    10: "Aborted",
    11: "Deleted",
    12: "Count",
}

# Compiled GUID pattern: accepts braces as optional.
_GUID_RE = re.compile(
    r"^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}"
    r"-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$"
)

# Matches the integer suffix in a DeviceObject path (e.g. HarddiskVolumeShadowCopy5).
_DEVICE_NUM_RE = re.compile(r"HarddiskVolumeShadowCopy(\d+)\s*$", re.IGNORECASE)


# =============================================================================
#  Pure utility helpers
# =============================================================================

def is_guid(value: str) -> bool:
    """Return ``True`` if *value* matches a VSS GUID pattern (braces optional)."""
    return bool(_GUID_RE.match(value.strip()))


def normalise_guid(value: str) -> str:
    """Return *value* as ``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`` (upper-case)."""
    return "{%s}" % value.strip("{}").upper()


def parse_wmi_datetime(raw: Optional[str]) -> str:
    """
    Convert a WMI DMTF datetime string to a human-readable ISO-8601 timestamp.

    WMI format::

        YYYYMMDDHHmmSS.ffffff+/-UUU   (UUU = UTC offset in minutes)

    Example: ``20240601143022.000000+000`` -> ``2024-06-01 14:30:22 UTC+0000``

    Returns the raw string on parse failure so no information is lost.
    """
    if not raw:
        return "N/A"
    try:
        sign   = 1 if raw[21] == "+" else -1
        offset = sign * int(raw[22:25])
        tz     = timezone(timedelta(minutes=offset))
        dt     = datetime.strptime(raw[:14], "%Y%m%d%H%M%S").replace(tzinfo=tz)
        return dt.strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception:
        return raw


def shadow_number_from_device(device_object: Optional[str]) -> Optional[int]:
    """
    Extract the integer *N* from a ``DeviceObject`` path of the form::

        \\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopyN

    Returns ``None`` if the path is empty or the pattern is absent.
    """
    if not device_object:
        return None
    match = _DEVICE_NUM_RE.search(device_object)
    return int(match.group(1)) if match else None


def safe_get(obj: Any, attr: str, default: Any = "") -> Any:
    """
    Retrieve attribute *attr* from a WMI object, returning *default* on any
    error.  Silently handles ``AttributeError`` and ``None`` property values.
    """
    try:
        value = getattr(obj, attr)
        return value if value is not None else default
    except Exception:
        return default


# =============================================================================
#  Data model
# =============================================================================

class ShadowCopyRecord:
    """
    Structured snapshot of a ``Win32_ShadowCopy`` WMI instance.

    All fields are populated via :func:`safe_get` so the constructor never
    raises even when properties are absent or null on the remote host.
    The *number* field is derived from the ``DeviceObject`` path suffix rather
    than a native WMI property.
    """

    __slots__ = (
        # Identification
        "shadow_id", "set_id", "number", "device_object",
        # Volume / timing
        "volume_name", "installed_on",
        # Origin
        "originating_machine", "service_machine",
        # State
        "state", "state_label",
        # Boolean flags
        "persistent", "client_accessible", "no_auto_release",
        "exposed_locally",  "exposed_remotely",
        "exposed_name",     "exposed_path",
        "hardware_assisted","transportable",
        "differential",     "plex",     "imported",
    )

    def __init__(self, wmi_obj: Any) -> None:
        device = safe_get(wmi_obj, "DeviceObject")
        state  = safe_get(wmi_obj, "State", 0)

        self.shadow_id           = safe_get(wmi_obj, "ID")
        self.set_id              = safe_get(wmi_obj, "SetID")
        self.number              = shadow_number_from_device(device)
        self.device_object       = device
        self.volume_name         = safe_get(wmi_obj, "VolumeName")
        self.installed_on        = parse_wmi_datetime(safe_get(wmi_obj, "InstallDate", None))
        self.originating_machine = safe_get(wmi_obj, "OriginatingMachine")
        self.service_machine     = safe_get(wmi_obj, "ServiceMachine")
        self.state               = int(state or 0)
        self.state_label         = _VSS_STATES.get(self.state, "Unknown")
        self.persistent          = bool(safe_get(wmi_obj, "Persistent",       False))
        self.client_accessible   = bool(safe_get(wmi_obj, "ClientAccessible", False))
        self.no_auto_release     = bool(safe_get(wmi_obj, "NoAutoRelease",    False))
        self.exposed_locally     = bool(safe_get(wmi_obj, "ExposedLocally",   False))
        self.exposed_remotely    = bool(safe_get(wmi_obj, "ExposedRemotely",  False))
        self.exposed_name        = safe_get(wmi_obj, "ExposedName")
        self.exposed_path        = safe_get(wmi_obj, "ExposedPath")
        self.hardware_assisted   = bool(safe_get(wmi_obj, "HardwareAssisted", False))
        self.transportable       = bool(safe_get(wmi_obj, "Transportable",    False))
        self.differential        = bool(safe_get(wmi_obj, "Differential",     False))
        self.plex                = bool(safe_get(wmi_obj, "Plex",             False))
        self.imported            = bool(safe_get(wmi_obj, "Imported",         False))

    # ── Display helpers ────────────────────────────────────────────────────

    def _num_tag(self) -> str:
        """Return a short label such as ``#5`` or ``#?`` when number is unknown."""
        return "#%d" % self.number if self.number is not None else "#?"

    def print_brief(self) -> None:
        """Single-row summary for tabular list output."""
        date_col = self.installed_on[:10] if self.installed_on not in ("N/A", "") else "N/A"
        print("  %-5s  %-40s  %-12s  %s" % (
            self._num_tag(), self.shadow_id, date_col, self.volume_name,
        ))

    def print_summary(self) -> None:
        """Compact multi-line block used for create output."""
        rows = [
            ("ID",          self.shadow_id),
            ("Device",      self.device_object),
            ("Volume",      self.volume_name),
            ("Created",     self.installed_on),
            ("Machine",     self.originating_machine),
            ("State",       "%d (%s)" % (self.state, self.state_label)),
            ("Persistent",  str(self.persistent)),
            ("Client Acc.", str(self.client_accessible)),
        ]
        print("  Shadow Copy %s" % self._num_tag())
        for label, value in rows:
            print("    %-14s %s" % (label + ":", value))

    def print_full(self) -> None:
        """All WMI properties in labelled rows for -query output."""
        rows = [
            ("Shadow #",          self._num_tag()),
            ("ID",                self.shadow_id),
            ("Set ID",            self.set_id),
            ("Device",            self.device_object),
            ("Volume",            self.volume_name),
            ("Created",           self.installed_on),
            ("Origin Machine",    self.originating_machine),
            ("Service Machine",   self.service_machine),
            ("State",             "%d (%s)" % (self.state, self.state_label)),
            ("Persistent",        self.persistent),
            ("Client Accessible", self.client_accessible),
            ("No Auto Release",   self.no_auto_release),
            ("Exposed Locally",   self.exposed_locally),
            ("Exposed Remotely",  self.exposed_remotely),
            ("Exposed Name",      self.exposed_name  or "(none)"),
            ("Exposed Path",      self.exposed_path  or "(none)"),
            ("Hardware Assisted", self.hardware_assisted),
            ("Transportable",     self.transportable),
            ("Differential",      self.differential),
            ("Plex",              self.plex),
            ("Imported",          self.imported),
        ]
        for label, value in rows:
            print("    %-22s %s" % (label + ":", value))


# =============================================================================
#  Core manager
# =============================================================================

class WMIVSSAdmin:
    """
    Remote VSS Shadow Copy manager via DCOM/WMI (Impacket >= 0.11).

    All four public actions (:meth:`action_create`, :meth:`action_delete`,
    :meth:`action_list`, :meth:`action_query`) share the same DCOM/WMI setup
    routine (:meth:`_connect`) and the same WQL result iterator
    (:meth:`_iter_query`).  Each action owns its ``dcom.disconnect()`` call
    inside a ``finally`` block to guarantee the transport is closed regardless
    of the outcome.
    """

    def __init__(
        self,
        username: str = "",
        password: str = "",
        domain:   str = "",
        options:  Any = None,
    ) -> None:
        self._username = username
        self._password = password
        self._domain   = domain
        self._options  = options
        self._lmhash   = ""
        self._nthash   = ""

        if options.hashes:
            parts        = options.hashes.split(":", 1)
            self._lmhash = parts[0]
            self._nthash = parts[1] if len(parts) > 1 else ""

    # ── DCOM / WMI setup ──────────────────────────────────────────────────

    def _connect(self, host: str) -> Tuple[dcomrt.DCOMConnection, Any]:
        """
        Open a DCOM connection to *host* and authenticate into the
        ``root/cimv2`` WMI namespace.

        Returns ``(dcom, IWbemServices)``.
        The caller **must** call ``dcom.disconnect()`` when finished.

        Connection sequence:
          1. ``DCOMConnection`` — TCP 135 endpoint mapper + dynamic RPC port
          2. ``CoCreateInstanceEx`` — activates ``CLSID_WbemLevel1Login``
          3. ``IWbemLevel1Login.NTLMLogin`` — obtains ``IWbemServices``
        """
        logging.debug("Opening DCOMConnection to %s", host)
        dcom = dcomrt.DCOMConnection(
            host,
            self._username,
            self._password,
            self._domain,
            self._lmhash,
            self._nthash,
            oxidResolver=True,
        )
        try:
            iface         = dcom.CoCreateInstanceEx(
                wmi.CLSID_WbemLevel1Login,
                wmi.IID_IWbemLevel1Login,
            )
            login         = wmi.IWbemLevel1Login(iface)
            wbem_services = login.NTLMLogin(_WMI_NAMESPACE, NULL, NULL)
            login.RemRelease()
            logging.debug("WMI session established on namespace %s", _WMI_NAMESPACE)
        except Exception:
            try:
                dcom.disconnect()
            except Exception:
                pass
            raise
        return dcom, wbem_services

    # ── WQL iteration ─────────────────────────────────────────────────────

    def _iter_query(self, wbem_services: Any, wql: str) -> Iterator[Any]:
        """
        Execute *wql* and yield each result object one at a time.

        Iteration ends silently when Impacket raises on an exhausted
        ``IEnumWbemClassObject`` — the conventional pattern across all
        Impacket WMI tools.
        """
        logging.debug("WQL -> %s", wql)
        ienum = wbem_services.ExecQuery(wql)
        while True:
            try:
                result = ienum.Next(0xFFFF, 1)
                yield result[0]
            except Exception:
                break

    # ── Reference resolution ──────────────────────────────────────────────

    def _resolve_to_guid(self, wbem_services: Any, ref: str) -> str:
        """
        Resolve *ref* to a canonical ``{GUID}`` string.

        *ref* is accepted as:

        - **GUID** (braces optional):
          ``{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}``
        - **Integer string** representing the shadow copy number (``N`` in
          ``HarddiskVolumeShadowCopyN``): resolved via a full WQL enumeration.

        Raises :class:`ValueError` for unrecognised format and
        :class:`LookupError` when a numeric reference matches no existing copy.
        """
        ref = ref.strip()

        if is_guid(ref):
            return normalise_guid(ref)

        if ref.lstrip("-").isdigit():
            target_number = int(ref)
            for obj in self._iter_query(
                wbem_services, "SELECT * FROM %s" % _WMI_CLASS
            ):
                record = ShadowCopyRecord(obj)
                if record.number == target_number:
                    return normalise_guid(record.shadow_id)
            raise LookupError(
                "No shadow copy with number %d found on the remote host."
                % target_number
            )

        raise ValueError(
            "Invalid reference %r — expected a GUID or an integer shadow number."
            % ref
        )

    def _fetch_record(self, wbem_services: Any, ref: str) -> ShadowCopyRecord:
        """
        Resolve *ref* to a GUID and return a populated :class:`ShadowCopyRecord`.

        Raises :class:`LookupError` when the shadow copy cannot be found.
        """
        guid  = self._resolve_to_guid(wbem_services, ref)
        query = "SELECT * FROM %s WHERE ID='%s'" % (_WMI_CLASS, guid)
        for obj in self._iter_query(wbem_services, query):
            return ShadowCopyRecord(obj)
        raise LookupError("Shadow copy %s not found on the remote host." % guid)

    # ── Public actions ────────────────────────────────────────────────────

    def action_create(self, host: str) -> None:
        """
        Invoke ``Win32_ShadowCopy::Create`` on *host* and print the resulting
        snapshot details.

        A single follow-up WQL query is issued after the Create RPC
        (``SELECT * FROM Win32_ShadowCopy WHERE ID='...'``) to retrieve
        the ``DeviceObject`` path and shadow number.
        """
        volume = self._options.volume
        logging.info("Connecting to %s ...", host)
        dcom, wbem_services = self._connect(host)

        try:
            shadow_class, _ = wbem_services.GetObject(_WMI_CLASS)

            logging.info(
                "Invoking Win32_ShadowCopy::Create"
                " (Volume=%s, Context=ClientAccessible) ...", volume
            )

            # Dispatch via Impacket >= 0.11 __getattr__ WMI method dispatcher.
            out_params = shadow_class.Create(volume, "ClientAccessible")

            if out_params is None:
                raise RuntimeError(
                    "Win32_ShadowCopy::Create returned no output parameters."
                )

            return_code = int(out_params.ReturnValue)
            if return_code != 0:
                description = _VSS_CREATE_CODES.get(return_code, "Unknown error")
                raise RuntimeError(
                    "Win32_ShadowCopy::Create failed — ReturnValue=%d: %s"
                    % (return_code, description)
                )

            shadow_id = str(out_params.ShadowID)
            logging.info("Creation succeeded — ShadowID: %s", shadow_id)
            logging.debug("Fetching full metadata for the new snapshot ...")

            record = self._fetch_record(wbem_services, shadow_id)

            num_label = " Shadow Copy %s" % record._num_tag() if record.number else ""
            print()
            print("[+] Created:%s" % num_label)
            print()
            record.print_summary()
            print()
            print("[i] Reference this snapshot later with:")
            print("    -query / -delete  %s" % record.shadow_id)
            if record.number is not None:
                print("    -query / -delete  %d" % record.number)

        finally:
            dcom.disconnect()

    def action_delete(self, host: str) -> None:
        """
        Delete the shadow copy identified by ``-delete <ref>`` on *host*.

        Retrieves the ``Win32_ShadowCopy`` instance via ``GetObject`` using the
        key-property object path (``Win32_ShadowCopy.ID="..."``) then dispatches
        the ``Delete`` method through Impacket's ``__getattr__`` WMI dispatcher.

        The trailing underscore in ``Delete_()`` follows the Impacket convention
        of disambiguating WMI method names from Python built-ins at the call site.
        """
        ref = self._options.delete
        logging.info("Connecting to %s ...", host)
        dcom, wbem_services = self._connect(host)

        try:
            guid = self._resolve_to_guid(wbem_services, ref)
            logging.info("Resolved '%s' to %s", ref, guid)

            # Build the WMI instance object path using the key property.
            # Format: ClassName.KeyProperty="value"
            obj_path = '%s.ID="%s"' % (_WMI_CLASS, guid)
            logging.info("Invoking Win32_ShadowCopy::Delete on %s ...", guid)
            
            wbem_services.DeleteInstance(obj_path)


            print()
            print("[+] Shadow copy %s deleted successfully." % guid)

        finally:
            dcom.disconnect()

    def action_list(self, host: str) -> None:
        """
        Enumerate all ``Win32_ShadowCopy`` instances on *host*, sorted by
        shadow copy number in ascending order.
        """
        logging.info("Connecting to %s ...", host)
        dcom, wbem_services = self._connect(host)

        try:
            records: List[ShadowCopyRecord] = []
            for obj in self._iter_query(
                wbem_services, "SELECT * FROM %s" % _WMI_CLASS
            ):
                records.append(ShadowCopyRecord(obj))

            # Sort by shadow number ascending; unknowns are pushed to the end.
            records.sort(
                key=lambda r: r.number if r.number is not None else 0x7FFFFFFF
            )

            print()
            if not records:
                print("[i] No VSS shadow copies found on %s." % host)
                return

            count = len(records)
            print("[+] %d shadow %s on %s:\n" % (
                count, "copy" if count == 1 else "copies", host,
            ))
            print("  %-5s  %-40s  %-12s  %s" % ("#", "GUID", "Date", "Volume"))
            print("  " + "-" * 78)
            for record in records:
                record.print_brief()
            print()
            print("[i] Use -query <GUID|N> for full details on any entry above.")

        finally:
            dcom.disconnect()

    def action_query(self, host: str) -> None:
        """Display all WMI properties of the shadow copy identified by *ref*."""
        ref = self._options.query
        logging.info("Connecting to %s ...", host)
        dcom, wbem_services = self._connect(host)

        try:
            record = self._fetch_record(wbem_services, ref)

            print()
            print("[+] Shadow copy details (Shadow Copy %s):" % record._num_tag())
            print()
            record.print_full()
            print()

        finally:
            dcom.disconnect()


# =============================================================================
#  Command-line interface
# =============================================================================

def _build_parser() -> argparse.ArgumentParser:
    """Construct and return the argument parser for wmi-vssadmin."""
    parser = argparse.ArgumentParser(
        prog="wmi_vssadmin.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=(
            "Remote VSS Shadow Copy Manager via DCOM/WMI.\n\n"
            #"  -create            Create a new ClientAccessible shadow copy\n"
            #"  -delete <GUID|N>   Delete shadow copy by GUID or shadow number\n"
            #"  -list              List all shadow copies on the remote host\n"
            #"  -query  <GUID|N>   Display full properties of one shadow copy\n"
        )
    )

    # Positional: target string in standard Impacket format
    parser.add_argument(
        "target",
        help="[[domain/]username[:password]@]<targetName|IP address>",
    )

    # Actions — mutually exclusive, exactly one is required
    actions = parser.add_mutually_exclusive_group(required=True)
    actions.add_argument(
        "-create",
        action="store_true",
        help="Create a new ClientAccessible VSS shadow copy",
    )
    actions.add_argument(
        "-delete",
        metavar="GUID|N",
        help="Delete a shadow copy by GUID ({xxxxxxxx-...}) or shadow number",
    )
    actions.add_argument(
        "-list",
        action="store_true",
        help="List all shadow copies present on the remote host",
    )
    actions.add_argument(
        "-query",
        metavar="GUID|N",
        help="Display all WMI properties of a specific shadow copy",
    )

    # Options specific to -create
    create_group = parser.add_argument_group("create options")
    create_group.add_argument(
        "-volume",
        default=r"C:\\",
        metavar="VOLUME",
        help=r"Volume to snapshot — trailing backslash required (default: C:\)",
    )

    # General options
    general_group = parser.add_argument_group("general options")
    general_group.add_argument(
        "-debug",
        action="store_true",
        help="Enable DEBUG-level logging",
    )

    # Authentication — mirrors Impacket's standard option set
    auth_group = parser.add_argument_group("authentication")
    auth_group.add_argument(
        "-hashes",
        metavar="LMHASH:NTHASH",
        help="NTLM hashes, format is LMHASH:NTHASH (LM hash may be empty: :NTHASH)",
    )
    auth_group.add_argument(
        "-no-pass",
        action="store_true",
        help="Don't ask for password (useful for -k)",
    )
    auth_group.add_argument(
        "-k",
        action="store_true",
        help="Use Kerberos authentication. Grabs credentials from ccache file (KRB5CCNAME) based on target parameters. If valid credentials cannot be found, it will use the ones specified in the command line",
    )
    auth_group.add_argument(
        "-aesKey",
        metavar="HEX",
        help="AES key to use for Kerberos Authentication (128 or 256 bits)",
    )

    return parser


def main() -> None:
    """Entry point: parse arguments, resolve credentials and dispatch action."""
    impacket_logger.init()
    print(impacket_version.BANNER)
    
    print(ascii_art)

    parser = _build_parser()
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    options = parser.parse_args()
    logging.getLogger().setLevel(logging.DEBUG if options.debug else logging.INFO)

    domain, username, password, host = parse_target(options.target)
    if domain is None:
        domain = ""

    # Prompt for a password only when no alternative credential is supplied.
    if (
        password  == ""
        and username  != ""
        and options.hashes  is None
        and not options.no_pass
        and options.aesKey  is None
    ):
        from getpass import getpass
        password = getpass("Password for %s: " % username)

    manager = WMIVSSAdmin(username, password, domain, options)

    try:
        if options.create:
            manager.action_create(host)
        elif options.delete:
            manager.action_delete(host)
        elif options.list:
            manager.action_list(host)
        elif options.query:
            manager.action_query(host)

    except LookupError as exc:
        logging.error("Not found: %s", exc)
        sys.exit(2)
    except ValueError as exc:
        logging.error("Invalid argument: %s", exc)
        sys.exit(2)
    except KeyboardInterrupt:
        logging.warning("Interrupted by user.")
        sys.exit(130)
    except Exception as exc:
        logging.error(str(exc))
        if options.debug:
            raise
        sys.exit(1)


if __name__ == "__main__":
    main()
