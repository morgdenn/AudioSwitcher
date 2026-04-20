"""
Audio Output Switcher — Windows 11 System Tray App
Uses Windows Core Audio API directly via comtypes (no pycaw for enumeration).
Left-click: cycle to next audio output device
Right-click: pick device from menu or Exit
"""

import sys
import ctypes
from ctypes import c_uint, c_ulong, c_ushort, c_wchar_p, Structure, Union, POINTER, byref

import comtypes
from comtypes import GUID, HRESULT, IUnknown, COMMETHOD, POINTER as COMPOINTER
from comtypes.client import CreateObject
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as Item, Menu


# ─── GUIDs ────────────────────────────────────────────────────────────────────

CLSID_MMDeviceEnumerator = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
CLSID_PolicyConfigClient = GUID("{870af99c-171d-4f9e-af0d-e63df40c2bc9}")

DEVICE_STATE_ACTIVE = 0x00000001
eRender  = 0   # output / playback
eConsole = 0   # default role
STGM_READ = 0


# ─── PROPVARIANT (for reading device friendly name) ───────────────────────────

class _PV_UNION(Union):
    _fields_ = [
        ("pwszVal", c_wchar_p),
        ("_pad",    ctypes.c_longlong),
    ]

class PROPVARIANT(Structure):
    _fields_ = [
        ("vt",         c_ushort),
        ("wReserved1", c_ushort),
        ("wReserved2", c_ushort),
        ("wReserved3", c_ushort),
        ("_u",         _PV_UNION),
    ]


# ─── PROPERTYKEY + PKEY_Device_FriendlyName ───────────────────────────────────

class PROPERTYKEY(Structure):
    _fields_ = [("fmtid", GUID), ("pid", c_ulong)]

PKEY_Device_FriendlyName = PROPERTYKEY()
PKEY_Device_FriendlyName.fmtid = GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}")
PKEY_Device_FriendlyName.pid   = 14

PKEY_DeviceInterface_FriendlyName = PROPERTYKEY()
PKEY_DeviceInterface_FriendlyName.fmtid = GUID("{026e516e-b814-414b-83cd-856d6fef4822}")
PKEY_DeviceInterface_FriendlyName.pid   = 2

VT_LPWSTR = 31


# ─── COM interfaces ───────────────────────────────────────────────────────────

class IPropertyStore(IUnknown):
    _iid_ = GUID("{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount",
            (["out"], COMPOINTER(c_ulong), "cProps")),
        COMMETHOD([], HRESULT, "GetAt",
            (["in"],  c_ulong, "iProp"),
            (["out"], COMPOINTER(PROPERTYKEY), "pkey")),
        COMMETHOD([], HRESULT, "GetValue",
            (["in"],  COMPOINTER(PROPERTYKEY), "key"),
            (["out"], COMPOINTER(PROPVARIANT), "pv")),
        COMMETHOD([], HRESULT, "SetValue",
            (["in"],  COMPOINTER(PROPERTYKEY), "key"),
            (["in"],  COMPOINTER(PROPVARIANT), "propvar")),
        COMMETHOD([], HRESULT, "Commit"),
    ]


class IMMDevice(IUnknown):
    _iid_ = GUID("{D666063F-1587-4E43-81F1-B948E807363F}")
    _methods_ = [
        COMMETHOD([], HRESULT, "Activate",
            (["in"],  COMPOINTER(GUID),            "iid"),
            (["in"],  c_uint,                      "dwClsCtx"),
            (["in"],  ctypes.c_void_p,             "pActivationParams"),
            (["out"], COMPOINTER(COMPOINTER(IUnknown)), "ppInterface")),
        COMMETHOD([], HRESULT, "OpenPropertyStore",
            (["in"],  c_uint, "stgmAccess"),
            (["out"], COMPOINTER(COMPOINTER(IPropertyStore)), "ppProperties")),
        COMMETHOD([], HRESULT, "GetId",
            (["out"], COMPOINTER(c_wchar_p), "ppstrId")),
        COMMETHOD([], HRESULT, "GetState",
            (["out"], COMPOINTER(c_uint), "pdwState")),
    ]


class IMMDeviceCollection(IUnknown):
    _iid_ = GUID("{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount",
            (["out"], COMPOINTER(c_uint), "pcDevices")),
        COMMETHOD([], HRESULT, "Item",
            (["in"],  c_uint, "nDevice"),
            (["out"], COMPOINTER(COMPOINTER(IMMDevice)), "ppDevice")),
    ]


class IMMDeviceEnumerator(IUnknown):
    _iid_ = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    _methods_ = [
        COMMETHOD([], HRESULT, "EnumAudioEndpoints",
            (["in"],  c_uint, "dataFlow"),
            (["in"],  c_uint, "dwStateMask"),
            (["out"], COMPOINTER(COMPOINTER(IMMDeviceCollection)), "ppDevices")),
        COMMETHOD([], HRESULT, "GetDefaultAudioEndpoint",
            (["in"],  c_uint, "dataFlow"),
            (["in"],  c_uint, "role"),
            (["out"], COMPOINTER(COMPOINTER(IMMDevice)), "ppEndpoint")),
        COMMETHOD([], HRESULT, "GetDevice",
            (["in"],  c_wchar_p, "pwstrId"),
            (["out"], COMPOINTER(COMPOINTER(IMMDevice)), "ppDevice")),
        COMMETHOD([], HRESULT, "RegisterEndpointNotificationCallback",
            (["in"], ctypes.c_void_p, "pClient")),
        COMMETHOD([], HRESULT, "UnregisterEndpointNotificationCallback",
            (["in"], ctypes.c_void_p, "pClient")),
    ]


class IPolicyConfig(IUnknown):
    _iid_ = GUID("{f8679f50-850a-41cf-9c72-430f290290c8}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetMixFormat"),
        COMMETHOD([], HRESULT, "GetDeviceFormat"),
        COMMETHOD([], HRESULT, "ResetDeviceFormat"),
        COMMETHOD([], HRESULT, "SetDeviceFormat"),
        COMMETHOD([], HRESULT, "GetProcessingPeriod"),
        COMMETHOD([], HRESULT, "SetProcessingPeriod"),
        COMMETHOD([], HRESULT, "GetShareMode"),
        COMMETHOD([], HRESULT, "SetShareMode"),
        COMMETHOD([], HRESULT, "GetPropertyValue"),
        COMMETHOD([], HRESULT, "SetPropertyValue"),
        COMMETHOD([], HRESULT, "SetDefaultEndpoint",
            (["in"], c_wchar_p, "wszDeviceId"),
            (["in"], c_uint,    "eRole")),
        COMMETHOD([], HRESULT, "SetEndpointVisibility"),
    ]


# ─── Audio helpers ─────────────────────────────────────────────────────────────

def _get_enumerator():
    return CreateObject(CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator)


def get_devices_friendly():
    """Return [(friendly_name, device_id), ...] for all active output devices."""
    results = []
    try:
        enum = _get_enumerator()
        collection = enum.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
        count = collection.GetCount()
        for i in range(count):
            device = collection.Item(i)
            try:
                dev_id = device.GetId()
                name   = _get_friendly_name(device) or f"Device {i+1}"
                results.append((name, dev_id))
            except Exception:
                pass
    except Exception as e:
        print(f"Enumeration error: {e}")
    return results


def _get_friendly_name(device):
    try:
        store = device.OpenPropertyStore(STGM_READ)
        for pkey in (PKEY_Device_FriendlyName, PKEY_DeviceInterface_FriendlyName):
            try:
                pv = store.GetValue(byref(pkey))
                if pv.vt == VT_LPWSTR and pv._u.pwszVal:
                    return pv._u.pwszVal
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_default_device_id():
    try:
        enum   = _get_enumerator()
        device = enum.GetDefaultAudioEndpoint(eRender, eConsole)
        return device.GetId()
    except Exception:
        return None


def set_default_device(device_id: str) -> bool:
    try:
        policy = CreateObject(CLSID_PolicyConfigClient, interface=IPolicyConfig)
        for role in range(3):   # eConsole, eMultimedia, eCommunications
            policy.SetDefaultEndpoint(device_id, role)
        return True
    except Exception as e:
        print(f"Switch error: {e}")
        return False


# ─── Tray icon ─────────────────────────────────────────────────────────────────

def make_icon() -> Image.Image:
    size = 64
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([0, 0, size-1, size-1], radius=12, fill=(30, 120, 255, 230))
    draw.polygon([(14,22),(26,22),(38,12),(38,52),(26,42),(14,42)], fill=(255,255,255,255))
    draw.arc([40,18,54,46], start=-60, end=60, fill=(255,255,255,200), width=3)
    draw.arc([44,24,58,40], start=-50, end=50, fill=(255,255,255,150), width=2)
    return img


# ─── App ───────────────────────────────────────────────────────────────────────

class AudioSwitcherApp:
    def __init__(self):
        self.devices      = []
        self.current_idx  = 0
        self.icon         = None
        self._refresh()

    def _refresh(self):
        self.devices = get_devices_friendly()
        default_id   = get_default_device_id()
        for i, (_, dev_id) in enumerate(self.devices):
            if dev_id == default_id:
                self.current_idx = i
                break

    def _current_name(self):
        return self.devices[self.current_idx][0] if self.devices else "No devices"

    def _cycle(self, icon=None, item=None):
        if not self.devices:
            return
        self.current_idx = (self.current_idx + 1) % len(self.devices)
        name, dev_id = self.devices[self.current_idx]
        if set_default_device(dev_id):
            if self.icon:
                self.icon.title = f"🔊 {name}"
            self._notify(f"Audio → {name}")

    def _select(self, index):
        def _handler(icon, item):
            self.current_idx = index
            name, dev_id = self.devices[index]
            if set_default_device(dev_id):
                if self.icon:
                    self.icon.title = f"🔊 {name}"
                    self.icon.menu = self._build_menu()
                    self.icon.update_menu()
                self._notify(f"Audio → {name}")
        return _handler

    def _notify(self, msg):
        try:
            if self.icon:
                self.icon.notify(msg, "Audio Switcher")
        except Exception:
            pass

    def _build_menu(self):
        """Build a fresh static menu (no lambdas — works across all pystray versions)."""
        default_id = get_default_device_id()
        items = [Item("Switch Device", self._cycle, default=True, visible=False)]
        for i, (name, dev_id) in enumerate(self.devices):
            tick = "✔  " if dev_id == default_id else "      "
            items.append(Item(tick + name, self._select(i)))
        items += [
            Menu.SEPARATOR,
            Item("🔄  Refresh", self._on_refresh),
            Menu.SEPARATOR,
            Item("Exit", self._on_exit),
        ]
        return pystray.Menu(*items)

    def _on_refresh(self, icon=None, item=None):
        self._refresh()
        if self.icon:
            self.icon.menu = self._build_menu()
            self.icon.update_menu()

    def _on_exit(self, icon=None, item=None):
        if self.icon:
            self.icon.stop()

    def run(self):
        self.icon = pystray.Icon(
            name  = "AudioSwitcher",
            icon  = make_icon(),
            title = f"🔊 {self._current_name()}",
            menu  = self._build_menu(),
        )
        self.icon.run()


# ─── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    comtypes.CoInitialize()
    app = AudioSwitcherApp()
    if not app.devices:
        ctypes.windll.user32.MessageBoxW(
            0,
            "No active audio output devices found.\nPlease check your audio settings.",
            "Audio Switcher",
            0x10,
        )
        sys.exit(1)
    app.run()
    comtypes.CoUninitialize()
