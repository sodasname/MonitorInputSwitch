# USB Monitor Input Switch

Automatically switch monitor input source using **DDC/CI** when a specific **USB HID device** is plugged in or unplugged on Windows.

This tool runs as a lightweight **Windows service** and listens for USB device events.  
When the configured USB device appears or disappears, the monitor input is switched via **DDC/CI (VCP command)**.

This is useful for setups such as:

- Switching monitor input when connecting a **KVM USB device**
- Automatically changing input when plugging in a **laptop docking USB**
- Multi-computer setups sharing the same monitor
- Simple USB-triggered display automation

---

## Features

- Detect **USB HID device plug/unplug events**
- Automatically **switch monitor input source**
- Uses **DDC/CI VCP commands**
- No external tools required
- Runs as a **Windows service**
- Includes **console debug mode**
- **Runtime loading of Dxva2.dll** (no SDK dependency)
- **Debounce protection** to avoid repeated triggers
- Optional **USB removal switching** (compile-time option)

---

## How It Works

1. The service registers for **HID device interface notifications**.
2. When a device matching the configured **VID/PID** appears or disappears:
3. The program sends a **DDC/CI VCP command** to the monitor.
4. The monitor switches to the configured input source.

DDC/CI communication is performed via the Windows **Dxva2 API**.

---

## Requirements

- Windows 10 / Windows 11
- Monitor with **DDC/CI support**
- DDC/CI **enabled in monitor OSD**
- Target USB device with known **VID/PID**

---

## Configuration

Configuration values are defined in the source code.

Example:

```cpp
static const wchar_t* kTargetVidPid = L"VID_24AE&PID_4056";
static const wchar_t* kTargetDisplayDevice = L"\\\\.\\DISPLAY1";

static constexpr BYTE kVcpCode = 0x60;
static constexpr DWORD kVcpValueOnArrival = 18;
```

## Common VCP Codes

| VCP Code | Function     |
| -------- | ------------ |
| 0x60     | Input Source |

Typical input values (monitor dependent):

| Value | Input       |
| ----- | ----------- |
| 17    | HDMI 1      |
| 18    | HDMI 2      |
| 15    | DisplayPort |

Check your monitor documentation for exact values.

## Optional: USB Removal Switching

To switch the monitor input when the USB device is unplugged:

```c++
#define USB_REMOVE_VCP_VALUE 17
```

If this macro is not defined, unplug events will be ignored.

## Build

Using MSVC:

```powershell
cl /EHsc /O2 UsbDisplaySwitch.cpp /link /SUBSYSTEM:CONSOLE
```

Required libraries are linked via `#pragma comment(lib, ...)`.

## Running

### Debug Mode (Console)

Run directly:

```
UsbDisplaySwitch.exe
```

The program will run in **console debug mode** and print switching events.

Stop with:

```
Ctrl + C
```

### Service Mode

Register the service using `sc`:

```powershell
sc create UsbDisplaySwitch binPath= "C:\path\UsbDisplaySwitch.exe" start= auto
```

Start service:

```powershell
sc start UsbDisplaySwitch
```

Stop service:

```powershell
sc stop UsbDisplaySwitch
```

## Finding Your Monitor Name

To determine the display device:

1. Run `EnumDisplayMonitors` example tools
2. Or check Windows display enumeration tools

Typical values:

```
\\.\DISPLAY1
\\.\DISPLAY2
```

## Finding USB VID/PID

Use **Device Manager**:

1. Open Device Manager
2. Locate your USB device
3. Properties → Details
4. Select **Hardware Ids**

Example:

> VID_24AE&PID_4056

## Limitations

- Monitor must support **DDC/CI**
- Some monitors restrict input switching
- Multi-monitor setups may require adjusting `DISPLAY1 / DISPLAY2`

## Why Not Use ControlMyMonitor?

This tool performs **native DDC/CI control directly through the Windows API**.

Advantages:

- No external dependency
- Smaller footprint
- Better integration with service mode

## License

MIT License