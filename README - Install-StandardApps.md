# Install-StandardApps

A PowerShell workstation-provisioning script that checks for and silently installs a baseline set of applications: **Microsoft 365 Apps (Office)**, **Google Chrome**, and **Adobe Acrobat Reader**.

It's built for repeatable, hands-off deployment — run it on a fresh machine (interactively or pushed through an RMM agent) and it brings the box up to a standard application baseline, skipping anything already present.

---

## What it does

For each application, the script **detects first and only installs if missing**, so it's safe to run repeatedly (idempotent):

- **Microsoft Office** — reports the status of Word, Excel, PowerPoint, and classic Outlook, plus whether the new Outlook for Windows is present. Installs Microsoft 365 Apps **only when no core Office app** (Word/Excel/PowerPoint) is detected.
- **Google Chrome** — installs via WinGet if not already present.
- **Adobe Acrobat Reader** — installs via WinGet if neither Acrobat nor Reader is present.

Everything is logged to the console and to a timestamped transcript file in `%TEMP%`.

> **Note on Outlook:** A missing *classic* Outlook does **not** trigger a full Office reinstall. As long as one core Office app is present, Office is treated as installed. This is deliberate — many machines run new Outlook or have classic Outlook removed without warranting a full suite redeploy.

---

## How detection works

The script uses several signals so it doesn't reinstall software that's already there:

- **App Paths registry keys** (`HKLM`/`HKCU\...\App Paths\<exe>`) for the app executables.
- **Uninstall registry entries** (display-name pattern match) as a secondary check for Chrome and Adobe.
- **On-disk fallback paths** in Program Files / Program Files (x86) / LocalAppData.
- **Appx package** lookup for the new Outlook for Windows.

Office detection covers modern Click-to-Run installs (Office 2016/2019/2021/365, the `Office16` folder). Older MSI Office (2013 and earlier) is not detected.

---

## Prerequisites

- **Windows 10 / 11**
- **PowerShell 5.1+** (enforced via `#Requires`)
- **Administrator rights** (enforced via `#Requires -RunAsAdministrator`)
- **WinGet / Microsoft App Installer** — required only for the Chrome and Adobe installs. If both are already present, WinGet isn't needed.
- **Internet access** to the Office CDN and the WinGet source.

---

## Usage

### Interactive (elevated PowerShell)

```powershell
# From an elevated PowerShell prompt:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\Install-StandardApps.ps1
```

### Via RMM / SYSTEM context

The script is designed to run as `SYSTEM` (e.g. pushed from NinjaOne, Datto RMM, ConnectWise Automate, Intune, etc.). It resolves the WinGet executable directly from the `WindowsApps` folder, so it works even when `winget` isn't on the SYSTEM account's PATH.

Just deploy and run the `.ps1` as a script through your tool of choice.

### Exit codes

| Code | Meaning                              |
|------|--------------------------------------|
| `0`  | Completed successfully               |
| `1`  | An error occurred (see the log file) |

---

## Configuration

Edit the variable near the top of the script to match your licensing:

| Variable           | Default              | Notes                                                            |
|--------------------|----------------------|------------------------------------------------------------------|
| `$OfficeProductId` | `O365BusinessRetail` | Use `O365ProPlusRetail` for *Microsoft 365 Apps for enterprise*. |

The Office install uses a generated ODT `configuration.xml` with: 64-bit edition, Current channel, `en-us` language, no UI (`Display Level="None"`), EULA auto-accepted, and updates enabled. Adjust the here-string in `Install-Microsoft365Apps` if you need a different edition, channel, or language.

---

## Logging

A transcript is written to:

```
%TEMP%\Standard-App-Install-<yyyyMMdd-HHmmss>.log
```

The log path is printed to the console at the end of a successful run and on error.

---

## WinGet package IDs

| Application          | Package ID                    |
|----------------------|-------------------------------|
| Google Chrome        | `Google.Chrome`               |
| Adobe Acrobat Reader | `Adobe.Acrobat.Reader.64-bit` |

Vendors occasionally change package IDs. If an install fails to find the package, confirm the current ID with:

```powershell
winget search Adobe.Acrobat.Reader
winget search Google.Chrome
```

---

## Safety / hardening notes

- The Office Deployment Tool bootstrapper (`setup.exe`) is downloaded from the Microsoft CDN and its **Authenticode signature is verified** (valid and signed by Microsoft Corporation) before it's executed.
- The Office install is capped at a **45-minute timeout**; the process is terminated and the run fails if it exceeds that.
- WinGet runs with `--silent --accept-package-agreements --accept-source-agreements --scope machine`.

---

## Limitations

- Detects only modern Click-to-Run Office (`Office16`), not Office 2013/MSI.
- Chrome's per-user fallback path checks the **running account's** `LOCALAPPDATA`, which is correct for machine provisioning but won't see a Chrome install under a different user profile.
- Relies on WinGet for Chrome/Adobe; if App Installer is broken or absent, those installs won't proceed.

---

## Disclaimer

Provided as-is, with no warranty. Test in a non-production environment before broad deployment, and confirm the application baseline and licensing (`$OfficeProductId`) match your organization.

## License

MIT (or your preference).
