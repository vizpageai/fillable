# Microsoft Store Publishing (Win32 EXE path)

This project now supports the Microsoft Store "unmodified installer" path for Win32 apps.

## What this repo now includes
- Inno Setup installer script: `packaging/installer/Fillable.iss`
- Build helper: `scripts/build_store_installer.ps1`
- Automatic right-click integration during install/uninstall:
  - Installer runs: `FillableDOC.exe --install-context-menu`
  - Uninstaller runs: `FillableDOC.exe --uninstall-context-menu`
- First-launch safety net: app also installs context menu once on first run.

## Build artifacts
1. Build EXE and installer:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_store_installer.ps1
```
2. Output:
- `dist/FillableDOC.exe`
- `dist/installer/FillableDOC-Setup.exe`

Optional MSI path (WiX):
```powershell
winget install WiXToolset.WiXToolset
powershell -ExecutionPolicy Bypass -File scripts/build_msi.ps1
```
Output:
- `dist/installer/FillableDOC.msi`

## Submit to Store (Partner Center)
1. Reserve app name in Partner Center.
2. Choose Win32 EXE/MSI submission path (unmodified installer).
3. Upload/host installer URL for `FillableDOC-Setup.exe` (or `FillableDOC.msi`) and provide metadata.
4. Publish submission.

## Important requirements
- Installer must be `.exe` or `.msi`.
- Installer should be offline.
- Hosted binary should remain unchanged after submission.
- Installer should only install intended product.

## AI billing and key handling (Store-safe pattern)
- Do not ship your OpenAI API key inside the EXE/MSI.
- Use `app_subscription` mode:
  - Client sends prompts to your backend endpoint (`/v1/openai-proxy`).
  - Backend validates subscription entitlement and rate limits.
  - Backend calls OpenAI with your server-side API key.
- `user_key` mode is supported for bring-your-own-key users; user key is secured with Windows DPAPI.
- For Store compliance, implement entitlement checks using Microsoft Store add-ons/subscriptions on your backend and deny proxy requests when entitlement is inactive.
- Reference policy (as of March 9, 2026): Microsoft Store Policy 7.18 (effective October 19, 2024), section `10.8.6` permits non-game PC products to use Microsoft recurring billing API or a secure third-party billing API for subscription digital goods.

## References
- Microsoft Store Policy 7.18 (effective October 19, 2024)  
  https://learn.microsoft.com/en-us/windows/apps/publish/store-policy-archive/store-policy-7-18
- Microsoft Learn: "How to distribute your Win32 application through Microsoft Store"  
  https://learn.microsoft.com/en-us/windows/apps/distribute-through-store/how-to-distribute-your-win32-app-through-microsoft-store
- Microsoft Learn: "Publish update to your MSI/EXE app on the Store"  
  https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msi/publish-update-to-your-app-on-store
