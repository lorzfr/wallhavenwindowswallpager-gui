# Wallpaper Puller

This is a Windows desktop app built with PowerShell and Windows Forms. It pulls a random wallpaper from the configured Wallhaven search, previews it, caches it locally, and applies it as your desktop wallpaper.

## Run it

Double-click `Launch Wallpaper Puller.cmd`, or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WallpaperPuller.ps1
```

## Configure it

Edit `config.json` to change the Wallhaven source URL or wallpaper style.

Supported `WallpaperStyle` values:

- `Fill`
- `Fit`
- `Stretch`
- `Center`
- `Span`
- `Tile`

## Notes

- Downloaded wallpapers are cached in `%LOCALAPPDATA%\WallpaperPuller\Cache`.
- The app uses the Windows wallpaper API plus a desktop refresh fallback.
- On some non-activated Windows installs, Microsoft may still restrict wallpaper changes. The app attempts the native route first and reports an error if Windows rejects it.

## Smoke test

To verify that the script loads without opening the GUI:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WallpaperPuller.ps1 -SmokeTest
```
