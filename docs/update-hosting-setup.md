# KKY Tool Update Hosting Setup

## Goal

Host two files on a normal HTTPS web hosting account:

- `latest.json`
- installer file such as `KKY_Tool_Revit(2019,21,23,25)_v2.03.exe`

The add-in checks `latest.json`, compares versions, then downloads the installer.

## Recommended Structure

Use a normal domain or subdomain over HTTPS.

Example URLs:

- `https://update.example.com/kky/latest.json`
- `https://update.example.com/kky/KKY_Tool_Revit(2019,21,23,25)_v2.03.exe`

Server folder structure:

```text
/public_html/
  /kky/
    latest.json
    KKY_Tool_Revit(2019,21,23,25)_v2.03.exe
```

If you want a backup mirror:

```text
Primary:
https://update.example.com/kky/latest.json

Backup:
https://www.example.com/kky-update/latest.json
```

## latest.json Example

```json
{
  "version": "2.03",
  "url": "https://update.example.com/kky/KKY_Tool_Revit(2019,21,23,25)_v2.03.exe",
  "publishedAt": "2026-03-20",
  "notes": "Hub update check and installer queue support"
}
```

Notes:

- `version` must match the version you want users to see.
- `url` can be absolute HTTPS or a relative path like `KKY_Tool_Revit(2019,21,23,25)_v2.03.exe`.
- If you use a relative path, it is resolved from the location of `latest.json`.

## Add-in Config

Update this file:

- `KKY_Tool_Revit_2019-2023/Resources/update-config.json`

Example:

```json
{
  "feedUrls": [
    "https://update.example.com/kky/latest.json",
    "https://www.example.com/kky-update/latest.json"
  ],
  "downloadDirectory": "%LOCALAPPDATA%\\KKY_Tool_Revit\\Updates"
}
```

Behavior:

- The add-in tries each URL in `feedUrls` in order.
- The first reachable feed is used.
- `feedUrl` still works for backward compatibility, but `feedUrls` is recommended.

## Release Steps

1. Build the installer `.exe`.
2. Upload the new installer file to the hosting folder.
3. Update `latest.json` with the new version and file URL.
4. Replace the old `latest.json` on the server.
5. Open the add-in and click `업데이트 확인`.

## Traffic Planning

For a 200 MB installer:

- 50 downloads is about 10 GB
- 100 downloads is about 20 GB
- 500 downloads is about 100 GB

Check monthly traffic limits before choosing a hosting plan.

## Practical Notes

- HTTPS on port 443 is strongly recommended.
- A normal web hosting site is usually less likely to be blocked than consumer cloud storage links.
- If some companies block one domain, add a second mirror URL in `feedUrls`.
- If `.exe` downloads are blocked by security software, consider also preparing a zipped delivery path later.
