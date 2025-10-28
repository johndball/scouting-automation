# docs/01-overview-setup.md

## Playbook Overview & Setup

This repository hosts sanitized, reusable Google Apps Script projects for Scouting units.

### Folder Layout (Google Drive)
```
/Troop Automation (root)
  /All Yall Weekly
  /RSVP & Transport
  /Leader Portal
  /Scout Portal
  /Shared Libs
  /Docs Output (PDF)
  Config (Google Sheet)
```

### Config Sheet (two columns)
See `config-template/Config.csv`. Create a Google Sheet named **Config** and paste the key/value pairs.

### Getting Started
1. Copy `shared-libs/*.gs` into your Apps Script projects.
2. Run **Install/Repair Triggers** from the custom menu.
3. Use `tools/sanitizer/` before publishing any forks.

### Safety Notes
- Replace real emails, calendar IDs, web app URLs, and Drive IDs with examples.
- Blur/replace QR codes and timestamps in screenshots.

