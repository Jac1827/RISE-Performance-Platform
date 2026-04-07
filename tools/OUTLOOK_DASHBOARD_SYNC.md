# Outlook Dashboard Sync

This workspace includes [`outlook_dashboard_sync.py`](/Users/jacheflin/Documents/Playground/tools/outlook_dashboard_sync.py), a Microsoft 365 attachment puller for dashboard update automations.

## What it does

1. Authenticates to Microsoft Graph with application credentials.
2. Scans a mailbox folder for recent messages with attachments.
3. Matches attachments against configured sender / subject / filename rules.
4. Saves matched files into the workspace.
5. Runs local post-update commands to repair or refresh dashboard inputs.

## Setup

1. Copy [`outlook_dashboard_sync.config.example.json`](/Users/jacheflin/Documents/Playground/tools/outlook_dashboard_sync.config.example.json) to `tools/outlook_dashboard_sync.config.json`.
2. Replace `mailbox_user` with the mailbox you want the automation to scan.
3. Update the `rules` array so each incoming report lands in the correct local file path.
4. Add any follow-up commands needed to rebuild dashboard sources.

## Required environment variables

- `MS365_TENANT_ID`
- `MS365_CLIENT_ID`
- `MS365_CLIENT_SECRET`

The Azure app registration needs Microsoft Graph application permissions for mail read access on the target mailbox, typically `Mail.Read`.

## Dry run

```bash
python3 tools/outlook_dashboard_sync.py --config tools/outlook_dashboard_sync.config.json --dry-run
```

## Automation prompt

Use a prompt like:

Pull the latest Microsoft 365 dashboard report emails, save matching attachments into the workspace using `tools/outlook_dashboard_sync.py`, run any configured post-update commands, and open an inbox item summarizing what changed or failed.
