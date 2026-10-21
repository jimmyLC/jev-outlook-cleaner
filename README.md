# Jev Outlook Cleaner

A local Web app for cautiously triaging older Outlook Inbox messages with TypeSafe Jev. The source can be public; credentials, mailbox content, review files, and personal rules stay on your computer. **Do not deploy this Web server to the public Internet.** It binds to `127.0.0.1` only.

## What it does

- Scans Inbox messages older than a chosen number of days. The Web interface caps each batch at 200 new messages.
- Skips messages matching your private `keep-list.json` and previously processed SHA-256 fingerprints of Internet Message IDs. It pages further to fill the batch.
- Classifies eligible messages as keep, review, or delete candidate. Attachments, high importance, and follow-up flags are protected.
- Automatically moves only explicit saved-rule matches or messages that Jev classifies as bulk/automated with confidence at least 0.90 **and** as delete candidates with confidence at least 0.85.
- Offers a manual review page where checked messages are kept and remembered; unchecked messages can be moved for that batch after you upload its downloaded selection file.
- Rechecks the Outlook account and each message's Inbox location, ID, subject, sender, received time, Internet Message ID, and content snapshot immediately before moving. Moves use **Deleted Items**, never permanent deletion.

The review page displays subject, sender, and received time. Jev receives the fields and up to 9,000 characters of plain-text body needed for classification. Review JSON and local run logs contain mailbox metadata and must stay private.

## Setup

Requires Python 3.10+, a TypeSafe API key, and a Microsoft Entra public client application. Configure the app for device-code flow, personal or organizational Microsoft accounts, and delegated Microsoft Graph permissions `Mail.ReadWrite` and `User.Read`. The app does not need a client secret or mail-send permission.

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
cp keep-list.example.json keep-list.json
cp delete-rules.example.json delete-rules.json
chmod 600 .env keep-list.json delete-rules.json
```

Edit `.env` locally to set `TYPESAFE_API_KEY` and `OUTLOOK_CLIENT_ID`. Never commit or paste the values. A system environment variable takes priority over `.env`. On first use, complete the Microsoft device sign-in shown in your terminal. The MSAL token cache is stored under `~/.jev-outlook-cleaner/` with owner-only permissions.

Check local configuration without displaying its values or connecting:

```sh
.venv/bin/python jev_outlook_cleaner.py check
```

## Web interface

```sh
.venv/bin/python web_app.py
```

Open `http://127.0.0.1:8765/` on the same computer. You can scan, inspect the generated preview, run the conservative auto-apply path, or download a manual selection from the preview and upload it on the dashboard. Scan and apply run in the background; use **Update preview list** after a scan completes. Only one operation runs at a time. The dashboard does not expose email bodies, API keys, tokens, raw review JSON, or run logs.

The Web process must remain running while you use it. Stop it with `Ctrl+C`. It does not alter any existing scheduler or run a scan on startup.

## Command line

```sh
.venv/bin/python jev_outlook_cleaner.py scan --older-than-days 30 --limit 200
.venv/bin/python jev_outlook_cleaner.py auto-apply reviews/review-<timestamp>.json
.venv/bin/python jev_outlook_cleaner.py apply-selection reviews/review-<timestamp>.json /path/to/outlook-cleanup-selection.json
```

The CLI also has an interactive `apply` command. Run `--help` for all options.

## Private data and safety

Never add `.env`, personal rule files, `reviews/`, `.processed-message-hashes.json`, token caches, logs, mailbox exports, backups, or credentials to Git. The supplied `.gitignore` protects common paths, but inspect `git diff --cached --name-only` and staged content before every push. Do not use `git add -f` for private files. The provided example rule files have empty arrays; add your own rules only to the ignored live JSON files.

Classification can be wrong. Keep and review decisions remain in Inbox. A one-time manual deletion choice does not create a future auto-delete rule. The local review and run files do not expire automatically; delete old ones from your own computer when no longer needed. Avoid placing this directory in a public sync folder.

## Tests

```sh
.venv/bin/python -m unittest -q test_cleaner test_web_app
```

Tests use synthetic mail and mock network actions; they do not connect to Outlook or Jev.

## License

MIT; see [LICENSE](LICENSE).
