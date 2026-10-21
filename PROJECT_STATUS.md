# Project status and agent handoff

Updated: 2026-10-21 (Asia/Singapore). Public repository: `https://github.com/jimmyLC/jev-outlook-cleaner`. This checkout is the **publishable source**. The owner keeps a separate private runtime directory with credentials, rules, review files, and processed history.

## Delivered

- A CLI cleaner (`jev_outlook_cleaner.py`) with configuration check, terminal Microsoft login, Inbox scan, three-way review, manual selection, and conservative auto-apply.
- A loopback-only Web interface (`web_app.py`) for scan jobs, preview navigation, auto-apply, and confirmed selection upload. Browser writes use Host, Origin, and CSRF checks; jobs call the existing CLI. Microsoft device-code login is started only by an explicit dashboard button (it wraps the CLI `login`); scan and apply jobs never start it.
- Empty example rule files, `.gitignore`, MIT license, setup documentation, and security guidance.

## Verification at handoff

- Last tested source commit before these handoff documents: `0c54ef05881b008651c31157da9c6cbcc7f478ff`.
- Python compilation and 23 offline tests passed in both source and private runtime copies. The local Web dashboard returned HTTP 200 and rendered in a browser on `127.0.0.1`; no scan was triggered during that check.
- The public file list was explicitly allowlisted. A local privacy scan found zero warnings and no match for live environment values or personal rule entries in tracked text. Repeat this check for every future publication.
- The Web scan, manual selection upload, and auto-apply have **not** been validated end to end against a real mailbox. No cloud or public Web deployment was performed.

## Data boundary

Graph supplies Inbox mail to the local process. A live scan sends the classification fields and up to 9,000 plain-text body characters to TypeSafe Jev. The process stores private review JSON/HTML, logs, and SHA-256 fingerprints locally. Microsoft tokens and TypeSafe credentials stay outside this Git repository. Only mail selected under the saved-rule or confidence gates can be moved, and only to recoverable Deleted Items after fresh validation.

## Next steps

1. Read `AGENTS.md`, `README.md`, and `SECURITY.md`; inspect the current diff and CLI help before making changes.
2. Improve and test the Web experience with synthetic data first. A real Inbox/Jev test requires the owner's current authorization and must be reported separately from offline checks.
3. Keep the Web server on `127.0.0.1`. This app has no multi-user authentication and is not ready for public Internet hosting.
4. Before a future push, inspect every staged filename and blob, run the secret scan, and request the owner's approval for the reviewed `main` diff.

Agent kickoff: “Use `PROJECT_STATUS.md` and `AGENTS.md` as the handoff. Work in this publishable source checkout; keep live mailbox data in the owner's private runtime workspace. Preserve the Outlook/Jev safety checks, run offline tests, and report the exact diff and remaining live-validation gap before proposing a public push.”
