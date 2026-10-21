# Security and privacy

This project handles private mailbox data. Keep the Web server on its default `127.0.0.1` binding; do not expose it through a tunnel, reverse proxy, container port mapping, or public hosting. The interface has local browser origin and CSRF checks, but it is not a multi-user authenticated service.

Keep `.env`, live rules, review JSON/HTML, logs, processed fingerprints, and `~/.jev-outlook-cleaner/` private. A public GitHub repository should contain only source, tests, docs, and empty example configuration. If a secret is accidentally committed, revoke and rotate it; removing the file from a later commit does not remove it from Git history.

Report security issues privately to the repository owner through GitHub's private vulnerability reporting feature if available. Do not post credentials or private email content in a public issue.
