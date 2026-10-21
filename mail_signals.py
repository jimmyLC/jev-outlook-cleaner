"""Deterministic local evidence about a message, used beside Jev's classification.

Nothing here contacts a network. Signals are short labels (never message text),
so they are safe to store in the private review file and to show in the UI.
"""

from __future__ import annotations

import re
from typing import Any

# Evidence that mail was mass-distributed or machine-generated.
_UNSUBSCRIBE = re.compile(
    r"unsubscribe|opt[- ]out|manage (?:your )?(?:email )?(?:preferences|subscription)|"
    r"view (?:this )?(?:email )?in (?:your )?browser|取消訂閱|退訂|不想再收到|電子報|訂閱設定",
    re.IGNORECASE,
)
_BULK_SENDER = re.compile(
    r"^(?:newsletter|news|promo|promotions?|marketing|edm|deals?|offers?|campaign|mailer|digest|updates?)\b",
    re.IGNORECASE,
)
_BULK_PRECEDENCE = {"bulk", "list", "junk"}

# Evidence that mail may matter later. Matches keep the message out of Jev.
_PROTECT_PHRASES: tuple[tuple[re.Pattern[str], str], ...] = tuple(
    (re.compile(pattern, re.IGNORECASE), label)
    for pattern, label in (
        (r"verification code|one[- ]time (?:pass|code)|\botp\b|驗證碼|認證碼|動態密碼", "驗證碼"),
        (r"password|passcode|密碼|reset your|重設|重置", "密碼／帳號存取"),
        (r"security alert|suspicious|unusual (?:sign|activity)|安全性?警示|異常登入|登入通知|新裝置登入", "安全通知"),
        (r"\breceipt\b|\binvoice\b|收據|發票|電子發票|請款", "收據／發票"),
        (r"order (?:confirmation|confirmed|#)|訂單(?:確認|成立)|付款(?:成功|完成)|payment (?:received|confirmation)", "訂單／付款"),
        (r"booking confirm|reservation|itinerary|boarding pass|訂位|訂房|電子機票|登機證", "預訂／行程"),
        (r"friend request|shared (?:a|an|with)|invited you|好友邀請|與您共用|邀請您", "邀請／共用"),
    )
)
_REPLY_PREFIX = re.compile(r"^\s*(?:re|fw|fwd|回覆|轉寄)\s*[:：]", re.IGNORECASE)
_HEAD_CHARS = 400


def header_map(message: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    for header in message.get("internetMessageHeaders") or []:
        name = str(header.get("name") or "").strip().casefold()
        if name and name not in result:
            result[name] = str(header.get("value") or "").strip()
    return result


def bulk_signals(message: dict[str, Any]) -> list[str]:
    signals: list[str] = []
    headers = header_map(message)
    if "list-unsubscribe" in headers:
        signals.append("List-Unsubscribe 標頭")
    if "list-id" in headers:
        signals.append("List-Id 標頭")
    if headers.get("precedence", "").casefold() in _BULK_PRECEDENCE:
        signals.append("Precedence: bulk")
    if headers.get("auto-submitted", "no").casefold() not in {"", "no"}:
        signals.append("自動產生郵件")
    body = ((message.get("body") or {}).get("content")) or ""
    if _UNSUBSCRIBE.search(body):
        signals.append("內文含取消訂閱")
    address = ((message.get("sender") or {}).get("emailAddress") or {}).get("address", "")
    if _BULK_SENDER.match(address.partition("@")[0]):
        signals.append("寄件者為電子報／行銷帳號")
    return signals


def protect_signals(message: dict[str, Any]) -> list[str]:
    signals: list[str] = []
    subject = message.get("subject") or ""
    if _REPLY_PREFIX.match(subject):
        signals.append("回覆／轉寄串")
    body = ((message.get("body") or {}).get("content")) or ""
    haystack = f"{subject}\n{body[:_HEAD_CHARS]}"
    for pattern, label in _PROTECT_PHRASES:
        if pattern.search(haystack):
            signals.append(label)
    return signals


def collect(message: dict[str, Any]) -> dict[str, list[str]]:
    return {"bulk": bulk_signals(message), "protect": protect_signals(message)}
