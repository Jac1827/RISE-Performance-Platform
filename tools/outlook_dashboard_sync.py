#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shlex
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
TOKEN_ROOT = "https://login.microsoftonline.com"
DEFAULT_LOOKBACK = 10
DEFAULT_PAGE_SIZE = 25


@dataclass
class AttachmentMatch:
    rule_name: str
    message_id: str
    message_subject: str
    received_at: str
    attachment_name: str
    target_path: Path
    saved: bool
    skipped_reason: str | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Pull report attachments from Outlook / Microsoft 365 and run local dashboard update commands."
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("tools/outlook_dashboard_sync.config.json"),
        help="Path to the JSON config file.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List matching emails and commands without saving files or running updates.",
    )
    parser.add_argument(
        "--lookback-days",
        type=int,
        default=DEFAULT_LOOKBACK,
        help="Only inspect messages received within this many days.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        help="Maximum number of recent messages to scan from the folder.",
    )
    return parser.parse_args()


def load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise SystemExit(f"Config not found: {path}")
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def env_required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def ensure_parent(path: Path, *, dry_run: bool) -> None:
    if dry_run:
        return
    path.parent.mkdir(parents=True, exist_ok=True)


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def parse_graph_datetime(value: str) -> datetime:
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    return datetime.fromisoformat(value)


def iso_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def compile_optional_regex(pattern: str | None) -> re.Pattern[str] | None:
    if not pattern:
        return None
    return re.compile(pattern, re.IGNORECASE)


def request_json(url: str, *, method: str = "GET", headers: dict[str, str] | None = None, data: bytes | None = None) -> Any:
    request = urllib.request.Request(url, method=method, headers=headers or {}, data=data)
    try:
        with urllib.request.urlopen(request) as response:
            payload = response.read()
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise SystemExit(f"HTTP {exc.code} calling {url}: {body}") from exc
    except urllib.error.URLError as exc:
        raise SystemExit(f"Request failed for {url}: {exc}") from exc
    if not payload:
        return {}
    return json.loads(payload.decode("utf-8"))


def get_access_token(config: dict[str, Any]) -> str:
    tenant_env = config.get("tenant_id_env", "MS365_TENANT_ID")
    client_env = config.get("client_id_env", "MS365_CLIENT_ID")
    secret_env = config.get("client_secret_env", "MS365_CLIENT_SECRET")

    payload = urllib.parse.urlencode(
        {
            "client_id": env_required(client_env),
            "client_secret": env_required(secret_env),
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
    ).encode("utf-8")
    token_url = f"{TOKEN_ROOT}/{env_required(tenant_env)}/oauth2/v2.0/token"
    response = request_json(
        token_url,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data=payload,
    )
    token = str(response.get("access_token", "")).strip()
    if not token:
        raise SystemExit("Microsoft token response did not include an access token.")
    return token


def list_messages(config: dict[str, Any], token: str, *, lookback_days: int, limit: int) -> list[dict[str, Any]]:
    mailbox = str(config.get("mailbox_user", "")).strip()
    if not mailbox:
        raise SystemExit("Config is missing 'mailbox_user'.")
    folder = str(config.get("mail_folder", "Inbox")).strip() or "Inbox"
    cutoff = utc_now().timestamp() - max(0, lookback_days) * 86400
    cutoff_dt = datetime.fromtimestamp(cutoff, tz=timezone.utc)
    cutoff_filter = urllib.parse.quote(f"receivedDateTime ge {iso_utc(cutoff_dt)}", safe=" ='():")
    select = urllib.parse.quote("id,subject,receivedDateTime,from,hasAttachments,isRead", safe=",")
    orderby = urllib.parse.quote("receivedDateTime desc", safe=" ")
    url = (
        f"{GRAPH_ROOT}/users/{urllib.parse.quote(mailbox)}/mailFolders/{urllib.parse.quote(folder)}/messages"
        f"?$select={select}&$filter={cutoff_filter}&$orderby={orderby}&$top={max(1, limit)}"
    )
    response = request_json(url, headers={"Authorization": f"Bearer {token}"})
    return list(response.get("value", []))


def list_attachments(mailbox: str, token: str, message_id: str) -> list[dict[str, Any]]:
    select = urllib.parse.quote("id,name,contentType,size,lastModifiedDateTime,@odata.type,contentBytes", safe=",@.")
    url = (
        f"{GRAPH_ROOT}/users/{urllib.parse.quote(mailbox)}/messages/{urllib.parse.quote(message_id)}/attachments"
        f"?$select={select}"
    )
    response = request_json(url, headers={"Authorization": f"Bearer {token}"})
    return list(response.get("value", []))


def sender_address(message: dict[str, Any]) -> str:
    return (
        message.get("from", {})
        .get("emailAddress", {})
        .get("address", "")
        .strip()
        .lower()
    )


def attachment_content(attachment: dict[str, Any]) -> bytes:
    content = attachment.get("contentBytes")
    if not content:
        return b""
    return base64.b64decode(content)


def normalize_command(command: Any, workspace: Path) -> list[str]:
    if isinstance(command, list) and all(isinstance(part, str) for part in command):
        return [part.replace("{workspace}", str(workspace)) for part in command]
    if isinstance(command, str):
        return [part.replace("{workspace}", str(workspace)) for part in shlex.split(command)]
    raise SystemExit(f"Invalid command entry: {command!r}")


def matches_rule(rule: dict[str, Any], message: dict[str, Any], attachment: dict[str, Any]) -> bool:
    subject = str(message.get("subject", ""))
    sender = sender_address(message)
    attachment_name = str(attachment.get("name", ""))

    sender_contains = str(rule.get("sender_contains", "")).strip().lower()
    subject_contains = str(rule.get("subject_contains", "")).strip().lower()
    attachment_regex = compile_optional_regex(rule.get("attachment_name_regex"))
    subject_regex = compile_optional_regex(rule.get("subject_regex"))

    if sender_contains and sender_contains not in sender:
        return False
    if subject_contains and subject_contains not in subject.lower():
        return False
    if attachment_regex and not attachment_regex.search(attachment_name):
        return False
    if subject_regex and not subject_regex.search(subject):
        return False
    return True


def target_path_for_rule(rule: dict[str, Any], attachment_name: str, workspace: Path) -> Path:
    raw_target = str(rule.get("target_path", "")).strip()
    if raw_target:
        return (workspace / raw_target).resolve()
    fallback_dir = str(rule.get("download_dir", "tmp/email-reports")).strip()
    return (workspace / fallback_dir / attachment_name).resolve()


def save_attachment(target_path: Path, content: bytes, *, dry_run: bool, overwrite: bool) -> tuple[bool, str | None]:
    if target_path.exists() and not overwrite:
        return False, "target exists and overwrite is disabled"
    if dry_run:
        return True, None
    ensure_parent(target_path, dry_run=False)
    target_path.write_bytes(content)
    return True, None


def collect_matches(
    config: dict[str, Any],
    token: str,
    *,
    lookback_days: int,
    limit: int,
    dry_run: bool,
    workspace: Path,
) -> list[AttachmentMatch]:
    mailbox = str(config.get("mailbox_user", "")).strip()
    messages = list_messages(config, token, lookback_days=lookback_days, limit=limit)
    rules = config.get("rules", [])
    if not isinstance(rules, list) or not rules:
        raise SystemExit("Config must include a non-empty 'rules' array.")

    matches: list[AttachmentMatch] = []
    matched_rule_names: set[str] = set()
    for message in messages:
        if not message.get("hasAttachments"):
            continue
        attachments = list_attachments(mailbox, token, str(message.get("id")))
        for attachment in attachments:
            if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
                continue
            for rule in rules:
                rule_name = str(rule.get("name", "unnamed-rule")).strip() or "unnamed-rule"
                if bool(rule.get("latest_only")) and rule_name in matched_rule_names:
                    continue
                if not matches_rule(rule, message, attachment):
                    continue
                target_path = target_path_for_rule(rule, str(attachment.get("name", "")), workspace)
                saved, skipped_reason = save_attachment(
                    target_path,
                    attachment_content(attachment),
                    dry_run=dry_run,
                    overwrite=bool(rule.get("overwrite", True)),
                )
                matches.append(
                    AttachmentMatch(
                        rule_name=rule_name,
                        message_id=str(message.get("id", "")),
                        message_subject=str(message.get("subject", "")),
                        received_at=str(message.get("receivedDateTime", "")),
                        attachment_name=str(attachment.get("name", "")),
                        target_path=target_path,
                        saved=saved,
                        skipped_reason=skipped_reason,
                    )
                )
                matched_rule_names.add(rule_name)
                break
    return matches


def run_commands(config: dict[str, Any], *, dry_run: bool, workspace: Path) -> list[dict[str, Any]]:
    commands = config.get("post_update_commands", [])
    if not isinstance(commands, list):
        raise SystemExit("'post_update_commands' must be an array when present.")
    results: list[dict[str, Any]] = []
    for command in commands:
        argv = normalize_command(command, workspace)
        if dry_run:
            results.append({"command": argv, "returncode": 0, "stdout": "", "stderr": "", "skipped": True})
            continue
        completed = subprocess.run(
            argv,
            cwd=workspace,
            text=True,
            capture_output=True,
            check=False,
        )
        results.append(
            {
                "command": argv,
                "returncode": completed.returncode,
                "stdout": completed.stdout.strip(),
                "stderr": completed.stderr.strip(),
                "skipped": False,
            }
        )
        if completed.returncode != 0:
            break
    return results


def print_summary(matches: list[AttachmentMatch], command_results: list[dict[str, Any]], *, dry_run: bool) -> None:
    mode_label = "DRY RUN" if dry_run else "LIVE RUN"
    print(f"[{mode_label}] Outlook dashboard sync")
    print("")
    if not matches:
        print("No matching report attachments found.")
    else:
        print("Attachment matches:")
        for match in matches:
            suffix = ""
            if match.skipped_reason:
                suffix = f" (skipped: {match.skipped_reason})"
            print(
                f"- {match.rule_name}: {match.attachment_name} from \"{match.message_subject}\" "
                f"at {match.received_at} -> {match.target_path}{suffix}"
            )
    print("")
    if command_results:
        print("Post-update commands:")
        for result in command_results:
            status = "skipped" if result.get("skipped") else f"exit {result['returncode']}"
            print(f"- {' '.join(shlex.quote(part) for part in result['command'])} [{status}]")
            if result.get("stderr"):
                print(f"  stderr: {result['stderr']}")
    else:
        print("No post-update commands configured.")


def main() -> int:
    args = parse_args()
    workspace = Path.cwd().resolve()
    config = load_json(args.config.resolve())
    token = get_access_token(config)
    matches = collect_matches(
        config,
        token,
        lookback_days=args.lookback_days,
        limit=args.limit,
        dry_run=args.dry_run,
        workspace=workspace,
    )
    command_results = run_commands(config, dry_run=args.dry_run, workspace=workspace)
    print_summary(matches, command_results, dry_run=args.dry_run)

    if any(result.get("returncode", 0) != 0 for result in command_results if not result.get("skipped")):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
