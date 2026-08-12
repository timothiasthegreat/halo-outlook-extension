#!/usr/bin/env python3
"""Pre-flight validation for the Halo Outlook Extension.

Verifies config, Halo connectivity, Graph connectivity, and action IDs
before starting the watcher. Safe to run multiple times — no mutations.

The script uses raw httpx calls (not the watcher package) so it works
regardless of whether the watcher is pip-installed or run from source.

Usage:
    python scripts/setup_check.py
    python scripts/setup_check.py --config my-config.yaml
    python scripts/setup_check.py --discover-actions
    python scripts/setup_check.py --list-custom-fields
"""

from __future__ import annotations

import argparse
import asyncio
import sys
import time
from pathlib import Path

# Add watcher dir to path so we can import config.py directly
_WATCHER = Path(__file__).resolve().parent.parent / "watcher"
sys.path.insert(0, str(_WATCHER))

import httpx
from config import load_config  # type: ignore[import-untyped]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def ok(msg: str) -> None:
    print(f"  \033[32m\u2713\033[0m {msg}")


def fail(msg: str) -> None:
    print(f"  \033[31m\u2717\033[0m {msg}")


def info(msg: str) -> None:
    print(f"  \033[36m?info?\033[0m {msg}")


async def _get_halo_token(config) -> str | None:
    """Acquire a Halo OAuth2 token via client credentials."""
    token_url = f"{config.halo.instance_url.rstrip('/')}/auth/token"
    async with httpx.AsyncClient(timeout=15.0) as c:
        try:
            r = await c.post(
                token_url,
                data={
                    "grant_type": "client_credentials",
                    "client_id": config.halo.client_id,
                    "client_secret": config.halo.client_secret,
                    "scope": "all",
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )
            if r.status_code == 200:
                return r.json()["access_token"]
        except Exception:
            pass
    return None


async def _get_graph_token(config) -> str | None:
    """Acquire a Graph OAuth2 token via client credentials."""
    token_url = (
        f"https://login.microsoftonline.com/{config.graph.tenant_id}"
        "/oauth2/v2.0/token"
    )
    async with httpx.AsyncClient(timeout=15.0) as c:
        try:
            r = await c.post(
                token_url,
                data={
                    "grant_type": "client_credentials",
                    "client_id": config.graph.client_id,
                    "client_secret": config.graph.client_secret,
                    "scope": "https://graph.microsoft.com/.default",
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )
            if r.status_code == 200:
                return r.json()["access_token"]
        except Exception:
            pass
    return None


async def check_halo(config) -> bool:
    """Verify HaloPSA connectivity and auth."""
    print("\n── HaloPSA ──")
    all_ok = True

    # Reachability
    try:
        async with httpx.AsyncClient(timeout=10.0) as c:
            r = await c.get(config.halo.instance_url)
            if r.status_code < 500:
                ok(f"HaloPSA reachable \u2192 {config.halo.instance_url}")
            else:
                fail(f"HaloPSA returned {r.status_code}")
                all_ok = False
    except Exception as e:
        fail(f"HaloPSA unreachable: {e}")
        return False

    # Token acquisition
    token = await _get_halo_token(config)
    if token:
        ok("OAuth2 token acquired (scope: all)")
    else:
        fail("OAuth2 failed \u2014 check client_id and client_secret")
        all_ok = False

    return all_ok


async def check_graph(config) -> bool:
    """Verify Microsoft Graph connectivity."""
    print("\n── Microsoft Graph ──")
    all_ok = True

    token = await _get_graph_token(config)
    if not token:
        fail("Graph token acquisition failed")
        info("Check tenant_id, client_id, and client_secret")
        return False

    ok("Token acquired (scope: https://graph.microsoft.com/.default)")

    # Test mailbox access
    try:
        async with httpx.AsyncClient(base_url=GRAPH_BASE, timeout=15.0) as c:
            r = await c.get(
                f"/users/{config.graph.user_email}",
                headers={"Authorization": f"Bearer {token}"},
            )
            if r.status_code == 200:
                ok(f"Mailbox accessible: {config.graph.user_email}")
            elif r.status_code == 404:
                fail(f"Mailbox not found: {config.graph.user_email}")
                info(
                    "Verify the email matches a licensed Exchange Online mailbox"
                )
                all_ok = False
            elif r.status_code == 403:
                fail("Permission denied \u2014 admin consent may be required")
                info("Go to Azure Portal \u2192 App Registration \u2192 API Permissions \u2192 Grant admin consent")
                all_ok = False
            else:
                fail(f"Unexpected status {r.status_code}")
                all_ok = False
    except Exception as e:
        fail(f"Mailbox check failed: {e}")
        all_ok = False

    return all_ok


async def discover_actions(config):
    """Auto-discover ticket action IDs from Halo."""
    print("\n── Ticket Action Discovery ──")

    token = await _get_halo_token(config)
    if not token:
        fail("Authentication required for action discovery")
        return

    try:
        api_url = f"{config.halo.instance_url.rstrip('/')}/api"
        async with httpx.AsyncClient(base_url=api_url, timeout=15.0) as c:
            r = await c.get(
                "/Actions",
                headers={"Authorization": f"Bearer {token}"},
            )
            if r.status_code != 200:
                fail(f"Could not fetch actions list (status {r.status_code})")
                return

            data = r.json()
            actions = data if isinstance(data, list) else data.get("actions", [])

            email_in = None
            email_out = None
            internal = None

            for a in actions:
                name = str(a.get("outcome", "")).lower()
                oid = a.get("outcome_id")
                if oid is None:
                    oid = a.get("id")
                if not oid:
                    continue
                if "email update" in name or "email received" in name:
                    email_in = (oid, a.get("outcome", "Email Update"))
                elif "email user" in name or "email sent" in name:
                    email_out = (oid, a.get("outcome", "Email User"))
                elif "recorded note" in name or "note" in name.lower().split():
                    internal = (oid, a.get("outcome", "Recorded Note"))

            if email_in:
                ok(f"email_received \u2192 {email_in[0]} ({email_in[1]})")
            else:
                info("email_received: not auto-detected (check manually; common ID: 0)")

            if email_out:
                ok(f"email_sent \u2192 {email_out[0]} ({email_out[1]})")
            else:
                info("email_sent: not auto-detected (check manually; common ID: 16)")

            if internal:
                ok(f"internal_note \u2192 {internal[0]} ({internal[1]})")
            else:
                info("internal_note: not auto-detected (check manually; common ID: 7)")

            print("\nAdd these to config.yaml under halo.actions:")
            print(f"  email_received: {email_in[0] if email_in else 0}")
            print(f"  email_sent:     {email_out[0] if email_out else 16}")
            print(f"  internal_note:  {internal[0] if internal else 7}")
    except Exception as e:
        fail(f"Action discovery failed: {e}")


async def list_custom_fields(config):
    """List custom fields to help find the conversationId field."""
    print("\n── Custom Fields ──")

    token = await _get_halo_token(config)
    if not token:
        fail("Authentication required for custom field listing")
        return

    try:
        api_url = f"{config.halo.instance_url.rstrip('/')}/api"
        async with httpx.AsyncClient(base_url=api_url, timeout=15.0) as c:
            r = await c.get(
                "/CustomFields",
                headers={"Authorization": f"Bearer {token}"},
            )
            if r.status_code != 200:
                fail(f"Could not fetch custom fields (status {r.status_code})")
                return

            data = r.json()
            fields = data if isinstance(data, list) else data.get("customfields", [])

            if not fields:
                info("No custom fields found")
                return

            print(f"  Found {len(fields)} custom field(s):\n")
            print(f"  {'ID':<6} {'Name':<40} {'Scope':<15} {'Type'}")
            print(f"  {'─'*6} {'─'*40} {'─'*15} {'─'*4}")
            for f in fields:
                fid = f.get("id", "?")
                name = str(f.get("name", "?"))[:39]
                scope = str(f.get("scope", "?"))[:14]
                ftype = f.get("type", "?")
                print(f"  {fid:<6} {name:<40} {scope:<15} {ftype}")

    except Exception as e:
        fail(f"Custom field listing failed: {e}")


async def run_checks(args) -> int:
    """Run pre-flight checks."""

    # Load config
    print("── Configuration ──")
    try:
        config = load_config(args.config)
        ok(f"config.yaml loaded and valid")
        ok(f"Halo URL: {config.halo.instance_url}")
        ok(f"Mailbox: {config.graph.user_email}")
    except FileNotFoundError:
        fail(f"Config file not found: {args.config}")
        info("Copy config.example.yaml to config.yaml and fill in your values")
        return 1
    except Exception as e:
        fail(f"Config validation failed: {e}")
        return 1

    # Discovery-only modes
    if args.discover_actions:
        await discover_actions(config)
        return 0

    if args.list_custom_fields:
        await list_custom_fields(config)
        return 0

    # Full pre-flight
    errors = 0

    if not await check_halo(config):
        errors += 1

    if not await check_graph(config):
        errors += 1

    # Summary
    print(f"\n{'─'*40}")
    if errors:
        print(f"  \033[31m{errors} check(s) failed\033[0m")
        print("  Fix the issues above, then re-run.")
        return 1
    else:
        print("  \033[32mAll checks passed\033[0m")
        print("  Watcher is ready to start:")
        print(f"    python -m watcher.watcher --config {args.config}")
        return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Pre-flight validation for Halo Outlook Extension"
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to config.yaml (default: config.yaml)",
    )
    parser.add_argument(
        "--discover-actions",
        action="store_true",
        help="Auto-discover ticket action IDs and exit",
    )
    parser.add_argument(
        "--list-custom-fields",
        action="store_true",
        help="List all custom fields on the Halo instance and exit",
    )
    args = parser.parse_args()

    sys.exit(asyncio.run(run_checks(args)))


if __name__ == "__main__":
    main()