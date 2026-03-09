#!/usr/bin/env python3
"""
Logical end-to-end test plan orchestrator for lib_Microsoft_Sharepoint sequences.

The plan runs in phases:
1) Bootstrap and resolver checks
2) List lifecycle checks
3) Drive/content lifecycle checks
4) Sharing/permissions checks
5) Cleanup checks

This script calls Convertigo over HTTP using:
- a provided Convertigo session (JSESSIONID / X-Convertigo-Authenticated), or
- an automatic login sequence call (default: ClientSDKtesting/login).

Graph credentials are optional:
- delegated mode: set ACCESS_TOKEN
- application mode: set AZ_TENANT_ID / AZ_CLIENT_ID / AZ_CLIENT_SECRET
- otherwise, sequence variables rely on server-side symbols/defaults.
"""

from __future__ import annotations

import base64
import datetime as dt
import json
import os
import re
import subprocess
import sys
import traceback
import urllib.parse
import urllib.request
from typing import Any, Callable, Dict, List, Optional


def _env(name: str, default: Optional[str] = None, required: bool = False) -> str:
    value = os.getenv(name, default)
    if required and (value is None or value == ""):
        raise RuntimeError(f"Missing required environment variable: {name}")
    return "" if value is None else value


def _env_bool(name: str, default: str) -> bool:
    return _env(name, default).strip().lower() in {"1", "true", "yes", "y", "on"}


def _normalize_convertigo_server_url(raw: str) -> str:
    value = raw.strip().rstrip("/")
    if value.endswith("/convertigo"):
        return value[: -len("/convertigo")]
    return value


# Convertigo session/auth context.
RAW_TEST_SERVER_ENDPOINT = os.getenv("TEST_SERVER_ENDPOINT", "")
RAW_C8O_SERVER_URL = os.getenv("C8O_SERVER_URL", "")
RAW_C8O_BASE_URL = os.getenv("C8O_BASE_URL", "")
TEST_SERVER_ENDPOINT = _env("TEST_SERVER_ENDPOINT", "")
C8O_SERVER_URL = _normalize_convertigo_server_url(
    _env("C8O_SERVER_URL", TEST_SERVER_ENDPOINT or "http://localhost:18080")
)
C8O_PROJECT = _env("C8O_PROJECT", "lib_Microsoft_Sharepoint")
C8O_BASE_URL = _env(
    "C8O_BASE_URL",
    f"{C8O_SERVER_URL}/convertigo/projects/{C8O_PROJECT}/.json",
)
if _env_bool("CI", "false") and RAW_TEST_SERVER_ENDPOINT == "" and RAW_C8O_SERVER_URL == "" and RAW_C8O_BASE_URL == "":
    raise RuntimeError(
        "No Convertigo endpoint provided in CI. Set TEST_SERVER_ENDPOINT or C8O_SERVER_URL or C8O_BASE_URL."
    )
C8O_TIMEOUT_SEC = int(_env("C8O_TIMEOUT_SEC", "240"))
C8O_ADMIN_INSTANCE = _env("C8O_ADMIN_INSTANCE", "")
C8O_XSRF_TOKEN = _env("C8O_XSRF_TOKEN", "")
C8O_JSESSIONID = _env("C8O_JSESSIONID", "")
C8O_XAUTH_TOKEN = _env("C8O_XAUTH_TOKEN", "")
C8O_AUTO_LOGIN = _env_bool("C8O_AUTO_LOGIN", "true")
C8O_LOGIN_PROJECT = _env("C8O_LOGIN_PROJECT", "ClientSDKtesting")
C8O_LOGIN_SEQUENCE = _env("C8O_LOGIN_SEQUENCE", "login")
C8O_LOGIN_BASE_URL = _env(
    "C8O_LOGIN_BASE_URL",
    f"{C8O_SERVER_URL}/convertigo/projects/{C8O_LOGIN_PROJECT}/.json",
)
C8O_LOGIN_EXTRA_FORM = _env("C8O_LOGIN_EXTRA_FORM", "")

# Token mode: delegated (ACCESS_TOKEN) or application (tenant/client/secret).
ACCESS_TOKEN = _env("ACCESS_TOKEN", "")
AZ_TENANT_ID = _env("AZ_TENANT_ID", "")
AZ_CLIENT_ID = _env("AZ_CLIENT_ID", "")
AZ_CLIENT_SECRET = _env("AZ_CLIENT_SECRET", "")
if any([AZ_TENANT_ID, AZ_CLIENT_ID, AZ_CLIENT_SECRET]) and not all(
    [AZ_TENANT_ID, AZ_CLIENT_ID, AZ_CLIENT_SECRET]
):
    raise RuntimeError("If AZ_* variables are used, provide all 3: AZ_TENANT_ID/AZ_CLIENT_ID/AZ_CLIENT_SECRET")

if ACCESS_TOKEN:
    TOKEN_MODE = "delegated_explicit"
elif AZ_TENANT_ID and AZ_CLIENT_ID and AZ_CLIENT_SECRET:
    TOKEN_MODE = "application_explicit"
else:
    TOKEN_MODE = "server_symbols_or_defaults"

TEST_PROVIDER = _env("TEST_PROVIDER", "").strip().lower()
if TEST_PROVIDER not in {"", "graph", "onprem", "both"}:
    raise RuntimeError("Invalid TEST_PROVIDER. Expected one of: graph, onprem, both, or empty")
ONPREM_SITE_BASE_URL = _env("ONPREM_SITE_BASE_URL", "")
ONPREM_PROTOCOL = _env("ONPREM_PROTOCOL", "")
ONPREM_USERNAME = _env("ONPREM_USERNAME", "")
ONPREM_PASSWORD = _env("ONPREM_PASSWORD", "")
ONPREM_COOKIE_HEADER = _env("ONPREM_COOKIE_HEADER", "")
ONPREM_HTTP_SERVER = _env("ONPREM_HTTP_SERVER", "")

# SharePoint test context.
SITE_ID = _env("SITE_ID", "")
SITE_HOSTNAME = _env("SITE_HOSTNAME", "")
SITE_PATH = _env("SITE_PATH", "")
LIST_ID = _env("LIST_ID", "")
LIST_NAME = _env("LIST_NAME", "")
DRIVE_ID = _env("DRIVE_ID", "")
DRIVE_NAME = _env("DRIVE_NAME", "")

GRAPH_DEFAULT_SITE_HOSTNAME = "convertigo.sharepoint.com"
GRAPH_DEFAULT_SITE_PATH = "/sites/test1"
GRAPH_DEFAULT_LIST_NAME = "Web Template Extensions"
GRAPH_DEFAULT_DRIVE_NAME = "Documents"

EXISTING_LIST_ITEM_ID = _env("EXISTING_LIST_ITEM_ID", "")
FALLBACK_LIST_ITEM_ID = _env("FALLBACK_LIST_ITEM_ID", "8")
EXISTING_DRIVE_ITEM_ID = _env("EXISTING_DRIVE_ITEM_ID", "")
FALLBACK_DRIVE_ITEM_ID = _env("FALLBACK_DRIVE_ITEM_ID", "01PETTHVHBULYXIKX2EZFIAFNZVAYMGFUP")
EXISTING_PERMISSION_ID = _env("EXISTING_PERMISSION_ID", "")
FALLBACK_PERMISSION_ID = _env("FALLBACK_PERMISSION_ID", "dGVzdDEgVmlzaXRvcnM")
INVITE_RECIPIENT_EMAILS = _env("INVITE_RECIPIENT_EMAILS", "test1@convertigo.onmicrosoft.com")
SUBSCRIPTION_NOTIFICATION_URL = _env("SUBSCRIPTION_NOTIFICATION_URL", "")
SUBSCRIPTION_RESOURCE = _env("SUBSCRIPTION_RESOURCE", "")
SUBSCRIPTION_EXPIRATION_DATETIME = _env("SUBSCRIPTION_EXPIRATION_DATETIME", "")
SUBSCRIPTION_CHANGE_TYPE = _env("SUBSCRIPTION_CHANGE_TYPE", "updated")
SUBSCRIPTION_CLIENT_STATE = _env("SUBSCRIPTION_CLIENT_STATE", "c8o-sharepoint-tests")

RUN_SHARE_OPERATIONS = _env_bool("RUN_SHARE_OPERATIONS", "true")
RUN_DESTRUCTIVE_CLEANUP = _env_bool("RUN_DESTRUCTIVE_CLEANUP", "true")
RUN_RESTORE_VERSION = _env_bool("RUN_RESTORE_VERSION", "false")
RUN_SUBSCRIPTION_OPERATIONS = _env_bool("RUN_SUBSCRIPTION_OPERATIONS", "false")
REPORT_FILE = _env("REPORT_FILE", "build/logical-test-plan-report.json")

# Mutable session state used by call_sequence.
SESSION: Dict[str, str] = {
    "jsessionid": C8O_JSESSIONID,
    "xauth": C8O_XAUTH_TOKEN,
}


def _json_dumps(value: Any) -> str:
    return json.dumps(value, separators=(",", ":"), ensure_ascii=True)


def _parse_form_pairs(raw_pairs: str) -> Dict[str, str]:
    parsed: Dict[str, str] = {}
    if raw_pairs.strip() == "":
        return parsed
    for chunk in raw_pairs.split("&"):
        if chunk == "":
            continue
        if "=" in chunk:
            key, value = chunk.split("=", 1)
        else:
            key, value = chunk, ""
        if key:
            parsed[key] = value
    return parsed


def _report_file_for_mode(base_report_file: str, mode: str) -> str:
    root, ext = os.path.splitext(base_report_file)
    if ext == "":
        ext = ".json"
    return f"{root}-{mode}{ext}"


def _run_both_provider_modes() -> int:
    script_path = os.path.abspath(__file__)
    providers = ["graph", "onprem"]
    run_outcomes: Dict[str, int] = {}
    run_reports: Dict[str, str] = {}
    combined_runs: Dict[str, Any] = {}

    for provider in providers:
        mode_report = _report_file_for_mode(REPORT_FILE, provider)
        env = os.environ.copy()
        env["TEST_PROVIDER"] = provider
        env["REPORT_FILE"] = mode_report
        print(f"=== Running provider mode: {provider} ===")
        completed = subprocess.run([sys.executable, script_path], env=env)
        run_outcomes[provider] = int(completed.returncode)
        run_reports[provider] = mode_report
        if os.path.isfile(mode_report):
            try:
                with open(mode_report, "r", encoding="utf-8") as f:
                    combined_runs[provider] = json.load(f)
            except Exception as exc:
                combined_runs[provider] = {"reportReadError": str(exc), "reportFile": mode_report}
        print()

    combined_summary = {
        "total": 0,
        "passed": 0,
        "failed": 0,
        "skipped": 0,
        "mandatoryFailed": 0,
    }
    for provider in providers:
        run_payload = combined_runs.get(provider, {})
        summary = run_payload.get("summary") if isinstance(run_payload, dict) else None
        if isinstance(summary, dict):
            combined_summary["total"] += int(summary.get("total", 0))
            combined_summary["passed"] += int(summary.get("passed", 0))
            combined_summary["failed"] += int(summary.get("failed", 0))
            combined_summary["skipped"] += int(summary.get("skipped", 0))
            combined_summary["mandatoryFailed"] += int(summary.get("mandatoryFailed", 0))

    combined_payload = {
        "timestamp": dt.datetime.now(dt.timezone.utc).isoformat(),
        "project": C8O_PROJECT,
        "baseUrl": C8O_BASE_URL,
        "tokenMode": TOKEN_MODE,
        "providerMode": "both",
        "runs": combined_runs,
        "runReports": run_reports,
        "runExitCodes": run_outcomes,
        "summary": combined_summary,
    }
    os.makedirs(os.path.dirname(REPORT_FILE) or ".", exist_ok=True)
    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        json.dump(combined_payload, f, indent=2, ensure_ascii=False)

    print("Combined provider run summary")
    print(
        "Summary: total={total} pass={passed} fail={failed} skip={skipped} mandatoryFail={mandatoryFailed}".format(
            **combined_summary
        )
    )
    print("Run exit codes: " + ", ".join([f"{k}={v}" for k, v in run_outcomes.items()]))
    print(f"Combined report: {REPORT_FILE}")

    return 1 if any(code != 0 for code in run_outcomes.values()) else 0


def _decode_payload(raw: str) -> Dict[str, Any]:
    try:
        return json.loads(raw)
    except Exception:
        return {"_raw": raw}


def _payload_is_login_ok(payload: Dict[str, Any]) -> bool:
    doc = payload.get("document")
    if isinstance(doc, dict):
        return str(doc.get("ok", "")).lower() == "true"
    resp = payload.get("response")
    return isinstance(resp, dict) and resp.get("ok") is True


def _auto_login() -> None:
    form: Dict[str, str] = {"__sequence": C8O_LOGIN_SEQUENCE}
    form.update(_parse_form_pairs(C8O_LOGIN_EXTRA_FORM))
    body = urllib.parse.urlencode(form).encode("utf-8")

    headers = {
        "Accept": "*/*",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    if C8O_ADMIN_INSTANCE:
        headers["Admin-Instance"] = C8O_ADMIN_INSTANCE
    if C8O_XSRF_TOKEN:
        headers["x-xsrf-token"] = C8O_XSRF_TOKEN

    req = urllib.request.Request(
        C8O_LOGIN_BASE_URL,
        data=body,
        method="POST",
        headers=headers,
    )
    with urllib.request.urlopen(req, timeout=C8O_TIMEOUT_SEC) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
        set_cookies = resp.headers.get_all("Set-Cookie", [])
        xauth = resp.headers.get("X-Convertigo-Authenticated", "")

    payload = _decode_payload(raw)
    if not _payload_is_login_ok(payload):
        raise RuntimeError(
            f"Auto-login failed on {C8O_LOGIN_PROJECT}/{C8O_LOGIN_SEQUENCE}: {response_error(payload)}"
        )

    jsessionid = ""
    for cookie in set_cookies:
        match = re.search(r"(?i)\bJSESSIONID=([^;]+)", cookie)
        if match:
            jsessionid = match.group(1)
            break
    if jsessionid == "":
        raise RuntimeError("Auto-login failed: missing JSESSIONID in Set-Cookie headers")

    SESSION["jsessionid"] = jsessionid
    SESSION["xauth"] = xauth


def _ensure_session() -> None:
    if SESSION.get("jsessionid", "") != "":
        return
    if C8O_AUTO_LOGIN:
        _auto_login()


def call_sequence(sequence: str, params: Dict[str, Any]) -> Dict[str, Any]:
    _ensure_session()

    form: Dict[str, str] = {"__sequence": sequence}
    for key, value in params.items():
        if value is None:
            continue
        form[key] = str(value)

    if TOKEN_MODE == "delegated_explicit":
        form.setdefault("accessToken", ACCESS_TOKEN)
    elif TOKEN_MODE == "application_explicit":
        form.setdefault("tenantId", AZ_TENANT_ID)
        form.setdefault("clientId", AZ_CLIENT_ID)
        form.setdefault("clientSecret", AZ_CLIENT_SECRET)

    body = urllib.parse.urlencode(form).encode("utf-8")
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    if C8O_ADMIN_INSTANCE:
        headers["Admin-Instance"] = C8O_ADMIN_INSTANCE
    if C8O_XSRF_TOKEN:
        headers["x-xsrf-token"] = C8O_XSRF_TOKEN
    if SESSION.get("jsessionid", ""):
        headers["Cookie"] = f"introjs-dontShowAgain=true; JSESSIONID={SESSION['jsessionid']}"
    if SESSION.get("xauth", ""):
        headers["X-Convertigo-Authenticated"] = SESSION["xauth"]

    req = urllib.request.Request(
        C8O_BASE_URL,
        data=body,
        method="POST",
        headers=headers,
    )
    with urllib.request.urlopen(req, timeout=C8O_TIMEOUT_SEC) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    return _decode_payload(raw)


def response_ok(payload: Dict[str, Any]) -> bool:
    resp = payload.get("response")
    return isinstance(resp, dict) and resp.get("ok") is True


def response_error(payload: Dict[str, Any]) -> str:
    err = payload.get("error")
    if isinstance(err, dict):
        return str(err.get("message", "Convertigo error"))
    process = payload.get("process")
    if isinstance(process, dict):
        p_error = str(process.get("error", "") or "")
        if p_error:
            return p_error
        return f"process exit={process.get('exit', '?')}"
    doc = payload.get("document")
    if isinstance(doc, dict) and isinstance(doc.get("error"), dict):
        return str(doc["error"].get("message", "Convertigo document error"))
    resp = payload.get("response")
    if not isinstance(resp, dict):
        raw = payload.get("_raw", "")
        return "Invalid payload" if not raw else raw[:240]
    msg = str(resp.get("errorMessage", "") or "")
    if msg:
        return msg
    return str(resp.get("status", "ERROR"))


def response_data(payload: Dict[str, Any]) -> Dict[str, Any]:
    resp = payload.get("response")
    if not isinstance(resp, dict):
        return {}
    data = resp.get("data")
    return data if isinstance(data, dict) else {}


def _first_non_empty(*values: Any) -> str:
    for value in values:
        text = str(value or "").strip()
        if text != "":
            return text
    return ""


def _normalize_site_path(path_value: str, fallback: str) -> str:
    path = _first_non_empty(path_value, fallback, "/")
    if not path.startswith("/"):
        path = "/" + path
    if len(path) > 1 and path.endswith("/"):
        path = path[:-1]
    return path


def _resolve_scope_defaults(is_onprem_mode: bool) -> Dict[str, str]:
    hostname = SITE_HOSTNAME.strip()
    path = SITE_PATH.strip()
    list_name = LIST_NAME.strip()
    drive_name = DRIVE_NAME.strip()

    if is_onprem_mode:
        if ONPREM_SITE_BASE_URL.strip() != "":
            try:
                parsed = urllib.parse.urlparse(ONPREM_SITE_BASE_URL.strip())
                if hostname == "" and parsed.hostname:
                    hostname = str(parsed.hostname)
                if path == "" and parsed.path:
                    path = str(parsed.path)
            except Exception:
                pass
        if hostname == "":
            hostname = ONPREM_HTTP_SERVER.strip()
        if path != "":
            path = _normalize_site_path(path, "/")
    else:
        hostname = _first_non_empty(hostname, GRAPH_DEFAULT_SITE_HOSTNAME)
        path = _normalize_site_path(path, GRAPH_DEFAULT_SITE_PATH)
        list_name = _first_non_empty(list_name, GRAPH_DEFAULT_LIST_NAME)
        drive_name = _first_non_empty(drive_name, GRAPH_DEFAULT_DRIVE_NAME)

    return {
        "siteHostname": hostname,
        "sitePath": path,
        "listName": list_name,
        "driveName": drive_name,
    }


def _extract_item_id_from_object(obj: Any) -> str:
    if not isinstance(obj, dict):
        return ""
    return _first_non_empty(obj.get("itemId"), obj.get("id"))


def _extract_first_item_id(data: Dict[str, Any]) -> str:
    items = data.get("items")
    if isinstance(items, list):
        for item in items:
            item_id = _extract_item_id_from_object(item)
            if item_id:
                return item_id
    item = data.get("item")
    return _extract_item_id_from_object(item)


def _extract_permission_id_from_data(data: Dict[str, Any]) -> str:
    permission = data.get("permission")
    if isinstance(permission, dict):
        value = _first_non_empty(permission.get("permissionId"), permission.get("id"))
        if value:
            return value
    permissions = data.get("permissions")
    if isinstance(permissions, list):
        for perm in permissions:
            if isinstance(perm, dict):
                value = _first_non_empty(perm.get("permissionId"), perm.get("id"))
                if value:
                    return value
    return ""


def _extract_version_id_from_data(data: Dict[str, Any]) -> str:
    versions = data.get("versions")
    if isinstance(versions, list):
        for version in versions:
            if isinstance(version, dict):
                value = _first_non_empty(version.get("versionId"), version.get("id"))
                if value:
                    return value
    version = data.get("version")
    if isinstance(version, dict):
        return _first_non_empty(version.get("versionId"), version.get("id"))
    return ""


def _extract_subscription_id_from_data(data: Dict[str, Any]) -> str:
    value = _first_non_empty(data.get("subscriptionId"), data.get("id"))
    if value:
        return value
    subscription = data.get("subscription")
    if isinstance(subscription, dict):
        return _first_non_empty(subscription.get("id"), subscription.get("subscriptionId"))
    return ""


class TestPlan:
    def __init__(self) -> None:
        self.results: List[Dict[str, Any]] = []
        self.ctx: Dict[str, Any] = {
            "graph_access_token": ACCESS_TOKEN,
            "site_id": SITE_ID,
            "list_id": LIST_ID,
            "drive_id": DRIVE_ID,
            "first_list_item_id": "",
            "created_list_item_id": "",
            "delete_list_item_id": "",
            "root_item_id": "root",
            "first_drive_item_id": "",
            "created_folder_item_id": "",
            "drive_item_id": "",
            "copy_monitor_url": "",
            "share_permission_id": "",
            "listed_permission_id": "",
            "invited_permission_id": "",
            "copy_operation_status_code": "",
            "drive_delta_link": "",
            "list_delta_link": "",
            "version_id": "",
            "subscription_id": "",
            "batch_failed_count": "",
        }
        self.mandatory_failed = False

    def add_result(self, phase: str, step: str, outcome: str, mandatory: bool, details: str = "") -> None:
        self.results.append(
            {
                "phase": phase,
                "step": step,
                "outcome": outcome,
                "mandatory": mandatory,
                "details": details,
            }
        )
        tag = "MANDATORY" if mandatory else "OPTIONAL"
        if outcome == "PASS":
            print(f"[PASS] [{phase}] {step} ({tag})")
        elif outcome == "SKIP":
            print(f"[SKIP] [{phase}] {step} ({tag}) {details}")
        else:
            print(f"[FAIL] [{phase}] {step} ({tag}) {details}")
            if mandatory:
                self.mandatory_failed = True

    def run_step(
        self,
        phase: str,
        step: str,
        sequence: str,
        params: Dict[str, Any],
        mandatory: bool = False,
        enabled: bool = True,
        validator: Optional[Callable[[Dict[str, Any]], bool]] = None,
        on_success: Optional[Callable[[Dict[str, Any], Dict[str, Any]], None]] = None,
    ) -> bool:
        if not enabled:
            self.add_result(phase, step, "SKIP", mandatory, "disabled by context/config")
            return False
        effective_params = dict(params)
        effective_token = _first_non_empty(self.ctx.get("graph_access_token", ""), "")
        if effective_token and _first_non_empty(effective_params.get("accessToken", ""), "") == "":
            effective_params["accessToken"] = effective_token
        try:
            payload = call_sequence(sequence, effective_params)
            ok = response_ok(payload) if validator is None else bool(validator(payload))
            if ok:
                if on_success is not None:
                    on_success(payload, self.ctx)
                self.add_result(phase, step, "PASS", mandatory)
                return True
            self.add_result(phase, step, "FAIL", mandatory, response_error(payload))
            return False
        except Exception as exc:
            self.add_result(phase, step, "FAIL", mandatory, f"{exc.__class__.__name__}: {exc}")
            return False

    def effective_list_item_id(self) -> str:
        return _first_non_empty(
            self.ctx.get("created_list_item_id", ""),
            self.ctx.get("first_list_item_id", ""),
            EXISTING_LIST_ITEM_ID,
            FALLBACK_LIST_ITEM_ID,
        )

    def effective_drive_item_id(self) -> str:
        return _first_non_empty(
            self.ctx.get("drive_item_id", ""),
            self.ctx.get("first_drive_item_id", ""),
            EXISTING_DRIVE_ITEM_ID,
            FALLBACK_DRIVE_ITEM_ID,
        )

    def effective_permission_id(self) -> str:
        return _first_non_empty(
            self.ctx.get("invited_permission_id", ""),
            self.ctx.get("share_permission_id", ""),
            self.ctx.get("listed_permission_id", ""),
            EXISTING_PERMISSION_ID,
            FALLBACK_PERMISSION_ID,
        )

    def write_report(self) -> None:
        os.makedirs(os.path.dirname(REPORT_FILE) or ".", exist_ok=True)
        payload = {
            "timestamp": dt.datetime.now(dt.timezone.utc).isoformat(),
            "project": C8O_PROJECT,
            "baseUrl": C8O_BASE_URL,
            "site": {
                "siteId": self.ctx.get("site_id", SITE_ID),
                "siteHostname": self.ctx.get("site_hostname", SITE_HOSTNAME),
                "sitePath": self.ctx.get("site_path", SITE_PATH),
                "listId": self.ctx.get("list_id", LIST_ID),
                "listName": self.ctx.get("list_name", LIST_NAME),
                "driveId": self.ctx.get("drive_id", DRIVE_ID),
                "driveName": self.ctx.get("drive_name", DRIVE_NAME),
            },
            "tokenMode": TOKEN_MODE,
            "providerMode": {
                "forcedProvider": TEST_PROVIDER,
                "resolved": TEST_PROVIDER if TEST_PROVIDER else "symbol/default",
                "onPremOverrides": {
                    "siteBaseUrl": ONPREM_SITE_BASE_URL,
                    "httpServer": ONPREM_HTTP_SERVER,
                    "protocol": ONPREM_PROTOCOL if TEST_PROVIDER == "onprem" else "",
                    "hasUsername": bool(ONPREM_USERNAME.strip()),
                    "hasPassword": bool(ONPREM_PASSWORD.strip()),
                    "hasCookieHeader": bool(ONPREM_COOKIE_HEADER.strip()),
                },
            },
            "convertigoSession": {
                "hasJSessionId": bool(SESSION.get("jsessionid")),
                "hasXAuthToken": bool(SESSION.get("xauth")),
                "autoLoginEnabled": C8O_AUTO_LOGIN,
                "loginProject": C8O_LOGIN_PROJECT,
                "loginSequence": C8O_LOGIN_SEQUENCE,
            },
            "results": self.results,
            "context": self.ctx,
            "summary": self.summary(),
        }
        with open(REPORT_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

    def summary(self) -> Dict[str, int]:
        total = len(self.results)
        passed = sum(1 for r in self.results if r["outcome"] == "PASS")
        failed = sum(1 for r in self.results if r["outcome"] == "FAIL")
        skipped = sum(1 for r in self.results if r["outcome"] == "SKIP")
        mandatory_failed = sum(1 for r in self.results if r["outcome"] == "FAIL" and bool(r["mandatory"]))
        return {
            "total": total,
            "passed": passed,
            "failed": failed,
            "skipped": skipped,
            "mandatoryFailed": mandatory_failed,
        }


def main() -> int:
    if TEST_PROVIDER == "both":
        return _run_both_provider_modes()

    plan = TestPlan()
    is_forced_provider = TEST_PROVIDER != ""
    is_onprem_mode = TEST_PROVIDER == "onprem"
    resolved_scope = _resolve_scope_defaults(is_onprem_mode)
    resolved_site_hostname = resolved_scope["siteHostname"]
    resolved_site_path = resolved_scope["sitePath"]
    resolved_list_name = resolved_scope["listName"]
    resolved_drive_name = resolved_scope["driveName"]
    plan.ctx["site_hostname"] = resolved_site_hostname
    plan.ctx["site_path"] = resolved_site_path
    plan.ctx["list_name"] = resolved_list_name
    plan.ctx["drive_name"] = resolved_drive_name
    ts = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d%H%M%S")
    folder_name = f"C8O_plan_folder_{ts}"
    file_name = f"c8o-plan-file-{ts}.txt"
    moved_name = f"c8o-plan-file-moved-{ts}.txt"
    copy_name = f"c8o-plan-copy-{ts}.txt"
    list_title = f"C8O List Item {ts}"
    list_updated_title = f"C8O List Item Updated {ts}"
    content_b64 = base64.b64encode(f"convertigo sharepoint plan {ts}".encode("utf-8")).decode("ascii")
    large_content_b64 = base64.b64encode(("0123456789ABCDEF" * 8192).encode("utf-8")).decode("ascii")
    subscription_expiration_iso = _first_non_empty(
        SUBSCRIPTION_EXPIRATION_DATETIME,
        (dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=50))
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z"),
    )

    print(f"Logical test plan for project {C8O_PROJECT}")
    print(f"Base URL: {C8O_BASE_URL}")
    print(f"Token mode: {TOKEN_MODE}")
    print(f"Provider mode: {TEST_PROVIDER if is_forced_provider else 'symbol/default'}")
    resolved_scope_host = resolved_site_hostname if resolved_site_hostname != "" else "<sequence default>"
    resolved_scope_path = resolved_site_path if resolved_site_path != "" else "<sequence default>"
    resolved_scope_list = resolved_list_name if resolved_list_name != "" else "<sequence default>"
    resolved_scope_drive = resolved_drive_name if resolved_drive_name != "" else "<sequence default>"
    print(
        f"Scope: host={resolved_scope_host} path={resolved_scope_path} list={resolved_scope_list} drive={resolved_scope_drive}"
    )
    if SESSION.get("jsessionid", ""):
        print("Convertigo session: provided by environment")
    elif C8O_AUTO_LOGIN:
        _ensure_session()
        print(f"Convertigo session: auto-login OK via {C8O_LOGIN_PROJECT}/{C8O_LOGIN_SEQUENCE}")
    else:
        print("Convertigo session: none (authenticated sequences may fail)")
    print()

    def site_scope() -> Dict[str, Any]:
        payload: Dict[str, Any] = {"siteId": plan.ctx.get("site_id", "")}
        if resolved_site_hostname != "":
            payload["siteHostname"] = resolved_site_hostname
        if resolved_site_path != "":
            payload["sitePath"] = resolved_site_path
        if is_forced_provider:
            payload["provider"] = TEST_PROVIDER
        if is_onprem_mode:
            if ONPREM_SITE_BASE_URL.strip() != "":
                payload["siteBaseUrl"] = ONPREM_SITE_BASE_URL
            if ONPREM_PROTOCOL.strip() != "":
                payload["onPremProtocol"] = ONPREM_PROTOCOL
            if ONPREM_USERNAME.strip() != "":
                payload["onPremUsername"] = ONPREM_USERNAME
            if ONPREM_PASSWORD.strip() != "":
                payload["onPremPassword"] = ONPREM_PASSWORD
            if ONPREM_COOKIE_HEADER.strip() != "":
                payload["cookieHeader"] = ONPREM_COOKIE_HEADER
        return payload

    def list_scope() -> Dict[str, Any]:
        payload = site_scope()
        payload["listId"] = plan.ctx.get("list_id", "")
        if resolved_list_name != "":
            payload["listName"] = resolved_list_name
        return payload

    def drive_scope() -> Dict[str, Any]:
        payload = site_scope()
        payload["driveId"] = plan.ctx.get("drive_id", "")
        if resolved_drive_name != "":
            payload["driveName"] = resolved_drive_name
        return payload

    def save_site(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["site_id"] = _first_non_empty(response_data(payload).get("siteId"), ctx.get("site_id", ""))

    def save_graph_access_token(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["graph_access_token"] = _first_non_empty(response_data(payload).get("token"), ctx.get("graph_access_token", ""))

    def save_list(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["list_id"] = _first_non_empty(response_data(payload).get("listId"), ctx.get("list_id", ""))

    def save_drive(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["drive_id"] = _first_non_empty(response_data(payload).get("driveId"), ctx.get("drive_id", ""))

    def save_first_list_item(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["first_list_item_id"] = _extract_first_item_id(response_data(payload))

    def save_created_list_item(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        item_id = _extract_first_item_id(response_data(payload))
        ctx["created_list_item_id"] = item_id
        ctx["delete_list_item_id"] = item_id

    def save_root_item(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        item_id = _extract_first_item_id(response_data(payload))
        if item_id:
            ctx["root_item_id"] = item_id

    def save_first_drive_item(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["first_drive_item_id"] = _extract_first_item_id(response_data(payload))

    def save_created_folder(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["created_folder_item_id"] = _extract_first_item_id(response_data(payload))

    def save_drive_item(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        item_id = _extract_first_item_id(response_data(payload))
        if item_id:
            ctx["drive_item_id"] = item_id

    def save_copy_monitor(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        data = response_data(payload)
        ctx["copy_monitor_url"] = _first_non_empty(data.get("monitorUrl"), data.get("operationLocation"))

    def save_copy_operation_status(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        data = response_data(payload)
        ctx["copy_operation_status_code"] = _first_non_empty(data.get("statusCode"), "")

    def save_list_delta(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["list_delta_link"] = _first_non_empty(response_data(payload).get("deltaLink"), ctx.get("list_delta_link", ""))

    def save_drive_delta(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["drive_delta_link"] = _first_non_empty(response_data(payload).get("deltaLink"), ctx.get("drive_delta_link", ""))

    def save_version(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["version_id"] = _extract_version_id_from_data(response_data(payload))

    def save_subscription(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["subscription_id"] = _extract_subscription_id_from_data(response_data(payload))

    def save_batch(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["batch_failed_count"] = _first_non_empty(response_data(payload).get("failedCount"), "0")

    def save_share_permission(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["share_permission_id"] = _extract_permission_id_from_data(response_data(payload))

    def save_listed_permission(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["listed_permission_id"] = _extract_permission_id_from_data(response_data(payload))

    def save_invited_permission(payload: Dict[str, Any], ctx: Dict[str, Any]) -> None:
        ctx["invited_permission_id"] = _extract_permission_id_from_data(response_data(payload))

    # Phase 1: Bootstrap and resolvers.
    plan.run_step(
        "phase-1-bootstrap",
        "GetGraphAccessToken",
        "GetGraphAccessToken",
        {"includeTokenPayload": "false"},
        mandatory=not is_onprem_mode,
        enabled=not is_onprem_mode,
        on_success=save_graph_access_token,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ResolveSite",
        "ResolveSite",
        site_scope(),
        mandatory=True,
        on_success=save_site,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ResolveList",
        "ResolveList",
        list_scope(),
        mandatory=True,
        on_success=save_list,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ResolveLibrary",
        "ResolveLibrary",
        drive_scope(),
        mandatory=True,
        on_success=save_drive,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ListSiteLists",
        "ListSiteLists",
        dict(site_scope(), top="20", maxPages="1"),
        mandatory=True,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ListSiteDrives",
        "ListSiteDrives",
        dict(site_scope(), top="20", maxPages="1"),
        mandatory=True,
    )
    plan.run_step(
        "phase-1-bootstrap",
        "ExecuteGraphBatch",
        "ExecuteGraphBatch",
        {
            "batchJson": _json_dumps(
                {
                    "requests": [
                        {
                            "id": "1",
                            "method": "GET",
                            "url": f"/sites/{plan.ctx.get('site_id', '')}?$select=id,displayName,webUrl",
                        },
                        {
                            "id": "2",
                            "method": "GET",
                            "url": f"/sites/{plan.ctx.get('site_id', '')}/drives?$top=1&$select=id,name",
                        },
                    ]
                }
            ),
            "includeRawResponse": "false",
        },
        mandatory=not is_onprem_mode,
        enabled=not is_onprem_mode,
        on_success=save_batch,
    )

    # Phase 2: List lifecycle.
    plan.run_step(
        "phase-2-list",
        "ListGetItems",
        "ListGetItems",
        dict(list_scope(), top="10", maxPages="1", expandFields="true"),
        mandatory=True,
        on_success=save_first_list_item,
    )
    plan.run_step(
        "phase-2-list",
        "ListGetItemsDelta",
        "ListGetItemsDelta",
        dict(list_scope(), top="10", maxPages="1", expandFields="true"),
        mandatory=False,
        enabled=True,
        on_success=save_list_delta,
    )
    plan.run_step(
        "phase-2-list",
        "GetListItem",
        "GetListItem",
        dict(list_scope(), itemId=plan.effective_list_item_id()),
        mandatory=False,
    )
    plan.run_step(
        "phase-2-list",
        "CreateListItem",
        "CreateListItem",
        dict(list_scope(), fieldsJson=_json_dumps({"Title": list_title})),
        mandatory=True,
        on_success=save_created_list_item,
    )
    plan.run_step(
        "phase-2-list",
        "UpdateListItem",
        "UpdateListItem",
        dict(list_scope(), itemId=plan.effective_list_item_id(), fieldsJson=_json_dumps({"Title": list_updated_title})),
        mandatory=True,
    )
    plan.run_step(
        "phase-2-list",
        "DeleteListItem",
        "DeleteListItem",
        dict(
            list_scope(),
            itemId=_first_non_empty(plan.ctx.get("delete_list_item_id", ""), plan.effective_list_item_id()),
            includeDeletedItemSnapshot="false",
        ),
        mandatory=True,
        enabled=RUN_DESTRUCTIVE_CLEANUP,
    )

    # Phase 3: Drive/content lifecycle.
    plan.run_step(
        "phase-3-drive",
        "GetItem",
        "GetItem",
        dict(drive_scope(), itemId="root"),
        mandatory=True,
        on_success=save_root_item,
    )
    plan.run_step(
        "phase-3-drive",
        "ListItems",
        "ListItems",
        dict(drive_scope(), parentItemId="root", top="20", maxPages="1"),
        mandatory=True,
        on_success=save_first_drive_item,
    )
    plan.run_step(
        "phase-3-drive",
        "ListItemsDelta",
        "ListItemsDelta",
        dict(drive_scope(), parentItemId="root", top="20", maxPages="1"),
        mandatory=False,
        enabled=True,
        on_success=save_drive_delta,
    )
    plan.run_step(
        "phase-3-drive",
        "CreateFolder",
        "CreateFolder",
        dict(drive_scope(), parentItemId="root", folderName=folder_name, conflictBehavior="rename"),
        mandatory=True,
        on_success=save_created_folder,
    )
    plan.run_step(
        "phase-3-drive",
        "UploadItemContent",
        "UploadItemContent",
        dict(
            drive_scope(),
            parentItemId=_first_non_empty(plan.ctx.get("created_folder_item_id", ""), "root"),
            fileName=file_name,
            contentBase64=content_b64,
            contentType="text/plain",
        ),
        mandatory=True,
        on_success=save_drive_item,
    )
    plan.run_step(
        "phase-3-drive",
        "DownloadItemContent",
        "DownloadItemContent",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), includeContentBase64="true"),
        mandatory=True,
    )
    plan.run_step(
        "phase-3-drive",
        "UpdateItem",
        "UpdateItem",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), updateJson=_json_dumps({"name": moved_name})),
        mandatory=True,
        on_success=save_drive_item,
    )
    plan.run_step(
        "phase-3-drive",
        "MoveItem",
        "MoveItem",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), destinationParentItemId="root", newName=moved_name),
        mandatory=True,
        on_success=save_drive_item,
    )
    plan.run_step(
        "phase-3-drive",
        "CopyItem",
        "CopyItem",
        dict(
            drive_scope(),
            itemId=plan.effective_drive_item_id(),
            destinationParentItemId="root",
            newName=copy_name,
            includeMonitorResponse="true",
        ),
        mandatory=True,
        on_success=save_copy_monitor,
    )
    plan.run_step(
        "phase-3-drive",
        "GetCopyItemOperation",
        "GetCopyItemOperation",
        dict(monitorUrl=_first_non_empty(plan.ctx.get("copy_monitor_url", ""), ""), includeMonitorBody="true"),
        mandatory=False,
        enabled=_first_non_empty(plan.ctx.get("copy_monitor_url", ""), "") != "",
        on_success=save_copy_operation_status,
    )
    plan.run_step(
        "phase-3-drive",
        "UploadItemLargeContent",
        "UploadItemLargeContent",
        dict(
            drive_scope(),
            itemId=plan.effective_drive_item_id(),
            contentBase64=large_content_b64,
            contentType="text/plain",
        ),
        mandatory=True,
        on_success=save_drive_item,
    )
    plan.run_step(
        "phase-3-drive",
        "ListItemVersions",
        "ListItemVersions",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), top="20", maxPages="1"),
        mandatory=False,
        on_success=save_version,
    )
    plan.run_step(
        "phase-3-drive",
        "RestoreItemVersion",
        "RestoreItemVersion",
        dict(
            drive_scope(),
            itemId=plan.effective_drive_item_id(),
            versionId=_first_non_empty(plan.ctx.get("version_id", ""), ""),
            includeItemSnapshot="false",
        ),
        mandatory=False,
        enabled=RUN_RESTORE_VERSION and _first_non_empty(plan.ctx.get("version_id", ""), "") != "",
    )

    # Phase 4: Sharing and permissions.
    plan.run_step(
        "phase-4-sharing",
        "CreateShareLink",
        "CreateShareLink",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), linkType="view", scope="organization"),
        mandatory=False,
        enabled=RUN_SHARE_OPERATIONS,
        on_success=save_share_permission,
    )
    plan.run_step(
        "phase-4-sharing",
        "ListItemPermissions",
        "ListItemPermissions",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), top="100", maxPages="2"),
        mandatory=False,
        enabled=RUN_SHARE_OPERATIONS,
        on_success=save_listed_permission,
    )
    plan.run_step(
        "phase-4-sharing",
        "InviteItemRecipients",
        "InviteItemRecipients",
        dict(
            drive_scope(),
            itemId=plan.effective_drive_item_id(),
            recipientEmails=INVITE_RECIPIENT_EMAILS,
            rolesCsv="read",
            sendInvitation="false",
        ),
        mandatory=False,
        enabled=RUN_SHARE_OPERATIONS and INVITE_RECIPIENT_EMAILS.strip() != "",
        on_success=save_invited_permission,
    )
    plan.run_step(
        "phase-4-sharing",
        "CreateGraphSubscription",
        "CreateGraphSubscription",
        dict(
            resource=_first_non_empty(
                SUBSCRIPTION_RESOURCE,
                f"/sites/{plan.ctx.get('site_id', '')}/drives/{plan.ctx.get('drive_id', '')}/root",
            ),
            changeType=SUBSCRIPTION_CHANGE_TYPE,
            notificationUrl=SUBSCRIPTION_NOTIFICATION_URL,
            expirationDateTime=subscription_expiration_iso,
            clientState=SUBSCRIPTION_CLIENT_STATE,
            includeResourceData="false",
        ),
        mandatory=False,
        enabled=(not is_onprem_mode) and RUN_SUBSCRIPTION_OPERATIONS and SUBSCRIPTION_NOTIFICATION_URL.strip() != "",
        on_success=save_subscription,
    )
    plan.run_step(
        "phase-4-sharing",
        "DeleteGraphSubscription",
        "DeleteGraphSubscription",
        dict(subscriptionId=_first_non_empty(plan.ctx.get("subscription_id", ""), "")),
        mandatory=False,
        enabled=(not is_onprem_mode)
        and RUN_SUBSCRIPTION_OPERATIONS
        and _first_non_empty(plan.ctx.get("subscription_id", ""), "") != "",
    )
    plan.run_step(
        "phase-4-sharing",
        "DeleteItemPermission",
        "DeleteItemPermission",
        dict(
            drive_scope(),
            itemId=plan.effective_drive_item_id(),
            permissionId=plan.effective_permission_id(),
            includeDeletedPermissionSnapshot="false",
        ),
        mandatory=False,
        enabled=RUN_SHARE_OPERATIONS,
    )

    # Phase 5: Cleanup.
    plan.run_step(
        "phase-5-cleanup",
        "DeleteItem",
        "DeleteItem",
        dict(drive_scope(), itemId=plan.effective_drive_item_id(), includeDeletedItemSnapshot="false"),
        mandatory=True,
        enabled=RUN_DESTRUCTIVE_CLEANUP,
    )
    plan.run_step(
        "phase-5-cleanup",
        "DeleteDriveFolderCleanup",
        "DeleteItem",
        dict(drive_scope(), itemId=_first_non_empty(plan.ctx.get("created_folder_item_id", ""), "root"), includeDeletedItemSnapshot="false"),
        mandatory=False,
        enabled=RUN_DESTRUCTIVE_CLEANUP and _first_non_empty(plan.ctx.get("created_folder_item_id", ""), "") != "",
    )

    summary = plan.summary()
    plan.write_report()
    print()
    print(
        "Summary: total={total} pass={passed} fail={failed} skip={skipped} mandatoryFail={mandatoryFailed}".format(
            **summary
        )
    )
    print(f"Report: {REPORT_FILE}")

    return 1 if plan.mandatory_failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"Fatal error: {exc}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(2)
