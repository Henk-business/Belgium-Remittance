"""
Per-customer output rules.

Each customer can have a set of rules that control how their sheet is built:

  chunk_size      float   Group rows into chunks where cumulative abs(amount)
                          approaches this value. 0 = no chunking (default).
  columns         list    Ordered list of column names to include.
                          Omit a column name to exclude it entirely.
  sort_by         list    Column names to sort by before chunking.
                          Default: ["Net due date", "due_date"]
  show_account    bool    Whether to include the Account column. Default True.
  total_position  str     "bottom" (normal) or "right" (yellow box to the right).

Example config for account 30111788:
{
    "chunk_size": 40000,
    "columns": ["Assignment","Document Number","Reference Key 3",
                "Document Date","Net due date","Document Type",
                "Amount in local currency"],
    "show_account": false,
    "total_position": "right"
}

Configs are stored as JSON files in templates/config_{account_id}.json in GitHub.
"""

import json
import io
import streamlit as st
from typing import Optional


DEFAULT_RULE = {
    "chunk_size":     0,
    "columns":        [],
    "sort_by":        ["Net due date", "due_date"],
    "show_account":   True,
    "total_position": "bottom",
}


def merge_rule(rule: dict) -> dict:
    """Merge a partial rule with defaults."""
    out = DEFAULT_RULE.copy()
    out.update(rule)
    return out


# ── GITHUB STORAGE ────────────────────────────────────────────────────────────

def _gh_headers():
    token = st.secrets.get("github", {}).get("token", "")
    if not token:
        return {}
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _gh_repo():
    return st.secrets.get("github", {}).get("repo", "")


def _api(path):
    return f"https://api.github.com/repos/{_gh_repo()}/contents/{path}"


def _gh_ok():
    return bool(_gh_repo() and st.secrets.get("github", {}).get("token", ""))


def load_rule_github(account_id: str) -> Optional[dict]:
    if not _gh_ok():
        return None
    import requests, base64
    try:
        resp = requests.get(_api(f"templates/config_{account_id}.json"),
                            headers=_gh_headers(), timeout=10)
        if not resp.ok:
            return None
        raw = base64.b64decode(resp.json()["content"].replace("\n", ""))
        return merge_rule(json.loads(raw))
    except Exception:
        return None


def save_rule_github(account_id: str, rule: dict) -> tuple[bool, str]:
    if not _gh_ok():
        return False, "GitHub not configured."
    import requests, base64
    raw    = json.dumps(rule, indent=2).encode()
    b64    = base64.b64encode(raw).decode()
    path   = f"templates/config_{account_id}.json"
    # Check if file exists (need SHA to update)
    sha    = None
    check  = requests.get(_api(path), headers=_gh_headers(), timeout=10)
    if check.ok:
        sha = check.json().get("sha")
    payload = {"message": f"Update rules for account {account_id}", "content": b64}
    if sha:
        payload["sha"] = sha
    resp = requests.put(_api(path),
                        headers={**_gh_headers(), "Content-Type": "application/json"},
                        json=payload, timeout=20)
    if resp.status_code in (200, 201):
        return True, f"Rules saved for account {account_id}."
    return False, f"GitHub error {resp.status_code}: {resp.json().get('message', '')}"


def delete_rule_github(account_id: str) -> tuple[bool, str]:
    if not _gh_ok():
        return False, "GitHub not configured."
    import requests
    path  = f"templates/config_{account_id}.json"
    check = requests.get(_api(path), headers=_gh_headers(), timeout=10)
    if not check.ok:
        return False, "Rule file not found."
    sha   = check.json().get("sha")
    resp  = requests.delete(_api(path),
                            headers={**_gh_headers(), "Content-Type": "application/json"},
                            json={"message": f"Delete rules for {account_id}", "sha": sha},
                            timeout=15)
    if resp.ok:
        return True, f"Rules deleted for account {account_id}."
    return False, f"GitHub error {resp.status_code}"


@st.cache_data(ttl=300, show_spinner=False)
def get_rule_cached(account_id: str, _repo: str) -> Optional[dict]:
    return load_rule_github(account_id)


def invalidate_rule_cache():
    get_rule_cached.clear()
