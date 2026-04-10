"""
GitHub-backed persistent storage for customer templates.

Templates are stored as .xlsx files in a `templates/` folder in your GitHub repo.
Reading works for anyone (public or private repo with a token).
Writing requires a personal access token with `contents: write` permission.

Setup (one time):
  1. Go to github.com → Settings → Developer settings → Personal access tokens → Fine-grained tokens
  2. Create a token with:
       - Repository access: your ar-suite repo only
       - Permissions: Contents → Read and write
  3. In Streamlit Cloud → your app → Settings → Secrets, add:
       [github]
       token = "github_pat_xxxxxxxxxxxx"
       repo  = "your-username/ar-suite"
"""

import base64
import json
import streamlit as st
import requests
import io
from typing import Optional


# ── GITHUB API HELPERS ────────────────────────────────────────────────────────

def _headers() -> dict:
    token = st.secrets.get("github", {}).get("token", "")
    if not token:
        return {}
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _repo() -> str:
    return st.secrets.get("github", {}).get("repo", "")


def _api(path: str) -> str:
    return f"https://api.github.com/repos/{_repo()}/contents/{path}"


def github_configured() -> bool:
    """Returns True if GitHub secrets are set up."""
    return bool(_repo() and st.secrets.get("github", {}).get("token", ""))


# ── READ ──────────────────────────────────────────────────────────────────────

def list_github_templates() -> list[dict]:
    """
    List all .xlsx files in the templates/ folder of the repo.
    Returns list of {account_id, filename, sha, size}.
    """
    if not github_configured():
        return []
    resp = requests.get(_api("templates"), headers=_headers(), timeout=10)
    if resp.status_code == 404:
        return []   # folder doesn't exist yet
    if not resp.ok:
        return []
    items = []
    for f in resp.json():
        if f.get("type") == "file" and f["name"].endswith(".xlsx"):
            acc_id = f["name"].replace(".xlsx", "")
            items.append({
                "account_id": acc_id,
                "filename":   f["name"],
                "sha":        f["sha"],
                "size":       f.get("size", 0),
            })
    return items


def load_github_template(account_id: str) -> Optional[bytes]:
    """Download a template file from GitHub. Returns bytes or None."""
    if not github_configured():
        return None
    filename = f"templates/{account_id}.xlsx"
    resp = requests.get(_api(filename), headers=_headers(), timeout=15)
    if not resp.ok:
        return None
    content_b64 = resp.json().get("content", "")
    # GitHub returns base64 with newlines
    return base64.b64decode(content_b64.replace("\n", ""))


def get_file_sha(account_id: str) -> Optional[str]:
    """Get the SHA of an existing template file (needed to update it)."""
    if not github_configured():
        return None
    filename = f"templates/{account_id}.xlsx"
    resp = requests.get(_api(filename), headers=_headers(), timeout=10)
    if resp.ok:
        return resp.json().get("sha")
    return None


# ── WRITE ─────────────────────────────────────────────────────────────────────

def save_github_template(account_id: str, file_bytes: bytes) -> tuple[bool, str]:
    """
    Upload or update a template file in the GitHub repo.
    Returns (success, message).
    """
    if not github_configured():
        return False, "GitHub not configured. Add your token and repo to Streamlit secrets."

    filename = f"templates/{account_id}.xlsx"
    content_b64 = base64.b64encode(file_bytes).decode("utf-8")
    sha = get_file_sha(account_id)

    payload = {
        "message": f"Update template for account {account_id}",
        "content": content_b64,
    }
    if sha:
        payload["sha"] = sha   # required when updating an existing file

    resp = requests.put(
        _api(filename),
        headers={**_headers(), "Content-Type": "application/json"},
        json=payload,
        timeout=20,
    )

    if resp.status_code in (200, 201):
        action = "Updated" if sha else "Saved"
        return True, f"{action} template for account {account_id} in GitHub."
    else:
        try:
            msg = resp.json().get("message", resp.text)
        except Exception:
            msg = resp.text
        return False, f"GitHub error {resp.status_code}: {msg}"


def delete_github_template(account_id: str) -> tuple[bool, str]:
    """Delete a template file from the GitHub repo."""
    if not github_configured():
        return False, "GitHub not configured."

    sha = get_file_sha(account_id)
    if not sha:
        return False, f"Template for {account_id} not found in GitHub."

    filename = f"templates/{account_id}.xlsx"
    payload = {
        "message": f"Delete template for account {account_id}",
        "sha": sha,
    }
    resp = requests.delete(
        _api(filename),
        headers={**_headers(), "Content-Type": "application/json"},
        json=payload,
        timeout=15,
    )
    if resp.ok:
        return True, f"Deleted template for account {account_id}."
    return False, f"GitHub error {resp.status_code}: {resp.text}"


# ── SYNC SESSION STATE ────────────────────────────────────────────────────────

TEMPLATE_CACHE_KEY = "gh_template_cache"


@st.cache_data(ttl=300, show_spinner=False)
def _cached_template(account_id: str, _repo: str) -> Optional[bytes]:
    """Cache template downloads for 5 minutes to avoid hammering GitHub API."""
    return load_github_template(account_id)


def get_template_cached(account_id: str) -> Optional[bytes]:
    """
    Get a template — checks session cache first, then GitHub.
    Uses st.cache_data with a 5-minute TTL.
    """
    if not github_configured():
        return None
    return _cached_template(account_id, _repo())


def invalidate_cache():
    """Call after saving/deleting to force fresh download next time."""
    _cached_template.clear()
