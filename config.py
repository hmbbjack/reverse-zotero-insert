#!/usr/bin/env python3
"""
Config for the "逆向 Zotero 插入" skill.

Design goal: the code contains NO hardcoded personal identifiers
(userID, emails, URIs). All personal values come from the environment.

Sources, in priority order:
  1. Environment variables (set globally, or via config.local.env / .env in this dir).
  2. Auto-derivation from ~/Zotero/zotero.sqlite (users table) - works when the
     DB is present and readable.
  3. Placeholder "<YOUR_USER_ID>" with a warning.
"""
import os
import shutil
import sqlite3
import tempfile

# 可选本地配置文件名（KEY=VALUE）；不提交到公开仓库
_LOCAL_ENV_NAMES = ("config.local.env", ".env")


def _load_local_env():
    """Load KEY=VALUE lines from an optional local .env next to this file."""
    here = os.path.dirname(os.path.abspath(__file__))
    for name in _LOCAL_ENV_NAMES:
        path = os.path.join(here, name)
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#") or "=" not in line:
                            continue
                        key, _, val = line.partition("=")
                        os.environ.setdefault(key.strip(), val.strip())
            except OSError:
                pass
            return


_load_local_env()


def _derive_user_id_from_sqlite():
    """Read userID from ~/Zotero/zotero.sqlite users table.

    Zotero holds an exclusive lock on the DB while it runs, so copy a snapshot
    before opening it. Returns the userID as str, or None.
    """
    db = os.path.expanduser("~/Zotero/zotero.sqlite")
    if not os.path.exists(db):
        return None
    tmp = tempfile.mktemp(suffix=".sqlite")
    try:
        shutil.copy2(db, tmp)
        conn = sqlite3.connect(tmp)
        try:
            cur = conn.cursor()
            cur.execute("SELECT userID FROM users LIMIT 1")
            row = cur.fetchone()
        finally:
            conn.close()
        return str(row[0]) if row else None
    except Exception:
        return None
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def get_user_id():
    """Return the Zotero userID (str). Placeholder if unknown."""
    v = os.environ.get("ZOTERO_USER_ID")
    if v:
        return v.strip()
    derived = _derive_user_id_from_sqlite()
    if derived:
        return derived
    return "<YOUR_USER_ID>"


def get_uri_prefix():
    """Return the item URI prefix, e.g. http://zotero.org/users/<userID>/items/"""
    v = os.environ.get("ZOTERO_URI_PREFIX")
    if v:
        return v.strip()
    uid = get_user_id()
    if uid.startswith("<"):
        return "http://zotero.org/users/<YOUR_USER_ID>/items/"
    return f"http://zotero.org/users/{uid}/items/"


def get_mcp_url():
    """Return the Zotero MCP server URL (Streamable HTTP endpoint)."""
    return os.environ.get("ZOTERO_MCP_URL", "http://127.0.0.1:23120/mcp").strip()


def get_crossref_mailto():
    """Return the CrossRef polite-pool mailto identifier, or None."""
    v = os.environ.get("CROSSREF_MAILTO", "").strip()
    return v if v else None


def get_ncbi_api_key():
    """Return the NCBI eutils API key, or None."""
    v = os.environ.get("NCBI_API_KEY", "").strip()
    return v if v else None


def is_personal():
    """True if a real user identity is configured (env var or derivable DB)."""
    if os.environ.get("ZOTERO_USER_ID"):
        return True
    return _derive_user_id_from_sqlite() is not None


def user_id_configured():
    """True if the userID is a real value, not the placeholder."""
    return not get_user_id().startswith("<")