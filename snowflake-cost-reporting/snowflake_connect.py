"""
Connect to a Snowflake account and pull a sample query into pandas.

Required environment variables (no defaults are baked in):
    SNOWFLAKE_ACCOUNT   account identifier, i.e. the part of the URL before
                        ".snowflakecomputing.com"
                        (https://abc12345.us-east-1.snowflakecomputing.com -> "abc12345.us-east-1")
    SNOWFLAKE_USER      your Snowflake login (often your work email)

Optional:
    SNOWFLAKE_PASSWORD  only needed for password auth. If it is NOT set, the script
                        uses browser-based SSO (authenticator="externalbrowser"),
                        which is what most corporate Okta/Azure AD setups use.
    SNOWFLAKE_WAREHOUSE / SNOWFLAKE_DATABASE / SNOWFLAKE_SCHEMA / SNOWFLAKE_ROLE
    SNOWFLAKE_QUERY     SQL to run (defaults to a session-info query)

Run:
    py snowflake_connect.py
"""

import os
import sys
from urllib.parse import urlparse

import snowflake.connector


# No hard-coded account or user: both must come from the environment.
DEFAULT_ACCOUNT = ""
DEFAULT_USER = ""


def _account_from_env() -> str:
    """Accept a bare account identifier or any Snowflake URL (path/query ignored)."""
    raw = os.environ.get("SNOWFLAKE_ACCOUNT", DEFAULT_ACCOUNT).strip()
    if not raw:
        sys.exit("SNOWFLAKE_ACCOUNT is not set.")
    host = urlparse(raw if "://" in raw else f"https://{raw}").hostname or raw
    return host.removesuffix(".snowflakecomputing.com")


def get_connection():
    user = os.environ.get("SNOWFLAKE_USER", DEFAULT_USER).strip()
    if not user:
        sys.exit("SNOWFLAKE_USER is not set.")

    params = {
        "account": _account_from_env(),
        "user": user,
        "warehouse": os.environ.get("SNOWFLAKE_WAREHOUSE"),
        "database": os.environ.get("SNOWFLAKE_DATABASE"),
        "schema": os.environ.get("SNOWFLAKE_SCHEMA"),
        "role": os.environ.get("SNOWFLAKE_ROLE"),
    }
    password = os.environ.get("SNOWFLAKE_PASSWORD")
    if password:
        params["password"] = password
    else:
        # Opens a browser window for SSO login; cache the token so later runs
        # don't pop the browser every time (if the account allows it)
        params["authenticator"] = "externalbrowser"
        params["client_store_temporary_credential"] = True

    # Drop unset optional values so Snowflake uses the user's defaults
    params = {k: v for k, v in params.items() if v}
    return snowflake.connector.connect(**params)


def main():
    query = os.environ.get(
        "SNOWFLAKE_QUERY",
        "SELECT CURRENT_USER(), CURRENT_ROLE(), CURRENT_WAREHOUSE(), "
        "CURRENT_DATABASE(), CURRENT_SCHEMA(), CURRENT_VERSION()",
    )

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute(query)

        # Option 1: straight into a pandas DataFrame
        df = cur.fetch_pandas_all()
        print(df.head())

        # Option 2: plain Python rows
        # rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
