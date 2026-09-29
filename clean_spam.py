import msal
import requests
import os
import sys

# Client ID and Refresh Token can be passed as environment variables (e.g. via GitHub Secrets)
CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "YOUR_CLIENT_ID_HERE")
REFRESH_TOKEN_ENV = os.environ.get("OUTLOOK_REFRESH_TOKEN", "")

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Mail.ReadWrite"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.bin")

def get_token_cache():
    """Loads the token cache from disk if it exists."""
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_token_cache(cache):
    """Saves the token cache to disk if it has changed."""
    if cache.has_state_changed and os.path.exists(os.path.dirname(CACHE_FILE)):
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def get_access_token():
    cache = get_token_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    # 1. Check if REFRESH_TOKEN was passed via environment variable (CI/GitHub Actions)
    if REFRESH_TOKEN_ENV:
        print("Using OUTLOOK_REFRESH_TOKEN from environment to authenticate...")
        result = app.acquire_token_by_refresh_token(REFRESH_TOKEN_ENV, scopes=SCOPES)
        if "access_token" in result:
            save_token_cache(cache)
            return result["access_token"]
        else:
            print(f"Error acquiring token using environment refresh token: {result.get('error_description')}")

    # 2. Attempt silent authentication using cached token/refresh token on disk
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_token_cache(cache)
            return result["access_token"]

    # 3. If running in CI (non-interactive) without valid tokens, fail early
    if os.environ.get("CI") or os.environ.get("GITHUB_ACTIONS"):
        raise RuntimeError("Missing or expired OUTLOOK_REFRESH_TOKEN secret in GitHub Actions environment.")

    # 4. Fallback for local run: Interactive Device Code flow
    print("No valid cached session found. Initializing device login...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError(f"Failed to create device flow: {flow.get('error_description')}")

    print("\n" + "=" * 60)
    print(flow["message"])
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_token_cache(cache)
        print("Authentication successful! Token cached for future runs.")
        return result["access_token"]
    else:
        raise ValueError(f"Failed to get token: {result.get('error_description')}")

def extract_sender_address(msg):
    """Extracts email address checking both 'from' and 'sender' fields."""
    for field in ('from', 'sender'):
        addr = msg.get(field, {}).get('emailAddress', {}).get('address', '')
        if addr:
            return addr.strip().strip('<>').strip()
    return ""

def is_no_tld_spam(sender_email):
    """
    Returns True if sender has no valid TLD after the @ symbol.
    Examples:
      - 'user@junkword' -> True (no dot after @)
      - 'user@domain.' -> True (trailing dot without TLD)
      - 'user@domain.com' -> False (valid TLD)
      - 'first.last@junkword' -> True (dot is before @, not after)
      - '' or 'nodomain' -> True (malformed sender)
    """
    if not sender_email:
        return True # Malformed / missing sender

    if '@' not in sender_email:
        return True # No @ symbol

    domain = sender_email.split('@')[-1].strip()

    # Check if domain has no dot, or dot is at the very beginning/end
    if '.' not in domain:
        return True

    parts = domain.split('.')
    if not parts[-1]: # e.g. domain ended with a dot
        return True

    return False

def clean_no_tld_spam():
    if CLIENT_ID == "YOUR_CLIENT_ID_HERE":
        print("Error: OUTLOOK_CLIENT_ID must be configured as an environment variable or secret.")
        sys.exit(1)

    print("Authenticating...")
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Use Microsoft Graph's well-known folder alias 'junkemail' directly
    print("Connecting to Junk folder...")
    url = f"{GRAPH_ENDPOINT}/me/mailFolders/junkemail/messages?$select=id,from,sender,subject&$top=100"

    all_messages = []
    print("Fetching message list from Junk Email...")

    # PHASE 1: Fetch ALL messages first without modifying/deleting them.
    page_num = 1
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"Error fetching messages: {resp.status_code} - {resp.text}")
            sys.exit(1)

        data = resp.json()
        messages_in_page = data.get('value', [])
        all_messages.extend(messages_in_page)
        print(f"  Page {page_num}: retrieved {len(messages_in_page)} messages...")

        url = data.get('@odata.nextLink')
        page_num += 1

    print(f"\nTotal messages in Junk Email: {len(all_messages)}")

    # PHASE 2: Filter matching spam
    to_delete = []
    ignored = []

    for msg in all_messages:
        sender = extract_sender_address(msg)
        subject = msg.get('subject', '<No Subject>')
        msg_id = msg.get('id')

        if is_no_tld_spam(sender):
            to_delete.append((msg_id, sender, subject))
        else:
            ignored.append((sender, subject))

    print(f"Messages identified without valid TLD: {len(to_delete)}")
    print(f"Messages kept (have valid TLD): {len(ignored)}\n")

    if not to_delete:
        print("No matching spam messages to delete.")
        return

    # PHASE 3: Delete the matching messages
    print(f"Deleting {len(to_delete)} messages...")
    deleted_count = 0
    for msg_id, sender, subject in to_delete:
        print(f"  [DELETING] {sender} | Subject: {subject}")
        del_url = f"{GRAPH_ENDPOINT}/me/messages/{msg_id}"
        del_resp = requests.delete(del_url, headers=headers)

        if del_resp.status_code == 204:
            deleted_count += 1
        else:
            print(f"    Failed to delete message ID {msg_id}: {del_resp.status_code}")

    print(f"\nSuccessfully deleted {deleted_count} messages.")

if __name__ == "__main__":
    clean_no_tld_spam()
