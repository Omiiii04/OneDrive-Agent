#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════╗
║           OneDrive Photo Importer Agent                      ║
║  Downloads every photo from OneDrive → local folder,         ║
║  then deletes each photo from OneDrive after safe download.  ║
╚══════════════════════════════════════════════════════════════╝

QUICK SETUP (one-time, ~5 minutes)
───────────────────────────────────
1. Go to: https://portal.azure.com
2. Search for "App registrations" → click "+ New registration"
3. Name it anything (e.g. "OneDrive Photo Importer")
4. Leave redirect URI blank → click Register
5. On the Overview page, copy the "Application (client) ID" → paste it as CLIENT_ID below
6. Go to "Authentication" → "+ Add a platform" → "Mobile and desktop applications"
   → Check: https://login.microsoftonline.com/common/oauth2/nativeclient → Save
   → Under "Advanced settings", set "Allow public client flows" to YES → Save
7. Go to "API permissions" → "+ Add a permission" → Microsoft Graph
   → Delegated → search "Files.ReadWrite" → check it → "Add permissions"
8. Done! Run the script.

INSTALL DEPENDENCIES
─────────────────────
   pip install msal requests tqdm

USAGE
──────
   python onedrive_photo_importer.py                  # Normal run
   python onedrive_photo_importer.py --dry-run        # Preview only, no changes
   python onedrive_photo_importer.py --no-delete      # Download but keep on OneDrive
   python onedrive_photo_importer.py --folder ~/pics  # Custom destination
   python onedrive_photo_importer.py --resume         # Skip already-downloaded photos
"""

from html import parser
import os
import sys
import json
import time
import logging
import argparse
from pathlib import Path

import requests
import msal
from tqdm import tqdm
import re

# ═══════════════════════════════════════════
#   CONFIGURATION  ← Edit these two lines
# ═══════════════════════════════════════════
CLIENT_ID       = "YOUR_CLIENT_ID_HERE"    # From Azure App Registration (step 5 above)
DOWNLOAD_FOLDER = "./onedrive_photos"      # Where photos will be saved locally
# ═══════════════════════════════════════════

SCOPES = ["Files.ReadWrite"]
GRAPH           = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE     = ".onedrive_token_cache.json"
PROGRESS_FILE   = ".onedrive_progress.json"
LOG_FILE        = "onedrive_importer.log"

# ──────────────── Logging ────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════
#   AUTHENTICATION
# ══════════════════════════════════════════

_msal_app = None

def acquire_token() -> str:
    global _msal_app
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE):
        cache.deserialize(Path(TOKEN_CACHE).read_text(encoding="utf-8"))

    _msal_app = msal.PublicClientApplication(       # ← assign to global
        CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )
    accounts = _msal_app.get_accounts()
    result = None

    if accounts:
        log.info(f"Reusing cached session for: {accounts[0]['username']}")
        result = _msal_app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        log.info("Starting interactive login (device code)…")
        flow = _msal_app.initiate_device_flow(scopes=SCOPES)
        if "message" not in flow:
            raise RuntimeError(f"Could not start device flow: {flow}")
        print("\n" + "─" * 62)
        print(flow["message"])
        print("─" * 62 + "\n")
        result = _msal_app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        err = result.get("error_description", result.get("error", "unknown"))
        raise RuntimeError(f"Authentication failed: {err}")

    Path(TOKEN_CACHE).write_text(cache.serialize(), encoding="utf-8")
    log.info("✓ Authenticated successfully.")
    return result["access_token"]

def get_fresh_token() -> str:
    """
    Silently refresh the access token using cached credentials.
    MSAL checks expiry locally — only makes a network call if token is actually expired.
    No login prompt needed.
    """
    global _msal_app
    accounts = _msal_app.get_accounts()
    if not accounts:
        raise RuntimeError("No cached account — please restart the script to log in again.")

    result = _msal_app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise RuntimeError("Silent token refresh failed — please restart the script.")

    # Persist the refreshed token to cache
    cache = _msal_app.token_cache
    Path(TOKEN_CACHE).write_text(cache.serialize(), encoding="utf-8")

    return result["access_token"]

# ══════════════════════════════════════════
#   GRAPH API HELPERS
# ══════════════════════════════════════════

def _get(token: str, url: str, stream: bool = False) -> requests.Response:
    """GET with automatic retry on rate-limit (HTTP 429)."""
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(6):
        resp = requests.get(url, headers=headers, stream=stream, timeout=60)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 10))
            log.warning(f"Rate-limited — waiting {wait}s before retry {attempt + 1}/6…")
            time.sleep(wait)
            continue
        return resp
    raise RuntimeError("Exceeded retry limit due to rate limiting.")


def _delete(token: str, item_id: str) -> bool:
    """DELETE a drive item. Returns True on success."""
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(4):
        resp = requests.delete(
            f"{GRAPH}/me/drive/items/{item_id}",
            headers=headers,
            timeout=30,
        )
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 5))
            time.sleep(wait)
            continue
        return resp.status_code == 204
    return False


# ══════════════════════════════════════════
#   PHOTO DISCOVERY
# ══════════════════════════════════════════

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp",
    ".heic", ".heif",                          # ← Apple/Samsung format
    ".tiff", ".tif", ".webp",
    ".raw", ".dng", ".cr2", ".nef",            # RAW formats
    ".arw", ".orf", ".rw2",
}

def is_image(item: dict) -> bool:
    """Detect images by facet (fast) OR file extension (fallback for HEIC etc.)"""
    if "image" in item or "photo" in item:
        return True
    ext = Path(item.get("name", "")).suffix.lower()
    return ext in IMAGE_EXTENSIONS

def iter_photos(token: str):
    """
    Recursively walk all OneDrive folders and yield every image/photo item.
    """
    fields = "id,name,size,image,photo,file,folder,parentReference,createdDateTime"

    def walk_folder(folder_id=None):
        if folder_id:
            url = f"{GRAPH}/me/drive/items/{folder_id}/children?$select={fields}&$top=200"
        else:
            url = f"{GRAPH}/me/drive/root/children?$select={fields}&$top=200"

        while url:
            resp = _get(token, url)
            if resp.status_code != 200:
                log.error(f"Failed to list folder (HTTP {resp.status_code}): {resp.text[:200]}")
                break

            data = resp.json()
            for item in data.get("value", []):
                if "folder" in item:
                    yield from walk_folder(item["id"])          # recurse into subfolders
                elif is_image(item) and "file" in item:
                    yield item                                   # it's a photo

            url = data.get("@odata.nextLink")

    yield from walk_folder()

# ══════════════════════════════════════════
#   DOWNLOAD & DELETE
# ══════════════════════════════════════════

def _unique_path(folder: Path, filename: str) -> Path:
    """Return a non-colliding file path (appends _1, _2, … if needed)."""
    dest = folder / filename
    stem, suffix = Path(filename).stem, Path(filename).suffix
    n = 1
    while dest.exists():
        dest = folder / f"{stem}_{n}{suffix}"
        n += 1
    return dest


def download_item(token: str, item: dict, folder: Path) -> Path:
    """Download with automatic retry — re-fetches a fresh URL on each attempt."""
    last_error = None

    for attempt in range(1, 6):  # up to 5 attempts
        try:
            # Always fetch a fresh pre-signed URL — old ones expire in ~15-30 min
            meta_resp = _get(token, f"{GRAPH}/me/drive/items/{item['id']}")
            meta_resp.raise_for_status()
            meta = meta_resp.json()

            download_url = meta.get("@microsoft.graph.downloadUrl")
            if not download_url:
                log.debug(f"No direct URL for {item['name']}, falling back to /content")
                resp = _get(token, f"{GRAPH}/me/drive/items/{item['id']}/content", stream=True)
                resp.raise_for_status()
            else:
                resp = requests.get(download_url, stream=True, timeout=120)
                resp.raise_for_status()

            dest = _unique_path(folder, item["name"])
            with open(dest, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=512 * 1024):
                    fh.write(chunk)

            if dest.stat().st_size == 0:
                dest.unlink()
                raise RuntimeError("Downloaded file is empty.")

            if attempt > 1:
                log.info(f"  ✓ Succeeded on attempt {attempt}: {item['name']}")

            return dest  # ← success, exit retry loop

        except (
            requests.exceptions.ConnectionError,
            requests.exceptions.ChunkedEncodingError,
            requests.exceptions.Timeout,
            ConnectionResetError,
        ) as e:
            last_error = e
            wait = 2 ** attempt  # 2s, 4s, 8s, 16s, 32s
            log.warning(
                f"  ⚠️  Network error on attempt {attempt}/5 for {item['name']} — "
                f"retrying in {wait}s…\n     ({type(e).__name__}: {e})"
            )
            time.sleep(wait)

    raise RuntimeError(f"Failed after 5 attempts: {last_error}")


# ══════════════════════════════════════════
#   PROGRESS / RESUME
# ══════════════════════════════════════════

def load_progress() -> dict:
    if os.path.exists(PROGRESS_FILE):
        try:
            return json.loads(Path(PROGRESS_FILE).read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"done_ids": [], "failed": []}


def save_progress(progress: dict):
    Path(PROGRESS_FILE).write_text(
        json.dumps(progress, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

def photo_date(item: dict) -> str:
    """
    Extract the photo's date as YYYY-MM-DD using 3 fallbacks:
      1. EXIF takenDateTime from the photo facet (most accurate)
      2. Date encoded in the filename  e.g. 20260518_190235.jpg → 2026-05-18
      3. createdDateTime from OneDrive metadata
    """
    # 1. EXIF date from Graph API photo facet
    taken = item.get("photo", {}).get("takenDateTime")
    if taken:
        return taken[:10]

    # 2. Parse date directly from filename (Samsung / Android naming: YYYYMMDD_HHMMSS)
    name = item.get("name", "")
    m = re.match(r"(\d{4})(\d{2})(\d{2})[_\-T]", name)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # 3. Fall back to OneDrive createdDateTime
    return item.get("createdDateTime", "")[:10]

# ══════════════════════════════════════════
#   MAIN
# ══════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Download all OneDrive photos locally and delete them from the cloud."
    )
    parser.add_argument(
        "--folder", default=DOWNLOAD_FOLDER,
        help=f"Destination folder (default: {DOWNLOAD_FOLDER})",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="List photos but do NOT download or delete anything.",
    )
    parser.add_argument(
        "--no-delete", action="store_true",
        help="Download photos but keep them on OneDrive.",
    )
    parser.add_argument(
        "--resume", action="store_true",
        help="Skip photos that were already downloaded in a previous run.",
    )
    parser.add_argument(
        "--limit", type=int, default=None,
        help="Only process this many photos (e.g. --limit 1 to test with a single photo).",
    )
    parser.add_argument(
        "--date", default=None,
        help="Only process photos taken/created on this date. Format: YYYY-MM-DD (e.g. --date 2024-03-21)",
    )
    parser.add_argument(
        "--date-from", default=None, dest="date_from",
        help="Process photos from this date onwards. Format: YYYY-MM-DD",
    )
    parser.add_argument(
        "--date-to", default=None, dest="date_to",
        help="Process photos up to this date. Format: YYYY-MM-DD",
    )
    args = parser.parse_args()

    # Guard against unconfigured script
    if CLIENT_ID == "YOUR_CLIENT_ID_HERE":
        print(
            "\n❌  CLIENT_ID is not set!\n"
            "   Open this file in a text editor and replace YOUR_CLIENT_ID_HERE\n"
            "   with your Azure Application (client) ID.\n"
            "   See the QUICK SETUP section at the top of this file.\n"
        )
        sys.exit(1)

    dest = Path(args.folder).expanduser().resolve()
    dest.mkdir(parents=True, exist_ok=True)

    # Print run summary
    action = "DRY RUN (no changes)" if args.dry_run else (
        "Download only (no delete)" if args.no_delete else
        "Download → Delete from OneDrive"
    )
    print()
    print("╔══════════════════════════════════════════════════╗")
    print("║         OneDrive Photo Importer                  ║")
    print("╚══════════════════════════════════════════════════╝")
    print(f"  📁  Destination : {dest}")
    print(f"  ⚙️   Mode        : {action}")
    print(f"  📄  Log file    : {LOG_FILE}")
    print()

    # Authenticate
    token = acquire_token()

    # Load resume progress
    progress = load_progress() if args.resume else {"done_ids": [], "failed": []}
    done_ids = set(progress["done_ids"])
    if done_ids:
        log.info(f"Resume mode: skipping {len(done_ids)} previously completed photos.")

    # Discover photos
    log.info("Scanning OneDrive for all photos (this may take a moment)…")
    all_photos = []
    for item in iter_photos(token):
        if item["id"] not in done_ids:
            all_photos.append(item)

    # CORRECT ORDER — filter first, then limit
    if args.date:
        all_photos = [p for p in all_photos if photo_date(p) == args.date]
        log.info(f"Date filter ({args.date}): {len(all_photos)} photos match.")

    if args.date_from or args.date_to:
        lo = args.date_from or "0000-00-00"
        hi = args.date_to   or "9999-99-99"
        all_photos = [p for p in all_photos if lo <= photo_date(p) <= hi]
        log.info(f"Date range ({lo} → {hi}): {len(all_photos)} photos match.")

    if args.limit:
        all_photos = all_photos[:args.limit]
        log.info(f"Limit applied: processing {len(all_photos)} photos.")

    total = len(all_photos)
    skipped = len(done_ids)
    log.info(f"Found {total} photos to process  (skipped {skipped} already done).")

    if total == 0:
        print("\n✅  Nothing to do — OneDrive photos are already empty or all imported.")
        return

    # Dry-run: just list
    if args.dry_run:
        print(f"\n{'NAME':<50} {'SIZE':>10}  PATH")
        print("─" * 80)
        total_size = 0
        for item in all_photos:
            size_mb = item.get("size", 0) / 1_000_000
            total_size += item.get("size", 0)
            parent = item.get("parentReference", {}).get("path", "")
            parent = parent.replace("/drive/root:", "")
            print(f"  {item['name']:<48} {size_mb:>8.1f}MB  {parent}")
        print("─" * 80)
        print(f"  Total: {total} photos  |  {total_size / 1_000_000:.1f} MB")
        print("\n(Dry run — nothing downloaded or deleted.)\n")
        return

    # Main import loop
    ok = fail = deleted = 0

    for item in tqdm(all_photos, desc="Importing", unit="photo", ncols=80):
        token = get_fresh_token()   # ← silently refresh before each photo
        name = item["name"]
        size_mb = item.get("size", 0) / 1_000_000
        log.debug(f"Processing: {name}  ({size_mb:.1f} MB)")

        try:
            # ── 1. Download ──────────────────────
            saved_path = download_item(token, item, dest)
            log.debug(f"  ✓ Saved → {saved_path.name}")
            ok += 1

            # ── 2. Delete from OneDrive ──────────
            if not args.no_delete:
                if _delete(token, item["id"]):
                    deleted += 1
                    log.debug(f"  🗑  Deleted from OneDrive: {name}")
                else:
                    log.warning(f"  ⚠️  Downloaded but could NOT delete: {name}")

            # ── 3. Record progress ───────────────
            progress["done_ids"].append(item["id"])
            save_progress(progress)

        except KeyboardInterrupt:
            print("\n\n⚠️  Interrupted by user. Progress has been saved.")
            print(f"   Run with --resume to continue where you left off.\n")
            save_progress(progress)
            sys.exit(0)

        except Exception as exc:
            log.error(f"  ✗ Failed — {name}: {exc}")
            progress["failed"].append({"id": item["id"], "name": name, "error": str(exc)})
            save_progress(progress)
            fail += 1

    # ── Summary ──
    print()
    print("═" * 50)
    print(f"  ✅  Downloaded : {ok}")
    if not args.no_delete:
        print(f"  🗑   Deleted   : {deleted}")
    if fail:
        print(f"  ❌  Failed     : {fail}  (see {PROGRESS_FILE} and {LOG_FILE})")
    print(f"  📁  Saved to   : {dest}")
    print("═" * 50)
    print()

    if fail:
        print(f"  {fail} photo(s) failed. Re-run with --resume to retry them.\n")
    else:
        # Clean up progress file on full success
        if os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)
        print("  All photos imported successfully! Progress file cleaned up.\n")


if __name__ == "__main__":
    main()
