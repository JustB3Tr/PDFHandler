import json
import urllib.request
import webbrowser

CURRENT_VERSION = "1.0.0"

# Fill these in later when you publish updates
VERSION_URL = "https://raw.githubusercontent.com/JustB3Tr/pdfhandler/main/version.json"
DOWNLOAD_URL = "https://github.com/JustB3Tr/pdfhandler/releases/latest"


def _version_tuple(v: str):
    parts = []
    for piece in v.strip().split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def check_for_updates():
    if not VERSION_URL:
        return None

    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as response:
            data = json.loads(response.read().decode("utf-8"))

        latest = data.get("version", "").strip()
        download_url = data.get("download_url", DOWNLOAD_URL).strip()

        if not latest:
            return {"ok": False, "message": "Update feed did not include a version."}

        has_update = _version_tuple(latest) > _version_tuple(CURRENT_VERSION)

        return {
            "ok": True,
            "has_update": has_update,
            "latest": latest,
            "download_url": download_url,
        }
    except Exception as exc:
        return {"ok": False, "message": str(exc)}


def open_download_page(url: str):
    if url:
        webbrowser.open(url)