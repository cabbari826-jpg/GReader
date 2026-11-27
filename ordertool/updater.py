from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
import hashlib
import json
import re
import tempfile


GITHUB_API = "https://api.github.com/repos/cabbari826-jpg/GReader/releases/latest"
USER_AGENT = "GReader-PDF-Order-Validator-Updater"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    tag_name: str
    installer_url: str
    checksum_url: str
    release_url: str
    notes: str


def _version_tuple(value: str):
    numbers = re.findall(r"\d+", str(value))
    return tuple(int(part) for part in numbers[:4]) or (0,)


def _get(url: str, timeout: int = 15):
    request = Request(url, headers={
        "User-Agent": USER_AGENT,
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    try:
        with urlopen(request, timeout=timeout) as response:
            return response.read()
    except (HTTPError, URLError, TimeoutError) as exc:
        raise UpdateError(f"Could not contact GitHub Releases: {exc}") from exc


def check_for_update(current_version: str) -> ReleaseInfo | None:
    payload = json.loads(_get(GITHUB_API).decode("utf-8"))
    tag = str(payload.get("tag_name") or "").strip()
    latest = tag.lstrip("vV")
    if not latest:
        raise UpdateError("The latest GitHub release has no version tag.")
    if _version_tuple(latest) <= _version_tuple(current_version):
        return None

    assets = payload.get("assets") or []
    installer = next((
        asset for asset in assets
        if str(asset.get("name", "")).lower() == "pdf_order_validator_setup.exe"
    ), None)
    checksum = next((
        asset for asset in assets
        if str(asset.get("name", "")).lower() == "pdf_order_validator_setup.exe.sha256"
    ), None)
    if not installer or not checksum:
        raise UpdateError(
            f"Release {tag} does not contain the installer and its SHA-256 checksum."
        )
    return ReleaseInfo(
        version=latest,
        tag_name=tag,
        installer_url=installer["browser_download_url"],
        checksum_url=checksum["browser_download_url"],
        release_url=payload.get("html_url", "https://github.com/cabbari826-jpg/GReader/releases"),
        notes=payload.get("body") or "",
    )


def download_update(release: ReleaseInfo, destination_dir: str | Path | None = None,
                    progress=None) -> Path:
    destination = Path(destination_dir) if destination_dir else Path(tempfile.gettempdir()) / "GReaderUpdates"
    destination.mkdir(parents=True, exist_ok=True)
    installer_path = destination / "PDF_Order_Validator_Setup.exe"
    checksum_data = _get(release.checksum_url).decode("utf-8", errors="replace").strip()
    expected = checksum_data.split()[0].lower() if checksum_data else ""
    if not re.fullmatch(r"[0-9a-f]{64}", expected):
        raise UpdateError("The release checksum file does not contain a valid SHA-256 value.")

    request = Request(release.installer_url, headers={"User-Agent": USER_AGENT})
    digest = hashlib.sha256()
    try:
        with urlopen(request, timeout=60) as response, open(installer_path, "wb") as output:
            total = int(response.headers.get("Content-Length") or 0)
            received = 0
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                output.write(chunk)
                digest.update(chunk)
                received += len(chunk)
                if progress:
                    progress(received, total)
    except Exception as exc:
        installer_path.unlink(missing_ok=True)
        raise UpdateError(f"Could not download the update installer: {exc}") from exc

    actual = digest.hexdigest().lower()
    if actual != expected:
        installer_path.unlink(missing_ok=True)
        raise UpdateError(
            f"Installer SHA-256 verification failed. Expected {expected}, received {actual}. "
            "The installer has been deleted."
        )
    return installer_path
