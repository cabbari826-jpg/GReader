from ordertool.updater import ReleaseInfo, UpdateError, check_for_update
import ordertool.updater as updater


def test_no_published_release_is_reported_clearly(monkeypatch):
    def missing_release(url):
        raise UpdateError("Could not contact GitHub Releases: HTTP Error 404: Not Found")

    monkeypatch.setattr(updater, "_get", missing_release)
    try:
        check_for_update("1.0.1")
    except UpdateError as exc:
        assert "No stable installer release has been published yet" in str(exc)
    else:
        raise AssertionError("A missing stable release must not look like a successful update check")


def test_current_version_is_not_offered_as_an_update(monkeypatch):
    payload = b'{"tag_name":"v1.0.1","assets":[],"html_url":"https://github.com/example/release"}'
    monkeypatch.setattr(updater, "_get", lambda url: payload)
    assert check_for_update("1.0.1") is None


def test_new_release_requires_installer_and_checksum(monkeypatch):
    payload = b'{"tag_name":"v1.0.2","assets":[]}'
    monkeypatch.setattr(updater, "_get", lambda url: payload)
    try:
        check_for_update("1.0.1")
    except UpdateError as exc:
        assert "installer and its SHA-256 checksum" in str(exc)
    else:
        raise AssertionError("Release without installer/checksum must be rejected")
