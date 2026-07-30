"""Browser smoke test: the generated dashboard must render without JS errors.

Marked `browser` — run in CI's browser job (or locally with playwright and a
chromium installed; set CHROME_BIN to point at a specific binary).
"""
import os
import subprocess
import sys

import pytest

pytestmark = pytest.mark.browser

playwright_sync = pytest.importorskip("playwright.sync_api")


@pytest.fixture(scope="module")
def dashboard(tmp_path_factory):
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = tmp_path_factory.mktemp("dash") / "d.html"
    subprocess.run([sys.executable, os.path.join(root, "generate_dashboard.py"),
                    os.path.join(root, "sample_data"), "-o", str(out),
                    "--no-datestamp"], check=True)
    return out


def test_dashboard_renders_without_errors(dashboard):
    errors = []
    with playwright_sync.sync_playwright() as p:
        launch = {}
        if os.environ.get("CHROME_BIN"):
            launch["executable_path"] = os.environ["CHROME_BIN"]
        browser = p.chromium.launch(**launch)
        page = browser.new_page()
        page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
        page.on("console",
                lambda m: errors.append(f"console.{m.type}: {m.text}")
                if m.type == "error" else None)
        page.goto(dashboard.as_uri())
        page.wait_for_timeout(600)

        assert page.eval_on_selector_all(".kpi", "els => els.length") >= 10
        assert page.eval_on_selector_all("#peopleTable tbody tr", "els => els.length") == 6
        assert page.eval_on_selector_all("#grantsTable tbody tr", "els => els.length") == 4
        assert page.eval_on_selector_all("#eventsTable tbody tr", "els => els.length") == 7

        # theme filter narrows the people table
        page.eval_on_selector_all(".chip", "els => els[1] && els[1].click()")
        page.wait_for_timeout(200)
        filtered = page.eval_on_selector_all("#peopleTable tbody tr", "els => els.length")
        assert 0 < filtered < 6

        # dark theme toggle doesn't blow up
        page.click("#themeBtn")
        page.wait_for_timeout(200)
        browser.close()
    assert errors == [], f"JS errors: {errors}"
