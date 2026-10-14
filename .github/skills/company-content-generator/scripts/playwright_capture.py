"""Thin Playwright-for-Python capture adapter.

OfficeCLI remains the sole document authoring/HTML-export/formula engine.
This module owns exactly one responsibility: turning OfficeCLI's own static
HTML export into bound, checked screenshot pixels using a managed Chromium
build. It never authors, edits, or exports Office documents itself, and it
never invents OfficeCLI subprocess evidence.

Importing this module does not require Playwright to be installed; the
dependency is only loaded lazily inside `CaptureEngine.__init__`, so default
unit checks (and any caller that never opens an engine) keep working without
Playwright. The capture path itself is optional/opt-in, exercised only by
integration tests and the real render/preflight path when a capability
probe has been explicitly requested.
"""

import hashlib
from pathlib import Path


class CaptureError(Exception):
    """Raised for any capture precondition/coverage failure. Callers (the
    runner) translate this into company.Invalid; this module stays free of
    the business-content validation vocabulary so it can be reused/tested
    in isolation."""


def file_sha256(path):
    path = Path(path)
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest() if hasattr(
            hashlib, "file_digest"
        ) else hashlib.sha256(stream.read()).hexdigest()


def _require(condition, message):
    if not condition:
        raise CaptureError(message)


def _import_playwright():
    try:
        import playwright  # noqa: F401
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise CaptureError(
            "Playwright capture adapter unavailable; install the pinned Playwright "
            "dependency and its matching managed Chromium build before capturing"
        ) from exc
    return playwright, sync_playwright


class CaptureEngine:
    """Owns exactly one managed Chromium browser process. Reused across every
    HTML export captured in a run so setup happens once, not per document."""

    def __init__(self, channel="chromium", headless=True, chromium_sandbox=True, timeout_ms=20000):
        playwright_module, sync_playwright = _import_playwright()
        self._cm = sync_playwright()
        self._playwright = self._cm.__enter__()
        try:
            self.browser = self._playwright.chromium.launch(
                channel=channel, headless=headless, chromium_sandbox=chromium_sandbox)
        except Exception as exc:
            self._cm.__exit__(type(exc), exc, exc.__traceback__)
            raise CaptureError(f"Managed Chromium launch failed: {exc}") from exc
        self.timeout_ms = timeout_ms
        executable = self._playwright.chromium.executable_path
        self.identity = {
            "engine": "playwright-chromium",
            "playwright_version": getattr(playwright_module, "__version__", None) or _package_version("playwright"),
            "channel": channel,
            "headless": headless,
            "chromium_sandbox": chromium_sandbox,
            "browser_version": self.browser.version,
            "executable_path": str(executable),
            "executable_sha256": file_sha256(executable),
        }
        _require(_text(self.identity["playwright_version"]), "Playwright version identity unavailable")

    def close(self):
        try:
            self.browser.close()
        finally:
            self._cm.__exit__(None, None, None)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def _open(self, html_path, viewport):
        context = self.browser.new_context(viewport=viewport)
        page = context.new_page()
        page.set_default_timeout(self.timeout_ms)
        page.goto(Path(html_path).resolve().as_uri())
        page.evaluate("document.fonts ? document.fonts.ready.then(() => true) : true")
        page.wait_for_load_state("networkidle")
        return context, page

    def capture_docx_pages(self, html_path, out_dir, expected_pages, viewport=None):
        """Capture every div.page[data-page='N'] for N in 1..expected_pages.
        Rejects natural-flow fixtures: requires the rendered page-container
        count to match the explicit expectation and every page box to be free
        of vertical DOM overflow (clipped/overflowing text is a failure, not
        an accepted partial page)."""
        _require(isinstance(expected_pages, int) and expected_pages > 0,
                 "DOCX capture needs an explicit positive expected page count")
        context, page = self._open(html_path, viewport or {"width": 900, "height": 1200})
        try:
            rendered = page.locator("div.page").count()
            _require(rendered == expected_pages,
                     f"DOCX page coverage mismatch: expected {expected_pages} explicit page(s), "
                     f"rendered {rendered}; natural-flow overflow is rejected, not accepted")
            out_dir = Path(out_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            results = []
            for number in range(1, expected_pages + 1):
                selector = f"div.page[data-page='{number}']"
                locator = page.locator(selector)
                _require(locator.count() == 1, f"DOCX page {number}: missing/duplicate page selector")
                locator.first.scroll_into_view_if_needed()
                overflow = bool(locator.first.evaluate("el => el.scrollHeight > el.clientHeight + 1"))
                out_png = out_dir / f"page-{number}.png"
                locator.first.screenshot(path=str(out_png))
                results.append({"page": number, "path": str(out_png), "selector": selector,
                                "overflow": overflow, "sha256": file_sha256(out_png)})
            overflowing = [r["page"] for r in results if r["overflow"]]
            _require(not overflowing,
                     f"DOCX page(s) {overflowing} overflow their rendered box; redistribute content "
                     "within the planned page budget instead of accepting clipped/dropped text")
            return {"rendered_pages": rendered, "pages": results}
        finally:
            context.close()

    def capture_pptx_slides(self, html_path, out_dir, expected_slides, viewport=None):
        """Capture every visible div.slide-container[data-slide='N']; hidden
        navigation duplicates are excluded via the `:visible` filter."""
        _require(isinstance(expected_slides, int) and expected_slides > 0,
                 "PPTX capture needs an explicit positive expected slide count")
        context, page = self._open(html_path, viewport or {"width": 1280, "height": 720})
        try:
            rendered = page.locator("div.slide-container").count()
            _require(rendered == expected_slides,
                     f"PPTX slide coverage mismatch: expected {expected_slides}, rendered {rendered}")
            out_dir = Path(out_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            results = []
            for number in range(1, expected_slides + 1):
                selector = f"div.slide-container[data-slide='{number}']:visible"
                locator = page.locator(selector)
                _require(locator.count() == 1, f"PPTX slide {number}: missing/duplicate visible slide selector")
                locator.first.scroll_into_view_if_needed()
                out_png = out_dir / f"slide-{number}.png"
                locator.first.screenshot(path=str(out_png))
                results.append({"slide": number, "path": str(out_png), "selector": selector,
                                "sha256": file_sha256(out_png)})
            return {"rendered_slides": rendered, "slides": results}
        finally:
            context.close()

    def capture_xlsx_region(self, html_path, out_png, sheet_name, start_cell, end_cell,
                            sheet_index=None, viewport=None):
        """Capture a bounded worksheet region using td[data-path='/Sheet/A1']
        corner cells, switching the active sheet tab first when needed."""
        context, page = self._open(html_path, viewport or {"width": 1600, "height": 1200})
        try:
            if sheet_index is not None:
                tab = page.locator(f"div.sheet-tab[data-sheet='{sheet_index}']")
                _require(tab.count() == 1, f"xlsx sheet tab {sheet_index} missing")
                tab.first.click()
                page.wait_for_selector(f"div.sheet-content[data-sheet='{sheet_index}'].active")
            prefix = f"/{sheet_name}"
            start = page.locator(f"td[data-path='{prefix}/{start_cell}']")
            end = page.locator(f"td[data-path='{prefix}/{end_cell}']")
            _require(start.count() >= 1, f"xlsx region start cell {prefix}/{start_cell} not emitted "
                                         "(likely merged away; pass the merge anchor cell instead)")
            _require(end.count() >= 1, f"xlsx region end cell {prefix}/{end_cell} not emitted "
                                       "(likely merged away; pass the merge anchor cell instead)")
            start_box = start.first.bounding_box()
            end_box = end.first.bounding_box()
            _require(start_box and end_box, f"xlsx region {prefix}/{start_cell}:{end_cell} not rendered")
            clip = {
                "x": min(start_box["x"], end_box["x"]),
                "y": min(start_box["y"], end_box["y"]),
                "width": max(start_box["x"] + start_box["width"], end_box["x"] + end_box["width"])
                         - min(start_box["x"], end_box["x"]),
                "height": max(start_box["y"] + start_box["height"], end_box["y"] + end_box["height"])
                          - min(start_box["y"], end_box["y"]),
            }
            out_png = Path(out_png)
            out_png.parent.mkdir(parents=True, exist_ok=True)
            page.screenshot(path=str(out_png), clip=clip)
            return {"sheet": sheet_name, "start_cell": start_cell, "end_cell": end_cell,
                    "clip": clip, "path": str(out_png), "sha256": file_sha256(out_png)}
        finally:
            context.close()


def _text(value):
    return isinstance(value, str) and bool(value.strip())


def _package_version(name):
    try:
        from importlib import metadata
        return metadata.version(name)
    except Exception:
        return None
