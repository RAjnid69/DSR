import os
import sys
import time
import re
import subprocess
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# Path helpers — must work both in dev (python dsr_pro.py) and when frozen
# inside a PyInstaller bundle (C:\Program Files\Agro DSR Professional\)
# ---------------------------------------------------------------------------

def _get_appdata_dir():
    """Return a writable per-user folder regardless of where the exe lives."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "Agro DSR Pro")
    os.makedirs(folder, exist_ok=True)
    return folder

# Session is stored in %APPDATA%\Agro DSR Pro\whatsapp_session  (always writable)
USER_DATA_DIR   = os.path.join(_get_appdata_dir(), "whatsapp_session")
# Exports land in %APPDATA%\Agro DSR Pro\chat export
EXPORT_BASE_DIR = os.path.join(_get_appdata_dir(), "chat export")

os.makedirs(USER_DATA_DIR,   exist_ok=True)
os.makedirs(EXPORT_BASE_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Playwright browser path — Playwright stores browsers in
#   %LOCALAPPDATA%\ms-playwright   (installed once via `playwright install chromium`)
# We must tell Playwright explicitly where to find them when running from an exe.
# ---------------------------------------------------------------------------
_MS_PLAYWRIGHT = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "ms-playwright"
)
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _MS_PLAYWRIGHT

# Number of upward scroll passes per day we go back
SCROLLS_PER_DAY = 5


def ensure_chromium_installed():
    """
    Run `playwright install chromium` if the browser folder is missing.
    This is a one-time setup step.
    """
    # Check if any chromium folder exists under ms-playwright
    if os.path.isdir(_MS_PLAYWRIGHT):
        for entry in os.listdir(_MS_PLAYWRIGHT):
            if "chromium" in entry.lower():
                return True   # already installed
    return False   # caller should offer to install


class WhatsAppExporter:
    def __init__(self, status_callback=None):
        self.status_callback = status_callback

    def log(self, message):
        print(message)
        if self.status_callback:
            self.status_callback(message)

    def install_chromium(self):
        """Run playwright install chromium and stream output to the log callback."""
        self.log("Installing Chromium browser — please wait...")
        try:
            # Find the playwright executable
            playwright_exe = os.path.join(
                os.path.dirname(sys.executable), "Scripts", "playwright.exe"
            )
            if not os.path.exists(playwright_exe):
                playwright_exe = "playwright"   # hope it's on PATH

            proc = subprocess.Popen(
                [playwright_exe, "install", "chromium"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )
            for line in proc.stdout:
                self.log(line.strip())
            proc.wait()
            if proc.returncode == 0:
                self.log("Chromium installed successfully!")
                return True
            else:
                self.log(f"Install failed (exit {proc.returncode}). Run manually: playwright install chromium")
                return False
        except Exception as e:
            self.log(f"Could not auto-install: {e}")
            self.log("Please open a terminal and run:  playwright install chromium")
            return False

    def run_export(self, group_name, days_back=0):
        """
        Export messages from a WhatsApp group.
        days_back=0  -> Today
        days_back=1  -> Yesterday
        days_back=2..5 -> N days ago
        """
        # --- Auto-install Chromium if missing ---
        if not ensure_chromium_installed():
            self.log("Chromium not found. Attempting auto-install...")
            if not self.install_chromium():
                return None

        target_date = datetime.now() - timedelta(days=days_back)
        target_date_str = target_date.strftime("%d/%m/%Y")
        date_label = self._label_for_days_back(days_back)

        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                self.log("Launching browser...")
                context = p.chromium.launch_persistent_context(
                    USER_DATA_DIR,
                    headless=False,
                    viewport={'width': 1280, 'height': 800},
                    # Prevent "restore session" popups that block WhatsApp
                    args=["--no-first-run", "--no-default-browser-check",
                          "--disable-default-apps", "--disable-infobars"]
                )
                page = context.new_page()
                self.log("Opening WhatsApp Web...")
                page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")

                # --- Wait for login ---
                self.log("Waiting for WhatsApp to load (scan QR code if prompted)...")
                try:
                    page.wait_for_selector(
                        'div[contenteditable="true"][data-tab="3"], '
                        'input[aria-label="Search or start a new chat"], '
                        '#pane-side',
                        timeout=60000
                    )
                    self.log("Logged in!")
                except Exception:
                    self.log("Action Required: Please scan the QR code in the browser window.")
                    try:
                        page.wait_for_selector(
                            'div[contenteditable="true"][data-tab="3"], #pane-side',
                            timeout=300000   # 5-minute QR scan window
                        )
                        self.log("QR scanned — logged in!")
                    except Exception:
                        self.log("Login timed out. Please try again.")
                        context.close()
                        return None

                # --- Navigate to group ---
                self.log(f"Searching for group: {group_name}")
                try:
                    search_box = page.locator(
                        'div[contenteditable="true"][data-tab="3"], '
                        'input[aria-label="Search or start a new chat"]'
                    ).first
                    search_box.wait_for(timeout=10000)
                    search_box.click()
                    search_box.fill(group_name)
                    time.sleep(1.5)
                    page.keyboard.press("Enter")
                    time.sleep(2)
                except Exception as e:
                    self.log(f"Could not find search box: {e}")
                    context.close()
                    return None

                # Verify we landed in the right chat
                try:
                    page.wait_for_selector('header[data-testid="conversation-header"]', timeout=12000)
                    header_text = page.locator('header[data-testid="conversation-header"]').inner_text()
                    if group_name.lower() not in header_text.lower():
                        # Try clicking the first search result
                        try:
                            page.locator(f'span[title="{group_name}"]').first.click()
                            time.sleep(2)
                        except Exception:
                            self.log(f"Error: Could not open chat '{group_name}'. Make sure the name matches exactly.")
                            context.close()
                            return None
                except Exception:
                    self.log(f"Error: Chat '{group_name}' not found after search.")
                    context.close()
                    return None

                self.log(f"Opened chat: {group_name}")

                # --- Scroll back far enough ---
                total_scroll_passes = max(4, SCROLLS_PER_DAY * (days_back + 1))
                self.log(f"Loading {date_label} messages (scrolling {total_scroll_passes}x)...")
                for i in range(total_scroll_passes):
                    page.mouse.wheel(0, -5000)
                    time.sleep(0.6)
                    if (i + 1) % 5 == 0:
                        self.log(f"  Scrolling... ({i + 1}/{total_scroll_passes})")
                time.sleep(1.5)

                export_folder = os.path.join(EXPORT_BASE_DIR, f"Chat Export {group_name}")
                os.makedirs(export_folder, exist_ok=True)

                export_lines = []
                media_count  = 1

                rows = page.locator('div[role="row"]').all()
                self.log(f"Found {len(rows)} rows. Filtering for {date_label} ({target_date_str})...")

                if days_back == 0:
                    target_messages = self._extract_by_date_marker(rows, "TODAY")
                elif days_back == 1:
                    target_messages = self._extract_by_date_marker(rows, "YESTERDAY")
                else:
                    target_messages = self._extract_by_exact_date(rows, target_date_str, target_date)

                # Fallback
                if not target_messages:
                    self.log(f"Date-marker not found, re-scanning by date string...")
                    target_messages = self._extract_by_exact_date(rows, target_date_str, target_date)

                if not target_messages:
                    self.log(f"No messages found for {date_label}. "
                             "Try scrolling further back by choosing a larger 'days ago' value.")
                    context.close()
                    return None

                self.log(f"Processing {len(target_messages)} messages for {date_label}...")

                for msg in target_messages:
                    try:
                        meta_elem = msg.locator('div.copyable-text[data-pre-plain-text]')
                        meta_text = meta_elem.get_attribute('data-pre-plain-text') if meta_elem.count() > 0 else ""

                        if not meta_text:
                            time_elem = msg.locator('div.x1rg5ocr.x12fk6hs').first
                            msg_time  = time_elem.inner_text() if time_elem.count() > 0 else ""
                            if msg_time:
                                meta_text = f"[{msg_time}, {target_date_str}] : "
                            else:
                                continue

                        text_elem = msg.locator('span.selectable-text').first
                        body      = text_elem.inner_text() if text_elem.count() > 0 else ""

                        # Media
                        img_btn = msg.locator('div[aria-label="Open picture"], div[role="button"]:has(img)').first
                        if img_btn.count() > 0:
                            self.log(f"Downloading media {media_count}...")
                            img_btn.click()
                            try:
                                page.wait_for_selector('div[aria-label="Download"]', timeout=5000)
                                with page.expect_download() as dl:
                                    page.locator('div[aria-label="Download"]').click()
                                d        = dl.value
                                ext      = d.suggested_filename.split('.')[-1]
                                fname    = f"IMG-{target_date.strftime('%Y%m%d')}-WA{media_count:04d}.{ext}"
                                d.save_as(os.path.join(export_folder, fname))
                                page.keyboard.press("Escape")
                                time.sleep(0.5)
                                body = f"{fname} (file attached)"
                                media_count += 1
                            except Exception:
                                self.log("Failed to download a media item — skipping.")
                                page.keyboard.press("Escape")

                        if body or "file attached" in body:
                            export_lines.append(f"{meta_text}{body}")
                    except Exception:
                        continue

                if not export_lines:
                    self.log("No text messages extracted.")
                    context.close()
                    return None

                date_suffix = target_date.strftime("%Y%m%d")
                txt_file    = os.path.join(export_folder,
                                           f"WhatsApp Chat with {group_name} [{date_suffix}].txt")
                with open(txt_file, "w", encoding="utf-8") as f:
                    f.write("\n".join(export_lines))

                self.log(f"Sync Complete! {len(export_lines)} msgs → {txt_file}")
                context.close()
                return txt_file

        except Exception as e:
            self.log(f"Sync Error: {e}")
            import traceback
            self.log(traceback.format_exc())
            return None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _label_for_days_back(self, days_back):
        if days_back == 0: return "Today"
        if days_back == 1: return "Yesterday"
        return f"{days_back} Days Ago"

    def _extract_by_date_marker(self, rows, marker_text):
        target_messages = []
        found_marker    = False
        for row in reversed(rows):
            row_text = row.inner_text().upper().strip()
            if marker_text in row_text and len(row_text) < 30:
                found_marker = True
                continue
            if found_marker:
                if self._is_date_separator(row_text):
                    break
                target_messages.insert(0, row)
        return target_messages

    def _extract_by_exact_date(self, rows, target_date_str, target_date):
        target_messages  = []
        in_target_block  = False
        short_yy         = target_date.strftime("%d/%m/%y")

        for row in rows:
            row_text  = row.inner_text().strip()
            row_upper = row_text.upper()

            if self._is_date_separator(row_upper):
                if (target_date_str in row_text or short_yy in row_text):
                    in_target_block = True
                else:
                    if in_target_block:
                        break
                    in_target_block = False
                continue

            if in_target_block:
                if (target_date_str in row_text or short_yy in row_text
                        or not self._row_has_date(row_text)):
                    target_messages.append(row)

        return target_messages

    def _is_date_separator(self, row_upper):
        if len(row_upper) > 40:
            return False
        return bool(
            "TODAY"     in row_upper or
            "YESTERDAY" in row_upper or
            re.search(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b', row_upper) or
            re.search(r'\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b', row_upper)
        )

    def _row_has_date(self, row_text):
        return bool(re.search(r'\d{1,2}/\d{1,2}/\d{2,4}', row_text))


# ------------------------------------------------------------------
# Public API
# ------------------------------------------------------------------

def get_whatsapp_exporter(group_name, days_back=0, status_callback=None):
    exporter = WhatsAppExporter(status_callback)
    return exporter.run_export(group_name, days_back=days_back)


if __name__ == "__main__":
    group = input("Group Name: ")
    days  = int(input("Days back (0=today, 1=yesterday, 2-5=older): ") or "0")
    res   = get_whatsapp_exporter(group, days_back=days)
    print(f"Result: {res}")
