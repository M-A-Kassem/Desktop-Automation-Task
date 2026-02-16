import pyautogui
import pyperclip
import os
import sys
import time
import json
import subprocess
import re
from pathlib import Path
from PIL import ImageDraw
# ── Configuration ─────────────────────────────────────────────────────────────
SCRIPT_DIR    = Path(__file__).parent
API_URL       = "https://jsonplaceholder.typicode.com/posts"
SAVE_DIR      = Path.home() / "Desktop" / "tjm-project"
ICON_TEMPLATE = SCRIPT_DIR / "notepad_icon.png"
SCREENSHOTS   = SCRIPT_DIR / "screenshots"
CONFIDENCE    = 0.9

# Move mouse to any screen corner to instantly abort the script
pyautogui.FAILSAFE = True
# ── Visual Grounding (Template Matching) ──────────────────────────────────────

def find_notepad_icon(retries=3):
    """
    Take a fresh screenshot and search it for our Notepad icon template.
    Retries up to 3 times with a 1-second gap between attempts.
    Returns the (x, y) center of the icon, or None if not found.
    """
    for attempt in range(1, retries + 1):
        print(f"  [{attempt}/{retries}] Scanning desktop for Notepad icon...")
        try:
            location = pyautogui.locateOnScreen(
                str(ICON_TEMPLATE), confidence=CONFIDENCE
            )
        except Exception:
            location = None

        if location:
            center = pyautogui.center(location)
            print(f"  Found at ({center.x}, {center.y})")
            return center

        if attempt < retries:
            print("  Not found — retrying in 1s...")
            time.sleep(1)

    print("  Icon not found after all attempts.")
    return None


def get_position_label(x, y, w=1920, h=1080):
    """Auto-label screenshot based on where the icon was detected."""
    if x < w / 3 and y < h / 3:
        return "top_left"
    if x > 2 * w / 3 and y > 2 * h / 3:
        return "bottom_right"
    if w / 3 <= x <= 2 * w / 3 and h / 3 <= y <= 2 * h / 3:
        return "center"
    return "detected"


def save_annotated_screenshot(pos):
    """Save a screenshot with a green box around the detected icon."""
    SCREENSHOTS.mkdir(exist_ok=True)
    screenshot = pyautogui.screenshot()
    draw = ImageDraw.Draw(screenshot)

    x, y = pos.x, pos.y
    draw.rectangle([(x - 35, y - 35), (x + 35, y + 35)], outline="lime", width=3)
    draw.text((x + 40, y - 10), f"Notepad ({x},{y})", fill="lime")

    label = get_position_label(x, y)
    name = f"icon_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
    screenshot.save(str(SCREENSHOTS / name))
    print(f"  Screenshot -> screenshots/{name}")


# ── Notepad Automation ────────────────────────────────────────────────────────

def show_desktop():
    """Win+D to minimize everything and show the desktop."""
    pyautogui.hotkey("win", "d")
    time.sleep(1.5)


def open_notepad():
    """
    Find the Notepad icon via template matching and double-click it.
    Falls back to launching notepad.exe directly if the icon isn't found.
    Validates that Notepad actually started before returning.
    """
    pos = find_notepad_icon()

    if pos:
        save_annotated_screenshot(pos)
        pyautogui.doubleClick(pos.x, pos.y)
    else:
        print("  FALLBACK: Launching notepad.exe directly...")
        subprocess.Popen(["notepad.exe"])

    # Poll for up to 10 seconds to confirm Notepad is running
    for _ in range(20):
        time.sleep(0.5)
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq notepad.exe"],
            capture_output=True, text=True,
        )
        if "notepad.exe" in result.stdout.lower():
            print("  Notepad is running.")
            time.sleep(1)
            return

    raise RuntimeError("Notepad did not start within 10 seconds.")


def paste(text):
    """Copy text to clipboard then Ctrl+V to paste it."""
    pyperclip.copy(text)
    time.sleep(0.15)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)


def save_file(post_id):
    """Ctrl+S -> type path into Save As dialog -> Enter -> handle overwrite."""
    target = str(SAVE_DIR / f"post_{post_id}.txt")

    pyautogui.hotkey("ctrl", "s")
    time.sleep(2)

    # Filename field is focused by default in the Save As dialog.
    # Select whatever is there, paste our full path over it.
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    paste(target)
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.5)

    # If file exists, Windows asks "Replace?" — Alt+Y clicks Yes.
    # Does nothing if the dialog didn't appear.
    pyautogui.hotkey("alt", "y")
    time.sleep(0.5)

    print(f"  Saved post_{post_id}.txt")


def close_notepad():
    """Close Notepad."""
    pyautogui.hotkey("alt", "F4")
    time.sleep(1)


# ── API ───────────────────────────────────────────────────────────────────────

def fetch_posts():
    """Fetch the first 10 posts from JSONPlaceholder API."""
    print("Fetching blog posts from API...")

    try:
        import requests
        resp = requests.get(API_URL, timeout=10)
        resp.raise_for_status()
        posts = resp.json()[:10]
        print(f"  Got {len(posts)} posts from API.\n")
        return posts
    except Exception:
        print("  Direct request failed, trying via Chrome...")

    chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    result = subprocess.run(
        [chrome, "--headless", "--disable-gpu", "--dump-dom", API_URL],
        capture_output=True, text=True, timeout=20,
    )
    match = re.search(r"\[.*\]", result.stdout, re.DOTALL)
    if match:
        posts = json.loads(match.group())[:10]
        print(f"  Got {len(posts)} posts from API.\n")
        return posts

    raise RuntimeError("Failed to fetch posts from API.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 50)
    print("  Notepad Desktop Automation")
    print("=" * 50 + "\n")

    if not ICON_TEMPLATE.is_file():
        sys.exit(f"ERROR: Template image not found: {ICON_TEMPLATE}")

    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    posts = fetch_posts()

    for post in posts:
        pid     = post["id"]
        content = f"Title: {post['title']}\n\n{post['body']}"

        print(f"── Post {pid}/10 ──")

        show_desktop()          # 1. Reveal desktop
        open_notepad()          # 2. Find icon -> double-click (or fallback)
        paste(content)          # 3. Paste the blog post text
        save_file(pid)          # 4. Save as post_{id}.txt
        close_notepad()         # 5. Close and repeat

        print()

    print("=" * 50)
    print(f"  Done — all 10 posts saved to {SAVE_DIR}")
    print("=" * 50)


if __name__ == "__main__":
    main()
