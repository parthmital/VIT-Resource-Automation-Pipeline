import re
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

URL = "https://unibud.in/VITQuestionBank"
STATE_FILE = Path("unibud_state.json")
DOWNLOAD_DIR = Path.cwd() / "downloads"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
HEADLESS = False
XP_SUBJECT_BUTTON = (
    "/html/body/div[1]/main/div/div/div/aside/div[2]/form/div[1]/div/button"
)
XP_INCLUDE_ANSWERS = (
    "/html/body/div[1]/main/div/div/div/aside/div[2]/form/div[4]/label/div"
)
XP_SEARCH_QUESTIONS = "/html/body/div[1]/main/div/div/div/aside/div[2]/form/button"
XP_GENERATE_PDF = "/html/body/div[1]/main/div/div/div/aside/div[3]/button"
LOC_MODULE_BUTTONS = (
    "xpath=/html/body/div[1]/main/div/div/div/aside/div[2]/form/div[2]/div/button"
)
LOC_Q_CHECKBOXES = "css=div.checkboxContainer input[name='selected_questions']"
LOC_QUESTION_CARD = "css=main[role='main'] div.mb-4.rounded-lg"
LOC_PAGINATION_BAR = "css=main[role='main'] div.flex-shrink-0.p-4.border-t div.flex.justify-center.items-center"
LOC_NEXT_BTN = f"{LOC_PAGINATION_BAR} > button:last-child"


def ask_subject_name():
    while True:
        s = input("Enter subject name (required): ").strip()
        if s:
            return s
        print("Subject name cannot be empty. Try again.")


def safe_click(page, selector, timeout_ms=20000):
    loc = page.locator(selector)
    loc.first.wait_for(state="visible", timeout=timeout_ms)
    loc.first.click()


def ensure_login_state(browser):
    context = browser.new_context()
    page = context.new_page()
    page.goto(URL, wait_until="domcontentloaded")
    print("\nOne-time login required.")
    print("1) Complete login in the opened browser window.")
    print("2) When you can access the Question Bank page, press Enter here.\n")
    input("Press Enter after login is complete...")
    context.storage_state(path=str(STATE_FILE))
    context.close()
    print(f"Saved session to: {STATE_FILE}\n")


def select_subject(page, subject_name):
    safe_click(page, f"xpath={XP_SUBJECT_BUTTON}")
    page.get_by_text(subject_name).first.click(timeout=15000)


def click_include_answers(page):
    try:
        safe_click(page, f"xpath={XP_INCLUDE_ANSWERS}", timeout_ms=5000)
    except PWTimeoutError:
        pass


def click_search_and_wait_for_results(page):
    safe_click(page, f"xpath={XP_SEARCH_QUESTIONS}")
    page.locator(LOC_Q_CHECKBOXES).first.wait_for(state="visible", timeout=20000)


def extract_first_question_id(page):
    card = page.locator(LOC_QUESTION_CARD).first
    card.wait_for(state="visible", timeout=20000)
    txt = card.inner_text() or ""
    m = re.search("Question ID:\\s*(\\d+)", txt)
    return m.group(1) if m else txt[:80].strip()


def wait_until_first_question_changes(page, before_id, timeout_ms=20000):
    deadline = page.evaluate("() => Date.now()") + timeout_ms
    while True:
        if page.evaluate("() => Date.now()") > deadline:
            raise PWTimeoutError("Timed out waiting for next page content to change.")
        after_id = extract_first_question_id(page)
        if after_id and after_id != before_id:
            return
        page.wait_for_timeout(200)


def check_all_questions_on_current_page(page):
    boxes = page.locator(LOC_Q_CHECKBOXES)
    n = boxes.count()
    if n == 0:
        raise RuntimeError("0 question checkboxes found on this page.")
    for i in range(n):
        cb = boxes.nth(i)
        try:
            if not cb.is_checked():
                cb.check()
        except Exception:
            try:
                cb.click()
            except Exception:
                pass


def click_next_if_possible(page):
    next_btn = page.locator(LOC_NEXT_BTN)
    if next_btn.count() == 0:
        return False
    if next_btn.first.get_attribute("disabled") is not None:
        return False
    try:
        next_btn.first.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        next_btn.first.click(timeout=5000)
    except Exception:
        next_btn.first.click(timeout=5000, force=True)
    return True


def paginate_next_until_end(page):
    while True:
        check_all_questions_on_current_page(page)
        before_id = extract_first_question_id(page)
        moved = click_next_if_possible(page)
        if not moved:
            break
        wait_until_first_question_changes(page, before_id, timeout_ms=20000)


def sanitize_filename(s):
    s = s.strip()
    s = re.sub('[\\\\/:*?"<>|]+', "_", s)
    return s if s else "Module"


def get_module_labels(page):
    btns = page.locator(LOC_MODULE_BUTTONS)
    n = btns.count()
    labels = []
    for i in range(n):
        t = (btns.nth(i).inner_text() or "").strip()
        labels.append(t if t else f"Module_{i+1}")
    return labels


def download_pdf_as(page, module_label):
    safe_name = sanitize_filename(module_label)
    target = DOWNLOAD_DIR / f"{safe_name}.pdf"
    with page.expect_download(timeout=60000) as dl_info:
        safe_click(page, f"xpath={XP_GENERATE_PDF}")
    dl_info.value.save_as(str(target))


def run_one_module(page, subject_name, module_idx, module_label):
    page.goto(URL, wait_until="domcontentloaded")
    select_subject(page, subject_name)
    page.locator(LOC_MODULE_BUTTONS).nth(module_idx).click()
    click_include_answers(page)
    click_search_and_wait_for_results(page)
    paginate_next_until_end(page)
    download_pdf_as(page, module_label)


def main():
    subject_name = ask_subject_name()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        if not STATE_FILE.exists():
            ensure_login_state(browser)
        context = browser.new_context(
            accept_downloads=True, storage_state=str(STATE_FILE)
        )
        page = context.new_page()
        page.goto(URL, wait_until="domcontentloaded")
        select_subject(page, subject_name)
        module_labels = get_module_labels(page)
        for idx, label in enumerate(module_labels):
            print(f"Processing: {label}")
            run_one_module(page, subject_name, idx, label)
        context.close()
        browser.close()
        print("Done.")


if __name__ == "__main__":
    main()
