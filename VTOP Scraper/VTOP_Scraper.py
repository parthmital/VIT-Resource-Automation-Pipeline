import os, time, glob, re
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

BASE_DOWNLOAD_ROOT = "D:\\Downloads"
CHROME_DRIVER_PATH = None
CHROME_USER_DATA_DIR = None
CHROME_PROFILE_DIR = None
PAGE_LOAD_TIMEOUT = 20
DOWNLOAD_TIMEOUT = 120
ZIP_EXT = ".zip"
DOWNLOAD_DIR = None


def sanitize_filename(name):
    name = name.strip()
    name = re.sub("[^A-Za-z0-9._ -]+", "_", name)
    name = name.strip()
    if not name:
        name = "name"
    return name


def build_subject_download_dir(driver):
    wait = WebDriverWait(driver, 20)
    code_cell = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//*[@id="getFacultyForCoursePage"]/div[2]/table/tbody/tr[2]/td[3]',
            )
        )
    )
    title_cell = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//*[@id="getFacultyForCoursePage"]/div[2]/table/tbody/tr[2]/td[4]',
            )
        )
    )
    course_code = code_cell.text.strip()
    course_title = title_cell.text.strip()
    folder_name = f"{course_code} {course_title}"
    folder_name = sanitize_filename(folder_name)
    subject_dir = Path(BASE_DOWNLOAD_ROOT) / folder_name
    subject_dir.mkdir(parents=True, exist_ok=True)
    print(f"Subject folder: {subject_dir}")
    return str(subject_dir)


def setup_driver():
    global DOWNLOAD_DIR
    base_dir = Path(BASE_DOWNLOAD_ROOT)
    base_dir.mkdir(parents=True, exist_ok=True)
    options = webdriver.ChromeOptions()
    if CHROME_USER_DATA_DIR:
        options.add_argument(f"--user-data-dir={CHROME_USER_DATA_DIR}")
    if CHROME_PROFILE_DIR:
        options.add_argument(f"--profile-directory={CHROME_PROFILE_DIR}")
    prefs = {
        "download.default_directory": str(base_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    driver_path = CHROME_DRIVER_PATH
    if driver_path and not Path(driver_path).is_file():
        print(
            f"[setup_driver] Warning: chromedriver not found at {driver_path}. Falling back to Selenium Manager."
        )
        driver_path = None
    if driver_path:
        service = Service(executable_path=driver_path)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    DOWNLOAD_DIR = str(base_dir)
    return driver


def set_subject_download_dir(subject_dir):
    global DOWNLOAD_DIR
    DOWNLOAD_DIR = subject_dir
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)
    print(f"DOWNLOAD_DIR for final files: {DOWNLOAD_DIR}")


def wait_for_new_zip(before_files, timeout=DOWNLOAD_TIMEOUT):
    end_time = time.time() + timeout
    before_files = set(before_files)
    base_dir = Path(BASE_DOWNLOAD_ROOT)
    while time.time() < end_time:
        zips = glob.glob(str(base_dir / f"*{ZIP_EXT}"))
        current_files = set(os.path.basename(z) for z in zips)
        new_files = [f for f in current_files if f not in before_files]
        if new_files:
            candidate = new_files[0]
            full_path = base_dir / candidate
            temp_path = str(full_path) + ".crdownload"
            if not os.path.exists(temp_path):
                return full_path
        time.sleep(1)
    raise TimeoutException("Timed out waiting for new ZIP download")


def get_current_zip_files():
    base_dir = Path(BASE_DOWNLOAD_ROOT)
    zips = glob.glob(str(base_dir / f"*{ZIP_EXT}"))
    return set(os.path.basename(z) for z in zips)


def normalize_slot(slot_text):
    s = slot_text.strip()
    if not s:
        return s
    parts = [p.strip() for p in s.split("+") if p.strip()]
    return parts[0] if parts else slot_text.strip()


def process_all_faculties(driver):
    wait = WebDriverWait(driver, 20)
    subject_dir = build_subject_download_dir(driver)
    set_subject_download_dir(subject_dir)
    faculty_selector = '//*[@id="getFacultyForCoursePage"]/div[2]/table/tbody/tr'
    faculties = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, faculty_selector))
    )
    faculty_count = len(faculties)
    print(f"Found {faculty_count} faculty rows (including any header/extra rows)")
    for index in range(faculty_count):
        print("=" * 50)
        print(f"Processing row {index+1}/{faculty_count}")
        faculties = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, faculty_selector))
        )
        fac = faculties[index]
        try:
            print("Row text:", fac.text)
        except StaleElementReferenceException:
            faculties = wait.until(
                EC.presence_of_all_elements_located((By.XPATH, faculty_selector))
            )
            fac = faculties[index]
            print("Row text (after re-find):", fac.text)
        try:
            open_button = fac.find_element(By.XPATH, ".//td[9]/button")
        except Exception:
            print("No td[9]/button in this row, skipping it.")
            continue
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", fac)
        time.sleep(0.5)
        try:
            wait.until(EC.element_to_be_clickable(open_button))
            open_button.click()
        except Exception as e:
            print(f"Normal click on row View button failed ({e}), trying JS click")
            driver.execute_script("arguments[0].click();", open_button)
        name_cell = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="CoursePageLectureDetail"]/div[2]/div/table/tbody/tr[2]/td[7]',
                )
            )
        )
        slot_cell = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="CoursePageLectureDetail"]/div[2]/div/table/tbody/tr[2]/td[6]',
                )
            )
        )
        raw_name = name_cell.text.strip()
        raw_slot = slot_cell.text.strip()
        slot_text = normalize_slot(raw_slot)
        parts = [p.strip() for p in raw_name.split("-")]
        if len(parts) >= 2:
            name_only = parts[1]
        else:
            name_only = raw_name
        combined_label = f"{name_only} {slot_text}".strip()
        safe_name = sanitize_filename(combined_label)
        print(
            f"Raw: {raw_name} | Slot raw: {raw_slot} | Slot norm: {slot_text} -> {combined_label} -> {safe_name}{ZIP_EXT}"
        )
        download_all_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="allMaterialDownload"]'))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", download_all_btn
        )
        time.sleep(0.5)
        before_files = get_current_zip_files()
        try:
            download_all_btn.click()
        except Exception as e:
            print(f"'allMaterialDownload' normal click failed ({e}), trying JS click")
            driver.execute_script("arguments[0].click();", download_all_btn)
        try:
            new_zip_path = wait_for_new_zip(before_files)
            print(f"Downloaded ZIP (raw): {new_zip_path}")
        except TimeoutException:
            print(f"Download timed out for: {raw_name}")
            try:
                back_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="backButton"]'))
                )
                back_btn.click()
            except TimeoutException:
                print("Could not click backButton after timeout.")
            continue
        subject_dir_path = Path(subject_dir)
        target_path = subject_dir_path / (safe_name + ZIP_EXT)
        counter = 1
        while target_path.exists():
            target_path = subject_dir_path / f"{safe_name}_{counter}{ZIP_EXT}"
            counter += 1
        new_zip_path.replace(target_path)
        print(f"Moved & renamed to: {target_path}")
        back_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="backButton"]'))
        )
        back_btn.click()


def main():
    driver = setup_driver()
    try:
        driver.get("https://vtop.vit.ac.in")
        input(
            "Login to VTOP, open the desired subject page (where faculty table is visible), then press Enter here..."
        )
        while True:
            process_all_faculties(driver)
            print("Finished processing this subject.")
            again_same = (
                input("Run again on the current subject page? (y/n): ").strip().lower()
            )
            if again_same == "y":
                continue
            cmd = (
                input(
                    "Open another subject in the same browser, then press Enter to run again, or type q to quit: "
                )
                .strip()
                .lower()
            )
            if cmd == "q":
                break
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
