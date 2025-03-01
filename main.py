import time
import tkinter
import tkinter as tk
import csv
import random
# ----- dependencies to pip install -----
# selenium
# win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.remote.webelement as wb
import tkinter.font
import os
from ctypes import windll
import win32com.client

# -----------------------------------------------------------------------
# ----- things you can adjust if you modify the source or resources -----
# -----------------------------------------------------------------------
NUM_PUBLIC_SCHOOLS: int = 107300  # the number of schools in the csv
TEXT_COLOR_ENTERED: str = "black"
TEXT_COLOR_PLACEHOLDER: str = "gray"
TEXT_COLOR_ERROR: str = "red"
PATH_TO_SCHOOL_DATASET: str = "resources/us-public-schools.csv"
DELIMITER_FOR_CSV_FILES: str = ';'
CURRENT_DRIVER_PATH: str = "resources/geckodriver.exe"
if os.name == "not":
    windll.gdi32.AddFontResource("resources/opendyslexic.otf")
OPEN_DYSLEXIC_FONT: tk.font.Font = None

# ---- xpaths ----
EMAIL_INPUT = '//*[@id="email"]'
SCHOOL_INPUT = '//*[@id="location"]'
ZIPCODE_INPUT = '//*[@id="zipcode"]'
DESCRIPTION_INPUT = '//*[@id="description"]'
SUBMIT_BUTTON = '//*[@id="submitButton"]'


class TextWrapper:
    def __init__(self, place_holder: str, element):
        self.PlaceHolder: str = place_holder
        self.Element = element
        self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
        self.Element.insert(tk.END, self.PlaceHolder)

        def on_focused(event) -> None:
            """Clear placeholder text when the user focuses on the Entry."""
            if self.get_stripped() == self.PlaceHolder:
                self.Element.delete("1.0", tk.END)
                self.Element.config(fg=TEXT_COLOR_ENTERED)

        def on_unfocused(event) -> None:
            """Restore placeholder text if Entry is empty."""
            if self.get_stripped() == "":
                self.Element.delete("1.0", tk.END)
                self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
                self.Element.insert(tk.END, self.PlaceHolder)

        self.Element.bind("<FocusIn>", on_focused)
        self.Element.bind("<FocusOut>", on_unfocused)

    def clear(self):
        self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
        self.Element.delete("1.0", tk.END)
        self.Element.insert(tk.END, self.PlaceHolder)

    def get_stripped(self) -> str:
        return self.Element.get("1.0", tk.END).strip()

    def has_user_text(self) -> bool:
        return self.get_stripped() != self.PlaceHolder


class EntryWrapper:
    def __init__(self, place_holder: str, element):
        self.PlaceHolder: str = place_holder
        self.Element = element
        self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
        self.Element.insert(tk.END, self.PlaceHolder)

        def on_focused(event) -> None:
            """Clear placeholder text when the user focuses on the Entry."""
            if self.get_stripped() == self.PlaceHolder:
                self.Element.delete(0, tk.END)
                self.Element.config(fg=TEXT_COLOR_ENTERED)

        def on_unfocused(event) -> None:
            """Restore placeholder text if Entry is empty."""
            if self.get_stripped() == "":
                self.Element.delete(0, tk.END)
                self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
                self.Element.insert(0, self.PlaceHolder)

        self.Element.bind("<FocusIn>", on_focused)
        self.Element.bind("<FocusOut>", on_unfocused)

    def clear(self):
        self.Element.config(fg=TEXT_COLOR_PLACEHOLDER)
        self.Element.delete(0, tk.END)
        self.Element.insert(0, self.PlaceHolder)

    def get_stripped(self) -> str:
        return self.Element.get().strip()

    def has_user_text(self) -> bool:
        return self.get_stripped() != self.PlaceHolder


debug_message_label: tk.Label = None


def debug_error(message: str) -> None:
    debug_message_label.config(text=message,  fg=TEXT_COLOR_ERROR)


def debug_message(message: str) -> None:
    debug_message_label.config(text=message, fg=TEXT_COLOR_ENTERED)


def get_target_from_lnk(shortcut_path) -> str | None:
    try:
        # Check if the shortcut exists
        if not os.path.exists(shortcut_path):
            print(f"Shortcut does not exist: {shortcut_path}")
            return None

        # Create a WScript.Shell COM object
        shell = win32com.client.Dispatch("WScript.Shell")

        # Load the shortcut using the COM object
        shortcut = shell.CreateShortcut(shortcut_path)

        # Get the target of the shortcut (the file or program it points to)
        target = shortcut.Target

        # If the target is a shortcut (another .lnk file), recursively resolve it
        if target.endswith('.lnk') and os.path.exists(target):
            print(f"Found a shortcut, resolving: {target}")
            return get_target_from_lnk(target)

        # If the target is a valid executable, return it
        elif os.path.isfile(target) and target.lower().endswith('.exe'):
            return target
        else:
            print(f"Target is not an executable: {target}")
            return None
    except Exception as e:
        print(f"Error reading shortcut {shortcut_path}: {e}")
        return None


def find_firefox_executable():
    # Expected locations where Firefox might be installed
    expected_locations = [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        r"C:\Firefox\firefox.exe",
    ]
    for location in expected_locations:
        if os.path.exists(location):
            return location

    start_menu_dirs = [
        os.path.join(os.environ["PROGRAMDATA"], "Microsoft", "Windows", "Start Menu", "Programs"),
        os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
    ]
    firefox_shortcut_name = "firefox.lnk"

    for directory in start_menu_dirs:
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower() == firefox_shortcut_name:
                    shortcut_path = os.path.join(root, file)
                    return get_target_from_lnk(shortcut_path)
    return None


if __name__ == "__main__":
    driver: webdriver.Firefox = None
    email_field: wb.WebElement = None
    school_field: wb.WebElement = None
    zipcode_field: wb.WebElement = None
    description_field: wb.WebElement = None
    submit_button: wb.WebElement = None

    root = tk.Tk()
    root.title("Department of Education DEIA Complaint Submission Helper")
    root.config(bg="black", pady=4, padx=4)
    root.geometry("800x800")

    OPEN_DYSLEXIC_FONT = tk.font.Font(family="OpenDyslexic", size=14)

    for i in range(9):
        root.rowconfigure(i, weight=1)
    for i in range(100):
        root.columnconfigure(i, weight=1)

    row_idx = 0

    # make this early for debug reasons
    debug_message_label = tk.Label(root, text="Press 'Load FireFox' Button", font=OPEN_DYSLEXIC_FONT)

    # Firefox install path
    firefox_path_entry = EntryWrapper("Enter the absolute path to your firefox.exe", tk.Entry(root))
    firefox_path_entry.Element.config(font=OPEN_DYSLEXIC_FONT)
    firefox_path_entry.Element.grid(column=0, columnspan=60, row=row_idx, sticky="nsew", padx=4, pady=4)

    path = find_firefox_executable()
    # path = None
    if path is None:
        debug_error("Please set the path to your firefox install")
    else:
        firefox_path_entry.Element.config(fg=TEXT_COLOR_ENTERED, state=tk.NORMAL)
        firefox_path_entry.Element.delete(0, tk.END)
        firefox_path_entry.Element.insert(0, path)
        firefox_path_entry.Element.config(state=tk.DISABLED)

    elements_button: tkinter.Button = None

    def try_load_elements() -> None:
        global email_field, school_field, zipcode_field, description_field, submit_button, submission_button
        elements_button.config(state=tk.DISABLED)
        try:
            email_field = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, EMAIL_INPUT))
            )
            print(type(email_field))
            school_field = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, SCHOOL_INPUT))
            )
            zipcode_field = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, ZIPCODE_INPUT))
            )
            description_field = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, DESCRIPTION_INPUT))
            )
            submit_button = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, SUBMIT_BUTTON))
            )
            elements_button.config(text="Elements Loaded", state=tk.DISABLED)
            submission_button.config(state=tk.NORMAL)
            debug_message("Fill out form info and press 'Submit'")
        except:
            debug_error("Timeout: Press 'Load Elements' when page loads")
            elements_button.config(state=tk.NORMAL)


    firefox_button: tkinter.Button = None

    def try_load_selenium() -> None:
        global driver, root
        options = Options()
        options.headless = True
        options.binary_location = firefox_path_entry.get_stripped()
        if options.binary_location.strip() == "":
            debug_error("Please Enter a path to FireFox in the top left")
            firefox_path_entry.Element.config(state=tk.NORMAL)
            return
        if not os.path.exists(options.binary_location):
            debug_error("Please Enter a path to FireFox in the top left")
            firefox_path_entry.Element.config(state=tk.NORMAL)
            return
        if not options.binary_location.endswith(".exe"):
            debug_error("The file at the end of the path should be an '.exe'")
            firefox_path_entry.Element.config(state=tk.NORMAL)
            return
        if not options.binary_location.lower().endswith("firefox.exe"):
            debug_error("Running non-'firefox.exe' executables is undefined behaviour")
            firefox_path_entry.Element.config(state=tk.NORMAL)
            return

        firefox_button.config(state=tk.DISABLED)
        service = Service(executable_path="resources/geckodriver.exe")
        try:
            driver = webdriver.Firefox(service=service, options=options)
            driver.get("https://enddei.ed.gov/")
            firefox_path_entry.Element.config(state=tk.DISABLED)
            elements_button.config(text="Load Elements", state=tk.NORMAL)
            firefox_button.config(text="Reload FireFox", state=tk.NORMAL)
            reset_button.config(state=tk.NORMAL)
            try_load_elements()
        except:
            debug_error("Invalid FireFox path")
            firefox_path_entry.Element.config(state=tk.NORMAL)
            return

        root.focus_force()


    # load selenium button
    firefox_button = tk.Button(root, text="Load FireFox", command=try_load_selenium, font=OPEN_DYSLEXIC_FONT)
    firefox_button.grid(column=60, columnspan=20, row=row_idx, sticky="nsew", padx=4, pady=4)

    # load elements button
    elements_button = tk.Button(root, text="Load Elements", command=try_load_elements, font=OPEN_DYSLEXIC_FONT)
    elements_button.config(state=tk.DISABLED)
    elements_button.grid(column=80, columnspan=20, row=row_idx, sticky="nsew", padx=4, pady=4)
    row_idx += 1

    # email
    email_entry = EntryWrapper("Enter a real email", tk.Entry(root))
    email_entry.Element.config(font=OPEN_DYSLEXIC_FONT)
    email_entry.Element.grid(column=0, columnspan=100, row=row_idx, sticky="nsew", padx=4, pady=4)
    row_idx += 1

    def get_related_school_row() -> dict:
        with open(PATH_TO_SCHOOL_DATASET, encoding="utf-8") as public_schools_csv:
            schools = csv.DictReader(public_schools_csv, delimiter=DELIMITER_FOR_CSV_FILES)

            index = random.randint(0, NUM_PUBLIC_SCHOOLS)
            for i, row in enumerate(schools):
                if i == index:
                    return row


    # get school or district
    school_label = tk.Label(root)
    school_label.config(state=tk.DISABLED, font=OPEN_DYSLEXIC_FONT)
    school_label.grid(column=0, columnspan=50, row=row_idx, sticky="nsew", padx=4, pady=4)

    # zipcode
    zipcode_label = tk.Label(root)
    zipcode_label.config(state=tk.DISABLED, font=OPEN_DYSLEXIC_FONT)
    zipcode_label.grid(column=50, columnspan=50, row=row_idx, sticky="nsew", padx=4, pady=4)
    row_idx += 1

    def fill_school_info(school_row: dict):
        global school_label, zipcode_label
        school_label.config(text=school_row["NAME"].title())
        zipcode_label.config(text=school_row['ZIP'])

    fill_school_info(get_related_school_row())

    # description
    __description = tk.Text(root, wrap=tk.WORD, font=OPEN_DYSLEXIC_FONT)
    description_entry = TextWrapper("Enter an example of the injustice you've found", __description)
    # description_entry.Element.tag_configure(TEXT_COLOR_ENTERED, foreground=TEXT_COLOR_ENTERED)
    description_entry.Element.config(fg=TEXT_COLOR_ENTERED)
    description_entry.Element.delete("1.0", tk.END)
    description_entry.Element.insert(tk.END,
   "In a small town, there was a school where the teachers, though well-meaning, believed that focusing too much on "
   "diversity and inclusion was the best way to prepare students for the world. Each day, classrooms were filled with "
   "posters and books that celebrated people from all over the world. The teachers spent a lot of time discussing "
   "different cultures, traditions, and histories, often making the students feel as if they needed to understand and "
   "appreciate things that were far removed from their own lives. Every year, the school held an event called "
   "\"Unity Day.\" While it was intended to bring the students together, it often felt like a forced celebration. "
   "Children were asked to share stories about their cultures and traditions, but some of them found it awkward. Not "
   "everyone had something unique to share, and those who didn’t were often left feeling out of place. Teachers would "
   "encourage students to create projects that highlighted cultural differences, but many of the kids didn’t fully "
   "understand why it was important. Lunchtime was another reminder of the school’s focus on diversity. The cafeteria "
   "served foods from various cultures, and while it seemed like a fun idea, many of the students weren’t always "
   "excited about the unfamiliar dishes. Some would refuse to try new foods altogether, while others would feel "
   "uncomfortable if they didn’t know much about the dish they were eating. Teachers would try to explain the cultural "
   "significance of the meals, but it didn’t always resonate with the students. Then there was the 'Buddy Program.' "
   "It was meant to help older students mentor younger ones, but it didn’t always work out as planned. The activities "
   "were intended to foster friendship, but it sometimes felt like a chore for both the older and younger students. "
   "The bonds they formed weren’t always as strong as expected, and many students didn’t feel like they had learned "
   "anything important. To outsiders, the school’s efforts might have seemed like an attempt to promote equality, but "
   "to some, it felt more like an imposed agenda. The emphasis on diversity often overshadowed the natural bonds that "
   "students might have formed through shared interests, rather than their differences. Some students began to "
   "question whether these programs were really about inclusion, or if they were just another set of rules they were "
   "expected to follow, without fully understanding the reason behind them. The school’s approach, though intended to "
   "help, often left students feeling confused and divided, as they struggled to understand what was truly important "
   "in their everyday lives.")

    description_entry.Element.grid(column=0, columnspan=100, row=row_idx, rowspan=4, sticky="nsew", padx=4, pady=4)
    row_idx += 4

    def reset_forum() -> None:
        global email_entry, description_entry, driver
        driver.get("https://enddei.ed.gov/")
        try_load_elements()

    def new_school() -> None:
        fill_school_info(get_related_school_row())

    # submission button
    def submit_info() -> None:
        global driver, email_entry, school_label, zipcode_label, description_entry

        if email_entry.get_stripped() == "" or not email_entry.has_user_text():
            debug_error("Must enter an email")
            return
        if school_label.cget("text").strip() == "":
            debug_error("Must enter a School name")
            return
        if zipcode_label.cget("text").strip() == "":
            debug_error("Must enter a zipcode")
            return
        if description_entry.get_stripped() == "" or not description_entry.has_user_text():
            debug_error("Must enter a description of your very valid complaint")
            return

        try:
            email_field.send_keys(email_entry.get_stripped())
            school_field.send_keys(school_label.cget("text").strip())
            zipcode_field.send_keys(zipcode_label.cget("text").strip())
            description_field.send_keys(description_entry.get_stripped())
            time.sleep(0.25)
            driver.execute_script("arguments[0].click();", submit_button)
        except:
            debug_error("Submission Failed. Click 'Reload Elements' in the top right")
            elements_button.config(state=tk.NORMAL, text="Reload Elements")

    reset_button = tk.Button(root, text="Reload Page", command=reset_forum, font=OPEN_DYSLEXIC_FONT)
    reset_button.config(state=tk.DISABLED)
    reset_button.grid(column=0, columnspan=30, row=row_idx, sticky="nsew", padx=4, pady=4)

    school_button = tk.Button(root, text="New School", command=new_school, font=OPEN_DYSLEXIC_FONT)
    school_button.grid(column=40, columnspan=30, row=row_idx, sticky="nsew", padx=4, pady=4)

    submission_button = tk.Button(root, text="Submit", command=submit_info, font=OPEN_DYSLEXIC_FONT)
    submission_button.config(state=tk.DISABLED)
    submission_button.grid(column=70, columnspan=30, row=row_idx, sticky="nsew", padx=4, pady=4)
    row_idx += 1

    debug_message_label.config(fg=TEXT_COLOR_ENTERED)
    debug_message_label.grid(column=10, columnspan=80, row=row_idx, sticky="nsew", padx=4, pady=4)

    # TODO: optional ? upload files (maximum of 10mb)

    root.mainloop()

    if driver is not None:
        driver.close()
