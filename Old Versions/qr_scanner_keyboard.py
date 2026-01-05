import keyboard
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# Global variables for accumulating keystrokes from the barcode scanner
accumulated = ""
last_time = 0
TIME_THRESHOLD = 0.1  # Reset accumulation if too much time passes between keystrokes.

def flash_color_for_duration(driver, color, duration, interval=0.3):
    """
    Flashes the background color of the page continuously for a given duration.
    
    :param driver: Selenium WebDriver instance.
    :param color: The color to flash (e.g., 'green' or 'red').
    :param duration: Total duration in seconds to flash.
    :param interval: Duration (in seconds) for each flash state.
    """
    start_time = time.time()
    while time.time() - start_time < duration:
        driver.execute_script("document.body.style.transition='background-color 0.2s';"
                              "document.body.style.backgroundColor = arguments[0];", color)
        time.sleep(interval)
        driver.execute_script("document.body.style.backgroundColor = 'white';")
        time.sleep(interval)

def flash_color_continuous(driver, color, interval=0.3):
    """
    Continuously flashes the background color of the page until the window is closed.
    
    :param driver: Selenium WebDriver instance.
    :param color: The color to flash (e.g., 'red').
    :param interval: Duration (in seconds) for each flash state.
    """
    print(f"Continuously flashing {color} until window is closed.")
    try:
        while True:
            driver.execute_script("document.body.style.transition='background-color 0.2s';"
                                  "document.body.style.backgroundColor = arguments[0];", color)
            time.sleep(interval)
            driver.execute_script("document.body.style.backgroundColor = 'white';")
            time.sleep(interval)
    except Exception as e:
        print("Browser window closed, stopping flashing.")

def open_and_handle_url(url):
    """
    Opens a new browser window (using Selenium) to load the URL.
    It then polls for up to 5 seconds for the word "Success" in the page source.
      - If "Success" is detected, it flashes green for 4 seconds then closes the window.
      - Otherwise, it flashes red until the window is manually closed.
    The browser window is maximized and brought to the front.
    """
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)  # Ensure chromedriver is in your PATH

    driver.get(url)
    # Maximize and bring window to front
    driver.maximize_window()
    driver.execute_script("window.focus();")
    
    print(f"Opened URL in new window: {url}")
    
    success_found = False
    poll_duration = 5  # total polling time in seconds
    poll_interval = 0.5  # interval between polls
    elapsed = 0

    # Poll for "Success" in the page source.
    while elapsed < poll_duration:
        page_source = driver.page_source
        if "Success" in page_source:
            success_found = True
            break
        time.sleep(poll_interval)
        elapsed += poll_interval

    if success_found:
        print("Success detected; flashing green for 4 seconds then closing the window.")
        flash_color_for_duration(driver, "green", duration=4, interval=0.3)
        driver.quit()
    else:
        print("No 'Success' found within polling period; continuously flashing red until the window is closed.")
        flash_color_continuous(driver, "red", interval=0.3)

def on_key_event(event):
    """
    Processes each keystroke from the scanner.
    Accumulates rapid keystrokes into a string and when the Enter key is pressed,
    treats the accumulated string as the scanned URL.
    """
    global accumulated, last_time

    if event.event_type != "down":
        return

    now = time.time()
    if now - last_time > TIME_THRESHOLD:
        accumulated = ""
    last_time = now

    if event.name == "enter":
        if accumulated:
            print("Scanned:", accumulated)
            open_and_handle_url(accumulated)
            accumulated = ""
    else:
        if len(event.name) == 1:
            accumulated += event.name
        elif event.name == "space":
            accumulated += " "

# Set up the global keyboard hook.
keyboard.hook(on_key_event)
print("QR scanner listener started. Scan a QR code with your barcode scanner...")

keyboard.wait()

