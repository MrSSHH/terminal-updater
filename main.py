import argparse
import pandas as pd
import pyautogui
from time import sleep
import logging
import cv2  # OpenCV for image matching
import numpy as np
from playsound import playsound
import ctypes
from pynput import mouse, keyboard
import json
import os
import sys

# Setup logging
logger = logging.getLogger('Logger')
logger.setLevel(logging.DEBUG)

# Create a file handler
file_handler = logging.FileHandler('log.txt')
file_handler.setFormatter(logging.Formatter('%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s', datefmt='%H:%M:%S'))
logger.addHandler(file_handler)

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Update devices based on UUIDs from an Excel file.")
parser.add_argument('-s', '--silent', action='store_true', help="Disable sound notifications.")
parser.add_argument('--short-wait', action='store_true', help="Use a shorter wait time (3 seconds) before starting.")
parser.add_argument('-v', '--verbose', action='store_true', help="Output all debug logs to terminal.")
args = parser.parse_args()

# Flag to disable sounds
disable_sounds = args.silent

# Set up console logging if verbose mode is enabled
if args.verbose:
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s', datefmt='%H:%M:%S'))
    logger.addHandler(console_handler)

# Load the UUIDs from the Excel file
df = pd.read_excel(r'Assets\Excels\data.xlsx')

# Global flag to signal Ctrl + Left Click
ctrl_left_click_detected = False
config_file_path = r'config.json'
stop_process = False  # Global flag to stop the updating process


def load_config():
    if os.path.exists(config_file_path):
        with open(config_file_path, 'r') as f:
            return json.load(f)
    return None


def save_config(data):
    with open(config_file_path, 'w') as f:
        json.dump(data, f, indent=4)
    logger.debug("Configuration saved.")


def get_cursor_position(message):
    global ctrl_left_click_detected
    ctrl_left_click_detected = False
    print(message + " Press Ctrl + Left Click at the position.")

    def on_click(x, y, button, pressed):
        if pressed and button == mouse.Button.left:
            if ctrl_listener.ctrl_pressed:
                global ctrl_left_click_detected
                ctrl_left_click_detected = True
                position = (x, y)
                print(f"Position confirmed at {position}.")
                logger.debug(f"Position registered: {position}")
                return False

    def on_press(key):
        if key == keyboard.Key.ctrl_l:
            ctrl_listener.ctrl_pressed = True

    def on_release(key):
        if key == keyboard.Key.ctrl_l:
            ctrl_listener.ctrl_pressed = False

    with mouse.Listener(on_click=on_click) as listener:
        with keyboard.Listener(on_press=on_press, on_release=on_release) as ctrl_listener:
            ctrl_listener.ctrl_pressed = False
            listener.join()
            play_sound(r'Sounds\beep.mp3')

    if ctrl_left_click_detected:
        return pyautogui.position()


def check_screen_for_image(image_path, retries=3, resize_factor=0.9, timeout=3):
    for _ in range(retries):
        screen = pyautogui.screenshot()
        screen = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2BGR)
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)

        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        logger.debug(f"Image recognition attempt for {image_path}, confidence: {max_val}")

        if max_val > 0.4:
            logger.debug(f"Image found at {max_loc} with confidence {max_val}")
            return max_loc
        sleep(timeout)
        template = cv2.resize(template, (int(template.shape[1] * resize_factor), int(template.shape[0] * resize_factor)))

    logger.debug(f"Failed to find image: {image_path} after {retries} retries.")
    return None


def show_alert(title, message):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x1)


def play_sound(file_path):
    if not disable_sounds:
        playsound(file_path)


def on_escape(key):
    global stop_process
    if key == keyboard.Key.esc and keyboard.Listener.ctrl_pressed:
        stop_process = True
        logger.debug("Update process has been stopped by the user.")
        return False  # Stop the listener


# Load the configuration file
config = load_config()

if config:
    use_previous = input("Use the previous settings (Y/N)? ").strip().lower() == 'y'
    if use_previous:
        uuid_field_pos = tuple(config['uuid_field_pos'])
        update_btn_pos = tuple(config['update_btn_pos'])
        version_field_pos = tuple(config['version_field_pos'])
        confirm_btn_pos = tuple(config['confirm_btn_pos'])
        desired_version = config['desired_version']
        logger.debug("Using previous settings.")
    else:
        desired_version = input("Enter the desired version (e.g., 802): ")
        uuid_field_pos = get_cursor_position("Place your cursor on the UUID input field.")
        update_btn_pos = get_cursor_position("Place your cursor on the Update button.")
        version_field_pos = get_cursor_position("Place your cursor on the version input field.")
        confirm_btn_pos = get_cursor_position("Place your cursor on the Confirm button.")

        config_data = {
            'uuid_field_pos': uuid_field_pos,
            'update_btn_pos': update_btn_pos,
            'version_field_pos': version_field_pos,
            'confirm_btn_pos': confirm_btn_pos,
            'desired_version': desired_version
        }
        save_config(config_data)

# Step 4: Prepare update window
wait_time = 3 if args.short_wait else 5
input(f"Once ready, press OK to start the update process. You will have {wait_time} seconds to open the update window.")
for i in range(wait_time, 0, -1):
    play_sound(r'Sounds\beep.mp3')
    sleep(1)
play_sound(r'Sounds\longbeep.mp3')

# Start listener for Ctrl + Esc
with keyboard.Listener(on_press=on_escape) as listener:
    # Check if the update window is open
    window_image_path = r'Images\OnScreen\openUpdateWindow.jpg'
    if not check_screen_for_image(window_image_path):
        show_alert("Error", "Error 1: Couldn't find the update window open. Make sure it's on screen and visible.")
        logger.error("Update window not found.")
        exit(1)

    # Step 5: Check for "Active" field
    active_field_image = r'Images\ActiveField.jpg'
    if not check_screen_for_image(active_field_image):
        show_alert("Error", "Error 2: Couldn't find the 'Active' field.")
        logger.error("Active field not found.")
        exit(2)

    # Process UUIDs
    for num, uuid in enumerate(df['Uuid']):
        if stop_process:
            logger.debug("Update process halted.")
            break

        logger.debug(f"Starting to upgrade UUID: {uuid}")

        # Step 6: Enter the UUID
        pyautogui.click(uuid_field_pos)
        logger.info(f"Click at ({uuid_field_pos})")
        sleep(0.6)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(uuid)
        pyautogui.press('enter')
        sleep(0.6)

        # Step 7: Click update button
        pyautogui.click(update_btn_pos)
        sleep(1)

        # Step 8: Open the version input window
        if not check_screen_for_image(window_image_path, retries=3, timeout=3):
            pyautogui.click(update_btn_pos)
            if not check_screen_for_image(window_image_path, retries=3):
                show_alert(f"Error 3: Couldn't find the update version window for UUID {uuid}.")
                logger.error(f"Update version window not found for UUID {uuid}")
                exit(3)

        # Step 9: Clear and update version field
        pyautogui.click(version_field_pos)
        pyautogui.hotkey('ctrl', 'a')
        sleep(0.6)

        pyautogui.press('backspace')

        # Step 10: Enter new version and confirm
        pyautogui.write(desired_version)
        pyautogui.click(confirm_btn_pos)

        # Verify the confirmation window disappeared
        if check_screen_for_image(window_image_path, retries=3, timeout=3):
            logger.debug(f"UUID {uuid} updated successfully to {desired_version} ({len(df['Uuid']) - num} Remaining).")
        else:
            logger.error(f"Confirmation window didn't disappear for UUID {uuid}")
            show_alert(f"Error 5: The confirmation window didn't disappear for UUID {uuid}.")
            exit(5)

# Finish the process
show_alert("Success", "All devices have been updated!")
logger.debug("Process completed successfully.")
