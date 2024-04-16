"""
    This script is used to automate the process of entering data into a web application.
    It reads data from an Excel file, performs operations based on the data, and interacts with the web application using PyAutoGUI.
    The script is modular and can be easily modified to work with different applications and data formats.
    The script is also robust and will handle errors gracefully.
    The script is well-documented and easy to understand.
    This script is blended with OOP concepts and is more modular and easy to maintain.
"""
import pyautogui
import time
import os
import pandas as pd
import logging 
import logging

logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')

def read_excel_file(file_path):
    """Reads an Excel file and returns a DataFrame.

    Args:
        file_path (str): Path to the Excel file.

    Returns:
        DataFrame: The data from the Excel file.
    """

    if not os.path.isfile(file_path):
        print(f"Error: File '{file_path}' not found.")
        exit(1)

    df = pd.read_excel(file_path)

    return df

def find_and_click(image_path, confidence=0.8, max_attempts=3):
    """Attempts to find and click an image on the screen.

    Args:
        image_path (str): Path to the image file.
        confidence (float): Minimum confidence level for image detection (0-1).
        max_attempts (int): Maximum number of attempts before giving up.
    """

    attempts = 0
    while attempts < max_attempts:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)

        if location is not None:
            pyautogui.click(location)
            return True  # Click successful

        attempts += 1
        time.sleep(0.5)  # Short pause between attempts

    return False  # Image not found after multiple attempts

def find_image_type_and_click(image_path, text_to_type, confidence=0.8, max_attempts=3):
    """Attempts to find an image on the screen, types the specified text, and then clicks.

    Args:
        image_path (str): Path to the image file.
        text_to_type (str): Text to type after finding the image.
        confidence (float): Minimum confidence level for image detection (0-1).
        max_attempts (int): Maximum number of attempts before giving up.
    """

    attempts = 0
    while attempts < max_attempts:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)

        if location is not None:
            pyautogui.click(location)
            pyautogui.write(text_to_type)
            pyautogui.press('enter')
            return True  

        attempts += 1
        time.sleep(0.5)  

    return False  

def sdal_reset(offset_x=-50):
    """
    Perform a reset operation for the SDAL application.
    Or use as a way to automate your way to the SDAL screen.
    """
    cimt_reset()
    image_paths = ['SEARCH2.jpg']
    for image_path in image_paths:
        if not os.path.isfile(image_path):
            print(f"Error: Image file '{image_path}' not found.")
            exit(1)  

    for image_path in image_paths:
        success = find_and_click(image_path)
        if not success:
            print(f"Warning: Image '{image_path}' not found on screen.")
        else:
            pyautogui.typewrite('SDAL')
            pyautogui.press('enter')
        time.sleep(1)  

def audt_reset(offset_x=-50):
    """
    Perform a reset operation for the AUDT application.
    Or use as a way to automate your way to the AUDT screen.
    """
    cimt_reset()
    image_paths = ['SEARCH2.jpg']
    for image_path in image_paths:
        if not os.path.isfile(image_path):
            print(f"Error: Image file '{image_path}' not found.")
            exit(1)  

    for image_path in image_paths:
        success = find_and_click(image_path)
        if not success:
            print(f"Warning: Image '{image_path}' not found on screen.")
        else:
            pyautogui.typewrite('AUDT')
            pyautogui.press('enter')
        time.sleep(1)  

def cimt_reset():
    """
    Perform a reset operation for the CIMT application.
    Or use as a way to automate your way to the CIMT screen.
    """
    image_paths = ['CUSM.jpg', 'CIMT.jpg']
    for image_path in image_paths:
        if not os.path.isfile(image_path):
            logging.error(f"Error: Image file '{image_path}' not found.")
            exit(1)  

    for image_path in image_paths:
        if not find_and_click(image_path):
            logging.warning(f"Warning: Image '{image_path}' not found on screen.")
        time.sleep(1)

def comt_reset():
    """
    Perform a reset operation for the COMT application.
    Or use as a way to automate your way to the COMT screen.
    """
    image_paths = ['CUSM.jpg', 'COMT.jpg']
    for image_path in image_paths:
        if not os.path.isfile(image_path):
            print(f"Error: Image file '{image_path}' not found.")
            exit(1)  

    for image_path in image_paths:
        success = find_and_click(image_path)
        if not success:
            print(f"Warning: Image '{image_path}' not found on screen.")
        time.sleep(1)  

def itpi_reset():
    """
    Perform a reset operation for the ITPI application.
    Or use as a way to automate your way to the ITPI screen.
    """
    image_paths = ['CUSM.jpg', 'ITPI.jpg']
    for image_path in image_paths:
        if not os.path.isfile(image_path):
            print(f"Error: Image file '{image_path}' not found.")
            exit(1)  

    for image_path in image_paths:
        success = find_and_click(image_path)
        if not success:
            print(f"Warning: Image '{image_path}' not found on screen.")
        time.sleep(1)  



def detect_image_and_break(image_path, timeout=30, check_interval=1, confidence=0.8):
    """
    Continuously checks for an image on the screen and stops the process if the image is found.

    Args:
        image_path (str): Path to the image file to be detected.
        timeout (int): Maximum time in seconds to wait for the image before timing out.
        check_interval (int): Time in seconds between checks for the image.
        confidence (float): Confidence level required to recognize the image (between 0 and 1).

    Returns:
        bool: True if the image was detected, False if the timeout was reached.
    """
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            logging.warning("Timeout reached, image not detected.")
            return False

        location = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if location:
            logging.info(f"Image detected at {location}.")
            return True

        time.sleep(check_interval)

def detect_color(region, target_color=(255, 0, 0), tolerance=10):
    screenshot = pyautogui.screenshot(region=region)
    for x in range(screenshot.width):
        for y in range(screenshot.height // 8 * 7, screenshot.height):  # Start from the bottom 8th of the image
            r, g, b = screenshot.getpixel((x, y))
            if all(abs(c - tc) <= tolerance for c, tc in zip((r, g, b), target_color)):
                return True
    return False


def cimt_clear(customer):
    """
    Clears the CIMT screen for a given customer.
    
    Args:
        customer (str): The customer's information to be cleared.
    """
    cimt_reset()
    pyautogui.typewrite(customer)
    
    print(f"Typing customer name: {customer}")
    print(f"Clearing CIMT for customer '{customer}'...")

    screen_width, screen_height = pyautogui.size()
    color_region = (0, 0, screen_width, screen_height)

    while True:
        if detect_color(color_region, target_color=(255, 0, 0)):
            break
        
        pyautogui.press('enter')
        print("Pressing 'enter'")
        time.sleep(0.01)  

        
        while True:
            if detect_color(color_region, target_color=(255, 0, 0)):
                print("Detected red breaking loop.")
                break
            pyautogui.press('backspace')
            #pyautogui.press('enter') commented out for testing
            pyautogui.press('down')
            print("Pressing 'down' key to navigate")
            
    logging.info("Red Detected, breaking loop.")        
    print("Done.")

if __name__ == '__main__':
    customer = 'ARGOCZ'
    #df = read_excel_file('EXAMPLE.xlsx')
    cimt_clear(customer)
    

