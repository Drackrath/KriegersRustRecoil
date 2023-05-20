import cv2
import numpy as np
from PIL import ImageGrab
from pynput.mouse import Button, Controller, Listener
import time
import os
from screeninfo import get_monitors

# Speed and delay settings
speed = 5
delay = 0.05  # Delay in seconds

# Get screen resolution
monitor = get_monitors()[0]
screen_width = monitor.width
screen_height = monitor.height

# Capture area settings
capture_width = 700
capture_height = 300

# Calculate the top left and bottom right points of the capture area
top_left = (int((screen_width - capture_width) / 2), int(screen_height - capture_height))
bottom_right = (int((screen_width + capture_width) / 2), screen_height)

# Create a mouse controller
mouse = Controller()

# Prepare the templates
templates = [
    {"path": "Semi-Automatic_Rifle_icon.png", "dx": -speed, "dy": speed},
    # Add more templates as needed
]
for template in templates:
    template_path = os.path.join('templates', template["path"])
    template_image = cv2.imread(template_path, 0)
    template["image"] = template_image
    template["shape"] = template_image.shape[::-1]

def move_mouse(dx, dy):
    x, y = mouse.position
    mouse.position = (x + dx, y + dy)

def check_image():
    # Capture screenshot in the defined area
    screenshot = ImageGrab.grab(bbox=(top_left[0], top_left[1], bottom_right[0], bottom_right[1]))

    # Convert the screenshot to a numpy array and then to grayscale
    img_gray = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2GRAY)

    for template in templates:
        # Perform template matching
        res = cv2.matchTemplate(img_gray, template["image"], cv2.TM_CCOEFF_NORMED)
        threshold = 0.8
        loc = np.where(res >= threshold)

        # If the template image is found in the screenshot, return the associated movements
        if np.any(loc[0]):
            return template["dx"], template["dy"]
    return None, None

def on_click(x, y, button, pressed):
    try:
        if button == Button.left and pressed:
            dx, dy = check_image()
            if dx is not None and dy is not None:
                move_mouse(dx, dy)
            time.sleep(delay)
    except NotImplementedError:
        pass  # Do nothing if NotImplementedError occurs

# Listen for mouse click events
with Listener(on_click=on_click) as listener:
    listener.join()