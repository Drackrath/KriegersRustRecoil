import cv2
import numpy as np
from PIL import ImageGrab
from pynput import mouse as pynput_mouse
from pynput.mouse import Button
from pynput.mouse import Listener as MouseListener
from pynput.keyboard import Key, Listener as KeyboardListener
import os
from screeninfo import get_monitors
import win32gui
import win32con
import win32api
import time
from pywinauto import mouse as win_mouse
from queue import Queue
import threading

# Speed and delay settings
speed = 50
delay = 1  # Delay in seconds

# Get screen resolution
monitor = get_monitors()[0]
screen_width = monitor.width
screen_height = monitor.height

# Capture area settings
capture_width = 750
capture_height = 150

# Calculate the top left and bottom right points of the capture area
top_left = (int((screen_width - capture_width) / 2), int(screen_height - capture_height))
bottom_right = (int((screen_width + capture_width) / 2), screen_height)

# Prepare the templates
templates = [
    {"path": "Semi-Automatic_Rifle_icon_100x100.png", "dx": 0, "dy": 100},
    # Add more templates as needed
]
for template in templates:
    template_path = os.path.join('templates', template["path"])
    template_image = cv2.imread(template_path, 0)
    template["image"] = template_image
    template["shape"] = template_image.shape[::-1]

# Keep track of the last time the image was checked
check_image_flag = True

from threading import Thread
from queue import Queue

# Initialize a task queue
task_queue = Queue()


# Worker function to process tasks from the queue
def worker_func():
    while True:
        task = task_queue.get()
        if task is None:
            break
        task()


# Start the worker thread
worker_thread = Thread(target=worker_func)
worker_thread.start()


# Adjustment in the on_key function
def on_key(key):
    print('Key pressed')  # This should be printed whenever any key is pressed
    global check_image_flag
    # Check if the 'F7' key was pressed
    if key == Key.f8:
        check_image_flag = not check_image_flag
        hwnd = win32gui.FindWindow(None, "TranspWindow")
        dx, dy = check_image()
        if dx is not None and dy is not None:
            current_pos = win32gui.GetCursorPos()
            new_pos = (current_pos[0] + dx, current_pos[1] + dy)
            win_mouse.move(coords=new_pos)
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)


def make_window_transparent(show):
    # Check if the window is currently visible
    hwnd = win32gui.FindWindow(None, "TranspWindow")
    print(f'hwnd: {hwnd}')  # print hwnd for debug
    if hwnd != 0:
        # If the window is found, make it transparent to clicks
        extended_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        print(f'Extended style before setting: {extended_style}')  # print extended style before setting for debug
        if show:
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, extended_style | win32con.WS_EX_TRANSPARENT)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        else:
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, extended_style & ~win32con.WS_EX_TRANSPARENT)
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

        extended_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        print(f'Extended style after setting: {extended_style}')  # print extended style after setting for debug
    else:
        print("Window 'TranspWindow' not found!")


def check_image():
    # Capture screenshot in the defined area
    screenshot = ImageGrab.grab(bbox=(top_left[0], top_left[1], bottom_right[0], bottom_right[1]))

    # Convert the screenshot to a numpy array and then to grayscale
    img_gray = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2GRAY)
    img_color = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2RGB)  # create a colored version for drawing

    for template in templates:
        # Perform template matching
        if img_gray.shape[0] < template["shape"][0] or img_gray.shape[1] < template["shape"][1]:
            print(f'Template image is larger than screenshot. Skipping template {template["path"]}.')
            continue

        res = cv2.matchTemplate(img_gray, template["image"], cv2.TM_CCOEFF_NORMED)
        print(f'Max match value for template {template["path"]}: {res.max()}')  # Print max match value for debug
        threshold = 0.30
        loc = np.where(res >= threshold)

        # Print the locations of the matches
        print(f'Match locations for template {template["path"]}: {loc}')

        # If the template image is found in the screenshot, draw a circle at the center of each match
        if np.any(loc[0]):
            print(f'Found match for template {template["path"]}. Moving mouse by {template["dx"]}, {template["dy"]}')

            for pt in zip(*loc[::-1]):
                center_x = pt[0] + template["shape"][0] // 2
                center_y = pt[1] + template["shape"][1] // 2

                # Draw a circle at the center of the match
                cv2.circle(img_color, (center_x, center_y), 10, (0, 255, 0), -1)

            cv2.imwrite("matches.png", cv2.cvtColor(img_color, cv2.COLOR_RGB2BGR))  # Save the image with matches

            return template["dx"], template["dy"]
    return None, None


def move_mouse(dx, dy):
    current_pos = win_mouse.get_cursor_position()
    new_pos = (current_pos.x + dx, current_pos.y + dy)
    win_mouse.move(coords=new_pos)


# Define the mouse callback function

def perform_mouse_movement(x, y):
    def mouse_move_thread(x, y):
        # Perform the mouse movement using pywinauto
        win_mouse.move(coords=(x, y))

    threading.Thread(target=mouse_move_thread, args=(x, y)).start()

# New function for pynput mouse event handling
def on_click(x, y, button, pressed):
    global check_image_flag
    if button == Button.left and pressed and check_image_flag:
        dx, dy = check_image()
        if dx is not None and dy is not None:
            # Perform mouse movement in a separate function
            perform_mouse_movement(x + dx, y + dy)

def wndProc(hwnd, message, wParam, lParam):
    return win32gui.DefWindowProc(hwnd, message, wParam, lParam)


def create_window():
    # Define window class
    wndclass = win32gui.WNDCLASS()
    wndclass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    wndclass.lpfnWndProc = {}
    wndclass.hIcon = win32gui.LoadIcon(None, win32con.IDI_APPLICATION)
    wndclass.hCursor = win32gui.LoadCursor(None, win32con.IDC_ARROW)
    wndclass.hbrBackground = win32gui.GetStockObject(win32con.WHITE_BRUSH)
    wndclass.lpszClassName = "TranspWindow"  # The class name you will use in CreateWindowEx

    # Register window class
    wndclass_atom = win32gui.RegisterClass(wndclass)

    # Create a window
    hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST,
        # Styles to make the window transparent and always on top
        wndclass_atom,  # The class name
        'TranspWindow',  # Window name
        win32con.WS_OVERLAPPEDWINDOW,  # Window style
        top_left[0], top_left[1], capture_width, capture_height,  # Position and size
        None,  # Parent window handle
        None,  # Menu handle
        0,  # Application instance handle
        None)  # Window creation data

    # Set the window's transparency
    win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0, 0, 0), 128, win32con.LWA_ALPHA)

    # Show the window
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

    # Run the message loop
    win32gui.PumpMessages()


# Run the event loop
if __name__ == "__main__":
    # Create a listener for keyboard events
    keyboard_listener = KeyboardListener(on_press=on_key)
    keyboard_listener.start()

    # Create a listener for mouse events
    mouse_listener = MouseListener(on_click=on_click)
    mouse_listener.start()

    create_window()

    while True:
        time.sleep(1)
