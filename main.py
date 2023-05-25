import sys

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
import threading
import ctypes
from threading import Thread
from queue import Queue
import keyboard

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

current_template = None  # Initialize at the start of your script

# Calculate the top left and bottom right points of the capture area
top_left = (int((screen_width - capture_width) / 2), int(screen_height - capture_height))
bottom_right = (int((screen_width + capture_width) / 2), screen_height)

# Prepare the templates
# Prepare the templates
#{0.000000,-2.257792},{0.323242,-2.300758},{0.649593,-2.299759},{0.848786,-2.259034},{1.075408,-2.323947},{1.268491,-2.215956},{1.330963,-2.236556},{1.336833,-2.218203},{1.505516,-2.143454},{1.504423,-2.233091},{1.442116,-2.270194},{1.478543,-2.204318},{1.392874,-2.165817},{1.480824,-2.177887},{1.597069,-2.270915},{1.449996,-2.145893},{1.369179,-2.270450},{1.582363,-2.298334},{1.516872,-2.235066},{1.498249,-2.238401},{1.465769,-2.331642},{1.564812,-2.242621},{1.517519,-2.303052},{1.422433,-2.211946},{1.553195,-2.248043},{1.510463,-2.285327},{1.553878,-2.240047},{1.520380,-2.221839},{1.553878,-2.240047},{1.553195,-2.248043} },
#[[0.000000,-2.257792],{0.323242,-2.300758},{0.649593,-2.299759},{0.848786,-2.259034},{1.075408,-2.323947},{1.268491,-2.215956},{1.330963,-2.236556},{1.336833,-2.218203},{1.505516,-2.143454},{1.504423,-2.233091},{1.442116,-2.270194},{1.478543,-2.204318},{1.392874,-2.165817},{1.480824,-2.177887},{1.597069,-2.270915},{1.449996,-2.145893},{1.369179,-2.270450},{1.582363,-2.298334},{1.516872,-2.235066},{1.498249,-2.238401},{1.465769,-2.331642},{1.564812,-2.242621},{1.517519,-2.303052},{1.422433,-2.211946},{1.553195,-2.248043},{1.510463,-2.285327},{1.553878,-2.240047},{1.520380,-2.221839},{1.553878,-2.240047},{1.553195,-2.248043} },
#[[-38, 52.3906], [12, 46], [-43, 42], [-58, 37], [0, 34], [0, 28], [34.5, 25], [22.5, 26],
#                    [42.5, 18], [36, 10], [39, 15], [39, 18], [28, 18], [24, 28], [5, 29], [-18, 32], [-30, 33],
#                    [-34, 32], [-36, 29], [-43, 24], [-45, 17], [-45, 8], [-43, 5], [-28, 14], [-19, 21],
#                    [0, 25], [0, 28], [40, 28], [53, 26], [48, 15], [38, 21]],


templates = [
    {
        "path": "Semi-Automatic_Rifle_Screenshot.png",
        "offsets": [ [0.000000,-2.257792],[0.323242,-2.300758],[0.649593,-2.299759],[0.848786,-2.259034],[1.075408,-2.323947],
  [1.268491,-2.215956],[1.330963,-2.236556],[1.336833,-2.218203],[1.505516,-2.143454],[1.504423,-2.233091],
  [1.442116,-2.270194],[1.478543,-2.204318],[1.392874,-2.165817],[1.480824,-2.177887],[1.597069,-2.270915],
  [1.449996,-2.145893],[1.369179,-2.270450],[1.582363,-2.298334],[1.516872,-2.235066],[1.498249,-2.238401],
  [1.465769,-2.331642],[1.564812,-2.242621],[1.517519,-2.303052],[1.422433,-2.211946],[1.553195,-2.248043],
  [1.510463,-2.285327],[1.553878,-2.240047],[1.520380,-2.221839],[1.553878,-2.240047],[1.553195,-2.248043] ],
        "delay": 0.13333,
        "speed": 1.0,
        "loop": True
    }
]
for template in templates:
    template_path = os.path.join('templates', template["path"])
    template_image = cv2.imread(template_path, cv2.IMREAD_UNCHANGED)
    template["image"] = template_image
    template["shape"] = template_image.shape[::-1]

# Keep track of the last time the image was checked
check_image_flag = True



# Initialize a task queue
task_queue = Queue()
should_exit = False

# Worker function to process tasks from the queue
def worker_func():
    while True:
        task = task_queue.get()
        if task is None:
            break
        task()



def check_for_exit_key():
    global should_exit
    while True:
        if keyboard.is_pressed('f3'):  # if key 'f3' is pressed
            print('Exiting Program')
            should_exit = True
            break



# Start the worker thread
worker_thread = Thread(target=worker_func)
worker_thread.start()

exit_thread = threading.Thread(target=check_for_exit_key, daemon=True)
exit_thread.start()



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

    # Convert the screenshot to a numpy array
    img_color = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    for template in templates:
        # Perform template matching
        if img_color.shape[0] < template["shape"][0] or img_color.shape[1] < template["shape"][1]:
            print(f'Template image is larger than screenshot. Skipping template {template["path"]}.')
            continue

        tmplt = template["image"]
        hh, ww = tmplt.shape[:2]

        # extract template mask as grayscale from alpha channel and make 3 channels
        tmplt_mask = tmplt[:, :, 3]
        tmplt_mask = cv2.merge([tmplt_mask, tmplt_mask, tmplt_mask])

        # extract templt2 without alpha channel from tmplt
        tmplt2 = tmplt[:, :, 0:3]

        # do template matching
        corrimg = cv2.matchTemplate(img_color, tmplt2, cv2.TM_CCORR_NORMED, mask=tmplt_mask)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(corrimg)
        max_val_ncc = '{:.3f}'.format(max_val)
        print("correlation match score: " + max_val_ncc)
        xx = max_loc[0]
        yy = max_loc[1]
        print('xmatch =', xx, 'ymatch =', yy)

        # draw red bounding box to define match location
        result = img_color.copy()
        pt1 = (xx, yy)
        pt2 = (xx + ww, yy + hh)
        cv2.rectangle(result, pt1, pt2, (0, 0, 255), 1)

        cv2.imwrite("matches.png", result)  # Save the image with matches

        if float(max_val_ncc) >= 0.9:
            print(f'Found match for template {template["path"]}')

            # Set the current_template to this template
            global current_template
            current_template = template

        return template["offsets"], template["delay"], template["speed"]
    return None, None, None


# Structure for mouse input event
class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)))

# Structure for input event
class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("mouse", MOUSEINPUT)]

def move_mouse(dx, dy):
    print(f'Moving mouse to: {dx} and {dy} ')
    # Create a new mouse input event
    mouse_input = MOUSEINPUT(round(dx), round(dy), 0, 0x0001, 0, None)

    # Create a new input event
    input_event = INPUT(0, mouse_input)

    # Call the SendInput function from user32.dll to send the input event
    ctypes.windll.user32.SendInput(1, ctypes.byref(input_event), ctypes.sizeof(input_event))


# Define the mouse callback function
mouse_pressed = False
current_thread = None
def perform_mouse_movement(offsets, delay, speed, offset_multiplier=17.2):
    def mouse_move_thread(offsets, delay, speed, offset_multiplier):
        global current_template
        global mouse_pressed
        idx = 0  # Index to keep track of current offset
        while current_template is not None and mouse_pressed:  # Only loop if mouse_pressed is True
            # Get the current offset, invert values, and apply multiplier
            offset = offsets[idx]
            offset = [-1 * element * offset_multiplier for element in offset]  # Invert offset values and apply multiplier
            # Perform the mouse movement
            move_mouse(*offset)
            # Wait for delay
            time.sleep(delay)
            idx = (idx + 1) % len(offsets)  # Move to next offset, loop back to start if at end

    global current_thread
    current_thread = threading.Thread(target=mouse_move_thread, args=(offsets, delay, speed, offset_multiplier))
    current_thread.start()

# New function for pynput mouse event handling
def on_click(x, y, button, pressed):
    print('{0} at {1}'.format('Pressed' if pressed else 'Released', (x, y)))

    global check_image_flag
    global mouse_pressed
    global current_thread
    if button == Button.left:
        if pressed and check_image_flag:
            mouse_pressed = True
            offsets, delay, speed = check_image()
            if offsets is not None and current_thread is None or not current_thread.is_alive():
                # Perform mouse movement in a separate function
                perform_mouse_movement(offsets, delay, speed)
        else:
            mouse_pressed = False

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
    win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0, 0, 0), 16, win32con.LWA_ALPHA)

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
        time.sleep(1000)
