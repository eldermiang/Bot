import keyboard
import pyautogui
from torch import fix
import win32api
import win32con
import ctypes
from threading import Event

# Setup variables
mode = None
fixedCycles = 0

# System Variables
exit = Event()
user32 = ctypes.windll.user32
covCount = 0
mysticCount = 0
refreshCount = 0

# Screen resolution
length = user32.GetSystemMetrics(78)
height = user32.GetSystemMetrics(79)

# Setup
def setup():
    global mode
    global fixedCycles
    while True:
        mode = input("Selected a mode FIXED or UNLIMITED?").lower()
        if mode not in ('fixed', 'unlimited'):
            print("Please enter a valid mode")
            continue
        else:
            break
    if mode == "fixed":
        while True:
            try:
                fixedCycles = int(input("Enter the number of refreshes: "))
                if fixedCycles <= 0:
                    raise ValueError("ValueError exception thrown")
            except ValueError:
                print("Please enter a valid number")
                continue
            else:
                break
    print("Mode: " + mode.upper())
    if mode == "fixed":
        print("Running for " + str(fixedCycles) + " refreshes.")
    # Refresh start delay
    exit.wait(3)

# Mouse click
def click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    exit.wait(0.5)

# Buy bookmark in popup
def buy():
    buy = pyautogui.locateCenterOnScreen('Images/Buy.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.65)
    if buy is not None:
        pyautogui.moveTo(buy)
        click()
        click()

# Locate bookmark and buy in shop
def buyBookmark():
    covenant = pyautogui.locateCenterOnScreen('Images/Covenant.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.85)
    mystic = pyautogui.locateCenterOnScreen('Images/Mystic.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.85)
    if covenant is not None:
        global covCount
        covCount += 1
        pyautogui.moveTo(covenant.x+225, covenant.y+20)
        click()
        buy()
    if mystic is not None:
        global mysticCount
        mysticCount += 1
        pyautogui.moveTo(mystic.x+225, mystic.y+20)
        click()
        buy()

# Print bot stats
def printStats():
    global refreshCount
    global covCount
    global mysticCount
    print("Program Terminated")
    print("************Stats************")
    print("Refresh Count: " + str(refreshCount))
    print("Covenant Count: " + str(covCount))
    print("Mystic Count: " + str(mysticCount))
    print("*****************************")

# Run program in loop as long as q is not pressed
print("Program Running")
setup()
print("Starting refresh")
while not exit.is_set():
    refresh = pyautogui.locateCenterOnScreen('Images/Refresh.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.70)
    confirm = pyautogui.locateCenterOnScreen('Images/Confirm.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.70)
    # Refresh shop and confirm
    if refresh is not None:
        pyautogui.moveTo(refresh)
        click()
        if confirm is not None:
            pyautogui.moveTo(confirm) 
            click()
            exit.wait(1)
            refreshCount += 1
            buyBookmark()
            pyautogui.dragTo(confirm.x, confirm.y-250, 0.2, button='left')
            buyBookmark()
            exit.wait(1)
    # Terminate program
    if keyboard.is_pressed('q') or (mode == "fixed" and refreshCount >= fixedCycles):
        printStats()
        exit.set()
        input()
