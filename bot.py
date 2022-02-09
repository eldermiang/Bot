import keyboard
import pyautogui
import win32api
import win32con
import ctypes
from threading import Event

exit = Event()
user32 = ctypes.windll.user32
covCount = 0
mysticCount = 0
refreshCount = 0
# Screen resolution
length = user32.GetSystemMetrics(0)
height = user32.GetSystemMetrics(1)

exit.wait(2)
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
    covenant = pyautogui.locateCenterOnScreen('Images/Covenant.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.70)
    mystic = pyautogui.locateCenterOnScreen('Images/Mystic.PNG', region=(0, 0, length, height), grayscale=False, confidence=0.70)
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
        

# Run program in loop as long as q is not pressed
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
    if keyboard.is_pressed('q'):
        print("************Stats************")
        print("Refresh Count:" + str(refreshCount))
        print("Covenant Count: " + str(covCount))
        print("Mystic Count: " + str(mysticCount))
        print("*****************************")
        exit.set()