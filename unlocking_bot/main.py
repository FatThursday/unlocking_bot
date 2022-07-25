import time
import pyautogui
import win32api
import win32con
import pandas
import pyperclip

data = pandas.read_csv("TEST_DATA.csv")

data_list = data["MAILE"].to_list()
GREEN = (106, 175, 87)
pyautogui.FAILSAFE = True

sleep_time = 2

def click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.5)


def click_and_mark():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)

#---------------CALIBRATION AREA----------------------------------------------------------

start = pyautogui.locateCenterOnScreen('GREEN_BANNER_ASIMUTE_PROD.png', confidence=0.9, grayscale=False) # znajduje pozycje przycisku search
green_banner_coords = ((start[0] + 250), (start[1]))
start = pyautogui.locateCenterOnScreen('SEARCH_FIELD_ASIMUTE_PROD.png', confidence=0.9, grayscale=False)
search_field = (start[0] + 150), (start[1] + 60) # ustawia pozycje pola wyszukiwania
search_button = ((green_banner_coords[0]), (green_banner_coords[1]+265)) #ustawia pozycje przycisku search
im = pyautogui.screenshot()
pixel = im.getpixel(search_button)
print(pixel)
pyperclip.copy("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX") #dummy adres do wklejenia
pyautogui.moveTo(search_field) # rusza kursor na pozycje pola wyszukiwania
click_and_mark() # klika w pole i zaznacza wszystko
pyautogui.hotkey('ctrl', 'v') # wkleja adres ze schowka
pyautogui.moveTo(search_button)
click()
pyautogui.moveTo((search_button[0], search_button[1] + 200))
while not pyautogui.pixelMatchesColor(int(search_button[0]), int(search_button[1]), (pixel[0], pixel[1], pixel[2])): # czeka az przycisk search bedzie znowu niebieski (czeka az Coupa przestanie szukac)
    time.sleep(1)
time.sleep(2)
unlock_position = pyautogui.locateCenterOnScreen('UNLOCK_BUTTON_PROD.png', region=(250, 250, 1366, 675),confidence=0.7,grayscale=False) # znajduje pozycje przycisku unlock
if unlock_position is not None:
    print("found it")
pyautogui.moveTo((unlock_position[0] - 30, unlock_position[1] + 30)) # rzusza kursor na przycisk unlock
wait_position_reference = pyautogui.position()
print(wait_position_reference) #tu konczy sie cykl kalibracyjny, w tym miejscu zapisuje pozycje kursora ktora bedzie sluzyc jako referencja w rozpoznaniu czy ma czekac na banner czy nie
#--------------- one loop test area ------------------------------------------------------


    # pyautogui.moveTo(search_field)
    # time.sleep(2)
    # pyperclip.copy(data_list[0]) #kopiuje adres z listy do schowka
    # pyautogui.moveTo(search_field) #rusza kursor na pozycje pola wyszukiwania
    # click_and_mark() #klika w pole i zaznacza wszystko
    # pyautogui.hotkey('ctrl', 'v') #wkleja adres ze schowka
    # pyautogui.moveTo(search_button_coords) #rusza kursor na przycisk search
    # click()
    # unlock_position = pyautogui.locateOnScreen('unlock_button3.png', confidence=0.80, grayscale=False) #znajduje pozycje przycisku unlock
    # if unlock_position is not None:
    # print("found it")
    # time.sleep(2)
    # pyautogui.moveTo((unlock_position[0]-30, unlock_position[1]+30))
    # print((unlock_position))


#-------------- live loop test area----------------------------------------------------


for i in data_list:
    pyperclip.copy(i) # kopiuje adres z listy do schowka
    pyautogui.moveTo(search_field) # rusza kursor na pozycje pola wyszukiwania
    click_and_mark() # klika w pole i zaznacza wszystko
    pyautogui.hotkey('ctrl', 'v') # wkleja adres ze schowka
    pyautogui.moveTo(search_button) # rusza kursor na przycisk search
    click()
    pyautogui.moveTo((search_button[0], search_button[1] + 200))
    while not pyautogui.pixelMatchesColor(int(search_button[0]), int(search_button[1]), (pixel[0], pixel[1], pixel[2])): # czeka az przycisk search bedzie znowu niebieski (czeka az Coupa przestanie szukac)
        time.sleep(1)
    time.sleep(2)
    unlock_position = pyautogui.locateCenterOnScreen('UNLOCK_BUTTON_PROD.png', region=(250, 250, 1366, 675), confidence=0.7, grayscale=False) # znajduje pozycje przycisku unlock
    if unlock_position is not None:
        print("found it")
    pyautogui.moveTo((unlock_position[0] - 30, unlock_position[1] + 30)) # rzusza kursor na przycisk unlock
    click()
    wait_position = pyautogui.position()
    time.sleep(sleep_time)
    if wait_position != wait_position_reference: # tu rozpoznaje czy odblokowal jakis adres i ma czekac czy nie, jesli odblokowal to czeka jesli nie to pomija
        while not pyautogui.pixelMatchesColor(int(green_banner_coords[0] - 200), int(green_banner_coords[1]), (GREEN[0], GREEN[1], GREEN[2])):
            time.sleep(1)
        pyautogui.moveTo(green_banner_coords) # rusza kursor na przycisk search
        click()
        time.sleep(sleep_time)
    else:
        pass
    pyautogui.moveTo(search_button)
    click()
    pyautogui.moveTo((search_button[0], search_button[1] + 200)) # rusza kursor z przycisku search bo najechanie na niego zmienia kolor a to istotne zeby nie byl zmieniony
    while not pyautogui.pixelMatchesColor(int(search_button[0]), int(search_button[1]), (pixel[0], pixel[1], pixel[2])): # czeka az przycisk search bedzie znowu niebieski (czeka az Coupa przestanie szukac)
        time.sleep(1)

