




import pyautogui
import time
import pyperclip
from datetime import datetime
print("Bot is starting ")
current_date = datetime.now().strftime("%Y-%m-%d")
current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
print("Date:", current_date)
print("Date and Time:", current_time)
# open Chrome
pyautogui.hotkey("win","r")
time.sleep(2)
pyautogui.write("chrome")
pyautogui.press("enter")
#
# Select chrome profile
pyautogui.moveTo(787,567, duration=1)
print("mouse is on your profile")
time.sleep(2)
pyautogui.click()
print("Chrome profile selected")
time.sleep(5)
#
#open chrome in new tab
#pyautogui.hotkey("ctrl","t",interval=0.1)
#time.sleep(2)
#
# Open weather website
pyautogui.typewrite("https://www.weather.com")
pyautogui.press("enter")
time.sleep(3)
print("website opened")
#
#Copy the webpage content
pyautogui.hotkey("ctrl", "a")
time.sleep(1)
pyautogui.hotkey("ctrl", "c")
time.sleep(2)
print("Weather page copied!")
# Copy the webpage content to clipboard using pyperclip
webpage_content = pyperclip.paste()
print("Webpage content copied to clipboard")
print(webpage_content[:500])
#
#Open Notepad
pyautogui.hotkey("win", "r")
time.sleep(2)
pyautogui.write("notepad")
pyautogui.press("enter")
time.sleep(3)
#open a new notepad tab
#pyautogui.hotkey("ctrl","n")
#time.sleep(2)
print("Notepad opened!")
#paste the copied content
pyautogui.hotkey("ctrl","v")
time.sleep(3)
print("weather data pasted into the notepad")
#
#create excel report
from openpyxl import Workbook
# Create current date and time
current_date = datetime.now().strftime("%Y-%m-%d")
current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Create Excel filename
excel_filename = f"daily_report_{current_date}.xlsx"

# Create a new Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Daily Report"

# Headers
ws["A1"] = "Date & Time"
ws["B1"] = "Weather Data"
ws["C1"] = "Comment"

# Report data
ws["A2"] = current_time
ws["B2"] = "Weather information captured from weather.com"
ws["C2"] = "Daily weather report captured successfully."

# Save Excel file
full_path = rf"C:\Users\jaiku\Gen AI Week 1\daily-report-bot\{excel_filename}"
wb.save(full_path)
print(f"Excel report saved at {full_path}")
# -------------------------------
# Open the saved Excel report
# -------------------------------

pyautogui.hotkey("win", "r")
time.sleep(2)

pyautogui.write(full_path)
pyautogui.press("enter")
time.sleep(7)

print("Excel report opened!")

# -------------------------------
# Take screenshot
# -------------------------------

screenshot_filename = rf"C:\Users\jaiku\Gen AI Week 1\daily-report-bot\daily_report_{current_date}.png"

pyautogui.screenshot(screenshot_filename)

print("Screenshot saved!")
print("Bot completed successfully!")

# -------------------------------
# Bot finished
# -------------------------------

print("Daily report bot completed successfully!")