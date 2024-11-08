import json
import time
import datetime
import math
import win32com.client
from playwright.sync_api import sync_playwright
from luxafor import Luxafor
import win32api
import win32gui
import win32con

# Load Config
with open('config.json') as config_file:
    config = json.load(config_file)

def log_message(message):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}")

def create_appointment(subject, attendees, duration):
    outlook = win32com.client.Dispatch('Outlook.Application')
    appointment = outlook.CreateItem(1)
    appointment.BusyStatus = 2
    start_time = time.strftime("%Y-%m-%d %H:%M")
    appointment.Start = start_time
    appointment.Duration = duration
    appointment.Subject = subject
    appointment.ReminderSet = False
    appointment.MeetingStatus = 1

    log_message(f"Creating appointment with subject: {subject}")
    log_message(f"Start time: {start_time}")
    end_time = (datetime.datetime.now() + datetime.timedelta(minutes=duration)).strftime("%Y-%m-%d %H:%M")
    log_message(f"End time: {end_time}")
    log_message(f"Duration: {duration} minutes")

    for attendee_email in attendees:
        recipient = appointment.Recipients.Add(attendee_email)
        recipient.Type = 1
        if not recipient.Resolve:
            log_message(f"Warning: Could not resolve recipient: {attendee_email}")
        else:
            log_message(f"Invite sent to: {attendee_email}")

    try:
        appointment.Save()
        appointment.Send()
        log_message("Appointment created and sent successfully.")
    except Exception as e:
        log_message(f"An error occurred while creating the appointment: {e}")

def start_pomodoro_timer(page):
    try:
        page.goto(config['pomofocus_url'])
        page.locator("text=START").first.click()
        page.on("dialog", lambda dialog: dialog.accept())
        log_message("Pomodoro timer started.")
        return True
    except Exception as e:
        log_message(f"An error occurred while starting the Pomodoro timer: {e}")
        return False

def wait_for_timer_completion(duration, buffer_time):
    log_message(f"Waiting for timer completion: {duration} minutes and {buffer_time} seconds.")
    time.sleep(duration * 60 + buffer_time)

def move_window_to_left_screen_and_maximize_upper_half():
    hWnd = win32gui.GetForegroundWindow()
    monitors = win32api.EnumDisplayMonitors()
    if len(monitors) > 1:
        left_monitor = monitors[1][2]

        # Move window to the left screen
        win32gui.SetWindowPos(hWnd, 0, left_monitor[0], left_monitor[1], 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

        # Get the window's current position after moving it
        rect = win32gui.GetWindowRect(hWnd)

        # Calculate the new height to fit the upper half of the left screen
        new_height = (left_monitor[3] - left_monitor[1]) // 2

        # Resize the window to the upper half of the left screen
        win32gui.SetWindowPos(hWnd, 0, rect[0], rect[1], rect[2] - rect[0], new_height, win32con.SWP_NOZORDER)

        log_message("Browser window moved to the left screen and maximized to the upper half.")
    else:
        log_message("Only one monitor detected. Browser window position unchanged.")

with sync_playwright() as p:
    Luxafor.set_color('red', config['luxafor_id'])
    log_message("Luxafor color set to red.")
    browser = p.firefox.launch(headless=False)

    context = browser.new_context(
        viewport={'width': config['viewport_width'], 'height': config['viewport_height']},
        device_scale_factor=float(config['zoom_level']),
        permissions=["notifications"]
    )
    log_message("Browser launched with notifications permission.")
    page = context.new_page()

    move_window_to_left_screen_and_maximize_upper_half()

    appointment_duration_minutes = math.ceil(config['pomodoro_duration_minutes'] + (config['buffer_time_seconds'] / 60.0))
    log_message(f"Calculated appointment duration: {appointment_duration_minutes} minutes")

    if start_pomodoro_timer(page):
        create_appointment(config['meeting_subject'], config['attendees'], appointment_duration_minutes)
        wait_for_timer_completion(config['pomodoro_duration_minutes'], config['buffer_time_seconds'])
    else:
        log_message("Pomodoro timer failed to start. Appointment not created.")

    browser.close()
    log_message("Browser closed.")
    Luxafor.set_color('green', config['luxafor_id'])
    log_message("Luxafor color set to green.")
