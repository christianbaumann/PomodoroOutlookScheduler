import json
import time
import win32com.client
from playwright.sync_api import sync_playwright
from luxafor import Luxafor

# Load Config
with open('config.json') as config_file:
    config = json.load(config_file)

def create_appointment(subject, attendees, duration):
    outlook = win32com.client.Dispatch('Outlook.Application')
    appointment = outlook.CreateItem(1)
    appointment.BusyStatus = 2
    appointment.Start = time.strftime("%Y-%m-%d %H:%M")
    appointment.Duration = duration
    appointment.Subject = subject
    appointment.ReminderSet = False
    appointment.MeetingStatus = 1

    for attendee_email in attendees:
        recipient = appointment.Recipients.Add(attendee_email)
        recipient.Type = 1
        if not recipient.Resolve:
            print(f"Warning: Could not resolve recipient: {attendee_email}")

    try:
        appointment.Save()
        appointment.Send()
        print("Appointment created and sent successfully.")
    except Exception as e:
        print(f"An error occurred while creating the appointment: {e}")

def start_pomodoro_timer(page):
    try:
        page.goto(config['pomofocus_url'])
        page.locator("text=START").first.click()
        page.on("dialog", lambda dialog: dialog.accept())
        return True
    except Exception as e:
        print(f"An error occurred: {e}")
        return False

def wait_for_timer_completion(duration, buffer_time):
    time.sleep(duration * 60 + buffer_time)

with sync_playwright() as p:
    Luxafor.set_color('red', config['luxafor_id'])
    browser = p.firefox.launch(headless=False)
    context = browser.new_context(viewport={'width': 1280, 'height': 1024}, device_scale_factor=float(config['zoom_level']))
    page = context.new_page()

    if start_pomodoro_timer(page):
        create_appointment(config['meeting_subject'], config['attendees'], config['appointment_duration_minutes'])
        wait_for_timer_completion(config['pomodoro_duration_minutes'], config['buffer_time_seconds'])
    else:
        print("Pomodoro timer failed to start. Appointment not created.")

    browser.close()
    Luxafor.set_color('green', config['luxafor_id'])
