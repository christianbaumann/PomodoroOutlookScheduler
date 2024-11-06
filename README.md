# PomodoroOutlookScheduler

Automate scheduling of Pomodoro timers with Outlook meeting integration and Luxafor light notifications.

## Overview

This tool is designed to help users efficiently manage their time by integrating the Pomodoro technique with Outlook calendar scheduling and Luxafor light notifications.

- **Pomodoro Technique**: A time management method where you work for a set period (typically 25 minutes) followed by a short break.
- **Outlook Integration**: Automatically schedules a meeting in Outlook during your Pomodoro session, marking your time as busy.
- **Luxafor Light Notifications**: Indicates your current status (e.g., busy, available) using a Luxafor light.

## Setup

1. Clone this repository.
2. (Optional) Set up a virtual environment:
    ```sh
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```
3. Install required dependencies:
    ```sh
    pip install -r requirements.txt
    ```

4. Install Playwright browsers:
    ```sh
    playwright install
    ```

5. Edit `config.json` with your specific configurations.
6. Run the script:
    ```sh
    python main.py
    ```
