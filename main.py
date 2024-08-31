#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import time
import psutil
import win32gui
import argparse
import threading
import logging.config
import re
from subprocess import Popen
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# GLOBAL VARIABLES
SPOTIFY_PID = 4300
SPOTIFY_PROCESS_NAME = 'Spotify.exe'
DEFAULT_CONFIG_FILE = 'ads.cfg'
LOGGING_CONFIG = {
    'version': 1,
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'level': 'INFO',
            'formatter': 'default'
        },
        'file': {
            'class': 'logging.FileHandler',
            'filename': 'spotify_ad_muter.log',
            'level': 'DEBUG',
            'formatter': 'detailed'
        }
    },
    'formatters': {
        'default': {
            'format': '%(levelname)s: %(message)s'
        },
        'detailed': {
            'format': '%(asctime)s %(levelname)s %(module)s %(message)s'
        }
    },
    'root': {
        'level': 'DEBUG',
        'handlers': ['console', 'file']
    }
}

# CONSTANTS
LOOP_SLEEP_TIME = 0.01

# FUNCTIONS
def set_mute(process_name, mute):
    """Mute or unmute the audio of the given process name."""
    try:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.name() == process_name:
                volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                volume.SetMute(mute, None)
    except Exception as e:
        logging.error(f"Error muting audio for process {process_name}: {str(e)}")

def is_process_running(process_name):
    """Check if the given process name is running."""
    try:
        for proc in psutil.process_iter():
            try:
                if process_name.lower() in proc.name().lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    except Exception as e:
        logging.error(f"Error checking if process {process_name} is running: {str(e)}")
    return False

def find_window(title):
    """Find the handle of the window with the given title."""
    try:
        hwnd = win32gui.FindWindowEx(0, 0, 0, title)
        return hwnd
    except Exception as e:
        logging.error(f"Error finding window with title '{title}': {str(e)}")
        return None

def clear():
    """Clear the console."""
    try:
        if os.name == 'nt':
            Popen('cls', shell=True)
    except Exception as e:
        logging.error(f"Error clearing console: {str(e)}")

def line_clear():
    """Clear the current line in the console."""
    try:
        sys.stdout.write('\033[2K\033[1G')
    except Exception as e:
        logging.error(f"Error clearing line in console: {str(e)}")
    return

def load_config(config_files):
    """Load the list of ad titles from the config files."""
    ad_titles = []
    for file in config_files:
        if os.path.isfile(file):
            with open(file, 'r') as f:
                ad_titles.extend(line.strip() for line in f if line.strip())
        else:
            logging.error(f'Config file not found: {file}')
    return ad_titles

def reload_config(ad_titles, config_files):
    """Reload the list of ad titles from the config files."""
    while True:
        new_ad_titles = load_config(config_files)
        if new_ad_titles != ad_titles:
            logging.info(f"Ad titles updated: {new_ad_titles}")
            ad_titles[:] = new_ad_titles
        time.sleep(5*60) # reload every 5 minutes

def main(args):
    """Main entry point of the program."""
    # Setup logging
    logging.config.dictConfig(LOGGING_CONFIG)

    # Set console size and title
    clear()
    Popen('title Spotify Ad Muter', shell=True)

    # Load ad titles from config file(s)
    ad_titles = load_config(args.config_files)

    # Start timer to reload ad titles periodically
    reload_thread = threading.Thread(target=reload_config, args=(ad_titles, args.config_files))
    reload_thread.daemon = True
    reload_thread.start()

    # Loop forever
    while True:
        ad_detected = False

        # Check if Spotify is running
        if is_process_running(SPOTIFY_PROCESS_NAME):
            for title in ad_titles:
                time.sleep(LOOP_SLEEP_TIME)

                # Check if the ad title is shown
                hwnd = find_window(title)
                if hwnd and re.match(title, win32gui.GetWindowText(hwnd), flags=re.IGNORECASE):
                    set_mute(SPOTIFY_PROCESS_NAME, True) # Mute

                    line_clear()
                    logging.info('Ads detected or Spotify is paused')
                    logging.debug(f'Ad title: {win32gui.GetWindowText(hwnd)}')

                    ad_detected = True
                    continue

                if ad_detected:
                    break

            else:
                set_mute(SPOTIFY_PROCESS_NAME, False) # Unmute
                
                line_clear()
                logging.info('No ads detected')

                ad_detected = False
        else:
            line_clear()
            logging.info('Spotify is closed')

        time.sleep(1)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Mute Spotify ads')
    parser.add_argument('-c', '--config-files', nargs='+', default=[DEFAULT_CONFIG_FILE],
                    type=argparse.FileType('r'), help='path(s) to the config file(s)')
    args = parser.parse_args()

    try:
        main(args)
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
        input("Press any key to exit") # pause before closing console window
        sys.exit(1)