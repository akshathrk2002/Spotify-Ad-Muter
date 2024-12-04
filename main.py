import os
import time
import psutil
import win32gui
import logging
import argparse
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# Constants
SPOTIFY = 'Spotify.exe'
CONFIG_FILE = 'ads.cfg'
SLEEP_TIME = 0.01
RELOAD_INTERVAL = 300  # 5 minutes

# Logging Setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('spotify_muter.log', mode='a')
    ]
)

def mute_spotify(mute):
    """Mute or unmute Spotify."""
    try:
        for session in AudioUtilities.GetAllSessions():
            if session.Process and session.Process.name() == SPOTIFY:
                session._ctl.QueryInterface(ISimpleAudioVolume).SetMute(mute, None)
    except Exception as e:
        logging.error(f"Error setting Spotify mute: {e}")

def is_running(process_name):
    """Check if a process is running."""
    return any(proc.name().lower() == process_name.lower() for proc in psutil.process_iter())

def load_ads(config_files):
    """Load ad titles from config files."""
    ads = []
    for file in config_files:
        if os.path.isfile(file):
            with open(file, 'r') as f:
                ads.extend(line.strip() for line in f if line.strip())
        else:
            logging.warning(f"Config file not found: {file}")
    return ads

def reload_ads(ads, config_files):
    """Reload ad titles periodically."""
    while True:
        ads[:] = load_ads(config_files)
        logging.info("Ad list reloaded.")
        time.sleep(RELOAD_INTERVAL)

def main(config_files):
    """Main function."""
    ads = load_ads(config_files)

    # Periodic ad reload in background
    import threading
    threading.Thread(target=reload_ads, args=(ads, config_files), daemon=True).start()

    while True:
        if is_running(SPOTIFY):
            muted = False
            for ad in ads:
                hwnd = win32gui.FindWindowEx(0, 0, 0, ad)
                if hwnd and ad.lower() in win32gui.GetWindowText(hwnd).lower():
                    mute_spotify(True)
                    logging.info("Ad detected OR Spotify muted.")
                    muted = True
                    break
            if not muted:
                mute_spotify(False)
                logging.info("No ads detected OR Spotify unmuted.")
        else:
            logging.info("Spotify is not running.")
        time.sleep(1)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Mute Spotify ads")
    parser.add_argument('-c', '--config-files', nargs='+', default=[CONFIG_FILE],
                        help="Paths to config file(s)")
    args = parser.parse_args()

    try:
        main(args.config_files)
    except Exception as e:
        logging.error(f"Error: {e}")
