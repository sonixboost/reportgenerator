import os
import sys
import requests
import shutil

GITHUB_REPO: str = "https://raw.githubusercontent.com/sonixboost/reportgenerator/main"
VERSION_FILE_URL: str = f"{GITHUB_REPO}/version.txt"
CURRENT_VERSION: str = "3.0.0"  # Update this with the current version of app
UPDATE_FILE_URL: str = f"{GITHUB_REPO}/NikkiReportGenerator_V2.1.1.exe"
EXE_NAME = "NikkiReportGenerator_V2.1.1.exe"

def check_for_update():
    try:
        print("Checking for update...")
        response = requests.get(VERSION_FILE_URL)
        latest_version = response.text.strip()
        if latest_version > CURRENT_VERSION:
            return latest_version
        else:
            return None
    except Exception as e:
        print(f"Error checking for updates: {e}")
        return None

def download_update():
    try:
        response = requests.get(UPDATE_FILE_URL, stream=True)
        with open("update.exe", "wb") as f:
            for chunk in response.iter_content(chunk_size=128):
                f.write(chunk)
        return "update.exe"
    except Exception as e:
        print(f"Error downloading update: {e}")
        return None

def apply_update(update_file):
    try:
        os.rename(EXE_NAME, EXE_NAME + ".old")  # Backup old file
        shutil.move(update_file, EXE_NAME)  # Replace old .exe with new one
        print("Update applied successfully.")
        # Restart application
        os.execl(EXE_NAME, EXE_NAME, *sys.argv)
    except Exception as e:
        print(f"Error applying update: {e}")
        if os.path.exists(EXE_NAME + ".old"):
            os.rename(EXE_NAME + ".old", EXE_NAME)  # Restore old file if update fails

def main():
    latest_version = check_for_update()
    if latest_version:
        print(f"Update available: {latest_version}")
        # update_file = download_update()
        # if update_file:
        #     apply_update(update_file)
    else:
        print("No update available.")

if __name__ == "__main__":
    main()
