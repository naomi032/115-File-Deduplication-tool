# CloudDrive2 Duplicate File Finder & Deleter

A graphical tool built with Python and Tkinter to scan your CloudDrive2 mount for duplicate files (based on SHA1 hash) and help you delete unwanted copies based on user-defined rules.

## Features

*   **Graphical User Interface:** Easy-to-use interface built with Tkinter.
*   **CloudDrive2 Integration:** Connects directly to your CloudDrive2 instance via its API (`clouddrive` library).
*   **Duplicate Detection:** Identifies duplicate video files based on their SHA1 hash. Only files with identical SHA1 hashes are considered duplicates.
*   **Configurable Scan Path:** Specify the root directory within your CloudDrive2 mount to scan.
*   **Rule-Based Deletion:** Choose which copy to keep from each duplicate set based on:
    *   Shortest file path
    *   Longest file path
    *   Oldest file (based on modification date)
    *   Newest file (based on modification date)
    *   Files ending with a specific suffix (e.g., keep `.mkv`)
*   **Visual Feedback:** The list clearly shows which files are marked to "Keep" and which are marked to "Delete" based on the selected rule.
*   **Safety Confirmation:** Prompts for confirmation before performing any deletions, clearly stating the rule being applied and the number of files affected.
*   **Logging:** Provides real-time feedback on the scanning, rule application, and deletion processes.
*   **Connection Testing:** Verify your CloudDrive2 connection details before starting a scan.
*   **Save Report:** Export the list of found duplicate sets (including paths, dates, sizes) to a text file.
*   **File Type Chart (Optional):** Visualize the distribution of file types in the scanned path (requires `matplotlib`).
*   **Multi-Language Support:** Includes English and Chinese (中文) interfaces. Language preference is saved.
*   **Customization:** Supports a custom background image (`background.png`) and application icon (`app_icon.ico`).
*   **Configuration File:** Saves connection details and paths to `config.ini` for easy reuse.

## Dependencies

*   **Python:** 3.7+ (due to f-strings, date parsing, and modern library usage)
*   **clouddrive:** The core library for interacting with CloudDrive2.
    ```bash
    pip install clouddrive
    ```
*   **Pillow:** Required for loading the custom background image and application icon.
    ```bash
    pip install Pillow
    ```
*   **matplotlib (Optional):** Required *only* for the "Show Cloud File Types" chart feature.
    ```bash
    pip install matplotlib
    ```

## Installation

1.  **Clone or Download:** Get the project files:
    ```bash
    git clone <repository_url> # Or download the ZIP
    cd <project_directory>
    ```
2.  **Create Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```
3.  **Install Dependencies:**
    ```bash
    pip install clouddrive Pillow matplotlib # Install all, including optional matplotlib
    # OR (if only required are needed):
    # pip install clouddrive Pillow
    ```
4.  **Prepare Custom Files (Optional):**
    *   Place your desired background image as `background.png` in the same directory as the script. For best transparency *effect*, pre-process the image by blending it with a light grey background (#F0F0F0) to your desired opacity level in an image editor.
    *   Place your desired application icon as `app_icon.ico` in the same directory.

## Configuration (`config.ini`)

Before running, you need to configure the connection details. You can either:

*   Fill the fields directly in the GUI and click "Save Config".
*   Create/edit the `config.ini` file manually in the same directory as the script.

The `config.ini` file should look like this:

```ini
[config]
clouddrvie2_address = http://127.0.0.1:19798  # Your CloudDrive2 API address
clouddrive2_account = your_username           # Your CloudDrive2 login username (if needed)
clouddrive2_passwd = your_password             # Your CloudDrive2 login password (if needed)
root_path = D:/CloudDrive/Media                # The path to scan AS SEEN BY THIS SCRIPT/OS
clouddrive2_root_path = D:/CloudDrive          # The CloudDrive2 Mount Point Path AS CONFIGURED IN CloudDrive2


Explanation of Paths (Important!):

clouddrvie2_address: The URL where your CloudDrive2 API is accessible.

clouddrive2_account, clouddrive2_passwd: Your CloudDrive2 login credentials (leave blank if not required by your setup).

root_path: This is the full path to the directory you want to start scanning from the perspective of the operating system running this script. For example, if your CloudDrive is mounted as D:\CloudDrive on Windows, and you want to scan the Media folder inside it, this would be D:/CloudDrive/Media or D:\CloudDrive\Media.

clouddrive2_root_path: This is the mount point path that you configured within the CloudDrive2 application itself. It's used to calculate the relative path within the cloud drive. Examples:

If CloudDrive2 is configured to mount the entire cloud as D:, this should be D: or D:/.

If CloudDrive2 is configured to mount the entire cloud at /mnt/clouddrive on Linux, this should be /mnt/clouddrive.

If CloudDrive2 only mounts a subfolder (e.g., MyMedia) from the cloud to D:\, this value should still likely represent the root mount point (D:\ in this case), and root_path would be D:\. The script calculates the relative path based on how root_path relates to clouddrive2_root_path.

Path Format Note: It's generally safer to use forward slashes (/) for paths in the config file, even on Windows, or use double backslashes (\\). Avoid copy-pasting paths directly from file explorers if possible, as it can sometimes include invisible characters. Manually typing paths is recommended if you encounter connection or scanning errors.

Usage
Launch: Run the Python script: python your_script_name.py.

Configure:

Fill in the CloudDrive2 Address, Account (if needed), Password (if needed), Root Path to Scan, and Mount Point fields.

Alternatively, if you have a config.ini file, click "Load Config".

Click "Save Config" to save the current settings to config.ini.

Test Connection: Click "Test Connection". Check the log area for success or failure messages. Address any configuration issues if it fails.

Find Duplicates: Once connected, click "Find Duplicates". The script will scan the specified path. This may take time depending on the number of files. Progress is shown in the log.

Review Results: When the scan completes, the list below will populate with sets of duplicate video files found.

Select Deletion Rule: Choose one of the rules under "Deletion Rules (Select ONE to Keep)".

Enter Suffix (If Applicable): If you selected "Keep files ending with:", enter the desired suffix (e.g., .mkv, -keep.mp4) in the "Suffix:" entry box.

Observe Actions: The "Action" column in the list will update to show "Keep" or "Delete" based on the selected rule. Files marked "Delete" (usually shown in red) are the ones targeted for removal.

Delete Files (Carefully!):

Verify the rule and the files marked for deletion.

Click the "Delete Files by Rule" button (usually styled red).

Confirm Deletion: A confirmation dialog will appear, stating the rule and the number of files to be deleted. Read this carefully. Click "Yes" to proceed with permanent deletion, or "No" to cancel.

Monitor Deletion: The log will show the progress of the deletion process.

Optional Actions:

Save Found Duplicates Report: Click this to save a text file listing all the duplicate sets found (before deletion).

Show Cloud File Types: Click this (if matplotlib is installed and you are connected) to see a pie chart of file extensions in the scanned path.

Building an Executable (Optional)
You can use PyInstaller to create a standalone executable. The script uses resource_path to help locate config.ini, background.png, app_icon.ico, and lang_pref.json when bundled.

A basic PyInstaller command might look like this:

# Make sure you have PyInstaller installed: pip install pyinstaller
# Navigate to your script's directory first

pyinstaller --onefile --windowed \
  --add-data "config.ini;." \
  --add-data "background.png;." \
  --add-data "app_icon.ico;." \
  --add-data "lang_pref.json;." \
  --icon="app_icon.ico" \
  your_script_name.py

# --onefile: Creates a single executable file.
# --windowed: Prevents a console window from appearing.
# --add-data "SOURCE;DEST": Bundles necessary data files. (Use ';' on Windows, ':' on Linux/macOS)
# --icon: Sets the executable's icon.
Use code with caution.
Bash
The executable will be located in the dist folder created by PyInstaller. Note that you might need to adjust --add-data paths and the separator (; or :) depending on your OS and file locations.

Important Notes & Warnings
DELETION IS PERMANENT: Files deleted by this tool are permanently removed from your CloudDrive2 storage. There is usually NO RECYCLE BIN. Use this tool with extreme caution.

BACKUP YOUR DATA: Always have backups of important data before running deletion tools.

SHA1 HASH ONLY: Duplicates are identified solely based on the SHA1 hash provided by CloudDrive2. Files with different content but the same name/size/date will not be flagged as duplicates. Files with identical content but different names will be flagged.

PATH CONFIGURATION IS CRITICAL: Ensure Root Path to Scan and CloudDrive2 Mount Point are set correctly according to the explanations above. Incorrect paths can lead to scan failures or unexpected behavior.

PERFORMANCE: Scanning large CloudDrive2 shares, especially over slower connections, can take a significant amount of time. The GUI might appear unresponsive during intense scanning or deletion phases – check the log for progress.

LANGUAGE PREFERENCE: The selected language is saved in lang_pref.json in the application's directory.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal


Acknowledgements
Built using Python, Tkinter.

Uses the clouddrive, Pillow, and matplotlib libraries.
