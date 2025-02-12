# GChat2Pdf

A Google Chat Takeout Json to PDF converter tool.
It is written in Python and uses ReportLab to generate the PDFs.

## Overview

- **Download your chats:**  
  Use Google Takeout to download all your chats.

- **Convert to PDF:**  
  Run this program to convert each chat into a separate PDF file.

- **Attached files:**  
  Any files attached to these chats (downloaded as part of the Takeout) are linked into the generated PDFs.

- **Output location:**  
  The script automatically creates a subfolder named `ChatPDFs` inside your specified output directory and saves all generated PDFs there.

- **Group and Participant Names:**  
  - If a chat’s group info lacks a name, it defaults to “Chat”.
  - If an individual participant’s record does not include a name, that person is listed as “Anonymous User”.

---

## Usage

```bash
usage: GChat2Pdf.py [-h] -i IN_DIR -o OUT_DIR [-l LOG_LEVEL] [-s START_DATE] [-e END_DATE] [-z TIME_ZONE] [-p PAPER_SIZE] [-m MAX_FILENAME_LEN] [-a | --all] [-ih MAX_IMG_HEIGHT_IN]
```

### Options

- **-h, --help**  
  Show this help message and exit.

- **-i IN_DIR, --in_dir IN_DIR**  
  Folder that contains **Users** and **Groups** from the Google Takeout folder.  
  *(Default: `Takeout/Google Chat`)*

- **-o OUT_DIR, --out_dir OUT_DIR**  
  Folder where the chat files will be saved.  
  **Note:** A new subfolder called `ChatPDFs` will be created inside this directory to hold all generated PDFs.  
  *(Default: current working directory)*

- **-l LOG_LEVEL, --log_level LOG_LEVEL**  
  Which logging level to show on the terminal.  
  *(Default: INFO)*

- **-s START_DATE, --start_date START_DATE**  
  Start date (YYYY-MM-DD).  
  *(Default: None — accepts any date)*

- **-e END_DATE, --end_date END_DATE**  
  End date (YYYY-MM-DD).  
  *(Default: None — up to today)*

- **-z TIME_ZONE, --time_zone TIME_ZONE**  
  Any pytz timezone (see [pytz.all_timezones](https://pythonhosted.org/pytz/)).  
  *(Default: UTC)*

- **-p PAPER_SIZE, --page_size PAPER_SIZE**  
  Either `'A4'` or `'letter'` are accepted.  
  *(Default: A4)*

- **-m MAX_FILENAME_LEN, --max_filename_length MAX_FILENAME_LEN**  
  Maximum filename length for the generated PDF files.

- **-a, --all**  
  Save files even if you did not participate in the chat.  
  *(Default: False)*

- **-ih MAX_IMG_HEIGHT_IN, --max_img_height_in MAX_IMG_HEIGHT_IN**  
  Maximum height in inches for embedded image thumbnails.  
  *(Default: 2 inches)*

---

## Example Commands

To run the script using the default Google Chat folder from Takeout and save the PDFs in your current directory:

```bash
python3 GChat2Pdf.py
```

Or, to specify custom input and output folders:

```bash
python3 GChat2Pdf.py -i "/path/to/Takeout/Google Chat" -o "/path/to/output"
```

All generated PDFs will be stored in the `/path/to/output/ChatPDFs` folder.

---

Enjoy converting your Google Chat history to nicely formatted PDFs!