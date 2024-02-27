# GChat2Pdf
A Google Takeout Chat to PDF converter which also handles Hebrew.
It is written in python and using reportlab.

First, download all your chats with Google Takeout.
Run the program to extract each chat you had to a different pdf file in the specified output directory.
Files added to these chats (and downloaded into the takeout folder) are linked into the generated pdfs.

This program handles Hebrew right-to-left reversal (although very crudely) if Hebrew is identified in the chats.


    usage: GChat2Pdf.py [-h] -i IN_DIR -o OUT_DIR [-l LOG_LEVEL] [-s START_DATE] [-e END_DATE] [-z TIME_ZONE] [-p PAPER_SIZE] [-m MAX_FILENAME_LEN]  
                    [-a | --all | --no-all] [-ih MAX_IMG_HEIGHT_IN]  

    options:  
      -h, --help            show this help message and exit  
      -i IN_DIR, --in_dir IN_DIR  
                            Google Chat folder within the Google Takeout folder  
      -o OUT_DIR, --out_dir OUT_DIR  
                            Folder where chat files will be saved.  
      -l LOG_LEVEL, --log_level LOG_LEVEL  
      -s START_DATE, --start_date START_DATE  
                            Start date (YYYY-MM-DD). 
                            Default: None (any)
      -e END_DATE, --end_date END_DATE  
                            End date (YYYY-MM-DD).  
                            Default: None (today)
      -z TIME_ZONE, --time_zone TIME_ZONE  
                            Any pytz timezone (look them up).  
      -p PAPER_SIZE, --page_size PAPER_SIZE  
                            'A4' or 'letter' are accepted.  
                            Default: A4
      -m MAX_FILENAME_LEN, --max_filename_length MAX_FILENAME_LEN  
                            Max filename length  
      -a, --all, --no-all   Save files which don't include me participating in the chats.  
                            Default: False (no-all)
      -ih MAX_IMG_HEIGHT_IN, --max_img_height_in MAX_IMG_HEIGHT_IN  
                            Maximum height in inches for embedded image thumbnails.  
                            Default: 2 inches.
