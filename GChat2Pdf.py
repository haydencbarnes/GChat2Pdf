from pathlib import Path
import argparse
import sys
import logging
import ijson  # streaming from file
import json  # all in memory
import datetime as dt
import pytz
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib import utils as rlutils
from reportlab.lib.units import inch
from io import BytesIO
from pillow_heif import register_heif_opener  # heic file reader
import fitz  # pdf thumbnails
import os

TRUNC_FILE_NAME = 47  # file names on disk are truncated to this basename (stem)

DM_PREFIX = "DM"
SPACE_PREFIX = "Space"
GROUPS_DIR = "Groups"
USERS_DIR = "Users"
USER_INFO_FILE = "user_info.json"
MESSAGES_FILE = "messages.json"
GROUP_INFO_FILE = "group_info.json"
MSG_STATE_DELETED = "DELETED"

PDF_TMP_FILE = "pdf_1st_page.png"


class HyperlinkedImage(Image, object):
    def __init__(
        self,
        filename,
        hyperlink=None,
        width=None,
        height=None,
        kind="direct",
        mask="auto",
        lazy=1,
        hAlign="CENTER",
    ):
        super(HyperlinkedImage, self).__init__(
            filename, width, height, kind, mask, lazy, hAlign=hAlign
        )
        self.hyperlink = hyperlink

    def drawOn(self, canvas, x, y, _sW=0):
        if self.hyperlink:
            x1 = self._hAlignAdjust(x, _sW)
            y1 = y
            x2 = x1 + self._width
            y2 = y1 + self._height
            canvas.linkURL(
                url=self.hyperlink, rect=(x1, y1, x2, y2), thickness=0, relative=1
            )
        super(HyperlinkedImage, self).drawOn(canvas, x, y, _sW)


class CChat2Pdf:
    def __init__(self, args):
        self.args = args
        logging.basicConfig(
            level=args.log_level,
            format=" {asctime}.{msecs:03.0f} {levelname: <9} {message}",
            datefmt="%Y-%m-%d %H:%M:%S",
            style="{",
        )
        self.logger = logging.getLogger("Chat2Pdf")
        if not self.args.in_dir.is_dir():
            self.logger.error(f"Can't open/find input folder {self.args.in_dir}.")
            sys.exit(0)
        try:
            self.args.out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.logger.error(
                f"Can't open/create output folder {self.args.out_dir}: {e}"
            )
            sys.exit(0)
        # Create a new folder inside the output folder for all PDFs.
        self.output_folder = self.args.out_dir / "ChatPDFs"
        try:
            self.output_folder.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.logger.error(
                f"Couldn't create output subfolder {self.output_folder}: {e}"
            )
            sys.exit(0)

        self.page_width = (letter if self.args.paper_size == "letter" else A4)[0]
        register_heif_opener()
        self.style_sheets = getSampleStyleSheet()
        self.style_sheets.add(
            ParagraphStyle(
                name="MeHeader", parent=self.style_sheets["Heading4"], alignment=TA_LEFT
            )
        )
        self.style_sheets.add(
            ParagraphStyle(
                name="OtherHeader",
                parent=self.style_sheets["Heading4"],
                alignment=TA_RIGHT,
            )
        )
        self.style_sheets.add(
            ParagraphStyle(
                name="MeNormal",
                parent=self.style_sheets["Normal"],
                alignment=TA_JUSTIFY,
                rightIndent=2.0 * inch,
            )
        )
        self.style_sheets.add(
            ParagraphStyle(
                name="OtherNormal",
                parent=self.style_sheets["Normal"],
                alignment=TA_JUSTIFY,
                leftIndent=2.0 * inch,
            )
        )
        self.unk_file_exts = set()
        self.logger.info("Init success.")

    def GetScaledImage(self, img_path_url, orig_path_url=None):
        # Convert Path objects to strings to avoid 'WindowsPath' object is not subscriptable error
        if isinstance(img_path_url, Path):
            img_path_url = str(img_path_url)
        
        try:
            img = rlutils.ImageReader(img_path_url)
            iw, ih = img.getSize()
            aspect = ih / float(iw)
            if ih > self.args.max_img_height_in * inch:  # shrink height first
                ih = self.args.max_img_height_in * inch
                iw = ih / aspect
            if (
                iw > self.page_width - 1.5 * inch
            ):  # shrink width if needed (1.5 is the left+right margins)
                iw = self.page_width - 1.5 * inch
                ih = iw * aspect
            if orig_path_url is None:
                orig_path_url = img_path_url
            elif isinstance(orig_path_url, Path):
                orig_path_url = str(orig_path_url)
            
            return HyperlinkedImage(
                img_path_url,
                hyperlink=str(orig_path_url),
                width=iw,
                height=ih,
                hAlign="CENTER",
            )
        except Exception as e:
            self.logger.warning(f"Error in GetScaledImage: {e}")
            raise


    def preprocess_text(self, text):
        """
        Preprocesses text for PDF display.
        Currently just replaces tabs and newlines for HTML formatting.
        """
        text = text.replace("\t", "&nbsp;" * 5)
        text = text.replace("\n", "<br />")
        return text

    def sanitize_filename(self, filename):
        """
        Sanitizes a filename to be safe for both Mac and Windows file systems.

        Args:
            filename: The original filename string

        Returns:
            A sanitized filename string
        """
        # Replace path separators with underscores
        filename = filename.replace("/", "_").replace("\\", "_")

        # Remove characters invalid in both Windows and Mac
        invalid_chars = '<>:"|?*'
        for char in invalid_chars:
            filename = filename.replace(char, "_")

        # Also handle Windows reserved names (CON, PRN, AUX, etc.)
        reserved_names = [
            "CON",
            "PRN",
            "AUX",
            "NUL",
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9",
            "LPT1",
            "LPT2",
            "LPT3",
            "LPT4",
            "LPT5",
            "LPT6",
            "LPT7",
            "LPT8",
            "LPT9",
        ]

        # Check if the filename (without extension) matches a reserved name
        name_parts = filename.split(".")
        if name_parts[0].upper() in reserved_names:
            name_parts[0] = name_parts[0] + "_"
            filename = ".".join(name_parts)

        # Remove leading and trailing spaces and periods (problematic in Windows)
        filename = filename.strip(" .")

        # Ensure the filename isn't empty after sanitization
        if not filename:
            filename = "unnamed_chat"

        # Limit filename length (255 is generally safe for most modern filesystems)
        # But we'll use a more conservative limit
        max_length = min(self.args.max_filename_len, 240)
        if len(filename) > max_length:
            # If there's an extension, preserve it
            if "." in filename[-5:]:  # Check last 5 chars for extension
                name, ext = filename.rsplit(".", 1)
                filename = name[: max_length - len(ext) - 1] + "." + ext
            else:
                filename = filename[:max_length]

        return filename

    def CreateOutput(self, dm_dir):
        msg_file_path = dm_dir.joinpath(MESSAGES_FILE)
        if not msg_file_path.exists():
            self.logger.debug(f"No messages in {dm_dir.name}.")
            return False
        
        grp_info_file_path = dm_dir.joinpath(GROUP_INFO_FILE)
        if not grp_info_file_path.exists():
            self.logger.warning(f"No group info file in {dm_dir.name}.")
            return False
        
        # Log the processing of this chat
        self.logger.info(f"Processing chat: {dm_dir.name}")
        
        # Generate a unique ID for this chat based on directory name
        chat_id = dm_dir.name.replace(" ", "_")
        
        try:
            with open(grp_info_file_path, "r", encoding="utf-8") as inf:
                group_info = json.load(inf)
                # Use group name if available; otherwise, default to "Chat"
                title_str = group_info.get("name", "").strip() or "Chat"
                
                # Keep track of unique participants to avoid duplicates
                unique_participants = set()
                unique_participants.add(self.user_name)  # Add yourself
                
                # Build participants string and collect unique names
                participants_str = f"<u>Participants:</u><br />\t{self.user_name} ({self.user_email})<br />"
                
                # Process all members, including anonymous ones
                for participant in group_info.get("members", []):
                    if isinstance(participant, dict):
                        # Only default to "Anonymous User" if no name is provided
                        name = participant.get("name") or "Anonymous User"
                        email = participant.get("email", "")
                    else:
                        name = "Anonymous User"
                        email = ""
                    
                    # Skip yourself (already added)
                    if name == self.user_name:
                        continue
                        
                    # Add to participants string if not already included
                    if name not in unique_participants:
                        unique_participants.add(name)
                        participants_str += f"\t{name}" + (f" ({email})" if email else "") + "<br />"
                
                # Create a unique filename that includes chat ID and participant count
                other_count = len(unique_participants) - 1  # Exclude yourself
                
                if other_count == 0:
                    # Handle case where you're the only participant (notes to self)
                    file_name = f"{title_str} (notes-{chat_id}).pdf"
                elif other_count == 1:
                    # For 1-on-1 chats, include the other person's name
                    other_name = next(name for name in unique_participants if name != self.user_name)
                    file_name = f"{title_str} with {other_name} ({chat_id}).pdf"
                else:
                    # For group chats, include the count of participants
                    file_name = f"{title_str} with {other_count} others ({chat_id}).pdf"
                
                participants_str = participants_str.replace("\t", "&nbsp;" * 5)
        except Exception as e:
            self.logger.error(f"Error processing group info for {dm_dir.name}: {e}")
            # Default filename if group info can't be processed
            file_name = f"Chat {chat_id}.pdf"
            participants_str = f"<u>Participants:</u><br />\t{self.user_name} ({self.user_email})<br />"

        # Always set I_participated to True if --all flag is used
        file_created = False
        I_participated = self.args.include_all  # Start with True if --all flag is used
        doc_components = []
        img_file_names = {}
        
        # Count total messages for debugging
        total_messages = 0
        processed_messages = 0
        
        try:
            with open(msg_file_path, "rb") as inf:
                for msg in ijson.items(inf, "messages.item"):
                    total_messages += 1
                    
                    if not (
                        "message_state" in msg and msg["message_state"] == MSG_STATE_DELETED
                    ):
                        try:
                            msg_dt = dt.datetime.strptime(
                                msg["created_date"].replace("\u202f", ""),
                                "%A, %B %d, %Y at %I:%M:%S%p %Z",
                            )
                            msg_dt = pytz.utc.localize(msg_dt, is_dst=None).astimezone(
                                pytz.timezone(self.args.time_zone)
                            )
                            msg_dt_str = msg_dt.strftime("%Y-%m-%d %H:%M:%S %Z%z")
                            msg_d = msg_dt.date()
                            
                            # Check date filters
                            if (
                                self.args.start_date is None or msg_d >= self.args.start_date
                            ) and (self.args.end_date is None or msg_d <= self.args.end_date):
                                processed_messages += 1
                                
                                # Check if you participated
                                if msg["creator"]["name"] == self.user_name:
                                    I_participated = True
                                    
                                # Create PDF components
                                if not file_created:
                                    pdf_io_buffer = BytesIO()
                                    output_buffer = SimpleDocTemplate(
                                        pdf_io_buffer,
                                        pagesize=letter
                                        if self.args.paper_size == "letter"
                                        else A4,
                                        rightMargin=0.75 * inch,
                                        leftMargin=0.75 * inch,
                                        topMargin=1.0 * inch,
                                        bottomMargin=0.75 * inch,
                                    )
                                    doc_components.append(
                                        Paragraph(title_str, self.style_sheets["Title"])
                                    )
                                    doc_components.append(
                                        Paragraph(
                                            participants_str, self.style_sheets["Heading5"]
                                        )
                                    )
                                    file_created = True

                                header_str = (
                                    msg["creator"]["name"]
                                    + (
                                        (" (" + msg["creator"]["email"] + ")")
                                        if "email" in msg["creator"]
                                        else ""
                                    )
                                    + " at "
                                    + msg_dt_str
                                    + ":"
                                )
                                doc_components.append(
                                    Paragraph(
                                        header_str,
                                        self.style_sheets["MeHeader"]
                                        if msg["creator"]["name"] == self.user_name
                                        else self.style_sheets["OtherHeader"],
                                    )
                                )
                                try:
                                    if "text" in msg:
                                        text = self.preprocess_text(
                                            msg["text"]
                                        )
                                        style_key = (
                                            "MeNormal"
                                            if msg["creator"]["name"] == self.user_name
                                            else "OtherNormal"
                                        )
                                        doc_components.append(
                                            Paragraph(text, self.style_sheets[style_key])
                                        )
                                    elif "attached_files" in msg:
                                        for i, f in enumerate(msg["attached_files"]):
                                            try:
                                                # Get the export name and strip any leading/trailing spaces and periods
                                                export_name = f["export_name"].strip(" .")
                                                
                                                # First try the direct path
                                                img_file_path = dm_dir.joinpath(export_name)
                                                
                                                # If the file doesn't exist, try with the truncated filename
                                                if not img_file_path.exists() and len(export_name) > TRUNC_FILE_NAME:
                                                    # Try with truncated filename (Google Takeout often truncates filenames)
                                                    truncated_name = export_name[:TRUNC_FILE_NAME] + export_name[export_name.rfind('.'):]
                                                    alt_path = dm_dir.joinpath(truncated_name)
                                                    if alt_path.exists():
                                                        img_file_path = alt_path
                                                        self.logger.debug(f"Using truncated filename: {truncated_name}")
                                                
                                                # Check if file exists before processing
                                                if not img_file_path.exists():
                                                    # Try to find the file by searching for similar filenames
                                                    potential_files = list(dm_dir.glob(f"*{img_file_path.suffix}"))
                                                    matching_files = [f for f in potential_files if export_name in f.name or f.name in export_name]
                                                    
                                                    if matching_files:
                                                        img_file_path = matching_files[0]
                                                        self.logger.debug(f"Found alternative file: {img_file_path.name}")
                                                    else:
                                                        # If file still not found, add a file link instead
                                                        file_link_str = (
                                                            f'<u>File attachment not found:</u> {export_name}'
                                                        )
                                                        doc_components.append(
                                                            Paragraph(
                                                                file_link_str,
                                                                self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                                else self.style_sheets["OtherNormal"],
                                                            )
                                                        )
                                                        continue  # Skip to next file
                                                
                                                # Process the file based on its type
                                                if img_file_path.suffix.lower() in [".jpg", ".png", ".jpeg", ".heic", ".gif", ".eps"]:
                                                    try:
                                                        doc_components.append(self.GetScaledImage(img_file_path))
                                                    except Exception as e:
                                                        self.logger.warning(f"Error processing image {img_file_path}: {e}")
                                                        file_link_str = (
                                                            f'<u>File attachment (error processing):</u> {export_name}'
                                                        )
                                                        doc_components.append(
                                                            Paragraph(
                                                                file_link_str,
                                                                self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                                else self.style_sheets["OtherNormal"],
                                                            )
                                                        )
                                                elif img_file_path.suffix.lower() == ".pdf":
                                                    try:
                                                        with fitz.open(img_file_path) as doc:
                                                            page = doc.load_page(0)
                                                            pix = page.get_pixmap()
                                                            pix.save(PDF_TMP_FILE)
                                                        doc_components.append(self.GetScaledImage(PDF_TMP_FILE, img_file_path))
                                                    except Exception as e:
                                                        self.logger.warning(f"Could not open PDF for thumbnail: {e}")
                                                        file_link_str = (
                                                            '<u>File attached:</u> <link href="'
                                                            + str(img_file_path)
                                                            + '">'
                                                            + img_file_path.name
                                                            + "</link>"
                                                        )
                                                        doc_components.append(
                                                            Paragraph(
                                                                file_link_str,
                                                                self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                                else self.style_sheets["OtherNormal"],
                                                            )
                                                        )
                                                else:
                                                    if img_file_path.suffix not in self.unk_file_exts:
                                                        suffix = img_file_path.suffix
                                                        self.logger.warning(f"File extension '{suffix}' without a thumbnail preview found.")
                                                        self.unk_file_exts.add(suffix)
                                                    file_link_str = (
                                                        '<u>File attached:</u> <link href="'
                                                        + str(img_file_path)
                                                        + '">'
                                                        + img_file_path.name
                                                        + "</link>"
                                                    )
                                                    doc_components.append(
                                                        Paragraph(
                                                            file_link_str,
                                                            self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                            else self.style_sheets["OtherNormal"],
                                                        )
                                                    )
                                            except Exception as e:
                                                self.logger.warning(f"Error processing attachment: {e}")
                                                file_link_str = f'<u>Error processing attachment:</u> {str(e)}'
                                                doc_components.append(
                                                    Paragraph(
                                                        file_link_str,
                                                        self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                        else self.style_sheets["OtherNormal"],
                                                    )
                                                )
                                    elif "annotations" in msg:
                                        if "video_call_metadata" in msg["annotations"][0]:
                                            doc_components.append(
                                                Paragraph(
                                                    "<u>Video call started.</u>",
                                                    self.style_sheets["MeNormal"]
                                                    if msg["creator"]["name"] == self.user_name
                                                    else self.style_sheets["OtherNormal"],
                                                )
                                            )
                                        elif (
                                            "gsuite_integration_metadata"
                                            in msg["annotations"][0]
                                        ):
                                            if (
                                                "call_data"
                                                in msg["annotations"][0][
                                                    "gsuite_integration_metadata"
                                                ]
                                            ):
                                                doc_components.append(
                                                    Paragraph(
                                                        f"<u>{msg['annotations'][0]['gsuite_integration_metadata']['call_data']['call_status']}</u>",
                                                        self.style_sheets["MeNormal"]
                                                        if msg["creator"]["name"]
                                                        == self.user_name
                                                        else self.style_sheets["OtherNormal"],
                                                    )
                                                )
                                            elif (
                                                "tasks_data"
                                                in msg["annotations"][0][
                                                    "gsuite_integration_metadata"
                                                ]
                                            ):
                                                task_str = (
                                                    'Task "'
                                                    + msg["annotations"][0][
                                                        "gsuite_integration_metadata"
                                                    ]["tasks_data"]["task_properties"]["title"]
                                                    + '"'
                                                )
                                                if (
                                                    "assignee"
                                                    in msg["annotations"][0][
                                                        "gsuite_integration_metadata"
                                                    ]["tasks_data"]["task_properties"]
                                                ):
                                                    task_str += (
                                                        " assigned to "
                                                        + msg["annotations"][0][
                                                            "gsuite_integration_metadata"
                                                        ]["tasks_data"]["task_properties"][
                                                            "assignee"
                                                        ]["id"]
                                                    )
                                                if (
                                                    "assignee_change"
                                                    in msg["annotations"][0][
                                                        "gsuite_integration_metadata"
                                                    ]["tasks_data"]
                                                ):
                                                    task_str += (
                                                        " removed from "
                                                        + msg["annotations"][0][
                                                            "gsuite_integration_metadata"
                                                        ]["tasks_data"]["assignee_change"][
                                                            "old_assignee"
                                                        ]["id"]
                                                    )
                                                if msg["annotations"][0][
                                                    "gsuite_integration_metadata"
                                                ]["tasks_data"]["task_properties"]["completed"]:
                                                    task_str += " completed."
                                                elif msg["annotations"][0][
                                                    "gsuite_integration_metadata"
                                                ]["tasks_data"]["task_properties"]["deleted"]:
                                                    task_str += " deleted."
                                                else:
                                                    task_str += "."
                                                doc_components.append(
                                                    Paragraph(
                                                        f"<u>{task_str}</u>",
                                                        self.style_sheets["MeNormal"]
                                                        if msg["creator"]["name"]
                                                        == self.user_name
                                                        else self.style_sheets["OtherNormal"],
                                                    )
                                                )
                                            else:
                                                self.logger.warning(
                                                    f"Unknown type under gsuite_integration_metadata. Message (complete):\n{msg}"
                                                )
                                        elif "url_metadata" in msg["annotations"][0]:
                                            try:
                                                # Check if image_url exists in the metadata
                                                if "image_url" in msg["annotations"][0]["url_metadata"]:
                                                    image_url = msg["annotations"][0]["url_metadata"]["image_url"]
                                                    
                                                    # Skip broken Google proxy URLs
                                                    if "googleusercontent.com/proxy" in image_url:
                                                        # Add a text link instead of trying to load the image
                                                        url_title = msg["annotations"][0]["url_metadata"].get("title", "Link")
                                                        url_link = msg["annotations"][0]["url_metadata"].get("url", image_url)
                                                        
                                                        url_link_str = f'<u>URL shared:</u> <link href="{url_link}">{url_title}</link>'
                                                        doc_components.append(
                                                            Paragraph(
                                                                url_link_str,
                                                                self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                                else self.style_sheets["OtherNormal"],
                                                            )
                                                        )
                                                    else:
                                                        # Try to load the image
                                                        try:
                                                            doc_components.append(self.GetScaledImage(image_url))
                                                        except Exception as e:
                                                            self.logger.warning(f"Could not load image from URL {image_url}: {e}")
                                                            # Fall back to text link
                                                            url_title = msg["annotations"][0]["url_metadata"].get("title", "Link")
                                                            url_link = msg["annotations"][0]["url_metadata"].get("url", image_url)
                                                            
                                                            url_link_str = f'<u>URL shared (image unavailable):</u> <link href="{url_link}">{url_title}</link>'
                                                            doc_components.append(
                                                                Paragraph(
                                                                    url_link_str,
                                                                    self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                                    else self.style_sheets["OtherNormal"],
                                                                )
                                                            )
                                                else:
                                                    # No image URL, just add a text link
                                                    url_title = msg["annotations"][0]["url_metadata"].get("title", "Link")
                                                    url_link = msg["annotations"][0]["url_metadata"].get("url", "#")
                                                    
                                                    url_link_str = f'<u>URL shared:</u> <link href="{url_link}">{url_title}</link>'
                                                    doc_components.append(
                                                        Paragraph(
                                                            url_link_str,
                                                            self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                            else self.style_sheets["OtherNormal"],
                                                        )
                                                    )
                                            except Exception as e:
                                                self.logger.warning(f"Error processing URL metadata: {e}")
                                                url_link_str = f'<u>URL shared (error processing):</u> {str(e)}'
                                                doc_components.append(
                                                    Paragraph(
                                                        url_link_str,
                                                        self.style_sheets["MeNormal"] if msg["creator"]["name"] == self.user_name
                                                        else self.style_sheets["OtherNormal"],
                                                    )
                                                )
                                        elif "drive_metadata" in msg["annotations"][0]:
                                            doc_components.append(
                                                Paragraph(
                                                    f"<u>File shared from google drive: {msg['annotations'][0]['drive_metadata']['title']} (file id:{msg['annotations'][0]['drive_metadata']['id']})</u>",
                                                    self.style_sheets["MeNormal"]
                                                    if msg["creator"]["name"] == self.user_name
                                                    else self.style_sheets["OtherNormal"],
                                                )
                                            )
                                        else:
                                            self.logger.warning(
                                                f"Unknown type under annotations. Message (complete):\n{msg}"
                                            )
                                    else:
                                        self.logger.warning(
                                            f"Unknown type. Message (complete): {msg}"
                                        )
                                except ValueError:
                                    pass
                                except Exception as e:
                                    self.logger.error(f"Error processing message content: {e}")
                        except Exception as e:
                            self.logger.warning(f"Error processing message in {dm_dir.name}: {e}")
                            continue
            
            # Log message counts
            self.logger.info(f"Chat {dm_dir.name}: Total messages: {total_messages}, Processed: {processed_messages}")
            
            # Check if we should create the file
            if file_created and (self.args.include_all or I_participated):
                self.logger.info(f"Creating PDF for {dm_dir.name}: {file_name}")
                output_buffer.build(doc_components)
                
                # Sanitize the filename before writing
                sanitized_file_name = self.sanitize_filename(file_name)
                
                # Create the output path
                output_path = self.output_folder.joinpath(sanitized_file_name)
                
                # Log if the filename was changed
                if sanitized_file_name != file_name:
                    self.logger.info(
                        f"Sanitized filename from '{file_name}' to '{sanitized_file_name}'"
                    )
                
                # Write the file
                with open(str(output_path), "wb") as outfile:
                    outfile.write(pdf_io_buffer.getbuffer())
                    return True  # Successfully created a file
            else:
                if not file_created:
                    self.logger.info(f"No file created for {dm_dir.name}: No messages matched date criteria")
                elif not I_participated and not self.args.include_all:
                    self.logger.info(f"No file created for {dm_dir.name}: You did not participate and --all flag not used")
                return False
                
        except Exception as e:
            self.logger.error(f"Error processing chat {dm_dir.name}: {e}")
            return False

    def run(self):
        users_dir = self.args.in_dir.joinpath(USERS_DIR)
        if not users_dir.is_dir():
            self.logger.error("Couldn't find users folder. Exiting.")
            sys.exit(0)
        subdirs = [p for p in users_dir.iterdir() if p.is_dir()]
        if not subdirs:
            self.logger.error(
                f"No valid user subdirectory found in {users_dir}. Exiting."
            )
            sys.exit(0)
        user_subdir = subdirs[0]
        user_info_file_path = user_subdir.joinpath(USER_INFO_FILE)
        if not user_info_file_path.exists():
            self.logger.error(
                f"Couldn't find {USER_INFO_FILE} in {user_subdir}. Exiting."
            )
            sys.exit(0)
        with open(user_info_file_path, "rb") as inf:
            user_info = json.load(inf)
            self.user_name = user_info["user"]["name"]
            self.user_email = user_info["user"]["email"]
        self.logger.info(f"You are {self.user_name} ({self.user_email})")
        groups_dir = self.args.in_dir.joinpath(GROUPS_DIR)
        if not groups_dir.is_dir():
            self.logger.error("Couldn't find groups folder. Exiting.")
            sys.exit(0)
        dirs = list(groups_dir.iterdir())
        num_dirs = len(dirs)
        self.logger.info(f"Found {num_dirs} chats/spaces. Generating output.")
        
        # Count successfully processed chats
        successful_chats = 0
        
        for i, group_path in enumerate(dirs):
            try:
                if group_path.name.find(DM_PREFIX) == 0:
                    self.logger.debug(f"{group_path.name} is a DM.")
                    if self.CreateOutput(group_path):
                        successful_chats += 1
                elif group_path.name.find(SPACE_PREFIX) == 0:
                    self.logger.debug(f"{group_path.name} is a space.")
                    if self.CreateOutput(group_path):
                        successful_chats += 1
                else:
                    self.logger.warning(
                        f"{group_path.name} is not a DM or a space. Ignoring."
                    )
                    continue
            except Exception as e:
                self.logger.error(f"Error processing {group_path.name}: {e}")
            
            print(
                f"\b\b\b\b\b{i * 100 / num_dirs:.0f}%",
                end="",
                flush=True,
                file=sys.stderr,
            )
        
        self.logger.info(f"Successfully processed {successful_chats} out of {num_dirs} chats/spaces.")


def main():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument(
        "-i",
        "--in_dir",
        type=Path,
        dest="in_dir",
        default=Path("Takeout") / "Google Chat",
        help='Folder that contains "Users" and "Groups" from Google Chat Takeout',
    )
    parser.add_argument(
        "-o",
        "--out_dir",
        type=Path,
        dest="out_dir",
        default=Path.cwd(),
        help="Folder where chat files will be saved",
    )
    parser.add_argument("-l", "--log_level", type=str, default=logging.INFO)
    parser.add_argument(
        "-s",
        "--start_date",
        type=dt.date.fromisoformat,
        dest="start_date",
        required=False,
    )
    parser.add_argument(
        "-e", "--end_date", type=dt.date.fromisoformat, dest="end_date", required=False
    )
    parser.add_argument("-z", "--time_zone", type=str, default="UTC")
    parser.add_argument("-p", "--paper_size", type=str, dest="paper_size", default="A4")
    parser.add_argument(
        "-m", "--max_filename_length", type=int, dest="max_filename_len", default=127
    )
    parser.add_argument(
        "-a",
        "--all",
        dest="include_all",
        default=False,
        action="store_true",
        help="Save files which don't include me participating in the chats.",
    )
    parser.add_argument("-ih", "--max_img_height_in", type=int, default=2)
    args = parser.parse_args()

    conv = CChat2Pdf(args)
    conv.run()


if __name__ == "__main__":
    main()
