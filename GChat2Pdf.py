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
        return HyperlinkedImage(
            img_path_url,
            hyperlink=str(orig_path_url),
            width=iw,
            height=ih,
            hAlign="CENTER",
        )

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
            self.logger.debug("No messages.")
            return
        grp_info_file_path = dm_dir.joinpath(GROUP_INFO_FILE)
        with open(grp_info_file_path, "r", encoding = "utf-8") as inf:
            group_info = json.load(inf)
            # Use group name if available; otherwise, default to "Chat"
            title_str = group_info.get("name", "").strip() or "Chat"
            file_name = title_str + " with"
            participants_str = f"<u>Participants:</u><br />\t{self.user_name} ({self.user_email})<br />"
            for participant in group_info.get("members", []):
                if isinstance(participant, dict):
                    # Only default to "Anonymous User" if no name is provided.
                    name = participant.get("name") or "Anonymous User"
                    email = participant.get("email")
                else:
                    name = "Anonymous User"
                    email = None
                if name != self.user_name:
                    participants_str += (
                        f"\t{name}" + (f" ({email})" if email else "") + "<br />"
                    )
                    if len(file_name) + len(name) + 1 < self.args.max_filename_len:
                        file_name += f" {name},"
            file_name = file_name[:-1] + ".pdf"
            participants_str = participants_str.replace("\t", "&nbsp;" * 5)

        file_created = False
        I_participated = False
        doc_components = []
        img_file_names = {}
        with open(msg_file_path, "rb") as inf:
            for msg in ijson.items(inf, "messages.item"):
                if not (
                    "message_state" in msg and msg["message_state"] == MSG_STATE_DELETED
                ):
                    msg_dt = dt.datetime.strptime(
                        msg["created_date"].replace("\u202f", ""),
                        "%A, %B %d, %Y at %I:%M:%S%p %Z",
                    )
                    msg_dt = pytz.utc.localize(msg_dt, is_dst=None).astimezone(
                        pytz.timezone(self.args.time_zone)
                    )
                    msg_dt_str = msg_dt.strftime("%Y-%m-%d %H:%M:%S %Z%z")
                    msg_d = msg_dt.date()
                    if (
                        self.args.start_date is None or msg_d >= self.args.start_date
                    ) and (self.args.end_date is None or msg_d <= self.args.end_date):
                        if msg["creator"]["name"] == self.user_name:
                            I_participated = True
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
                                )  # Using preprocess_text instead of FixHebrewText
                                style_key = (  # Removed Hebrew style selection logic
                                    "MeNormal"
                                    if msg["creator"]["name"] == self.user_name
                                    else "OtherNormal"
                                )
                                doc_components.append(
                                    Paragraph(text, self.style_sheets[style_key])
                                )
                            elif "attached_files" in msg:
                                for i, f in enumerate(msg["attached_files"]):
                                    # Always use the original file path for Windows or if the file is a PDF or PNG
                                    if os.name == "nt" or f["export_name"].lower().endswith((".pdf", ".png")):
                                        img_file_path = dm_dir.joinpath(f["export_name"])
                                    else:
                                        img_file_path = dm_dir.joinpath(f["export_name"])
                                        fn = img_file_path.stem[:TRUNC_FILE_NAME] + img_file_path.suffix
                                        img_file_path = img_file_path.parent.joinpath(fn)
                                        if fn not in img_file_names:
                                            img_file_names[fn] = 1
                                        else:
                                            img_file_path = img_file_path.parent.joinpath(
                                                img_file_path.stem + f"({img_file_names[fn]})" + img_file_path.suffix
                                            )
                                            img_file_names[fn] += 1

                                    # Then continue processing the file type...
                                    if img_file_path.suffix.lower() in [".jpg", ".png", ".jpeg", ".heic", ".gif", ".eps"]:
                                        doc_components.append(self.GetScaledImage(img_file_path))
                                    elif img_file_path.suffix.lower() == ".pdf":
                                        # Check file existence before trying to generate a thumbnail
                                        if not img_file_path.exists():
                                            self.logger.warning(f"PDF file not found: {img_file_path}. Adding file link instead.")
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
                                    doc_components.append(
                                        self.GetScaledImage(
                                            msg["annotations"][0]["url_metadata"][
                                                "image_url"
                                            ]
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
                            print(msg)
                            raise e
        if file_created and (self.args.include_all or I_participated):
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
        for i, group_path in enumerate(dirs):
            if group_path.name.find(DM_PREFIX) == 0:
                self.logger.debug(f"{group_path.name} is a DM.")
            elif group_path.name.find(SPACE_PREFIX) == 0:
                self.logger.debug(f"{group_path.name} is a space.")
            else:
                self.logger.warning(
                    f"{group_path.name} is not a DM or a space. Ignoring."
                )
                continue
            self.CreateOutput(group_path)
            print(
                f"\b\b\b\b\b{i * 100 / num_dirs:.0f}%",
                end="",
                flush=True,
                file=sys.stderr,
            )


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
