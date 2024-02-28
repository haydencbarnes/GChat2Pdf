"""
Copyright (C) 2024  Michael Dovrat

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

from pathlib import Path
import argparse
import sys
import logging
import ijson #streaming from file
import json #all in memory
import datetime as dt
import pytz
import re
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib import utils as rlutils
from reportlab.lib.units import inch
from io import BytesIO
from pillow_heif import register_heif_opener #heic file reader
import fitz #pdf thumbnails
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 


TRUNC_FILE_NAME = 47 #file names on disk are truncated to this basename (stem) for some reason even though the name is full in the google json. google takeout did this
TRUNC_HEB_LINE = 60

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
  def __init__(self, filename, hyperlink=None, width=None, height=None, kind='direct',
              mask='auto', lazy=1, hAlign='CENTER'):
    super(HyperlinkedImage, self).__init__(filename, width, height, kind, mask, lazy,
                                            hAlign=hAlign)
    self.hyperlink = hyperlink

  def drawOn(self, canvas, x, y, _sW=0):
    if self.hyperlink:  
      x1 = self._hAlignAdjust(x, _sW)
      y1 = y
      x2 = x1 + self._width
      y2 = y1 + self._height
      canvas.linkURL(url=self.hyperlink, rect=(x1, y1, x2, y2), thickness=0, relative=1)
    super(HyperlinkedImage, self).drawOn(canvas, x, y, _sW)


class CChat2Pdf:
  def __init__(self, args):
    self.args = args
    logging.basicConfig(level=args.log_level, 
                        format=" {asctime}.{msecs:03.0f} {levelname: <9} {message}", 
                        datefmt="%Y-%m-%d %H:%M:%S", style='{')
    self.logger = logging.getLogger('Chat2Pdf')
    if not self.args.in_dir.is_dir():
      self.logger.error(f"Can't open input folder {self.args.in_dir}.")
      sys.exit(0)
    try:
      self.args.out_dir.mkdir(parents = True, exist_ok = True)
    except:
      self.logger.error(f"Can't open/create output folder {self.args.out_dir}")
      sys.exit(0)
    self.page_width = (letter if self.args.paper_size == 'letter' else A4)[0]
    register_heif_opener()
    pdfmetrics.registerFont(TTFont('Hebrew', 'ARIAL.ttf'))
    self.style_sheets = getSampleStyleSheet()
    self.style_sheets.add(ParagraphStyle(name='MeHeader', parent = self.style_sheets['Heading4'], alignment=TA_LEFT))
    self.style_sheets.add(ParagraphStyle(name='OtherHeader', parent = self.style_sheets['Heading4'], alignment=TA_RIGHT))
    self.style_sheets.add(ParagraphStyle(name='MeNormal', parent = self.style_sheets['Normal'], alignment=TA_JUSTIFY, rightIndent=2.0*inch))
    self.style_sheets.add(ParagraphStyle(name='OtherNormal', parent = self.style_sheets['Normal'], alignment=TA_JUSTIFY, leftIndent=2.0*inch))
    self.style_sheets.add(ParagraphStyle(name='MeNormalHeb', parent = self.style_sheets['Normal'], alignment=TA_JUSTIFY, wordWrap="RTL", rightIndent=2.0*inch, fontName="Hebrew", fontSize = 12))
    self.style_sheets.add(ParagraphStyle(name='OtherNormalHeb', parent = self.style_sheets['Normal'], alignment=TA_JUSTIFY, wordWrap="RTL", leftIndent=2.0*inch, fontName="Hebrew", fontSize = 12))
    self.unk_file_exts = set()
    self.logger.info("Init success.")


  def GetScaledImage(self, img_path_url, orig_path_url = None):
    img = rlutils.ImageReader(img_path_url)
    iw, ih = img.getSize()
    aspect = ih / float(iw)
    if ih > self.args.max_img_height_in * inch:  #shrink height first
      ih = self.args.max_img_height_in * inch
      iw = ih / aspect
    if iw > self.page_width - 1.5*inch: #shrink width if needed. #1.5 is the left+right margins. Hard coded at the moment
      iw = self.page_width - 1.5*inch
      ih = iw * aspect
    if orig_path_url is None:
      orig_path_url = img_path_url
    return HyperlinkedImage(img_path_url, hyperlink=str(orig_path_url), 
                            width = iw, height = ih, hAlign = 'CENTER')


  def FixHebrewText(self, text):
    char_replace = { "(":")", ")":"(", "[":"]", "]":"[", "{":"}", "}":"{", "<" : ">", ">" : "<" } #that last one <> is dangerous
    if all(c < "\u0590" or c > "\u05EA" for c in text): #No Hebrew
      return False, text
    """ handle hebrew very crudely (not perfectly):
        1. Reverse the content of the message (paragraph)
        2. Re-reverse the parentheses ( to ) and ) to (, and also curly and also square brackets 
        3. Re-reverse all chunks of non-Hebrew sequences (since English words and word orders was reversed)
        4. Cut the words to new lines from the end of the string to the beginning, making sure we don't spill over to a new line (that would
           make the first words appear on the 2nd line, which will be out of order again)
    """
    text_r = "".join([char_replace[c] if c in char_replace else c for c in text[::-1]])  #steps 1, 2 at once
    pos = 0
    re_pat = re.compile(r"[ -'*-;=?-Z\\^-z|~]+",re.ASCII)  #all ascii characters without the parentheses () {} [] <>
    #step 3
    text_r_r = text_r
    while pos < len(text_r_r):
      ascii_seq = re_pat.search(text_r_r, pos)
      if ascii_seq is None:
        break
      span = list(ascii_seq.span())
      match_str = ascii_seq.group()
      """ Since we have included the space in the matching ascii, we are trying to recognize sequences of English words together, not word by word,
          but we have gobbled up spaces as well as punctuations from the neighboring (hebrew) words, we must "return" them to the neighboring words
          and not reverse them.
      """
      while len(match_str)>0 and match_str[0] in " ?-.\"'":
        match_str = match_str[1:]
        span[0] += 1
      while len(match_str)>0 and match_str[-1] in " ?-.\"'":
        match_str = match_str[:-1]
        span[1] -= 1
      if span[1]-span[0] > 1: #still needs reversing
        text_r_r = text_r_r[:span[0]] + match_str[::-1] + text_r_r[span[1]:]
      pos = ascii_seq.end(0)
    #step 4 (break to lines and break long sentences to lines that will not spill over to the next line)
    lines = text_r_r.splitlines()
    final_str = ""
    for line in lines[::-1]: #reverse order of lines
      if len(line) <= TRUNC_HEB_LINE: #line is short enough as is
        final_str += line + '\n'
      else:
        words = line.split()
        sub_line_len = 0
        sub_line = ""
        while len(words) > 0 and sub_line_len <= TRUNC_HEB_LINE: #take words from end to start
          if sub_line_len + 1 + len(words[-1]) <= TRUNC_HEB_LINE:
            sub_line = words[-1] + ' ' + sub_line
            sub_line_len += len(words[-1]) + 1
            words.pop()
            if len(words) == 0: #we are done with this line
              final_str += sub_line + '\n'
          else: #new sub-line
            final_str += sub_line + '\n'
            sub_line_len = 0
            sub_line = ""
    #remove the last newline we may have added
    if final_str[-1] == '\n':
      final_str = final_str[:-1]
    return True, final_str


  def CreateOutput(self, dm_dir):
    msg_file_path = dm_dir.joinpath(MESSAGES_FILE)
    if not msg_file_path.exists():
      self.logger.debug("No messages.")
      return
    grp_info_file_path = dm_dir.joinpath(GROUP_INFO_FILE)
    with open(grp_info_file_path, "r") as inf:
      group_info = json.load(inf)
      title_str = group_info['name'] if 'name' in group_info else "Chat"
      file_name = title_str + " with"
      #replace characters that can't be in the filename
      file_name = re.sub(r"[/:\\*?\"<>|]", '-', file_name)
      participants_str = f"<u>Participants:</u><br />\t{self.user_name} ({self.user_email})<br />"
      for participant in group_info['members']:
        if participant['name'] != self.user_name:
          participants_str += f"\t{participant['name']}" + (f" ({participant['email']})" if 'email' in participant else "") + "<br />"
          if len(file_name) + len(participant['name']) + 1 < self.args.max_filename_len:
            file_name += f" {participant['name']},"
      file_name = file_name[:-1] + ".pdf" #remove the last ","
      participants_str = participants_str.replace('\t','&nbsp;'*5)
    file_created = False
    I_participated = False
    hebrew_seen = False
    doc_components = []
    img_file_names = {}
    with open(msg_file_path, "rb") as inf:
      for msg in ijson.items(inf, "messages.item"):
        if not ('message_state' in msg and msg['message_state'] == MSG_STATE_DELETED): #message was not deleted
          #remove narrow no-break space (between the time and the AM/PM) and extract message date
          msg_dt = dt.datetime.strptime(msg['created_date'].replace(u'\u202f',''), "%A, %B %d, %Y at %I:%M:%S%p %Z")
          #change from utc to local timezone
          msg_dt = pytz.utc.localize(msg_dt, is_dst=None).astimezone(pytz.timezone(self.args.time_zone))
          msg_dt_str = msg_dt.strftime('%Y-%m-%d %H:%M:%S %Z%z') 

          #filter dates
          msg_d = msg_dt.date()
          if ((self.args.start_date is None or msg_d >= self.args.start_date) and 
              (self.args.end_date is None or msg_d <= self.args.end_date)):
            if msg['creator']['name'] == self.user_name:
              I_participated = True
            if not file_created:
              pdf_io_buffer = BytesIO()
              output_buffer = SimpleDocTemplate(pdf_io_buffer, 
                                              pagesize=letter if self.args.paper_size == 'letter' else A4,
                                              rightMargin=0.75*inch, leftMargin=0.75*inch,
                                              topMargin=1.0*inch, bottomMargin=0.75*inch)
              doc_components.append(Paragraph(title_str, self.style_sheets['Title']))
              doc_components.append(Paragraph(participants_str, self.style_sheets['Heading5']))
              file_created = True

            header_str = (msg['creator']['name'] + ((" (" + msg['creator']['email'] + ")") if 'email' in msg['creator'] else "") + " at " +
                        msg_dt_str + ":")
            doc_components.append(Paragraph(header_str, self.style_sheets[
              'MeHeader' if msg['creator']['name'] == self.user_name else 'OtherHeader'
              ]))
            
            try:
              if 'text' in msg:
                is_hebrew, text = self.FixHebrewText(msg['text'])
                #if res and not hebrew_seen:
                #  self.logger.info(f"{dm_dir} has Hebrew in it.")
                #  hebrew_seen = True
                text = text.replace('\t','&nbsp;'*5)
                text = text.replace('\n', '<br />')
                if is_hebrew:
                  doc_components.append(Paragraph(text, self.style_sheets[
                  'MeNormalHeb' if msg['creator']['name'] == self.user_name else 'OtherNormalHeb' 
                  ]))
                else:
                  doc_components.append(Paragraph(text, self.style_sheets[
                  'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                  ]))
              elif 'attached_files' in msg:
                for i, f in enumerate(msg['attached_files']):
                  img_file_path = dm_dir.joinpath(f['export_name'])
                  fn = img_file_path.stem[:TRUNC_FILE_NAME] + img_file_path.suffix #truncate file name to 47 chars for some reason
                  img_file_path = img_file_path.parent.joinpath(fn)
                  if fn not in img_file_names: 
                    img_file_names[fn] = 1 
                  else:
                    img_file_path = img_file_path.parent.joinpath(img_file_path.stem + 
                      f"({img_file_names[fn]})" + img_file_path.suffix) 
                    img_file_names[fn] += 1
                  if img_file_path.suffix in ['.jpg', '.png', '.jpeg', '.JPG', '.PNG', '.heic', '.dng', '.gif', '.HEIC', '.eps', '.EPS']: #graphics directly embedded
                    doc_components.append(self.GetScaledImage(img_file_path))
                  elif img_file_path.suffix in ['.pdf', '.PDF']:
                    with fitz.open(img_file_path) as doc:
                      page = doc.load_page(0)  # number of page
                      pix = page.get_pixmap()
                      pix.save(PDF_TMP_FILE)
                    doc_components.append(self.GetScaledImage(PDF_TMP_FILE, img_file_path))
                  else: #not a file we want to or can show a thumbnail of
                    if img_file_path.suffix not in self.unk_file_exts:
                      suffix = img_file_path.suffix
                      self.logger.warning(f"File extension '{suffix}' without a thumbnail preview found.")
                      self.unk_file_exts.add(suffix)
                    file_link_str = '<u>File attached:</u> <link href="' + str(img_file_path) + '">' + img_file_path.name + '</link>'
                    doc_components.append(Paragraph(file_link_str, self.style_sheets[
                    'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                    ]))
              elif 'annotations' in msg:
                if 'video_call_metadata' in msg['annotations'][0]: #video meeting, it's in a list of annotations but always seems to be the first
                  doc_components.append(Paragraph("<u>Video call started.</u>", self.style_sheets[
                    'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                  ]))
                elif 'gsuite_integration_metadata' in msg['annotations'][0]:
                  if 'call_data' in msg['annotations'][0]['gsuite_integration_metadata']:  #call status
                    doc_components.append(Paragraph(f"<u>{msg['annotations'][0]['gsuite_integration_metadata']['call_data']['call_status']}</u>", self.style_sheets[
                      'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                    ]))
                  elif 'tasks_data' in msg['annotations'][0]['gsuite_integration_metadata']: #tasks data
                    task_str = "Task \"" + msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['task_properties']['title'] + "\""
                    if 'assignee' in msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['task_properties']:
                      task_str += " assigned to " + msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['task_properties']['assignee']['id'] 
                    if 'assignee_change' in msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']:
                      task_str += " removed from " + msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['assignee_change']['old_assignee']['id']
                    if msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['task_properties']['completed']:
                      task_str += " completed."
                    elif msg['annotations'][0]['gsuite_integration_metadata']['tasks_data']['task_properties']['deleted']:
                      task_str += " deleted."
                    else:
                      task_str += "."
                    doc_components.append(Paragraph(f"<u>{task_str}</u>", self.style_sheets[
                      'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                    ]))
                  else:
                    self.logger.warning(f"Unknown type under gsuite_integration_metadata. Message (complete):\n{msg}")
                elif 'url_metadata' in msg['annotations'][0]: #it's an image added with a URL
                  doc_components.append(self.GetScaledImage(msg['annotations'][0]['url_metadata']['image_url']))
                elif 'drive_metadata' in msg['annotations'][0]:#link to a file on the drive
                  doc_components.append(Paragraph(f"<u>File shared from google drive: {msg['annotations'][0]['drive_metadata']['title']}" 
                                                  " (file id:{msg['annotations'][0]['drive_metadata']['id']})</u>", self.style_sheets[
                      'MeNormal' if msg['creator']['name'] == self.user_name else 'OtherNormal' 
                    ]))
                else:
                  self.logger.warning(f"Unknown type under annotations. Message (complete):\n{msg}")
              else:
                self.logger.warning(f"Unknown type. Message (complete): {msg}")
            except ValueError:
              pass #quietly ignore some parsing errors with unquoted strings with junk in them
            except:
                #this is really unexpected - find out what's going on
                print(msg)
                raise
                

    if file_created and (self.args.include_all or I_participated):
      output_buffer.build(doc_components)  
      with open(str(self.args.out_dir.joinpath(file_name)), "wb") as outfile:
        outfile.write(pdf_io_buffer.getbuffer())
        

  def run(self):
    users_dir = self.args.in_dir.joinpath(USERS_DIR)
    if not users_dir.is_dir():
      self.logger.error(f"Couldn't find users folder. Exiting.")
      sys.exit(0)
    user_subdir = list(users_dir.iterdir())[0]
    user_info_file_path = user_subdir.joinpath(USER_INFO_FILE)
    with open(user_info_file_path, "rb") as inf:
      user_info = json.load(inf) 
      self.user_name = user_info['user']['name']
      self.user_email = user_info['user']['email']
      #self.user_memberships = user_info['membership_info']
    self.logger.info(f"You are {self.user_name} ({self.user_email})")

    groups_dir = self.args.in_dir.joinpath(GROUPS_DIR)
    if not groups_dir.is_dir():
      self.logger.error(f"Couldn't find groups folder. Exiting.")
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
        self.logger.warning(f"{group_path.name} is not a DM or a space. Ignoring.")
        continue #ignore for now
      self.CreateOutput(group_path)
      print(f"\b\b\b\b\b{i*100/num_dirs:.0f}%" , end = "", flush = True, file = sys.stderr)


def __del__(self):
  p = Path(PDF_TMP_FILE)
  if p.exists():
    p.unlink()



def main():
  parser = argparse.ArgumentParser()
  parser.add_argument('-i', '--in_dir', type=Path, dest='in_dir', required = True, help='Google Chat folder within the Google Takeout folder')
  parser.add_argument('-o', '--out_dir', type=Path, dest='out_dir', required = True, help='Folder where chat files will be saved.')
  parser.add_argument('-l', '--log_level', type = str, required = False, default = logging.INFO)
  parser.add_argument('-s', '--start_date', type=dt.date.fromisoformat, dest='start_date', required = False, help = 'Start date (YYYY-MM-DD).')
  parser.add_argument('-e', '--end_date', type=dt.date.fromisoformat, dest = 'end_date', required = False, help = 'End date (YYYY-MM-DD).')
  parser.add_argument('-z', '--time_zone', type = str, required = False, default = 'UTC', help = "Any pytz timezone (look them up).")
  parser.add_argument('-p', '--page_size', type = str, dest = 'paper_size', required = False, default = 'A4', help = "'A4' or 'letter' are accepted.")
  parser.add_argument('-m', '--max_filename_length', type = int, dest = 'max_filename_len', required = False, default = 127, help = "Max filename length")
  parser.add_argument('-a', '--all', dest = 'include_all', default = False, action=argparse.BooleanOptionalAction, help = "Save files which don't include me participating in the chats.")
  parser.add_argument('-ih', '--max_img_height_in', type = int, dest = 'max_img_height_in', required = False, default = 2, help = 'Maximum height for embedded image thumbnails.')
  args = parser.parse_args()

  conv = CChat2Pdf(args)
  conv.run()




if __name__ == "__main__":
  main()
