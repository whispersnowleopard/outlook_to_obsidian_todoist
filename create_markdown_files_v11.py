#!/usr/bin/env python3
# written by Dave Gilbert & Claude 2025.11.12
# Oracle NetSuite NSGBU Channel Sales
# Version 11 - Complete rewrite with thread consolidation, attachment extraction, and Todoist integration
#
# Usage:
#   python3 create_markdown_files_v11.py                # Process emails, create CSV (no send)
#   python3 create_markdown_files_v11.py --send-tasks   # Send tasks from existing CSV
#   python3 create_markdown_files_v11.py --auto-send    # Process emails and auto-send to Todoist

import os
import sys
import shutil
import re
import logging
import urllib.parse
import csv
import uuid
import smtplib
from email import policy
from email.parser import BytesParser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from bs4 import BeautifulSoup
from email.utils import parsedate_to_datetime
from datetime import datetime
from pathlib import Path

# ============================================================================
# CONFIGURATION
# ============================================================================

# Paths
VAULT_NAME = "daveg-nsgbu"
VAULT_DIR = Path("/Users/degilber/obsidian/vaults/daveg-nsgbu")
EMAIL_INBOX_DIR = Path("/Users/degilber/obsidian/email-inbox")
PROCESSED_DIR = EMAIL_INBOX_DIR / "processed"
DEST_DIR = VAULT_DIR / "Archive-Email"
ATTACHMENTS_BASE = VAULT_DIR / "attachments"
SCRIPT_DIR = Path("/Users/degilber/obsidian/scripts")
CSV_PATH = SCRIPT_DIR / "todoist_queue.csv"

# Todoist
TODOIST_EMAIL = "add.task.9z98cy96979zb987@todoist.net"
TODOIST_PROJECT = "#Professional"

# Gmail credentials from environment
GMAIL_ADDRESS = os.environ.get('GMAIL_ADDRESS')
GMAIL_APP_PASSWORD = os.environ.get('GMAIL_APP_PASSWORD')

# Settings
MAX_FILENAME_LENGTH = 80
MIN_IMAGE_SIZE = 5000  # Skip images smaller than 5KB (likely tracking pixels/logos)

# Ensure directories exist
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
DEST_DIR.mkdir(parents=True, exist_ok=True)
ATTACHMENTS_BASE.mkdir(parents=True, exist_ok=True)
SCRIPT_DIR.mkdir(parents=True, exist_ok=True)

# ============================================================================
# LOGGING SETUP
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)

# ============================================================================
# KEYWORD TAGS (from v10)
# ============================================================================

KEYWORD_TAGS = {
    'training': ['training'],
    'oracle champions': ['Oracle_Champions'],
    'APC': ['APC'],
    'Solution Provider': ['SP_Program'],
    'SP Partner': ['SP_Program'],
    'Advanced Partner Center': ['APC'],
    'WGPT': ['WGPT'],
    'Weighted Growth Planning': ['WGPT'],
    'Weighted Growth Calculator': ['WGPT'],
    'CAMO': ['CAMO'],
    'Eide': ['Eide_Bailey'],
    'RSM US': ['RSM_US'],
    'SPAC': ['SPAC'],
    'Doozy': ['Crafted_ERP'],
    'Crafted ERP': ['Crafted_ERP'],
    'Vested Group': ['Vested_Group'],
    'Citrin': ['Citrin_Cooperman'],
    'Citrin Cooperman': ['Citrin_Cooperman'],
    'Meridian': ['Meridian'],
    'govirtualoffice': ['GVO'],
    'GVO': ['GVO'],
    'Oasis': ['Oasis'],
    'Sikich': ['Sikitch'],
    'Cumula 3': ['Cumula_3'],
    'PSP': ['PSP'],
    'Partner Strategic Planning': ['PSP'],
    'Domo': ['Domo'],
    'Whitespace': ['Whitespace'],
    'New Logo': ['Logo_Sales'],
    'Initial Sale': ['Logo_Sales'],
    'Logo Sales': ['Logo_Sales'],
    'SuiteCharts': ['SuiteCharts'],
    'Churn': ['Churn'],
    'Certifications': ['Certs_&_Auths'],
    'Authorizations': ['Certs_&_Auths'],
    'Multibook Authorized': ['Certs_&_Auths'],
    'SuiteSuccess': ['SuiteSuccess'],
    'SQL': ['SQL'],
    'SuitePulse': ['SuitePulse'],
    'Alliance': ['Alliance'],
    'Channel': ['Channel'],
    'Translation': ['Translation'],
    'PRC': ['PRC'],
    'Partner Resource Center': ['PRC'],
    'APC Login Traffic': ['APC_Login_Traffic'],
    'Enablement': ['Enablement'],
    'APS': ['APS'],
    'LCS': ['LCS'],
    'Onboarding': ['Onboarding'],
    'SuiteLife': ['SuiteLife'],
    'Analytics': ['Analytics'],
    'Data Visualization': ['Data_Viz'],
    'Charting': ['Data_Viz'],
    'Graph': ['Data_Viz'],
    'NSN': ['NSN'],
    'NetSuite Next': ['NSN'],
    'MCP': ['MCP'],
    'NLCorp': ['NLCorp'],
    'NSCorp': ['NLCorp'],
    'Systools': ['Systools'],
    'NetSuite on NetSuite': ['Systools'],
    'Portfolio Value': ['Portfolio_Value'],
    'PVuM': ['Portfolio_Value'],
    '411': ['411'],
    'NSCR': ['Systools'],
    'SuiteWorld': ['SuiteWorld'],
    'SPA/UIF': ['SPA_UIF'],
    'JET': ['JET'],
    'HighCharts': ['Data_Viz'],
    'Whitespace AI': ['Whitespace_AI'],
    'all hands': ['all_hands'],
    'Wrike': ['Wrike'],
    'PDR': ['PDR'],
    'NSGBU': ['NSGBU'],
    'BPO': ['BPO'],
    'Walker': ['Channel_Sales_Leadership'],
    'Mike Walker': ['Channel_Sales_Leadership','AVP'],
    'Craig West': ['Channel_Sales_Leadership','GVP'],
    'John Curtain': ['Channel_Sales_Leadership','AMO','RVP'],
    'Greg Goldstein': ['Channel_Sales_Leadership','AMO','RVP'],
    'Michael Kulisch': ['Channel_Sales_Leadership','BPO','RVP'],
    'Brent Pieterick': ['Channel_Sales_Leadership','Logo_Sales','RVP'],
    'Steve Mullen': ['Channel_Sales_Leadership','Logo_Sales','RVP'],
    'Will Chan': ['Channel_Programs'],
    'Larry Becker': ['Channel_Programs'],
    'Chris Hering': ['Channel_Programs','SuiteLife'],
    'Erin Cermak': ['SuiteLife'],
    'Matt Guggemos': ['Systools'],
    'Abhishek Purkayastha': ['Systools'],
    'Hazel Deniece Alano': ['Systools'],
    'Erik Rackwitz': ['Systools'],
    'Kristine Guardian': ['Systools'],
    'Alan Liu': ['Systools'],
    'Pioquinto Pangilinan': ['Systools'],
    'Jakub Durica': ['Systools'],
    'Kelly Young': ['CAMO','RSD'],
    'Stephen Spess': ['CAMO','RSD'],
    'Caitlin Sasso': ['CAMO','RSD'],
    'Ben Alford': ['CAMO','RSD'],
    'Bill Oriordan': ['CAMO','RSD'],
    'Kathy Garin': ['CAMO','RSD'],
    'Megan Robinson': ['CAMO','RSD'],
    'Regional Sales Manager': ['Logo_Sales','RSM'],
    'Ted Harding': ['Logo_Sales','RSM'],
    'C.B Wetmore': ['Logo_Sales','RSM'],
    'Charlie Fairbourn': ['Logo_Sales','RSM'],
    'Tony Stein': ['Logo_Sales','RSM'],
    'Skylar Luke': ['Logo_Sales','RSM'],
    'Robert Fierros': ['Logo_Sales','RSM'],
    'Jim Traynor': ['Logo_Sales','RSM'],
    'Aqsa Javed': ['Logo_Sales','RSM'],
    'Nicole Marsolais': ['Logo_Sales','RSM'],
    'Meghan Herlehy': ['Logo_Sales','RSM'],
    'Tom Butowicz': ['Logo_Sales','RSM'],
    'Nicholas Jarufe': ['Logo_Sales','RSM'],
    'Matt Uritis': ['Logo_Sales','RSM'],
    'Mike Cronin': ['Logo_Sales','RSM'],
}

APPROVED_ACRONYMS = {tag for tags in KEYWORD_TAGS.values() for tag in tags if tag.isupper() and len(tag) >= 3}
ACRONYM_PATTERN = r'\b[A-Z]{3,}\b'

# Priority keywords
URGENT_KEYWORDS = ['urgent', 'immediately', 'asap', 'critical', 'emergency', 'time sensitive']
HIGH_PRIORITY_KEYWORDS = ['important', 'priority', 'deadline', 'due today', 'due tomorrow']

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def safe(val):
    """Escape quotes in YAML values"""
    return val.replace('"', "'") if val else ""

def normalize_subject(subject):
    """Remove RE:, FW:, etc. from subject for thread matching"""
    subject = re.sub(r'^(RE|FW|FWD):\s*', '', subject, flags=re.IGNORECASE).strip()
    return subject

def should_skip_attachment(filename, content, size):
    """Determine if attachment should be skipped"""
    if not filename:
        return True
    
    filename_lower = filename.lower()
    
    # Skip common signature/logo patterns
    skip_patterns = ['signature', 'logo', 'icon', 'pixel', 'spacer', 'banner', 'badge']
    if any(pattern in filename_lower for pattern in skip_patterns):
        return True
    
    # Skip tiny images (tracking pixels)
    if size < MIN_IMAGE_SIZE and filename_lower.endswith(('.png', '.jpg', '.jpeg', '.gif')):
        return True
    
    return False

def extract_attachments(msg, email_date_str, subject):
    """Extract and save attachments from email"""
    attachments = []
    
    # Create attachment directory
    safe_subject = "".join(x for x in subject if x.isalnum() or x in (" ", "_", "-")).strip().replace(" ", "_")
    if len(safe_subject) > 40:
        safe_subject = safe_subject[:40]
    
    attach_dir = ATTACHMENTS_BASE / f"{email_date_str}_{safe_subject}"
    
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        
        filename = part.get_filename()
        if not filename:
            continue
        
        # Get content and size
        try:
            content = part.get_payload(decode=True)
            if not content:
                continue
            size = len(content)
        except Exception as e:
            logging.warning(f"Error decoding attachment {filename}: {e}")
            continue
        
        # Check if we should skip this attachment
        if should_skip_attachment(filename, content, size):
            logging.debug(f"Skipping attachment: {filename} (size: {size} bytes)")
            continue
        
        # Save attachment
        attach_dir.mkdir(parents=True, exist_ok=True)
        filepath = attach_dir / filename
        
        # Handle duplicate filenames
        counter = 1
        while filepath.exists():
            name, ext = os.path.splitext(filename)
            filepath = attach_dir / f"{name}_{counter}{ext}"
            counter += 1
        
        try:
            with open(filepath, 'wb') as f:
                f.write(content)
            
            # Store relative path for Obsidian linking
            rel_path = filepath.relative_to(VAULT_DIR)
            attachments.append({
                'filename': filename,
                'path': str(rel_path),
                'is_image': filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp'))
            })
            logging.info(f"Extracted attachment: {filename}")
        except Exception as e:
            logging.error(f"Error saving attachment {filename}: {e}")
    
    return attachments

def convert_html_to_markdown(html):
    """Convert HTML email to markdown, removing images and cleaning up links"""
    soup = BeautifulSoup(html, "html.parser")

    # Remove ALL images
    for img in soup.find_all('img'):
        img.decompose()
    
    # Remove any remaining image references
    for tag in soup.find_all(attrs={'src': True}):
        if 'fabric-cdn' in tag.get('src', '') or '.svg' in tag.get('src', ''):
            tag.decompose()
    
    # Replace <a> tags with markdown, skipping cid: links
    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.startswith("cid:") or 'fabric-cdn' in href or '.svg' in href:
            a.decompose()
            continue
        text = a.get_text(strip=True)
        # Clean up the text
        text = re.sub(r'^\d+\.\s*', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        if not text or len(text) < 3:
            text = "link"
        a.replace_with(f"[{text}]({href})")

    # Get text
    text = soup.get_text(separator="\n", strip=True)
    
    return text.strip()

def clean_email_body(body):
    """Clean up email body text"""
    if not body:
        return ""
    
    # Remove [cid:...] patterns
    body = re.sub(r'\[cid:[^\]]+\]', '', body)
    
    # Fix filename<url> pattern - allow spaces in filenames
    body = re.sub(
        r'([^\n<]+?\.(?:docx?|xlsx?|pdf|csv|pptx?))<(https?://[^>]+)>',
        r'[\1](\2)',
        body
    )
    
    # Fix bare <url> patterns
    body = re.sub(
        r'<(https?://[^>]+)>',
        r'[\1](\1)',
        body
    )
    
    # Remove URLs that are clearly image/icon references
    body = re.sub(r'\[https?://[^\]]*fabric-cdn[^\]]*\]', '', body)
    body = re.sub(r'\[https?://[^\]]*\.svg[^\]]*\]', '', body)
    
    # Clean up malformed markdown links
    body = re.sub(
        r'\[\s*\d*\.\s*\[https?://[^\]]+\]\s*([^\]]+)\]\((https?://[^\)]+)\)',
        r'[\1](\2)',
        body
    )
    
    # Remove pattern where image URL and link URL are the same
    body = re.sub(
        r'\[[^\]]*\]\[(https?://[^\]]+)\]\(\1\)',
        r'',
        body
    )
    
    # Normalize excessive newlines
    body = re.sub(r'\n{3,}', '\n\n', body)
    
    return body.strip()

def extract_email_fields(eml_path):
    """Extract all fields from .eml file"""
    with open(eml_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    
    # Get recipients
    recipients = []
    if msg.get_all('to'):
        recipients += msg.get_all('to')
    if msg.get_all('cc'):
        recipients += msg.get_all('cc')
    recipient_str = ", ".join(recipients)
    
    # Get date
    subject = msg['subject'] or 'untitled'
    date_raw = msg['date'] or ''
    try:
        date_obj = parsedate_to_datetime(date_raw)
        date_str = date_obj.strftime("%Y-%m-%d")
    except Exception:
        date_str = "unknown-date"
    
    # Get message ID and references for threading
    message_id = msg.get('message-id', '')
    references = msg.get('references', '')
    in_reply_to = msg.get('in-reply-to', '')
    
    # Extract body
    body = ""
    html_body = ""
    
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    body = part.get_payload(decode=True).decode(part.get_content_charset('utf-8'), errors='replace')
                elif part.get_content_type() == 'text/html':
                    html_body = part.get_payload(decode=True).decode(part.get_content_charset('utf-8'), errors='replace')
        else:
            if msg.get_content_type() == 'text/plain':
                body = msg.get_payload(decode=True).decode(msg.get_content_charset('utf-8'), errors='replace')
            elif msg.get_content_type() == 'text/html':
                html_body = msg.get_payload(decode=True).decode(msg.get_content_charset('utf-8'), errors='replace')
    except Exception as e:
        logging.warning(f"Error parsing payload for {eml_path}: {e}")
    
    # Convert HTML to markdown if available
    if html_body:
        try:
            md_body = convert_html_to_markdown(html_body)
            if not body or (len(md_body) > len(body) + 20):
                body = md_body
        except Exception as e:
            logging.error(f"Error converting HTML to markdown: {e}")
    
    # Clean up body
    body = clean_email_body(body)
    
    # Extract attachments
    attachments = extract_attachments(msg, date_str, subject)
    
    return {
        'subject': subject,
        'from': msg['from'] or '',
        'recipients': recipient_str,
        'date': date_raw,
        'date_str': date_str,
        'body': body,
        'html_body': html_body,
        'message_id': message_id,
        'references': references,
        'in_reply_to': in_reply_to,
        'attachments': attachments
    }

def extract_action_links(body, html_body):
    """Extract Wrike, OneDrive, SharePoint, Confluence, and NetSuite links"""
    links = {}
    combined = body + ' ' + html_body
    
    # Wrike
    wrike_match = re.search(r'https://www\.wrike\.com/open\.htm\?id=(\d+)', combined)
    if wrike_match:
        links['wrike'] = f"https://www.wrike.com/open.htm?id={wrike_match.group(1)}"
    
    # OneDrive
    onedrive_match = re.search(r'https://(?:1drv\.ms/[^\s<>"]+|[^/]*\.onedrive\.live\.com[^\s<>"]+)', combined)
    if onedrive_match:
        links['onedrive'] = onedrive_match.group(0)
    
    # SharePoint
    sharepoint_match = re.search(r'https://[^/]*\.sharepoint\.com[^\s<>"]+', combined)
    if sharepoint_match:
        links['sharepoint'] = sharepoint_match.group(0)
    
    # Confluence
    confluence_match = re.search(r'https://[^/]*\.atlassian\.net/wiki[^\s<>"]+', combined)
    if confluence_match:
        links['confluence'] = confluence_match.group(0)
    
    # NetSuite
    netsuite_match = re.search(r'https://nlcorp\.app\.netsuite\.com[^\s<>"]+', combined)
    if netsuite_match:
        links['netsuite'] = netsuite_match.group(0)
    
    return links

def extract_keywords_and_tags(subject, body):
    """Extract tags based on keywords in subject and body"""
    tags = set(['email'])
    lower_text = f"{subject.lower()} {body.lower()}"
    
    for word, taglist in KEYWORD_TAGS.items():
        if word.lower() in lower_text:
            tags.update(taglist)
    
    acronyms_found = set(re.findall(ACRONYM_PATTERN, f"{subject} {body}"))
    for acronym in acronyms_found:
        if acronym in APPROVED_ACRONYMS:
            tags.add(acronym)
    
    project_tags = set(re.findall(r'@(\w+)', subject + '\n' + body))
    tags.update(project_tags)
    
    return sorted(tags)

def determine_priority(subject, body):
    """Determine Todoist priority based on keywords"""
    lower_text = f"{subject.lower()} {body.lower()}"
    
    for keyword in URGENT_KEYWORDS:
        if keyword in lower_text:
            return 'p1'
    
    for keyword in HIGH_PRIORITY_KEYWORDS:
        if keyword in lower_text:
            return 'p2'
    
    return 'p3'

def find_existing_thread_note(subject, message_id, references, in_reply_to):
    """Find existing note for this email thread"""
    normalized_subject = normalize_subject(subject)
    
    # Search Archive-Email directory for matching notes
    for note_path in DEST_DIR.glob("*.md"):
        with open(note_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
            # Check frontmatter for matching subject or message IDs
            frontmatter_match = re.search(r'^---\n(.*?)\n---', content, re.DOTALL)
            if frontmatter_match:
                frontmatter = frontmatter_match.group(1)
                
                # Check subject match
                subject_match = re.search(r'^subject:\s*"(.+)"', frontmatter, re.MULTILINE)
                if subject_match:
                    note_subject = normalize_subject(subject_match.group(1))
                    if note_subject == normalized_subject:
                        return note_path
                
                # Check message ID in references
                if references and references in content:
                    return note_path
                if in_reply_to and in_reply_to in content:
                    return note_path
    
    return None

def update_existing_note(note_path, new_email_data, new_eml_filename):
    """Append new email to existing thread note"""
    with open(note_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find the end of "## My Notes" section if it exists
    my_notes_pattern = r'(## My Notes.*?)(\n---\n|\n## Email Thread)'
    my_notes_match = re.search(my_notes_pattern, content, re.DOTALL)
    
    # Build new email section
    new_section = f"\n### Update - {new_email_data['date_str']}\n\n"
    new_section += f"**From:** {new_email_data['from']}\n"
    new_section += f"**Email file:** [[processed/{new_eml_filename}]]\n\n"
    
    # Add attachments if any
    if new_email_data['attachments']:
        new_section += "**Attachments:**\n"
        for att in new_email_data['attachments']:
            if att['is_image']:
                new_section += f"![[{att['path']}]]\n"
            else:
                new_section += f"[[{att['path']}]]\n"
        new_section += "\n"
    
    new_section += new_email_data['body'] + "\n"
    
    # Insert new section
    if '## Email Thread' in content:
        # Insert after "## Email Thread" header
        content = content.replace('## Email Thread', f"## Email Thread\n{new_section}", 1)
    elif my_notes_match:
        # Insert after "## My Notes" section
        insert_pos = my_notes_match.end()
        content = content[:insert_pos] + f"\n---\n\n## Email Thread\n{new_section}" + content[insert_pos:]
    else:
        # Add to end
        content += f"\n---\n\n## Email Thread\n{new_section}"
    
    # Write updated content
    with open(note_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    logging.info(f"Updated existing thread note: {note_path.name}")
    return note_path

def create_new_note(email_data, eml_filename, unique_id, action_links, tags):
    """Create new Obsidian note"""
    safe_title = "".join(x for x in email_data['subject'] if x.isalnum() or x in (" ", "_", "-")).strip().replace(" ", "_")
    if len(safe_title) > MAX_FILENAME_LENGTH:
        safe_title = safe_title[:MAX_FILENAME_LENGTH].rstrip('_')
    
    note_filename = f"{email_data['date_str']}_{safe_title or 'untitled'}.md"
    note_path = DEST_DIR / note_filename
    
    # Build Obsidian vault URI
    rel_path = f"Archive-Email/{note_filename}"
    vault_uri = f"obsidian://vault/{urllib.parse.quote(VAULT_NAME)}/{urllib.parse.quote(rel_path)}"
    
    # Build task line with links
    task_line = f"{email_data['subject']} [KB {unique_id}]({vault_uri})"
    
    # Add action links
    if action_links.get('wrike'):
        task_line += f" [Wrike]({action_links['wrike']})"
    if action_links.get('onedrive'):
        task_line += f" [OneDrive]({action_links['onedrive']})"
    if action_links.get('sharepoint'):
        task_line += f" [SharePoint]({action_links['sharepoint']})"
    if action_links.get('confluence'):
        task_line += f" [Confluence]({action_links['confluence']})"
    if action_links.get('netsuite'):
        task_line += f" [NetSuite]({action_links['netsuite']})"
    
    # Add tags
    tag_string = ' '.join(f'@{tag}' for tag in tags)
    task_line += f" {tag_string} {TODOIST_PROJECT}"
    
    # Build frontmatter
    frontmatter_lines = [
        "---",
        f'id: {unique_id}',
        f'subject: "{safe(email_data["subject"])}"',
        f'date: {email_data["date"]}',
        f'from: "{safe(email_data["from"])}"',
        f'recipients: "{safe(email_data["recipients"])}"',
        f'email_file: "[[processed/{eml_filename}]]"',
        'tags:'
    ]
    
    for tag in tags:
        if re.match(r'^[a-zA-Z_]', tag):
            frontmatter_lines.append(f'  - {tag}')
    
    frontmatter_lines.append("---")
    frontmatter_lines.append("")
    
    # Add task callout
    frontmatter_lines.append("> [!todo] PASTE THIS INTO TODOIST QUICK ADD")
    frontmatter_lines.append(f"> {task_line}")
    frontmatter_lines.append("")
    
    # Add My Notes section
    frontmatter_lines.append("## My Notes")
    frontmatter_lines.append("")
    frontmatter_lines.append("---")
    frontmatter_lines.append("")
    
    # Add Email Thread section
    frontmatter_lines.append("## Email Thread")
    frontmatter_lines.append("")
    frontmatter_lines.append(f"### Original Email - {email_data['date_str']}")
    frontmatter_lines.append("")
    
    # Add attachments if any
    if email_data['attachments']:
        frontmatter_lines.append("**Attachments:**")
        for att in email_data['attachments']:
            if att['is_image']:
                frontmatter_lines.append(f"![[{att['path']}]]")
            else:
                frontmatter_lines.append(f"[[{att['path']}]]")
        frontmatter_lines.append("")
    
    # Write note
    md_content = "\n".join(frontmatter_lines) + "\n" + email_data['body']
    
    with open(note_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
    
    logging.info(f"Created new note: {note_filename}")
    return note_path, task_line, vault_uri, action_links

def process_emails():
    """Main email processing function"""
    tasks = []
    
    for eml_file in EMAIL_INBOX_DIR.glob("*.eml"):
        try:
            logging.info(f"Processing: {eml_file.name}")
            
            # Extract email data
            email_data = extract_email_fields(eml_file)
            
            # Generate unique ID
            unique_id = str(uuid.uuid4())[:8]
            
            # Extract tags and priority
            tags = extract_keywords_and_tags(email_data['subject'], email_data['body'])
            priority = determine_priority(email_data['subject'], email_data['body'])
            
            # Extract action links
            action_links = extract_action_links(email_data['body'], email_data['html_body'])
            
            # Check for existing thread
            existing_note = find_existing_thread_note(
                email_data['subject'],
                email_data['message_id'],
                email_data['references'],
                email_data['in_reply_to']
            )
            
            if existing_note:
                # Update existing thread note
                note_path = update_existing_note(existing_note, email_data, eml_file.name)
                
                # Still create task for new email in thread
                rel_path = note_path.relative_to(VAULT_DIR)
                vault_uri = f"obsidian://vault/{urllib.parse.quote(VAULT_NAME)}/{urllib.parse.quote(str(rel_path))}"
                
                # Build task line
                task_line = f"{email_data['subject']} [KB {unique_id}]({vault_uri})"
                for link_type, url in action_links.items():
                    task_line += f" [{link_type.title()}]({url})"
                tag_string = ' '.join(f'@{tag}' for tag in tags)
                task_line += f" {tag_string} {TODOIST_PROJECT}"
            else:
                # Create new note
                note_path, task_line, vault_uri, action_links = create_new_note(
                    email_data, eml_file.name, unique_id, action_links, tags
                )
            
            # Build links section for Todoist email body
            links_list = [f"[KB {unique_id}]({vault_uri})"]
            for link_type, url in action_links.items():
                links_list.append(f"[{link_type.title()}]({url})")
            
            # Add to tasks list
            tasks.append({
                'subject': email_data['subject'],
                'tags': ','.join(tags),
                'priority': priority,
                'links': '\n'.join(links_list),
                'comment': ''
            })
            
            # Move processed email
            shutil.move(str(eml_file), str(PROCESSED_DIR / eml_file.name))
            
        except Exception as e:
            logging.error(f"Failed to process {eml_file.name}: {e}", exc_info=True)
    
    return tasks

def save_tasks_to_csv(tasks):
    """Save tasks to CSV for editing/sending"""
    with open(CSV_PATH, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['subject', 'tags', 'priority', 'links', 'comment'])
        writer.writeheader()
        writer.writerows(tasks)
    
    logging.info(f"Saved {len(tasks)} tasks to: {CSV_PATH}")

def send_task_to_todoist(subject, tags, priority, links, comment=''):
    """Send a single task to Todoist via Gmail"""
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        logging.error("Gmail credentials not found in environment variables")
        return False
    
    try:
        # Build email subject with tags and priority
        email_subject = f"{subject} {tags.replace(',', ' @')} {priority}"
        if not tags.startswith('@'):
            email_subject = f"{subject} @{tags.replace(',', ' @')} {priority}"
        
        # Build email body
        body_parts = [links]
        if comment:
            body_parts.append(f"\n{comment}")
        email_body = '\n'.join(body_parts)
        
        # Create message
        msg = MIMEText(email_body)
        msg['Subject'] = email_subject
        msg['From'] = GMAIL_ADDRESS
        msg['To'] = TODOIST_EMAIL
        
        # Send via Gmail SMTP
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            smtp.send_message(msg)
        
        logging.info(f"Sent task to Todoist: {subject}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send task to Todoist: {e}")
        return False

def send_tasks_from_csv():
    """Read CSV and send all tasks to Todoist"""
    if not CSV_PATH.exists():
        logging.error(f"CSV file not found: {CSV_PATH}")
        return
    
    with open(CSV_PATH, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        tasks = list(reader)
    
    sent_count = 0
    for task in tasks:
        if send_task_to_todoist(
            task['subject'],
            task['tags'],
            task['priority'],
            task['links'],
            task.get('comment', '')
        ):
            sent_count += 1
    
    # Rename CSV to mark as sent
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    sent_path = SCRIPT_DIR / f"todoist_sent_{timestamp}.csv"
    CSV_PATH.rename(sent_path)
    
    logging.info(f"Sent {sent_count}/{len(tasks)} tasks. CSV archived as: {sent_path.name}")

# ============================================================================
# MAIN
# ============================================================================

def main():
    # Parse command line arguments
    auto_send = '--auto-send' in sys.argv
    send_only = '--send-tasks' in sys.argv
    
    if send_only:
        # Just send tasks from existing CSV
        logging.info("Sending tasks from CSV...")
        send_tasks_from_csv()
    else:
        # Process emails
        logging.info("Processing emails...")
        tasks = process_emails()
        
        if tasks:
            save_tasks_to_csv(tasks)
            
            if auto_send:
                logging.info("Auto-sending tasks to Todoist...")
                send_tasks_from_csv()
            else:
                logging.info(f"Tasks saved to CSV. Edit if needed, then run with --send-tasks to send.")
        else:
            logging.info("No emails to process.")
    
    logging.info("Done!")

if __name__ == "__main__":
    main()
