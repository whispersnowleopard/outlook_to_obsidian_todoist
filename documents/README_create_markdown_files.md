# Email → Obsidian → Todoist Automation

**Script:** `create_markdown_files_v11.py`  
**Author:** Dave Gilbert & Claude  
**Last Updated:** November 2025

---

## What This Does

Processes Outlook `.eml` files and:
- Converts them to Obsidian markdown notes
- Extracts tasks/action items
- Consolidates email threads into single notes
- Sends tasks to Todoist via Gmail

---

## Quick Start

### 1. Drop emails here:
```
/Users/degilber/obsidian/email-inbox/
```

### 2. Run the script (pick one):

```bash
# Process emails → Create CSV (review before sending)
python3 create_markdown_files_v11.py

# Send tasks from CSV (after reviewing/editing)
python3 create_markdown_files_v11.py --send-tasks

# Process emails → Auto-send to Todoist (no review)
python3 create_markdown_files_v11.py --auto-send
```

### 3. Done!
- Processed emails → `/Users/degilber/obsidian/email-inbox/processed/`
- Markdown notes → `/Users/degilber/obsidian/vaults/daveg-nsgbu/Archive-Email/`
- CSV queue → `/Users/degilber/obsidian/scripts/todoist_queue.csv`

---

## Requirements

### Gmail Credentials (for Todoist email integration)

Set these environment variables:

```bash
export GMAIL_ADDRESS="your-email@gmail.com"
export GMAIL_APP_PASSWORD="your-app-specific-password"
```

**To make permanent (add to `~/.zshrc` or `~/.bash_profile`):**
```bash
echo 'export GMAIL_ADDRESS="your-email@gmail.com"' >> ~/.zshrc
echo 'export GMAIL_APP_PASSWORD="your-app-password"' >> ~/.zshrc
source ~/.zshrc
```

**Get Gmail App Password:**
1. Go to https://myaccount.google.com/apppasswords
2. Create new app password
3. Copy the 16-character password (no spaces)

---

## What Gets Created

### Markdown Notes
Location: `/Users/degilber/obsidian/vaults/daveg-nsgbu/Archive-Email/`

**Features:**
- YAML frontmatter with metadata
- Unique stable ID for linking
- Link to original `.eml` file
- Extracted attachments (real ones, not signatures/logos)
- Action item links (Wrike, OneDrive, Confluence, etc.)
- Task callout block with Obsidian link
- Thread consolidation (new emails append to existing notes)
- "My Notes" section at top (safe from overwrites)

### CSV Queue
Location: `/Users/degilber/obsidian/scripts/todoist_queue.csv`

**Columns:**
- `subject` - Task title
- `tags` - Auto-extracted tags (@person, @project, etc.)
- `priority` - p1/p2/p3 (based on keywords like "urgent")
- `links` - All action links from email
- `comment` - Email sender info

**You can edit this before sending!**

---

## Typical Workflow

### Standard (with review):
```bash
# 1. Process emails
python3 create_markdown_files_v11.py

# 2. Review/edit the CSV if needed
open /Users/degilber/obsidian/scripts/todoist_queue.csv

# 3. Send to Todoist
python3 create_markdown_files_v11.py --send-tasks
```

### Fast (auto-send):
```bash
python3 create_markdown_files_v11.py --auto-send
```

---

## Troubleshooting

### "No emails to process"
- Check: Are there `.eml` files in `/Users/degilber/obsidian/email-inbox/`?
- Already processed files are in `email-inbox/processed/`

### "Gmail credentials not found"
- Set environment variables (see Requirements above)
- Check: `echo $GMAIL_ADDRESS` and `echo $GMAIL_APP_PASSWORD`

### CSV not found when running --send-tasks
- Run without flags first to create the CSV
- Or check: `/Users/degilber/obsidian/scripts/todoist_queue.csv`
- After sending, CSV is archived as `todoist_sent_TIMESTAMP.csv`

### Tasks not showing up in Todoist
- Check Todoist spam/inbox
- Verify Gmail credentials work
- Check script output for send errors

### Attachments too large
- Script skips images < 5KB (tracking pixels, signatures)
- Real attachments saved to: `/Users/degilber/obsidian/vaults/daveg-nsgbu/attachments/`

---

## Smart Features

### Priority Detection (p1/p2/p3)
- **p1** - Keywords: urgent, immediately, asap, critical
- **p2** - Keywords: important, soon, priority
- **p3** - Default for everything else

### Tag Extraction
Auto-detects and adds tags for:
- People (Oracle Champions members, key contacts)
- Partners (RSM, Eide Bailey, Citrin Cooperman, etc.)
- Programs (APC, SPAC, WGPT, PSP)
- Topics (training, implementation, SuiteSuccess)

### Link Extraction
Finds and includes:
- Wrike tasks
- OneDrive/SharePoint docs
- Confluence pages
- NetSuite records
- General URLs (filtered to real action items)

### Thread Consolidation
- Groups emails by subject + sender
- Appends new emails to existing notes
- Preserves your "My Notes" section at top
- Never overwrites your work

---

## File Structure

```
/Users/degilber/obsidian/
├── email-inbox/              # Drop .eml files here
│   └── processed/            # Processed files moved here
├── vaults/daveg-nsgbu/
│   ├── Archive-Email/        # Markdown notes created here
│   └── attachments/          # Extracted attachments
└── scripts/
    ├── create_markdown_files_v11.py  # The script
    ├── todoist_queue.csv             # Tasks to send
    └── todoist_sent_*.csv            # Archived sent tasks
```

---

## Configuration (in script)

If you need to change settings, edit these at the top of the script:

```python
VAULT_NAME = "daveg-nsgbu"
VAULT_DIR = Path("/Users/degilber/obsidian/vaults/daveg-nsgbu")
EMAIL_INBOX_DIR = Path("/Users/degilber/obsidian/email-inbox")
TODOIST_EMAIL = "add.task.9z98cy96979zb987@todoist.net"
TODOIST_PROJECT = "#Professional"
MIN_IMAGE_SIZE = 5000  # Skip images smaller than 5KB
```

---

## Version History

**v11** (November 2025)
- Thread consolidation
- Attachment extraction (smart filtering)
- Todoist integration via Gmail
- Action link extraction
- "My Notes" safe zone
- Unique stable IDs for Obsidian links

---

## Support

Script written collaboratively with Claude.  
For issues/updates, reference chat: https://claude.ai/chat/e5a3a6f9-37c0-40f9-8edd-7e97d4ed4207

**Quick command reference card:**
```bash
# Process only
python3 create_markdown_files_v11.py

# Send tasks
python3 create_markdown_files_v11.py --send-tasks

# Process + auto-send
python3 create_markdown_files_v11.py --auto-send
```

**Check what happened:**
```bash
# See processed emails
ls -lht /Users/degilber/obsidian/email-inbox/processed/

# See created notes
ls -lht /Users/degilber/obsidian/vaults/daveg-nsgbu/Archive-Email/

# See task queue
cat /Users/degilber/obsidian/scripts/todoist_queue.csv

# See sent archives
ls -lht /Users/degilber/obsidian/scripts/todoist_sent_*.csv
```
