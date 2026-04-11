# drivemap.py

A Python tool that recursively maps Google Drive folders shared with you, reads inline Excel spreadsheet content, detects duplicate files, flags empty folders, and outputs a searchable, interactive HTML report plus a Markdown summary.

> 💡 **New to coding?** This project was built with the help of [Claude](https://claude.ai) and [Claude Code](https://claude.ai/code), and it's a great first project to adapt using those same tools. See the [Using Claude & Claude Code](#using-claude--claude-code) section below.

---

## What It Does

- 🔍 Finds folders in your Google Drive **"Shared with me"** that match a folder name you specify
- 📂 Recursively walks every subfolder, cataloguing all files
- 📊 Reads the contents of Excel (`.xlsx`) files whose names start with `1 LINKS` and displays them inline
- 🟡 Flags **duplicate** filenames across the folder tree
- 🟣 Flags **empty folders**
- 📄 Outputs two files:
  - `drive_map.html` — a fully interactive, searchable, filterable report with a sidebar folder tree
  - `drive_map.md` — a plain Markdown table version of the same data

---

## Example Output

The HTML report includes:
- A collapsible **folder tree sidebar** for navigation
- A **search bar** to filter by filename, type, or description
- **Type filter buttons** (PDF, Word, Excel, Audio, Video, Folder)
- **Summary stats** (total files, folders, duplicates, empty folders)
- Inline display of spreadsheet resource rows
- Scroll-spy sidebar highlighting and a back-to-top button

---

## Requirements

- Python 3.10 or later
- A Google account with files in **Shared with me**
- A Google Cloud project with the Drive API enabled (see setup below)

### Python packages

```
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl
```

---

## Google API Setup

This is a one-time setup. You only need to do this once per Google account.

1. Go to [https://console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. In the left menu, go to **APIs & Services → Library**
4. Search for **Google Drive API** and click **Enable**
5. Go to **APIs & Services → Credentials**
6. Click **Create Credentials → OAuth client ID**
7. Choose **Desktop app** as the application type
8. Download the resulting file and rename it to `credentials.json`
9. Place `credentials.json` in the same folder as `drivemap.py`

The first time you run the script, a browser window will open asking you to authorize access. After you approve, a `token.json` file is saved locally so you won't need to authorize again.

> ⚠️ **Never share or commit `credentials.json` or `token.json`** — these give access to your Google Drive. Add both to your `.gitignore` if you're using Git.

---

## Running the Script

```
python drivemap.py
```

The script will print progress as it walks the folder tree, then write `drive_map.html` and `drive_map.md` to the same directory. Open `drive_map.html` in any browser.

---

## Customization Guide

These are the key settings at the top of `drivemap.py` that you'll want to adjust for your own use case.

### 1. Change the target folder name

```python
TARGET_NAME = "NEW CFJ RESOURCE FOLDER Rev 2024"
```

Replace this with the **exact name** of the folder (or folders) in your "Shared with me" that you want to map. The match is case-insensitive.

**Example:**
```python
TARGET_NAME = "2025 Team Resources"
```

### 2. Change which Excel files get read inline

```python
if mime == XLSX_MIME and name.lower().startswith("1 links"):
```

This line controls which `.xlsx` files have their contents expanded inline in the report. Change `"1 links"` to match whatever naming convention your spreadsheets use.

**Example — read any file starting with "resources":**
```python
if mime == XLSX_MIME and name.lower().startswith("resources"):
```

**Example — read ALL Excel files inline:**
```python
if mime == XLSX_MIME:
```

### 3. Change the output filenames

```python
MD_OUTPUT   = "drive_map.md"
HTML_OUTPUT = "drive_map.html"
```

Rename these to whatever suits your project.

### 4. Adjust the Excel column mapping

Inside the `read_xlsx_links()` function, the script looks for these column headers in your spreadsheet:

| Column | Looks for |
|--------|-----------|
| Resource Type | "resource type", "type" |
| Title | "title", "name" |
| Link | "resource link", "link", "url" |
| Add'l Info | "add'l", "additional", "notes", "description" |

If your spreadsheet uses different headers, find the `find_col()` calls in `read_xlsx_links()` and add your header names to the list.

**Example — your sheet uses "Category" instead of "Resource Type":**
```python
col_type = find_col("resource type", "type", "category")
```

### 5. Add new file type badges or filters

File type display names are defined in the `mime_label()` function and filter categories in `type_key()`. You can add new MIME types or rename existing labels there.

---

## File Structure

```
your-project/
├── drivemap.py          # The script
├── credentials.json     # Your Google OAuth credentials (DO NOT share)
├── token.json           # Auto-generated auth token (DO NOT share)
├── drive_map.html       # Generated output (HTML report)
└── drive_map.md         # Generated output (Markdown table)
```

---

## Using Claude & Claude Code

This script was developed with the help of [Claude](https://claude.ai), Anthropic's AI assistant. If you want to adapt it for your own use case, Claude is an excellent starting point — even if you're new to Python.

### What is Claude Code?

[Claude Code](https://claude.ai/code) is a command-line tool that lets Claude read, edit, and run code directly on your machine. It's particularly useful for projects like this one, where you want to make targeted changes without needing to understand every line.

### Getting started with Claude Code

1. Install Claude Code:
```
npm install -g @anthropic/claude-code
```
2. In your terminal, navigate to the folder containing `drivemap.py`:
```
cd path/to/your/project
```
3. Start a Claude Code session:
```
claude
```

### Example prompts to try

Once you're in a Claude Code session, you can describe what you want in plain English. Here are some examples based on common adaptations of this script:

**Change the target folder:**
> "Change the TARGET_NAME to match my folder called '2025 Board Resources' and update the output filenames to board_map.html and board_map.md"

**Add a new file type to the filter bar:**
> "Add a Google Slides filter button to the type filter bar in the HTML output"

**Read a differently-named Excel file inline:**
> "Update the script so that any Excel file whose name contains the word 'index' gets read inline, not just ones starting with '1 LINKS'"

**Add a new column to the HTML table:**
> "Add a 'File Size' column to the HTML output table"

### Using Claude on claude.ai (no coding required)

You can also paste the script directly into [claude.ai](https://claude.ai) and ask questions in plain language:

> "I have a folder called 'Staff Shared Resources 2025' — what do I need to change in this script to map that folder instead?"

Claude will point you to exactly the right lines.

---

## Security Notes

- The script only requests **read-only** access to Google Drive (`drive.readonly` scope)
- No data is sent anywhere except to the Google Drive API
- All output is written locally to your machine
- `credentials.json` and `token.json` should never be shared or committed to version control

---

## License

MIT License — free to use, modify, and share. Attribution appreciated but not required.

---

## Contributing

Found a bug or have an improvement? Open an issue or pull request on GitHub. This project welcomes contributions of all kinds, including documentation improvements and new feature ideas.
