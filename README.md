# Local Notion AI Agent

A local-first AI agent for macOS that combines Python, Ollama, local file access, SQLite memory, and the Notion API to manage Notion workspaces through natural-language instructions.

Current version: V5

---

## Overview

This project is a local AI workflow designed to interact with Notion through natural-language commands.

The agent runs locally on macOS, uses Ollama for local LLM inference, reads supported files from the Mac, stores lightweight operational memory with SQLite, and performs structured operations through the Notion REST API.

The project was developed incrementally from basic Notion page operations into a more general local AI agent capable of file reading, database querying, persistent state, batch-update previews, and spreadsheet-to-Notion import workflows.

---

## Key Features

- Natural-language interaction through a terminal interface
- Local LLM inference using Ollama
- Notion page search
- Notion child-page creation
- Notion block creation
- Database schema inspection
- Database record creation
- Database record querying
- Multi-condition AND / OR filtering
- Persistent operational memory using SQLite
- Contextual references such as "the previous record"
- Batch-update preview before execution
- Explicit confirmation for multi-record updates
- Local macOS file search
- Local directory browsing
- TXT reading
- Markdown reading
- PDF reading
- DOCX reading
- XLSX reading
- CSV reading
- Spreadsheet-to-Notion import preview
- Explicit confirmation before spreadsheet import
- Local-file path allowlist
- Sensitive-file protection
- Read-before-write safety design

---

## Architecture

```text
User
  |
  v
Terminal Interface
  |
  v
Python Agent
  |
  +----> Ollama / qwen3:8b
  |
  +----> Local Mac Files
  |
  +----> SQLite Persistent State
  |
  +----> Notion REST API
```

---

## Tech Stack

- Python
- Ollama
- qwen3:8b
- Notion REST API
- SQLite
- Requests
- prompt_toolkit
- PyPDF
- python-docx
- openpyxl
- python-dotenv

---

## Supported Local Files

The current V5 version can read:

- `.txt`
- `.md`
- `.csv`
- `.xlsx`
- `.docx`
- `.pdf`

Local file access is currently read-only.

The agent does not currently delete, rename, move, or overwrite local files.

---

## Local File Access

The current version is designed to read files from approved local directories such as:

```text
~/Desktop
~/Documents
~/Desktop/AgentWorkspace
```

Sensitive paths and files are blocked by design.

Examples include:

```text
.env
.ssh
.aws
.config
.pem
.key
.p12
.pfx
```

---

## Safety Design

The agent includes multiple safeguards.

### Local Files

- Local files are read-only
- File access is restricted to approved directories
- Sensitive file names and extensions are blocked
- The agent does not modify or delete local files

### Notion Operations

- Notion can be read and written
- Multi-record updates require a preview first
- Spreadsheet imports require a preview first
- Batch actions require explicit user confirmation
- Ambiguous multi-record operations are not executed automatically
- The agent does not claim success unless the Notion API confirms success

### Credentials

API credentials are loaded through environment variables rather than being hard-coded in source code.

Never commit:

```text
.env
API tokens
passwords
private keys
state.db
sensitive personal documents
```

---

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/wchsun03/notion_local_agent.git
cd notion_local_agent
```

### 2. Create a Python Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

```bash
python3 -m pip install -r requirements.txt
```

### 4. Install Ollama

Install Ollama on macOS.

Then download the model:

```bash
ollama pull qwen3:8b
```

Confirm that the model is available:

```bash
ollama list
```

You should see:

```text
qwen3:8b
```

### 5. Configure Environment Variables

Copy the example environment file:

```bash
cp .env.example .env
```

Then edit `.env` and add your own Notion integration token:

```env
NOTION_TOKEN=your_notion_token_here
OLLAMA_MODEL=qwen3:8b
```

Never commit the real `.env` file.

### 6. Configure Notion Access

Create a Notion integration and grant it access to the pages and databases that the agent should manage.

The agent can only access Notion content explicitly shared with the integration.

If a page is moved, made private, or newly created as a top-level page, the integration may need to be granted access again.

---

## Running the Agent

If you are using the cloned GitHub repository:

```bash
cd notion_local_agent
```

Activate the virtual environment:

```bash
source .venv/bin/activate
```

Start the agent:

```bash
python3 agent.py
```

When you see:

```text
You >
```

the agent is ready for natural-language instructions.

---

## Shell vs Agent Interface

When the terminal prompt ends with `%`, you are in the macOS shell.

Use shell commands such as:

```text
cd
source
python3
pip
git
```

When the prompt looks like:

```text
You >
```

you are inside the AI agent.

Use natural-language instructions.

---

## Built-in Agent Commands

- `/state` — show persistent operational state
- `/clear` — clear current conversation context
- `/forget` — clear persistent operational memory
- `/cancel` — cancel a pending batch update
- `/cancel-import` — cancel a pending spreadsheet import
- `exit` — exit the agent

---

## Example Workflow: Local File Search

```text
幫我在桌面找名稱包含 CV 的檔案
```

## Example Workflow: Read a PDF

```text
讀取桌面的 Job Description.pdf，
整理主要職責、技能要求與重要關鍵字
```

## Example Workflow: Create a Database Record

```text
在 Applications 表格新增一筆資料：

Company：EY
Role：Business Analyst
Deadline：2028-12-24
Status：Preparing

其餘欄位留空
```

## Example Workflow: Use Conversation Memory

```text
把剛才那筆 Status 改成 Applied
```

## Example Workflow: Query a Database

```text
在 Applications 找出 Deadline 是 2028-12-24 的資料
```

---

## Batch Update Workflow

For operations that may affect multiple records, the agent uses a preview-first workflow.

Example:

```text
把 Applications 裡 Status 是 Preparing 的所有資料改成 Applied
```

The user must then explicitly enter:

```text
確認
```

To cancel:

```text
/cancel
```

---

## Spreadsheet-to-Notion Import

Example:

```text
讀取桌面的 Applications.xlsx，
準備匯入 Applications database，
但先不要真的匯入
```

The workflow is:

```text
Read spreadsheet
↓
Read headers
↓
Inspect Notion schema
↓
Build column mapping
↓
Preview rows
↓
Wait for confirmation
```

To confirm:

```text
確認匯入
```

To cancel:

```text
/cancel-import
```

---

## Persistent Memory

The agent uses SQLite to remember selected operational context across restarts.

Examples include:

```text
last_page
last_database
last_record
last_file
```

---

## Notion Integration Permissions

If the agent says that a page or database cannot be found, but the object exists in Notion, check the integration permissions.

Common causes:

- The page was moved
- The page was moved into Private
- A new top-level page was created
- The integration was not granted access

---

## Ollama Troubleshooting

```bash
ollama list
ollama run qwen3:8b
```

---

## Python Troubleshooting

```bash
which python3
python3 -m py_compile agent.py
python3 agent.py
```

---

## Version History

### V1
- Basic Notion page search
- Child-page creation

### V2
- Added Notion blocks
- Added block validation

### V3
- Added database creation

### V4
- Added database record CRUD

### V4.1
- Added session conversation memory
- Added contextual references such as "the previous record"

### V4.2
- Added SQLite persistent operational memory
- Added non-title record lookup
- Added multi-field queries

### V4.3
- Added AND / OR queries
- Added batch update preview
- Added explicit confirmation
- Added pending batch state

### V5
- Added local macOS file access
- Added local file search
- Added directory browsing
- Added PDF, DOCX, XLSX, CSV, TXT, and Markdown reading
- Added spreadsheet-to-Notion import preview
- Added explicit import confirmation
- Added local path allowlist
- Added sensitive-file protection

---

## Current Limitations

- Local files are currently read-only
- Very large documents are not automatically chunked
- Large PDFs may exceed practical model context
- Spreadsheet column mapping mainly depends on header-name matching
- The terminal is currently the primary interface
- Background scheduling is not implemented
- Automatic synchronization is not implemented
- GUI support is not implemented
- Full Notion API coverage is not yet implemented

---

## Roadmap

Potential future improvements include:

- Large-document chunking
- Retrieval-based document processing
- Improved spreadsheet schema matching
- More robust import validation
- Local operation logs
- Background scheduling
- macOS LaunchAgent support
- Desktop GUI
- Expanded Notion API coverage
- More granular permissions
- Automated testing
- Plugin architecture

---

## Recommended Daily Workflow

```bash
cd ~/Desktop/"Notion Agent"
source .venv/bin/activate
python3 agent.py
```

Then use natural-language commands at `You >`.

---

## Development Workflow

```bash
cp agent.py agent_v5_backup.py
python3 -m py_compile agent.py
python3 agent.py
```

Recommended cycle:

```text
Modify
↓
Test
↓
git status
↓
git add
↓
git commit
↓
git push
```

---

## Git Workflow

```bash
git status
git add .
git commit -m "Describe the update"
git push
```

---

## Repository Security

The `.gitignore` excludes local-only and sensitive files such as:

```text
.env
.venv/
state.db
__pycache__/
*.pyc
.DS_Store
backup files
local test files
```

The repository includes `.env.example` instead of the real `.env`.

Example:

```env
NOTION_TOKEN=your_notion_token_here
OLLAMA_MODEL=qwen3:8b
```

---

## Project Goal

The long-term goal of this project is to build a flexible local AI workspace operator that can:

- understand natural-language instructions
- interact with local documents
- manage structured Notion content
- preserve operational context
- perform safe batch actions
- remain local-first
- minimize unnecessary cloud dependencies

The current V5 version represents the first complete local Mac + Notion workflow.

---

## License

No license has been selected yet.
