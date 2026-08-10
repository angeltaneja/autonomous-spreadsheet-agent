# Autonomous Spreadsheet Automation Agent

An autonomous, production-quality AI Agent built to automate spreadsheet workflows. The agent accepts natural language commands to generate realistic datasets, load and style them in Microsoft Excel, upload them to Google Sheets, and verify completion—all driven by dynamic tool selection and reasoning.

It comes equipped with a premium, glassmorphic dark-mode web control center to run prompts, monitor progress, preview generated data, read logs, and open files.

---

## Key Features
* **Autonomous Reasoning**: Uses the Gemini 2.5 Flash model to interpret instructions, formulate a step-by-step execution plan, select the correct tools dynamically, and handle retries.
* **Microsoft Excel COM Automation**: Opens local Excel, imports the data, applies high-quality typography (Segoe UI), adds headers (Navy background, bold white text), inserts borders, applies alternating row zebra-striping, formats currencies, and auto-fits columns.
* **Google Sheets API Integration**: Dynamically connects using Google's APIs, creates a new spreadsheet, structures the data, applies identical formatting, and optionally opens drive access.
* **Premium Web Dashboard**: A glassmorphic web UI featuring:
  - A command input box.
  - A real-time log terminal showing agent thoughts and tool parameters.
  - A visual stepper showing step-by-step progress.
  - An interactive preview table of the generated dataset.
  - Instant click-to-open links for Excel and Google Sheets.

---

## Prerequisites
* **Operating System**: Windows (required for local Excel COM automation).
* **Python**: Version 3.12 or higher.
* **Office**: Microsoft Excel installed on the computer.

---

## Setup & Installation

### 1. Install Dependencies
Open your terminal in the project directory and run:
```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables
Copy `.env.template` to `.env`:
```bash
copy .env.template .env
```
Open `.env` and enter your Gemini API Key:
```env
GEMINI_API_KEY=your_actual_gemini_api_key_here
```

### 3. Configure Google Sheets API Credentials
The application supports two authentication methods for Google Sheets. Place **one** of these files in the root of the project:

#### Option A: Service Account Key (Recommended)
1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project, enable the **Google Sheets API** and the **Google Drive API**.
3. Create a **Service Account** under *IAM & Admin > Service Accounts*.
4. Generate a new JSON Key for the Service Account and download it.
5. Save the file in the project root as `service_account.json`.

#### Option B: OAuth Client ID (Interactive User Login)
1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Enable the **Google Sheets API** and **Google Drive API**.
3. Configure the OAuth Consent Screen (Internal or Testing) and add your email.
4. Go to *Credentials > Create Credentials > OAuth client ID*, select **Desktop App** as the application type.
5. Download the client secrets JSON.
6. Save the file in the project root as `credentials.json`.
   * *Note*: On the first run, the agent will launch a browser window asking you to authorize the application. It will then cache the token in `token.json` so you do not need to log in again.

---

## Running the Application

### Method 1: The Quick Startup Script (Recommended)
Simply double-click the **`run.bat`** file in the project root. This will:
1. Initialize the `.env` file if missing.
2. Automatically launch your browser to `http://127.0.0.1:8000`.
3. Start the FastAPI backend server.

### Method 2: Manual Start
Run the uvicorn server manually:
```bash
python -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
```
Then navigate to [http://127.0.0.1:8000](http://127.0.0.1:8000) in your browser.

---

## Interacting with the Agent

1. Type the following command in the Command Center input box:
   ```text
   Create an employee CSV and import it into Excel and Google Sheets.
   ```
2. Click **Run Agent**.
3. Watch the progress:
   - The status stepper will light up as steps run.
   - The log terminal will show Gemini planning: *"First, I need to generate... then I will..."*
   - Excel will open in the background, style and save your workbook, and close.
   - The Google Sheet link and "Open Excel Locally" buttons will activate once completed.
   - The preview table will display the generated employee rows.
