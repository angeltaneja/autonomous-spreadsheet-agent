import os
import subprocess
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pathlib import Path

from backend.config import HOST, PORT, EXCEL_OUTPUT_PATH
from backend.state import state, reset_state, add_log
# We will import agent run function when we create it

app = FastAPI(title="Agentic Spreadsheet Automation API")

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class RunPromptRequest(BaseModel):
    prompt: str

# Define placeholder for agent running function (implemented in agent.py)
agent_runner = None

def run_agent_in_background(prompt: str):
    global agent_runner
    if agent_runner is None:
        try:
            from backend.agent import run_agent
            agent_runner = run_agent
        except ImportError as e:
            add_log(f"Error loading agent: {e}")
            state.status = "failed"
            state.error = f"Agent not implemented: {e}"
            return
            
    try:
        agent_runner(prompt)
    except Exception as e:
        add_log(f"Unhandled agent error: {e}")
        state.status = "failed"
        state.error = str(e)

@app.post("/api/run")
async def run_agent_endpoint(request: RunPromptRequest, background_tasks: BackgroundTasks):
    if state.status == "running" or state.status == "planning":
        raise HTTPException(status_code=400, detail="Agent is already running a task.")
    
    reset_state()
    state.status = "planning"
    add_log(f"Received prompt: '{request.prompt}'")
    background_tasks.add_task(run_agent_in_background, request.prompt)
    return {"status": "started"}

@app.get("/api/state")
async def get_state():
    return state

@app.post("/api/excel/open")
async def open_excel():
    if not os.path.exists(EXCEL_OUTPUT_PATH):
        raise HTTPException(status_code=404, detail="Excel file does not exist yet. Please run the agent first.")
    
    try:
        # Use subprocess start on Windows — more reliable than os.startfile in threaded context
        subprocess.Popen(["start", "", EXCEL_OUTPUT_PATH], shell=True)
        return {"status": "success", "message": "Excel workbook opened.", "path": EXCEL_OUTPUT_PATH}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to open Excel file: {str(e)}")

@app.get("/api/sheets/url")
async def get_sheets_url():
    if not state.google_sheets_url:
        raise HTTPException(status_code=404, detail="No Google Sheet has been created yet.")
    return {"url": state.google_sheets_url}

@app.get("/api/excel/path")
async def get_excel_path():
    if not state.excel_path:
        raise HTTPException(status_code=404, detail="No Excel file has been created yet.")
    return {"path": state.excel_path}

# Mount the frontend static files if directory exists
frontend_dir = Path(__file__).resolve().parent.parent / "frontend"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
else:
    # If frontend doesn't exist yet, we will mount it later
    pass
