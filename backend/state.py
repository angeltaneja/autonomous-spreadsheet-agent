from pydantic import BaseModel
from typing import List, Dict, Any, Optional

class Step(BaseModel):
    id: str
    name: str
    status: str = "pending"  # pending, running, completed, failed
    message: str = ""

class AgentState(BaseModel):
    status: str = "idle"  # idle, planning, running, completed, failed
    error: Optional[str] = None
    steps: List[Step] = [
        Step(id="csv", name="Generate Employee CSV"),
        Step(id="excel", name="Import to Microsoft Excel"),
        Step(id="sheets", name="Import to Google Sheets"),
        Step(id="verify", name="Verify and Finalize")
    ]
    logs: List[str] = []
    csv_data: List[Dict[str, Any]] = []
    excel_path: Optional[str] = None
    google_sheets_url: Optional[str] = None

# Global agent state singleton
state = AgentState()

def reset_state():
    global state
    state.status = "idle"
    state.error = None
    state.logs = []
    state.csv_data = []
    state.excel_path = None
    state.google_sheets_url = None
    for step in state.steps:
        step.status = "pending"
        step.message = ""

def add_log(message: str):
    print(f"[Agent Log] {message}")
    state.logs.append(message)

def update_step(step_id: str, status: str, message: str = ""):
    for step in state.steps:
        if step.id == step_id:
            step.status = status
            if message:
                step.message = message
