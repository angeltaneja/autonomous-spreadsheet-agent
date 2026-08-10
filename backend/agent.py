import os
from google import genai
from google.genai import types

from backend.config import (
    GEMINI_API_KEY,
    CSV_OUTPUT_PATH,
    EXCEL_OUTPUT_PATH,
    GOOGLE_SHEET_TITLE
)
from backend.state import state, add_log, update_step
from backend.tools import (
    generate_employee_csv,
    import_csv_to_excel,
    import_csv_to_google_sheets
)

# Tool function map for execution
TOOL_MAP = {
    "generate_employee_csv": generate_employee_csv,
    "import_csv_to_excel": import_csv_to_excel,
    "import_csv_to_google_sheets": import_csv_to_google_sheets
}

# Models that reliably support function/tool calling
CAPABLE_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
    "gemini-pro-latest",
    "gemini-flash-lite-latest",
]


def _execute_tool(fn_name: str, fn_args: dict) -> str:
    """Run a single tool and update state. Returns result string."""
    # Always enforce configured paths
    if fn_name == "generate_employee_csv":
        fn_args["filename"] = CSV_OUTPUT_PATH
        update_step("csv", "running", "Generating employee CSV data...")
    elif fn_name == "import_csv_to_excel":
        fn_args["csv_path"] = CSV_OUTPUT_PATH
        fn_args["excel_path"] = EXCEL_OUTPUT_PATH
        update_step("excel", "running", "Importing CSV into Microsoft Excel...")
    elif fn_name == "import_csv_to_google_sheets":
        fn_args["csv_path"] = CSV_OUTPUT_PATH
        fn_args["sheet_title"] = GOOGLE_SHEET_TITLE
        update_step("sheets", "running", "Uploading to Google Sheets...")

    result = TOOL_MAP[fn_name](**fn_args)

    if fn_name == "generate_employee_csv":
        state.csv_data = result
        update_step("csv", "completed", f"Generated {len(result)} employee records.")
        add_log(f"Tool generate_employee_csv: created {len(result)} records.")
    elif fn_name == "import_csv_to_excel":
        state.excel_path = result
        update_step("excel", "completed", f"Excel saved to {result}")
        add_log(f"Tool import_csv_to_excel: saved to '{result}'")
    elif fn_name == "import_csv_to_google_sheets":
        state.google_sheets_url = result
        update_step("sheets", "completed", "Google Sheet created and ready.")
        add_log(f"Tool import_csv_to_google_sheets: sheet at '{result}'")

    return str(result)


def _run_direct(prompt: str):
    """
    Fallback: execute the full workflow directly without LLM tool-calling.
    Used when all models are quota-exhausted or hallucinate without calling tools.
    """
    add_log("[Direct Mode] Running tool pipeline directly...")
    add_log(f"[Direct Mode] Task: {prompt}")

    add_log("[Direct Mode] Step 1: Generating employee CSV...")
    csv_result = _execute_tool("generate_employee_csv", {})

    add_log("[Direct Mode] Step 2: Importing CSV into Excel...")
    excel_result = _execute_tool("import_csv_to_excel", {})

    add_log("[Direct Mode] Step 3: Uploading to Google Sheets...")
    sheets_result = _execute_tool("import_csv_to_google_sheets", {})

    add_log("[Direct Mode] All tools executed successfully.")


def _run_with_llm(prompt: str, api_key: str) -> bool:
    """
    Try to run via Gemini tool-calling loop.
    Returns True if LLM actually called tools and completed the task.
    Returns False if LLM didn't call any tools (hallucinated).
    """
    system_instruction = (
        "You are an autonomous spreadsheet automation AI agent.\n"
        "Your objective is to complete the user's task by calling the available tools.\n\n"
        "STRICT RULES:\n"
        "1. You MUST call tools to complete the task. Do NOT describe what you did without calling tools.\n"
        "2. Execute in this exact order:\n"
        "   a. Call `generate_employee_csv` to create the data file.\n"
        "   b. Call `import_csv_to_excel` to import it into Excel.\n"
        "   c. Call `import_csv_to_google_sheets` to upload to Google Sheets.\n"
        "3. Report progress after each tool call completes.\n"
        "4. NEVER fabricate tool results. Always use actual return values."
    )

    tools_list = [generate_employee_csv, import_csv_to_excel, import_csv_to_google_sheets]
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        tools=tools_list,
        temperature=0.1
    )

    for model_name in CAPABLE_MODELS:
        try:
            add_log(f"Trying model: {model_name}...")
            client = genai.Client(api_key=api_key)
            chat = client.chats.create(model=model_name, config=config)
            response = chat.send_message(prompt)
            add_log(f"Agent connected: {model_name}")

            tools_called = 0

            # Tool-calling loop
            while True:
                if response.text:
                    # Truncate long agent thoughts to keep logs clean
                    thought = response.text.strip()[:300]
                    add_log(f"Agent: {thought}")

                function_calls = response.function_calls
                if not function_calls:
                    add_log("Agent: no more tool calls.")
                    break

                tool_responses = []
                for call in function_calls:
                    fn_name = call.name
                    fn_args = dict(call.args) if call.args else {}
                    call_id = getattr(call, "id", None)

                    add_log(f"Calling tool: {fn_name}")
                    if fn_name in TOOL_MAP:
                        try:
                            result = _execute_tool(fn_name, fn_args)
                            tools_called += 1
                            tool_responses.append(
                                types.Part.from_function_response(
                                    name=fn_name,
                                    response={"status": "success", "result": result},
                                    id=call_id
                                )
                            )
                        except Exception as tool_err:
                            err_msg = str(tool_err)
                            add_log(f"Tool error ({fn_name}): {err_msg}")
                            # Mark step failed
                            step_map = {
                                "generate_employee_csv": "csv",
                                "import_csv_to_excel": "excel",
                                "import_csv_to_google_sheets": "sheets"
                            }
                            if fn_name in step_map:
                                update_step(step_map[fn_name], "failed", err_msg)
                            tool_responses.append(
                                types.Part.from_function_response(
                                    name=fn_name,
                                    response={"status": "failed", "error": err_msg},
                                    id=call_id
                                )
                            )
                    else:
                        tool_responses.append(
                            types.Part.from_function_response(
                                name=fn_name,
                                response={"status": "failed", "error": f"Unknown tool: {fn_name}"},
                                id=call_id
                            )
                        )

                add_log("Sending tool results back to agent...")
                response = chat.send_message(tool_responses)

            # If the model responded but called 0 tools, it hallucinated
            if tools_called == 0:
                add_log(f"Model {model_name} did not call any tools (hallucination). Using direct mode.")
                return False

            return True

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                add_log(f"  SKIP {model_name}: quota exhausted")
            elif "404" in err_str or "NOT_FOUND" in err_str:
                add_log(f"  SKIP {model_name}: not found")
            else:
                add_log(f"  SKIP {model_name}: {err_str[:100]}")

    add_log("All LLM models exhausted. Switching to direct execution mode.")
    return False


def run_agent(prompt: str):
    """
    Main agent entry point. Tries Gemini tool-calling first,
    falls back to direct tool execution if all models fail or hallucinate.
    """
    state.status = "running"
    add_log("Starting agent workflow...")

    api_key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")

    try:
        # Try LLM-driven tool calling first
        llm_succeeded = False
        if api_key:
            llm_succeeded = _run_with_llm(prompt, api_key)
        else:
            add_log("No API key — running in direct mode.")

        # If LLM didn't call tools, run directly
        if not llm_succeeded:
            _run_direct(prompt)

        # Final verification
        csv_ok = bool(state.csv_data)
        excel_ok = bool(state.excel_path) and os.path.exists(state.excel_path)
        sheets_ok = bool(state.google_sheets_url)

        add_log(f"Verify: CSV={csv_ok} | Excel={excel_ok} | Sheets={sheets_ok}")

        if csv_ok and excel_ok and sheets_ok:
            update_step("verify", "completed", "All 3 tasks completed successfully.")
            state.status = "completed"
            add_log("Workflow completed successfully!")
        elif csv_ok and excel_ok:
            update_step("verify", "completed", "CSV + Excel completed. Google Sheets skipped.")
            state.status = "completed"
            add_log("Workflow completed (Excel only).")
        else:
            failed = []
            if not csv_ok: failed.append("CSV")
            if not excel_ok: failed.append("Excel")
            if not sheets_ok: failed.append("Google Sheets")
            msg = f"Failed steps: {', '.join(failed)}"
            update_step("verify", "failed", msg)
            state.status = "failed"
            state.error = msg
            add_log(f"Workflow failed: {msg}")

    except Exception as ex:
        err_str = str(ex)
        add_log(f"Fatal error: {err_str}")
        state.status = "failed"
        state.error = err_str
        update_step("verify", "failed", f"Error: {err_str}")
