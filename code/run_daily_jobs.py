#!/usr/bin/env python3
import json
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
STATE_PATH = BASE_DIR / "daily_run_state.json"
UPDATE_SCRIPT = BASE_DIR / "update_jobs.py"
MIN_UPTIME_SECONDS = 600


def load_state():
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_state(state):
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def uptime_seconds():
    return time.monotonic()


def main():
    today = date.today().isoformat()
    state = load_state()

    if state.get("last_success_date") == today:
        print(f"{datetime.now().isoformat(timespec='seconds')} already updated today")
        return 0

    if uptime_seconds() < MIN_UPTIME_SECONDS:
        print(f"{datetime.now().isoformat(timespec='seconds')} skipped: waiting for 10 minutes after boot")
        return 0

    result = subprocess.run([sys.executable, str(UPDATE_SCRIPT)], cwd=str(BASE_DIR))
    if result.returncode == 0:
        state["last_success_date"] = today
        state["last_success_at"] = datetime.now().isoformat(timespec="seconds")
        save_state(state)
    else:
        state["last_failure_at"] = datetime.now().isoformat(timespec="seconds")
        state["last_failure_code"] = result.returncode
        save_state(state)
    return result.returncode


if __name__ == "__main__":
    sys.exit(main())
