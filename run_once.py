"""
run_once.py — called by GitHub Actions, runs the job exactly once.
"""
from mis_automation import run_job

if __name__ == "__main__":
    run_job()
