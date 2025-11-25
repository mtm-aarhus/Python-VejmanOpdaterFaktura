"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from robot_framework.process import process
import os
import time
# pylint: disable-next=unused-argum

start = time.perf_counter()
print("Starting")
orchestrator_connection = OrchestratorConnection(
    "VejmanOpdaterFaktura",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
)

process(orchestrator_connection)

# --- TIMER END ---
end = time.perf_counter()
elapsed = end - start

print(f"\n=== Robot runtime: {elapsed:.2f} seconds ({elapsed/60:.2f} minutes) ===")