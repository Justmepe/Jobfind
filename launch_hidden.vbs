' Launches the jobwatch runner with NO visible console window.
' Window style 0 = hidden; True = wait for completion so the scheduled
' task reflects the real exit and honors its time limit.
CreateObject("WScript.Shell").Run """d:\SBS\Resume\job-pipeline\run_jobwatch.cmd""", 0, True
