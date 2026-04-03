import os
import re
import subprocess

out = subprocess.check_output("netstat -ano | findstr :8000", shell=True).decode()
pids = []
for line in out.splitlines():
    if "LISTENING" in line:
        parts = line.split()
        pid = parts[-1]
        pids.append(pid)

for pid in set(pids):
    try:
        os.system(f"taskkill /PID {pid} /F")
        print(f"Killed PID {pid}")
    except Exception as e:
        print(e)
