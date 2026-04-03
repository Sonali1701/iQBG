import psutil
print("Scanning for orphaned uvicorn processes...")
for p in psutil.process_iter(['pid', 'name', 'cmdline']):
    try:
        if p.info['name'] and 'python.exe' in p.info['name'] and p.info['cmdline'] and 'uvicorn' in p.info['cmdline']:
            print(f"Killing PID {p.info['pid']}")
            p.kill()
    except Exception as e:
        pass
print("Done.")
