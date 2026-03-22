# run.py
import subprocess, sys, time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class Reloader(FileSystemEventHandler):
    def __init__(self):
        self.proc = self.start()

    def start(self):
        return subprocess.Popen([sys.executable, "app.py"])

    def on_modified(self, event):
        if event.src_path.endswith("app.py"):
            self.proc.kill()
            time.sleep(0.3)
            self.proc = self.start()

observer = Observer()
observer.schedule(Reloader(), ".", recursive=False)
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()