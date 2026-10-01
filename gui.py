"""Simple desktop UI for the audio transcriber.

Pick an input folder (or a single file), pick where the transcripts go, press
Start. Progress and log output appear as the work happens.

Run it with:  .venv\\Scripts\\pythonw.exe gui.py      (no console window)
          or  .venv\\Scripts\\python.exe gui.py       (keeps a console)

This is a wrapper around main.py - all the transcription settings still live
at the top of main.py.
"""

import os
import re
import sys
import queue
import threading
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
MAIN = os.path.join(HERE, "main.py")
MEDIA_EXT = (".pptx", ".mp3", ".mp4", ".docx")

# Lines main.py prints that we turn into progress
RE_PERCENT = re.compile(r"(\d+)\s*%")
RE_FOUND = re.compile(r"\[Files\] Total: (\d+)")
RE_START = re.compile(r"\[(?:PPTX|MP3|MP4|DOCX)\] Processing (.+?)\.\.\.")
RE_SAVED = re.compile(r"\[OK\] Saved .*? to (.+)")
RE_ENGINE = re.compile(r"\[Auto\] Using (.+)")


class TranscriberUI:
    def __init__(self, root):
        self.root = root
        self.proc = None
        self.q = queue.Queue()
        self.total_files = 0
        self.done_files = 0

        root.title("Audio Transcriber")
        root.minsize(680, 460)

        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(root)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        # --- input ---
        ttk.Label(main, text="Input").grid(row=0, column=0, sticky="w", **pad)
        self.in_var = tk.StringVar(value=os.path.join(HERE, "presentations"))
        ttk.Entry(main, textvariable=self.in_var).grid(row=0, column=1, sticky="ew", **pad)
        btns = ttk.Frame(main)
        btns.grid(row=0, column=2, sticky="e", padx=(0, 10))
        ttk.Button(btns, text="Folder...", width=9, command=self.pick_in_folder).pack(side="left")
        ttk.Button(btns, text="File...", width=8, command=self.pick_in_file).pack(side="left", padx=(4, 0))

        # --- output ---
        ttk.Label(main, text="Output").grid(row=1, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar(value=os.path.join(HERE, "output"))
        ttk.Entry(main, textvariable=self.out_var).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(main, text="Folder...", width=9,
                   command=self.pick_out_folder).grid(row=1, column=2, sticky="e", padx=(0, 10))

        # --- options ---
        opts = ttk.Frame(main)
        opts.grid(row=2, column=0, columnspan=3, sticky="w", padx=10)
        self.clean_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Remove repeated text",
                        variable=self.clean_var).pack(side="left")
        self.backup_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Keep raw copy (_backup.txt)",
                        variable=self.backup_var).pack(side="left", padx=(16, 0))

        # --- progress ---
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(main, textvariable=self.status).grid(row=3, column=0, columnspan=3,
                                                       sticky="w", padx=10, pady=(12, 2))
        self.bar = ttk.Progressbar(main, maximum=100)
        self.bar.grid(row=4, column=0, columnspan=3, sticky="ew", padx=10)

        self.overall = tk.StringVar(value="")
        ttk.Label(main, textvariable=self.overall).grid(row=5, column=0, columnspan=3,
                                                        sticky="w", padx=10, pady=(2, 0))

        # --- log ---
        logf = ttk.Frame(main)
        logf.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=10, pady=(10, 4))
        main.rowconfigure(6, weight=1)
        self.log = tk.Text(logf, height=12, wrap="word", state="disabled",
                           font=("Consolas", 9))
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # --- buttons ---
        bar = ttk.Frame(main)
        bar.grid(row=7, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        self.start_btn = ttk.Button(bar, text="Start", command=self.start)
        self.start_btn.pack(side="left")
        self.cancel_btn = ttk.Button(bar, text="Cancel", command=self.cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Open output folder",
                   command=self.open_output).pack(side="right")

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.drain)

    # ---------- pickers ----------
    def pick_in_folder(self):
        d = filedialog.askdirectory(title="Choose the folder with your videos/audio",
                                    initialdir=self._start_dir(self.in_var.get()))
        if d:
            self.in_var.set(os.path.normpath(d))

    def pick_in_file(self):
        f = filedialog.askopenfilename(
            title="Choose a file to transcribe",
            initialdir=self._start_dir(self.in_var.get()),
            filetypes=[("Supported", "*.pptx *.mp3 *.mp4 *.docx"), ("All files", "*.*")])
        if f:
            self.in_var.set(os.path.normpath(f))

    def pick_out_folder(self):
        d = filedialog.askdirectory(title="Where should the transcripts go?",
                                    initialdir=self._start_dir(self.out_var.get()))
        if d:
            self.out_var.set(os.path.normpath(d))

    @staticmethod
    def _start_dir(path):
        if os.path.isdir(path):
            return path
        parent = os.path.dirname(path)
        return parent if os.path.isdir(parent) else HERE

    def open_output(self):
        out = self.out_var.get()
        if not os.path.isdir(out):
            messagebox.showinfo("Not there yet", f"{out}\n\nThis folder doesn't exist yet.")
            return
        if sys.platform == "win32":
            os.startfile(out)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", out])
        else:
            subprocess.Popen(["xdg-open", out])

    # ---------- logging ----------
    def write(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------- run ----------
    def start(self):
        src = self.in_var.get().strip()
        dst = self.out_var.get().strip()
        if not os.path.exists(src):
            messagebox.showerror("Input not found", f"Can't find:\n{src}")
            return
        if os.path.isdir(src):
            hits = [f for f in os.listdir(src) if f.lower().endswith(MEDIA_EXT)]
            if not hits:
                messagebox.showerror(
                    "Nothing to do",
                    f"No .pptx, .mp3, .mp4 or .docx files in:\n{src}")
                return
        if not dst:
            messagebox.showerror("No output folder", "Choose where the transcripts should go.")
            return

        cmd = [sys.executable, "-u", MAIN, src, "-o", dst]
        if not self.clean_var.get():
            cmd.append("--no-clean")
        if self.backup_var.get():
            cmd.append("--backup")

        self.total_files = self.done_files = 0
        self.bar["value"] = 0
        self.overall.set("")
        self.status.set("Starting...")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        try:
            self.proc = subprocess.Popen(
                cmd, cwd=HERE,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True, encoding="utf-8", errors="replace", bufsize=0,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
        except Exception as e:
            messagebox.showerror("Could not start", str(e))
            self.finish()
            return

        threading.Thread(target=self.reader, args=(self.proc,), daemon=True).start()

    def reader(self, proc):
        """Read child output and push whole lines onto the queue.

        tqdm redraws with \\r, so split on both carriage returns and newlines.
        """
        buf = ""
        while True:
            chunk = proc.stdout.read(256)
            if not chunk:
                break
            buf += chunk
            parts = re.split(r"[\r\n]", buf)
            buf = parts.pop()
            for line in parts:
                if line.strip():
                    self.q.put(("line", line.rstrip()))
        if buf.strip():
            self.q.put(("line", buf.strip()))
        self.q.put(("done", proc.wait()))

    def drain(self):
        """Pull queued output on the Tk thread and update the widgets."""
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "line":
                    self.handle(payload)
                else:
                    self.finish(payload)
        except queue.Empty:
            pass
        self.root.after(100, self.drain)

    def handle(self, line):
        m = RE_FOUND.search(line)
        if m:
            self.total_files = int(m.group(1))
            self.overall.set(f"0 of {self.total_files} files done")
            return  # the "[Files] Found ..." line above already says this

        m = RE_START.search(line)
        if m:
            self.bar["value"] = 0
            name = os.path.basename(m.group(1))
            n = self.done_files + 1
            self.status.set(f"Transcribing {name}"
                            + (f"  ({n} of {self.total_files})" if self.total_files else ""))
            self.write(line)
            return

        if RE_SAVED.search(line):
            self.done_files += 1
            self.bar["value"] = 100
            if self.total_files:
                self.overall.set(f"{self.done_files} of {self.total_files} files done")
            self.write(line)
            return

        if "Transcribing" in line and "%" in line:
            m = RE_PERCENT.search(line)
            if m:
                self.bar["value"] = int(m.group(1))
            return  # don't spam the log with every bar redraw

        m = RE_ENGINE.search(line)
        if m:
            self.status.set(m.group(1))

        self.write(line)

    def cancel(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            self.status.set("Cancelling...")

    def finish(self, code=None):
        self.proc = None
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        if code == 0:
            self.bar["value"] = 100
            self.status.set("Done.")
        elif code is not None:
            self.status.set(f"Stopped (exit code {code}).")

    def on_close(self):
        if self.proc and self.proc.poll() is None:
            if not messagebox.askokcancel("Still running",
                                          "Transcription is still running. Stop it and quit?"):
                return
            self.proc.terminate()
        self.root.destroy()


def main():
    if not os.path.isfile(MAIN):
        print(f"main.py not found next to gui.py ({MAIN})")
        return
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform == "win32" else "clam")
    except tk.TclError:
        pass
    TranscriberUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
