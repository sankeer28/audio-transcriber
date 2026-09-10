"""One-step installer for GPU (Vulkan) transcription.

Downloads the whisper.cpp Vulkan binaries and a GGML model so main.py can use
your GPU instead of the CPU. Everything it fetches lands in whispercpp/ and
models/, both gitignored.

    python setup.py               # interactive
    python setup.py --yes         # accept defaults, no prompts
    python setup.py --model medium

Uses only the Python standard library, so it works before `pip install`.
"""

import argparse
import os
import shutil
import subprocess
import sys
import urllib.request
import zipfile

WHISPERCPP_ZIP = ("https://github.com/jerryshell/whisper.cpp-windows-vulkan-bin"
                  "/releases/download/v1.0.0/whisper.cpp-windows-vulkan.zip")
MODEL_BASE = "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/"

# name -> (filename, approx size for the prompt)
MODELS = {
    "tiny":     ("ggml-tiny.bin", "75 MB"),
    "base":     ("ggml-base.bin", "142 MB"),
    "small":    ("ggml-small.bin", "466 MB"),
    "medium":   ("ggml-medium.bin", "1.5 GB"),
    "large-v3": ("ggml-large-v3.bin", "2.9 GB"),
    "large-v3-turbo": ("ggml-large-v3-turbo.bin", "1.6 GB"),
}

BIN_DIR = "whispercpp"
MODEL_DIR = "models"


def human(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f}{unit}"
        n /= 1024


def download(url, dest):
    """Download with a single-line progress indicator."""
    print(f"  -> {os.path.basename(dest)}")
    tmp = dest + ".part"

    def hook(blocks, block_size, total):
        done = blocks * block_size
        if total > 0:
            pct = min(100, done * 100 / total)
            bar = "#" * int(pct / 2.5)
            sys.stdout.write(f"\r     [{bar:<40}] {pct:5.1f}%  {human(done)}/{human(total)}")
        else:
            sys.stdout.write(f"\r     {human(done)}")
        sys.stdout.flush()

    try:
        urllib.request.urlretrieve(url, tmp, hook)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    sys.stdout.write("\n")
    os.replace(tmp, dest)


def confirm(question, assume_yes):
    if assume_yes:
        return True
    return input(f"{question} [Y/n] ").strip().lower() in ("", "y", "yes")


def install_binaries(assume_yes):
    exe = "whisper-cli.exe" if sys.platform == "win32" else "whisper-cli"
    target = os.path.join(BIN_DIR, exe)
    if os.path.isfile(target):
        print(f"[OK] whisper.cpp already installed ({target})")
        return True

    if sys.platform != "win32":
        print("[!] Prebuilt Vulkan binaries are only packaged for Windows.")
        print("    On Linux/macOS, build whisper.cpp with -DGGML_VULKAN=1 and put")
        print(f"    whisper-cli in {BIN_DIR}/, or just use the CPU engine (no setup needed).")
        return False

    print(f"\nwhisper.cpp Vulkan binaries (~18 MB)")
    print(f"  source: {WHISPERCPP_ZIP}")
    if not confirm("Download?", assume_yes):
        return False

    os.makedirs(BIN_DIR, exist_ok=True)
    zip_path = os.path.join(BIN_DIR, "_whispercpp.zip")
    download(WHISPERCPP_ZIP, zip_path)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(BIN_DIR)
    os.remove(zip_path)
    if not os.path.isfile(target):
        print(f"[!] Expected {target} after extraction but it is missing")
        return False
    print(f"[OK] Installed to {BIN_DIR}/")
    return True


def install_model(name, assume_yes):
    filename, size = MODELS[name]
    dest = os.path.join(MODEL_DIR, filename)
    if os.path.isfile(dest):
        print(f"[OK] Model already present ({dest})")
        return dest

    print(f"\nModel '{name}' ({size})")
    if not confirm("Download?", assume_yes):
        return None

    os.makedirs(MODEL_DIR, exist_ok=True)
    download(MODEL_BASE + filename, dest)
    print(f"[OK] Saved to {dest}")
    return dest


def verify(model_path):
    """Run whisper-cli so the user can see which GPU it picked up."""
    exe = "whisper-cli.exe" if sys.platform == "win32" else "whisper-cli"
    binary = os.path.join(BIN_DIR, exe)
    if not os.path.isfile(binary) or not model_path:
        return
    print("\nDetecting GPU devices...")
    try:
        out = subprocess.run([binary, "--help"], capture_output=True, text=True,
                             timeout=60).stderr
    except Exception as e:
        print(f"[!] Could not run {binary}: {e}")
        return
    devices = [ln for ln in out.splitlines() if "ggml_vulkan" in ln]
    if devices:
        for ln in devices:
            print("  " + ln.strip())
        print("\n[OK] Vulkan is working. If the GPU you want is not device 0,")
        print("     set WHISPERCPP_GPU_DEVICE in main.py to its index.")
    else:
        print("[!] No Vulkan devices reported. Your GPU driver may not support Vulkan;")
        print("    main.py will fall back to the CPU engine automatically.")


def main():
    ap = argparse.ArgumentParser(description="Install GPU transcription support")
    ap.add_argument("--yes", "-y", action="store_true", help="accept defaults, no prompts")
    ap.add_argument("--model", default="large-v3", choices=sorted(MODELS),
                    help="GGML model to download (default: large-v3, most accurate)")
    args = ap.parse_args()

    print("=" * 62)
    print(" Audio Transcriber - GPU setup")
    print("=" * 62)
    print(f"Platform: {sys.platform}   Python: {sys.version.split()[0]}")

    if shutil.which("ffmpeg") is None:
        print("\n[!] ffmpeg was not found on PATH. It is required to read mp4/mp3 files.")
        print("    Windows: choco install ffmpeg   macOS: brew install ffmpeg")
        print("    Linux:   sudo apt install ffmpeg")
    else:
        print("[OK] ffmpeg found")

    have_bin = install_binaries(args.yes)
    model_path = install_model(args.model, args.yes) if have_bin else None

    if have_bin and model_path:
        verify(model_path)
        print("\nDone. Put files in presentations/ and run:  python main.py")
    else:
        print("\nGPU setup incomplete - main.py will use the CPU engine instead,")
        print("which needs no setup (it downloads its own model on first run).")


if __name__ == "__main__":
    main()
