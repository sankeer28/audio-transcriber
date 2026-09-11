#  Audio Transcriber

Extract text and transcribe audio from PowerPoint presentations, MP4 videos, and MP3 files using Whisper.

## Features

- **Text Extraction**: Extracts all text content from PowerPoint slides
- **Audio Transcription**: Uses faster-whisper to transcribe audio from multiple sources:
  - PowerPoint files (.pptx) - embedded audio recordings
  - Video files (.mp4) - audio track extraction
  - Audio files (.mp3) - direct transcription
- **Checkpoint/Resume Support**: Automatically saves progress during long transcriptions
  - Resume from where you left off if interrupted - the decoder seeks to the last
    saved timestamp, so already-transcribed audio is not processed again
  - Checkpoints saved every 10 segments
  - Safe to stop with Ctrl+C anytime
- **Live Progress Tracking**: Real-time progress bar and live checkpoint file
  - While a transcription is running, view progress in `output/[filename]_checkpoint.json`
  - See completed segments as they're processed
  - Track timestamp and text in real-time
  - The checkpoint is deleted once the file finishes successfully
- **GPU and CPU Support**: Automatic device detection with intelligent fallback
- **Multiple Models**: Supports various Whisper model sizes (tiny, base, small, medium, large)
- **Configurable**: Easy-to-modify settings for performance and quality tuning

## Quick Start

```bash
# 1. Install Python dependencies
pip install -r requirements.txt

# 2. (Optional but recommended) install GPU acceleration
python setup.py

# 3. Drop your .pptx / .mp4 / .mp3 files into presentations/
# 4. Transcribe - results appear in output/
python main.py
```

That's it. `main.py` picks the fastest engine available on your machine
automatically; no configuration is required.

You also need **ffmpeg** on your PATH (see [Installing ffmpeg](#installing-ffmpeg)).

### Should I run setup.py?

`setup.py` downloads whisper.cpp with the Vulkan GPU backend plus a model
(~3 GB). It is worth it if you have any dedicated GPU:

| | Without setup.py | With setup.py |
|---|---|---|
| AMD / Intel GPU | CPU only | GPU - much faster |
| NVIDIA GPU | GPU via CUDA | GPU via Vulkan |
| No GPU | CPU only | CPU only (skip it) |

On an AMD RX 9070 XT, a 15.8-minute lecture took **43 seconds** with `setup.py`
versus an estimated ~12 minutes on CPU alone.

Skipping it is fine - `main.py` falls back to the CPU engine, which downloads
its own model on first run and needs no setup at all.

If the GPU binaries are installed but the model is missing, `main.py` offers to
download it for you rather than quietly dropping to CPU. Set
`AUTO_DOWNLOAD_MODEL = False` to turn that prompt off.

## Requirements

- Python 3.8 or higher
- ffmpeg
- A GPU is optional - everything works on CPU, just slower

<a name="installing-ffmpeg"></a>
### Installing ffmpeg

```bash
# Windows
choco install ffmpeg          # or download from https://ffmpeg.org/download.html

# macOS
brew install ffmpeg

# Linux
sudo apt install ffmpeg
```

### Optional: openai-whisper engine

The default engines do not need PyTorch. Only install these extras if you
specifically want `TRANSCRIPTION_ENGINE = "standard"` (several GB):

```bash
pip install -r requirements-standard.txt
```

## Usage

### 1. Prepare Your Files

- Place your files in the `presentations` folder:
  - PowerPoint presentations (.pptx)
  - Video files (.mp4)
  - Audio files (.mp3)

### 2. Run the Transcriber

```bash
python main.py
```

## Configuration

Edit the configuration settings at the top of `main.py`:

### Transcription Engine
```python
TRANSCRIPTION_ENGINE = "auto"   # "auto", "whisper.cpp", "faster-whisper", "standard"
```
`"auto"` (the default) picks whisper.cpp on the GPU if `setup.py` has been run,
then an NVIDIA GPU via CUDA, then CPU. Set it explicitly to override.

### Folder Settings
```python
PPTX_FOLDER = "presentations"   # Input folder
OUTPUT_FOLDER = "output"        # Output folder
```

### Whisper Model Settings
```python
WHISPER_MODEL = "large-v3"    # Options: "tiny", "base", "small", "medium", "large-v3", "large-v3-turbo"
FORCE_LANGUAGE = "en"         # Force language ("en", "es", "fr", etc.) or None
```

### Performance Settings
```python
FORCE_DEVICE = "cpu"          # Options: None (auto), "cpu", "cuda"
USE_HALF_PRECISION = False    # Enable fp16 for speed boost (NVIDIA GPU only)
CPU_THREADS = 0               # CPU worker threads (0 = auto-detect)
CPU_COMPUTE_TYPE = "int8"     # "int8" or "float32"
```

## Model Size Guide

CPU speeds below were measured on a Ryzen 9 9900X3D (12 threads, int8) against a
123-second clip. With `setup.py` installed, GPU transcription is far faster - see
[GPU Support](#gpu-support).

| Model  | Speed | Quality | Memory | Best For |
|--------|-------|---------|--------|----------|
| tiny   | Fastest | Good | ~1GB | Quick drafts, testing |
| base   | Fast | Better | ~1GB | General use |
| small  | ~20x realtime | Good | ~2GB | Fast turnaround |
| medium | Slower | Very Good | ~5GB | Middle ground |
| large-v3 | ~1.3x realtime | Best | ~3GB on disk | **Default** - maximum accuracy |
| large-v3-turbo | ~4x faster than large-v3 | Near-best | ~1.6GB | When large-v3 is too slow |

A one-hour recording takes roughly 45 minutes on `large-v3`, versus about three
minutes on `small`. Switch models via `WHISPER_MODEL` in `main.py`.

## GPU Support

Which engine you need depends on your GPU vendor:

| Your GPU | Engine | Notes |
|---|---|---|
| NVIDIA | `faster-whisper` | Set `FORCE_DEVICE = "cuda"` and `USE_HALF_PRECISION = True` |
| AMD / Intel | `whisper.cpp` | Vulkan backend - vendor neutral |
| None | `faster-whisper` | CPU with int8 |

**faster-whisper cannot use an AMD GPU.** It runs on CTranslate2, which is
CUDA-only; there is no ROCm backend, and the community ROCm forks target
gfx900-gfx1151, which excludes RDNA4 (gfx1201). Setting `FORCE_DEVICE = "cuda"`
on an AMD card silently falls back to CPU.

### whisper.cpp (Vulkan) setup

Measured on an AMD RX 9070 XT with `large-v3`: **~22x realtime** (a 15.8-minute
lecture transcribed in 43 seconds), versus ~1.3x realtime on CPU.

Run `python setup.py` and it handles all of this for you. To do it by hand:

1. Download a Windows Vulkan build of whisper.cpp and extract it to `whispercpp/`
   so that `whispercpp/whisper-cli.exe` exists. The official whisper.cpp releases
   ship CPU/BLAS/cuBLAS builds only - no Vulkan - so use a prebuilt Vulkan
   package such as
   [jerryshell/whisper.cpp-windows-vulkan-bin](https://github.com/jerryshell/whisper.cpp-windows-vulkan-bin),
   or build from source with `-DGGML_VULKAN=1`.
2. Download a GGML model into `models/`:
   ```bash
   curl -L -o models/ggml-large-v3.bin      https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-large-v3.bin
   ```
3. Nothing to configure - `TRANSCRIPTION_ENGINE = "auto"` detects the files and
   switches to the GPU engine on its own. To pin it explicitly:
   ```python
   TRANSCRIPTION_ENGINE = "whisper.cpp"
   WHISPERCPP_MODEL = "models/ggml-large-v3.bin"
   WHISPERCPP_GPU_DEVICE = 0   # Vulkan device index
   ```

On startup whisper.cpp lists the Vulkan devices it found. Confirm your discrete
GPU is the one at `WHISPERCPP_GPU_DEVICE`, since integrated graphics often
enumerate alongside it.

Note that GGML models are a different format from the CTranslate2 models used by
`faster-whisper` - switching engines means downloading the model again.

## Cleaning Transcripts

Whisper sometimes stutters or loops on a phrase. `clean_transcripts.py` removes
those artifacts from every `.txt` in `output/` (a `_backup.txt` copy is written
first):

```bash
python clean_transcripts.py
```

By default it is conservative: legitimate doubled words such as "had had",
"that that" and "New York, New York" are preserved, and only words sharing a
stem are treated as duplicates. To restore the older, more destructive
behaviour, set `AGGRESSIVE = True` at the top of the script.
