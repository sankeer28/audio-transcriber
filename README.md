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

## Requirements

- Python 3.8 or higher
- ffmpeg (for MP4 video processing)
- CUDA-compatible GPU (optional)
- PyTorch + openai-whisper (optional, only for `TRANSCRIPTION_ENGINE = "standard"`)

## Installation

### 1. Clone or Download the Project

### 2. Install Python Dependencies

```bash
python -m venv .venv
.venv\Scripts\activate      # Windows
source .venv/bin/activate    # macOS / Linux

pip install -r requirements.txt
```

The default `faster-whisper` engine does **not** need PyTorch. Only install the
optional extras if you want `TRANSCRIPTION_ENGINE = "standard"` (openai-whisper),
which pulls in several GB of dependencies:

```bash
pip install -r requirements-standard.txt
```

### 3. Install ffmpeg

#### Windows:
```bash
# Using chocolatey:
choco install ffmpeg

# Or download from: https://ffmpeg.org/download.html
```

#### macOS:
```bash
brew install ffmpeg
```

#### Linux:
```bash
sudo apt install ffmpeg
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
TRANSCRIPTION_ENGINE = "faster-whisper"  # Options: "standard", "faster-whisper"
```

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
CPU_THREADS = 12              # CPU worker threads (0 = library default)
CPU_COMPUTE_TYPE = "int8"     # "int8" or "float32"
```

## Model Size Guide

Measured on this machine (Ryzen 9 9900X3D, 12 threads, int8, CPU-only) against a
123-second clip:

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

### Whisper Model Settings
```python
WHISPER_MODEL = "large-v3"   # Most accurate
CPU_THREADS = 12             # Match your core count
CPU_COMPUTE_TYPE = "int8"    # "float32" for reference quality (much slower)
CPU_BEAM_SIZE = 5            # Whisper's reference beam size
```

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
3. Set the engine in `main.py`:
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
