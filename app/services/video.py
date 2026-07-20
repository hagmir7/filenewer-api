# ── Standard Library ──────────────────────────────────────────────────────────
import io
import os
import re
import json
import tempfile
import subprocess

# ── Third Party ───────────────────────────────────────────────────────────────
try:
    import yt_dlp

    YTDLP_AVAILABLE = True
except ImportError:
    YTDLP_AVAILABLE = False


def add_voice_to_video(
    video_source,
    audio_source,
    speed_adjust_target: str = None,
    volume_video: float = 1.0,
    volume_audio: float = 1.0,
    replace_audio: bool = True,
    output_format: str = "mp4",
    video_filename: str = "video.mp4",
    audio_filename: str = "audio.mp3",
) -> dict:
    """
    Add/replace audio/voice to a video file.

    If audio and video durations don't match, user must
    specify speed_adjust_target to resolve the mismatch:
        'video' → stretch/shrink video audio to match voice
        'audio' → stretch/shrink voice to match video

    Args:
        video_source         : file object | bytes | URL string
        audio_source         : file object | bytes | URL string
        speed_adjust_target  : 'video' | 'audio' | None
                               Required when durations mismatch
        volume_video         : video original audio volume 0.0-2.0
                               (only if replace_audio=False)
        volume_audio         : new audio/voice volume 0.0-2.0
        replace_audio        : True  → replace video audio entirely
                               False → mix voice with original audio
        output_format        : mp4 | avi | mkv | webm   (default: mp4)
        video_filename       : original video filename
        audio_filename       : original audio filename

    Returns:
        {
            'bytes'              : bytes,
            'video_duration'     : float,
            'audio_duration'     : float,
            'output_duration'    : float,
            'duration_match'     : bool,
            'speed_adjusted'     : bool,
            'speed_ratio'        : float,
            'method'             : str,
            'size_kb'            : float,
            'size_mb'            : float,
            'output_format'      : str,
        }

    Raises:
        DurationMismatchError : when durations differ and
                                speed_adjust_target is not specified.
    """
    import tempfile
    import shutil
    import subprocess
    import math

    TOLERANCE = 0.5  # seconds — acceptable duration difference

    # ── Read sources ──────────────────────────────
    video_bytes = _read_media_source(video_source, "video")
    audio_bytes = _read_media_source(audio_source, "audio")

    if not video_bytes:
        raise ValueError("Video source is empty or could not be read.")
    if not audio_bytes:
        raise ValueError("Audio source is empty or could not be read.")

    # ── Validate options ──────────────────────────
    if output_format not in ("mp4", "avi", "mkv", "webm"):
        output_format = "mp4"

    if not (0.0 <= volume_video <= 2.0):
        raise ValueError("volume_video must be between 0.0 and 2.0.")
    if not (0.0 <= volume_audio <= 2.0):
        raise ValueError("volume_audio must be between 0.0 and 2.0.")

    if speed_adjust_target and speed_adjust_target not in ("video", "audio"):
        raise ValueError("speed_adjust_target must be: video or audio.")

    # ── Check ffmpeg ───────────────────────────────
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        raise RuntimeError(
            "ffmpeg is required but not found. "
            "Install ffmpeg: https://ffmpeg.org/download.html"
        )

    with tempfile.TemporaryDirectory() as tmp:

        # ── Write to temp files ────────────────────
        video_ext = _get_extension(video_filename, "mp4")
        audio_ext = _get_extension(audio_filename, "mp3")

        video_path = os.path.join(tmp, f"input_video.{video_ext}")
        audio_path = os.path.join(tmp, f"input_audio.{audio_ext}")
        output_path = os.path.join(tmp, f"output.{output_format}")

        with open(video_path, "wb") as f:
            f.write(video_bytes)
        with open(audio_path, "wb") as f:
            f.write(audio_bytes)

        # ── Get durations ──────────────────────────
        video_duration = _get_media_duration(video_path, ffmpeg)
        audio_duration = _get_media_duration(audio_path, ffmpeg)

        if video_duration <= 0:
            raise ValueError("Cannot determine video duration.")
        if audio_duration <= 0:
            raise ValueError("Cannot determine audio duration.")

        duration_diff = abs(video_duration - audio_duration)
        duration_match = duration_diff <= TOLERANCE
        speed_adjusted = False
        speed_ratio = 1.0

        # ── Handle duration mismatch ───────────────
        if not duration_match:
            if speed_adjust_target is None:
                # ── Return mismatch info — no processing ──
                return {
                    "bytes": None,
                    "video_duration": round(video_duration, 2),
                    "audio_duration": round(audio_duration, 2),
                    "output_duration": 0,
                    "duration_match": False,
                    "speed_adjusted": False,
                    "speed_ratio": 1.0,
                    "method": "none",
                    "size_kb": 0,
                    "size_mb": 0,
                    "output_format": output_format,
                    "mismatch": True,
                    "mismatch_info": {
                        "video_duration": round(video_duration, 2),
                        "audio_duration": round(audio_duration, 2),
                        "difference_sec": round(duration_diff, 2),
                        "video_duration_str": _seconds_to_time(video_duration),
                        "audio_duration_str": _seconds_to_time(audio_duration),
                        "suggestion": (
                            f"Durations differ by {round(duration_diff, 2)}s. "
                            f'Set speed_adjust_target to "video" to adjust '
                            f'video speed, or "audio" to adjust audio speed.'
                        ),
                    },
                }

            # ── Adjust speed ───────────────────────
            if speed_adjust_target == "audio":
                # Stretch/shrink audio to match video
                speed_ratio = audio_duration / video_duration
                adjusted_path = os.path.join(tmp, f"adjusted_audio.{audio_ext}")
                _adjust_audio_speed(audio_path, adjusted_path, speed_ratio, ffmpeg)
                audio_path = adjusted_path
                speed_adjusted = True

            elif speed_adjust_target == "video":
                # Stretch/shrink video to match audio
                speed_ratio = video_duration / audio_duration
                adjusted_path = os.path.join(tmp, f"adjusted_video.{video_ext}")
                _adjust_video_speed(video_path, adjusted_path, speed_ratio, ffmpeg)
                video_path = adjusted_path
                speed_adjusted = True

        # ── Build ffmpeg command ───────────────────
        if replace_audio:
            cmd = _build_replace_audio_cmd(
                video_path,
                audio_path,
                output_path,
                volume_audio,
                ffmpeg,
                output_format,
            )
        else:
            cmd = _build_mix_audio_cmd(
                video_path,
                audio_path,
                output_path,
                volume_video,
                volume_audio,
                ffmpeg,
                output_format,
            )

        # ── Run ffmpeg ─────────────────────────────
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
        )

        if not os.path.exists(output_path):
            raise RuntimeError(f"ffmpeg failed: {result.stderr[-1000:]}")

        # ── Read output ────────────────────────────
        with open(output_path, "rb") as f:
            output_bytes = f.read()

    output_duration = _get_duration_from_bytes(output_bytes, output_format)

    return {
        "bytes": output_bytes,
        "video_duration": round(video_duration, 2),
        "audio_duration": round(audio_duration, 2),
        "output_duration": round(output_duration, 2),
        "duration_match": duration_match,
        "speed_adjusted": speed_adjusted,
        "speed_ratio": round(speed_ratio, 4),
        "method": "replace" if replace_audio else "mix",
        "size_kb": round(len(output_bytes) / 1024, 2),
        "size_mb": round(len(output_bytes) / (1024 * 1024), 2),
        "output_format": output_format,
        "mismatch": False,
        "mismatch_info": None,
    }


def _read_media_source(source, media_type: str) -> bytes:
    """
    Read media from file object, bytes, or URL string.
    Supports: local file, raw bytes, HTTP URL, YouTube URL.
    """
    import urllib.request
    import re

    if source is None:
        raise ValueError(f"{media_type} source is required.")

    # ── File object ────────────────────────────────
    if hasattr(source, "read"):
        return source.read()

    # ── Raw bytes ──────────────────────────────────
    if isinstance(source, bytes):
        return source

    # ── String: URL or path ────────────────────────
    if isinstance(source, str):
        source = source.strip()

        # ── YouTube URL ────────────────────────────
        youtube_pattern = re.compile(
            r"(https?://)?(www\.)?"
            r"(youtube\.com/watch\?v=|youtu\.be/|youtube\.com/shorts/)"
            r"[\w-]+"
        )
        if youtube_pattern.match(source):
            return _download_youtube_audio(source)

        # ── HTTP/HTTPS URL ─────────────────────────
        if source.startswith(("http://", "https://")):
            return _download_url(source)

        # ── Local file path ────────────────────────
        if os.path.exists(source):
            with open(source, "rb") as f:
                return f.read()

        raise ValueError(
            f"Cannot read {media_type} source: " f"not a file, URL, or valid path."
        )

    raise ValueError(f"Invalid {media_type} source type: {type(source)}")


def _download_url(url: str) -> bytes:
    """Download file from HTTP/HTTPS URL."""
    import urllib.request

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as response:
            return response.read()
    except Exception as e:
        raise ValueError(f"Failed to download URL: {url}\n{e}")


def _download_youtube_audio(url: str) -> bytes:
    """Download YouTube audio using yt-dlp."""
    import yt_dlp
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        audio_path = os.path.join(tmp, "audio.mp3")
        ydl_opts = {
            "format": "bestaudio/best",
            "outtmpl": os.path.join(tmp, "audio.%(ext)s"),
            "quiet": True,
            "no_warnings": True,
            "postprocessors": [
                {
                    "key": "FFmpegExtractAudio",
                    "preferredcodec": "mp3",
                    "preferredquality": "128",
                }
            ],
        }

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])

        for f in os.listdir(tmp):
            if f.endswith(".mp3"):
                audio_path = os.path.join(tmp, f)
                break

        if not os.path.exists(audio_path):
            raise ValueError("Failed to download YouTube audio.")

        with open(audio_path, "rb") as f:
            return f.read()


def _get_media_duration(path: str, ffmpeg: str) -> float:
    """Get media duration in seconds using ffprobe."""
    import subprocess
    import shutil
    import json as _json

    ffprobe = shutil.which("ffprobe") or ffmpeg.replace("ffmpeg", "ffprobe")

    result = subprocess.run(
        [
            ffprobe,
            "-v",
            "quiet",
            "-print_format",
            "json",
            "-show_streams",
            path,
        ],
        capture_output=True,
        text=True,
        timeout=30,
    )

    try:
        info = _json.loads(result.stdout)
        for stream in info.get("streams", []):
            duration = float(stream.get("duration", 0))
            if duration > 0:
                return duration
    except Exception:
        pass

    # Fallback: parse stderr
    import re

    duration_pattern = re.search(
        r"Duration: (\d+):(\d+):(\d+\.\d+)",
        result.stderr + result.stdout,
    )
    if duration_pattern:
        h, m, s = duration_pattern.groups()
        return int(h) * 3600 + int(m) * 60 + float(s)

    return 0.0


def _get_duration_from_bytes(file_bytes: bytes, ext: str) -> float:
    """Get duration from bytes by writing to temp file."""
    import tempfile
    import shutil

    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        return 0.0

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, f"output.{ext}")
        with open(path, "wb") as f:
            f.write(file_bytes)
        return _get_media_duration(path, ffmpeg)


def _adjust_audio_speed(
    input_path: str,
    output_path: str,
    speed_ratio: float,
    ffmpeg: str,
):
    """
    Adjust audio speed using ffmpeg atempo filter.
    atempo supports 0.5 to 2.0 — chain for outside range.
    """
    import subprocess
    import math

    # Build atempo filter chain for extreme ratios
    # atempo range: 0.5 to 2.0
    filters = []
    ratio = speed_ratio

    if ratio > 2.0:
        # Chain multiple atempo filters
        while ratio > 2.0:
            filters.append("atempo=2.0")
            ratio /= 2.0
        if ratio > 1.0:
            filters.append(f"atempo={ratio:.4f}")
    elif ratio < 0.5:
        while ratio < 0.5:
            filters.append("atempo=0.5")
            ratio *= 2.0
        if ratio < 1.0:
            filters.append(f"atempo={ratio:.4f}")
    else:
        filters.append(f"atempo={ratio:.4f}")

    filter_str = ",".join(filters)

    cmd = [
        ffmpeg,
        "-y",
        "-i",
        input_path,
        "-filter:a",
        filter_str,
        output_path,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    if not os.path.exists(output_path):
        raise RuntimeError(f"Audio speed adjustment failed: {result.stderr[-500:]}")


def _adjust_video_speed(
    input_path: str,
    output_path: str,
    speed_ratio: float,
    ffmpeg: str,
):
    import subprocess

    pts_factor = 1.0 / speed_ratio

    audio_filters = []
    ratio = speed_ratio
    if ratio > 2.0:
        while ratio > 2.0:
            audio_filters.append("atempo=2.0")
            ratio /= 2.0
        if ratio > 1.0:
            audio_filters.append(f"atempo={ratio:.4f}")
    elif ratio < 0.5:
        while ratio < 0.5:
            audio_filters.append("atempo=0.5")
            ratio *= 2.0
        if ratio < 1.0:
            audio_filters.append(f"atempo={ratio:.4f}")
    else:
        audio_filters.append(f"atempo={ratio:.4f}")

    audio_filter_str = ",".join(audio_filters)

    cmd = [
        ffmpeg,
        "-y",
        "-i",
        input_path,
        "-filter_complex",
        f"[0:v]setpts={pts_factor:.4f}*PTS[v];[0:a]{audio_filter_str}[a]",
        "-map",
        "[v]",
        "-map",
        "[a]",
        output_path,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    # Check BOTH exit code and that a non-trivial file was actually written
    if (
        result.returncode != 0
        or not os.path.exists(output_path)
        or os.path.getsize(output_path) < 1024
    ):
        raise RuntimeError(f"Video speed adjustment failed: {result.stderr[-1000:]}")
    
    


def _build_replace_audio_cmd(
    video_path: str,
    audio_path: str,
    output_path: str,
    volume_audio: float,
    ffmpeg: str,
    output_format: str,
) -> list:
    """Build ffmpeg command to replace video audio entirely."""
    cmd = [
        ffmpeg,
        "-y",
        "-i",
        video_path,
        "-i",
        audio_path,
        "-map",
        "0:v:0",  # video from input 0
        "-map",
        "1:a:0",  # audio from input 1
        "-c:v",
        "copy",  # copy video stream
    ]

    # Volume filter for new audio
    if volume_audio != 1.0:
        cmd += [
            "-filter:a",
            f"volume={volume_audio:.2f}",
        ]
    else:
        cmd += ["-c:a", "aac"]

    # Shortest flag
    cmd += ["-shortest"]

    # Output format options
    if output_format == "mp4":
        cmd += ["-movflags", "+faststart"]

    cmd.append(output_path)
    return cmd


def _build_mix_audio_cmd(
    video_path: str,
    audio_path: str,
    output_path: str,
    volume_video: float,
    volume_audio: float,
    ffmpeg: str,
    output_format: str,
) -> list:
    """Build ffmpeg command to mix voice with existing video audio."""
    cmd = [
        ffmpeg,
        "-y",
        "-i",
        video_path,
        "-i",
        audio_path,
        "-filter_complex",
        (
            f"[0:a]volume={volume_video:.2f}[original];"
            f"[1:a]volume={volume_audio:.2f}[voice];"
            f"[original][voice]amix=inputs=2:duration=shortest[mixed]"
        ),
        "-map",
        "0:v:0",
        "-map",
        "[mixed]",
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-shortest",
    ]

    if output_format == "mp4":
        cmd += ["-movflags", "+faststart"]

    cmd.append(output_path)
    return cmd


def _get_extension(filename: str, default: str = "mp4") -> str:
    """Get file extension from filename."""
    if "." in filename:
        return filename.rsplit(".", 1)[-1].lower()
    return default


def _seconds_to_time(seconds: float) -> str:
    """Convert seconds to a human-readable HH:MM:SS.ms string."""
    seconds = max(0, float(seconds))
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    if hours > 0:
        return f"{hours:02d}:{minutes:02d}:{secs:05.2f}"
    return f"{minutes:02d}:{secs:05.2f}"
