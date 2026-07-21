import os


def youtube_to_transcript(
    url: str,
    language: str = "en",
    output_format: str = "text",
    include_timestamps: bool = False,
    auto_captions: bool = True,
    translate_to: str = None,
) -> dict:
    """
    Transcribe a YouTube video.

    Strategy:
        1. Try YouTube Transcript API (fast, free, no audio download)
        2. Fallback → yt-dlp + Whisper (download audio + transcribe)

    Args:
        url               : YouTube video URL
        language          : transcript language code    (default: en)
        output_format     : 'text' | 'srt' | 'vtt' | 'json'
        include_timestamps: include timestamps in text  (default: False)
        auto_captions     : use auto-generated captions (default: True)
        translate_to      : translate to language code  (default: None)

    Returns:
        {
            'transcript'  : str,
            'segments'    : list,
            'language'    : str,
            'method'      : str,
            'title'       : str,
            'duration'    : int,
            'word_count'  : int,
            'char_count'  : int,
        }
    """
    import re

    # ── Validate URL ──────────────────────────────
    if not url or not url.strip():
        raise ValueError("URL is required.")

    url = url.strip()

    youtube_pattern = re.compile(
        r"(https?://)?(www\.)?"
        r"(youtube\.com/watch\?v=|youtu\.be/|youtube\.com/shorts/)"
        r"[\w-]+"
    )

    if not youtube_pattern.match(url):
        raise ValueError(
            "Invalid YouTube URL. "
            "Supported: youtube.com/watch?v=..., youtu.be/..., "
            "youtube.com/shorts/..."
        )

    # ── Validate output_format ─────────────────────
    valid_formats = ("text", "srt", "vtt", "json")
    if output_format not in valid_formats:
        raise ValueError(f"output_format must be one of: {valid_formats}")

    # ── Extract video ID ───────────────────────────
    video_id = _extract_youtube_id(url)
    if not video_id:
        raise ValueError("Cannot extract video ID from URL.")

    # ── Strategy 1: YouTube Transcript API ────────
    api_error_msg = None
    try:
        return _transcribe_via_api(
            video_id=video_id,
            url=url,
            language=language,
            output_format=output_format,
            include_timestamps=include_timestamps,
            auto_captions=auto_captions,
            translate_to=translate_to,
        )
    except Exception as api_error:
        # Save the message now — exception variables are auto-deleted
        # once this except block ends, so we can't reference `api_error`
        # later (that caused the original UnboundLocalError bug).
        api_error_msg = str(api_error)

    # ── Strategy 2: yt-dlp + Whisper ──────────────
    try:
        return _transcribe_via_whisper(
            url=url,
            video_id=video_id,
            language=language,
            output_format=output_format,
            include_timestamps=include_timestamps,
            translate_to=translate_to,
        )
    except Exception as whisper_error:
        raise RuntimeError(
            f"Both transcription methods failed.\n"
            f"API error: {api_error_msg}\n"
            f"Whisper error: {whisper_error}"
        )


def _extract_youtube_id(url: str) -> str:
    """Extract YouTube video ID from any URL format."""
    import re

    patterns = [
        r"(?:v=|/)([0-9A-Za-z_-]{11}).*",
        r"(?:youtu\.be/)([0-9A-Za-z_-]{11})",
        r"(?:shorts/)([0-9A-Za-z_-]{11})",
        r"(?:embed/)([0-9A-Za-z_-]{11})",
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)

    return ""


def _transcribe_via_api(
    video_id: str,
    url: str,
    language: str,
    output_format: str,
    include_timestamps: bool,
    auto_captions: bool,
    translate_to: str,
) -> dict:
    """
    Transcribe using YouTube Transcript API.
    No audio download needed — uses YouTube's existing captions.

    NOTE: youtube-transcript-api v1.0+ changed its API:
      - YouTubeTranscriptApi.list_transcripts(...)  ->  YouTubeTranscriptApi().list(...)
      - transcript.fetch() now returns a FetchedTranscript object,
        not a plain list of dicts, so we call .to_raw_data() to convert it
        into the list[dict] shape the rest of this code expects.
    """
    from youtube_transcript_api import YouTubeTranscriptApi

    # ── Get video metadata ────────────────────────
    title = _get_youtube_title(url)
    duration = 0

    # ── Fetch transcript ──────────────────────────
    ytt_api = YouTubeTranscriptApi()
    transcript_list = ytt_api.list(video_id)

    transcript = None

    # Try requested language first
    try:
        transcript = transcript_list.find_transcript([language])
    except Exception:
        pass

    # Try auto-generated captions
    if transcript is None and auto_captions:
        try:
            transcript = transcript_list.find_generated_transcript([language])
        except Exception:
            pass

    # Fallback to any available language
    if transcript is None:
        try:
            for t in transcript_list:
                transcript = t
                break
        except Exception:
            pass

    if transcript is None:
        raise RuntimeError("No transcript found for this video.")

    # ── Translate if requested ─────────────────────
    if translate_to and translate_to != transcript.language_code:
        try:
            transcript = transcript.translate(translate_to)
        except Exception:
            pass

    fetched = transcript.fetch()
    segments = fetched.to_raw_data()  # list[dict] with 'text', 'start', 'duration'
    detected_language = translate_to or transcript.language_code

    # ── Calculate duration ─────────────────────────
    if segments:
        last = segments[-1]
        duration = int(last.get("start", 0) + last.get("duration", 0))

    # ── Format output ──────────────────────────────
    formatted = _format_transcript(segments, output_format, include_timestamps)
    word_count = len(formatted.split())
    char_count = len(formatted)

    return {
        "transcript": formatted,
        "segments": segments,
        "language": detected_language,
        "method": "youtube-transcript-api",
        "title": title,
        "video_id": video_id,
        "url": url,
        "duration": duration,
        "duration_str": _seconds_to_time(duration),
        "word_count": word_count,
        "char_count": char_count,
        "output_format": output_format,
        "segment_count": len(segments),
    }


def _transcribe_via_whisper(
    url: str,
    video_id: str,
    language: str,
    output_format: str,
    include_timestamps: bool,
    translate_to: str,
) -> dict:
    """
    Download audio with yt-dlp then transcribe with Whisper.
    Slower but works when captions are not available.
    """
    import shutil
    import yt_dlp
    import whisper
    import tempfile

    # ── Check ffmpeg is installed ──────────────────
    # yt-dlp's audio postprocessing and Whisper both shell out to the
    # ffmpeg/ffprobe binaries. pip installing yt-dlp/whisper does NOT
    # install these — they must come from your OS package manager:
    #   Ubuntu/Debian : sudo apt-get install ffmpeg
    #   macOS         : brew install ffmpeg
    #   Windows       : download from ffmpeg.org and add 'bin' to PATH
    if shutil.which("ffmpeg") is None or shutil.which("ffprobe") is None:
        raise RuntimeError(
            "ffmpeg/ffprobe not found on this system. Install it first: "
            "Ubuntu/Debian: 'sudo apt-get install ffmpeg' | "
            "macOS: 'brew install ffmpeg' | "
            "Windows: download from ffmpeg.org and add to PATH."
        )

    title = _get_youtube_title(url)
    duration = 0

    with tempfile.TemporaryDirectory() as tmp:
        audio_path = os.path.join(tmp, f"{video_id}.mp3")

        # ── Download audio ─────────────────────────
        ydl_opts = {
            "format": "bestaudio/best",
            "outtmpl": os.path.join(tmp, f"{video_id}.%(ext)s"),
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
            info = ydl.extract_info(url, download=True)
            title = info.get("title", title)
            duration = info.get("duration", 0)

        # ── Find downloaded audio file ─────────────
        audio_file = None
        for f in os.listdir(tmp):
            if f.endswith(".mp3"):
                audio_file = os.path.join(tmp, f)
                break

        if not audio_file or not os.path.exists(audio_file):
            raise RuntimeError("Audio download failed.")

        # ── Load Whisper model ─────────────────────
        model = whisper.load_model("base")

        # ── Transcribe ─────────────────────────────
        whisper_opts = {
            "verbose": False,
            "task": "translate" if translate_to == "en" else "transcribe",
        }
        if language and language != "auto":
            whisper_opts["language"] = language

        result = model.transcribe(audio_file, **whisper_opts)

    # ── Build segments ────────────────────────────
    segments = [
        {
            "text": seg["text"].strip(),
            "start": seg["start"],
            "duration": seg["end"] - seg["start"],
        }
        for seg in result.get("segments", [])
    ]

    detected_language = result.get("language", language)

    # ── Format output ──────────────────────────────
    formatted = _format_transcript(segments, output_format, include_timestamps)
    word_count = len(formatted.split())
    char_count = len(formatted)

    return {
        "transcript": formatted,
        "segments": segments,
        "language": detected_language,
        "method": "whisper",
        "title": title,
        "video_id": video_id,
        "url": url,
        "duration": duration,
        "duration_str": _seconds_to_time(duration),
        "word_count": word_count,
        "char_count": char_count,
        "output_format": output_format,
        "segment_count": len(segments),
    }


def _get_youtube_title(url: str) -> str:
    """Get YouTube video title without downloading."""
    try:
        import yt_dlp

        ydl_opts = {
            "quiet": True,
            "no_warnings": True,
            "skip_download": True,
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            return info.get("title", "Unknown Title")
    except Exception:
        return "Unknown Title"


def _format_transcript(
    segments: list,
    output_format: str,
    include_timestamps: bool,
) -> str:
    """Format transcript segments into desired output format."""
    import json as _json

    if output_format == "json":
        return _json.dumps(segments, ensure_ascii=False, indent=2)

    if output_format == "srt":
        return _to_srt(segments)

    if output_format == "vtt":
        return _to_vtt(segments)

    # ── Plain text ────────────────────────────────
    parts = []
    for seg in segments:
        text = seg.get("text", "").strip()
        if not text:
            continue
        if include_timestamps:
            start = _seconds_to_time(int(seg.get("start", 0)))
            parts.append(f"[{start}] {text}")
        else:
            parts.append(text)

    return " ".join(parts) if not include_timestamps else "\n".join(parts)


def _to_srt(segments: list) -> str:
    """Convert segments to SRT subtitle format."""
    lines = []
    for i, seg in enumerate(segments, start=1):
        start = seg.get("start", 0)
        duration = seg.get("duration", 2)
        end = start + duration
        text = seg.get("text", "").strip()

        lines.append(str(i))
        lines.append(
            f"{_seconds_to_srt_time(start)} --> " f"{_seconds_to_srt_time(end)}"
        )
        lines.append(text)
        lines.append("")

    return "\n".join(lines)


def _to_vtt(segments: list) -> str:
    """Convert segments to WebVTT subtitle format."""
    lines = ["WEBVTT", ""]
    for seg in segments:
        start = seg.get("start", 0)
        duration = seg.get("duration", 2)
        end = start + duration
        text = seg.get("text", "").strip()

        lines.append(
            f"{_seconds_to_vtt_time(start)} --> " f"{_seconds_to_vtt_time(end)}"
        )
        lines.append(text)
        lines.append("")

    return "\n".join(lines)


def _seconds_to_time(seconds: int) -> str:
    """Convert seconds to HH:MM:SS string."""
    seconds = int(seconds)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def _seconds_to_srt_time(seconds: float) -> str:
    """Convert seconds to SRT time format HH:MM:SS,mmm."""
    ms = int((seconds % 1) * 1000)
    s = int(seconds)
    h = s // 3600
    m = (s % 3600) // 60
    s = s % 60
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def _seconds_to_vtt_time(seconds: float) -> str:
    """Convert seconds to WebVTT time format HH:MM:SS.mmm."""
    ms = int((seconds % 1) * 1000)
    s = int(seconds)
    h = s // 3600
    m = (s % 3600) // 60
    s = s % 60
    return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"


def download_youtube_video(
    url: str,
    quality: str = "best",
    format_type: str = "mp4",
    max_height: int = None,
) -> dict:
    """
    Download a YouTube video/short.

    Args:
        url         : YouTube video/shorts URL
        quality     : 'best' | 'worst' | specific height e.g. '1080' | '720'
        format_type : mp4 | webm | mkv          (default: mp4)
        max_height  : cap resolution e.g. 1080, 720, 480 (default: None = highest available)

    Returns:
        {
            'bytes'      : bytes,
            'title'      : str,
            'duration'   : int,
            'width'      : int,
            'height'     : int,
            'fps'        : float,
            'filesize_mb': float,
            'video_id'   : str,
            'ext'        : str,
        }
    """
    import yt_dlp
    import tempfile
    import re

    if not url or not url.strip():
        raise ValueError("URL is required.")

    url = url.strip()

    youtube_pattern = re.compile(
        r"(https?://)?(www\.)?"
        r"(youtube\.com/watch\?v=|youtu\.be/|youtube\.com/shorts/)"
        r"[\w-]+"
    )
    if not youtube_pattern.match(url):
        raise ValueError(
            "Invalid YouTube URL. Supported: youtube.com/watch?v=..., "
            "youtu.be/..., youtube.com/shorts/..."
        )

    if format_type not in ("mp4", "webm", "mkv"):
        format_type = "mp4"

    # ── Build format selector ─────────────────────
    if max_height:
        height_filter = f"[height<={max_height}]"
    else:
        height_filter = ""

    if quality == "worst":
        fmt = f"worst{height_filter}"
    else:
        fmt = f"bestvideo{height_filter}+bestaudio/" f"best{height_filter}"

    with tempfile.TemporaryDirectory() as tmp:
        outtmpl = os.path.join(tmp, "%(id)s.%(ext)s")

        ydl_opts = {
            "format": fmt,
            "outtmpl": outtmpl,
            "quiet": True,
            "no_warnings": True,
            "merge_output_format": format_type,
            "postprocessors": [
                {
                    "key": "FFmpegVideoConvertor",
                    "preferedformat": format_type,
                }
            ],
        }

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)

        video_id = info.get("id", "")
        title = info.get("title", "video")
        duration = info.get("duration", 0)
        width = info.get("width", 0)
        height = info.get("height", 0)
        fps = info.get("fps", 0)

        # ── Find downloaded file ────────────────────
        output_path = None
        for f in os.listdir(tmp):
            if f.startswith(video_id):
                output_path = os.path.join(tmp, f)
                break

        if not output_path or not os.path.exists(output_path):
            raise RuntimeError("Download failed — output file not found.")

        with open(output_path, "rb") as f:
            video_bytes = f.read()

    return {
        "bytes": video_bytes,
        "title": title,
        "duration": duration,
        "width": width,
        "height": height,
        "fps": fps,
        "filesize_mb": round(len(video_bytes) / (1024 * 1024), 2),
        "video_id": video_id,
        "ext": format_type,
    }


def get_youtube_video_info(url: str) -> dict:
    """
    Get available formats/qualities for a YouTube video without downloading.

    Args:
        url : YouTube video/shorts URL

    Returns:
        { 'title', 'duration', 'formats': [ {height, ext, filesize, format_id} ] }
    """
    import yt_dlp
    import re

    if not url or not url.strip():
        raise ValueError("URL is required.")

    youtube_pattern = re.compile(
        r"(https?://)?(www\.)?"
        r"(youtube\.com/watch\?v=|youtu\.be/|youtube\.com/shorts/)"
        r"[\w-]+"
    )
    if not youtube_pattern.match(url.strip()):
        raise ValueError("Invalid YouTube URL.")

    ydl_opts = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)

    formats = []
    seen_heights = set()

    for f in info.get("formats", []):
        height = f.get("height")
        if not height or height in seen_heights:
            continue
        if f.get("vcodec") == "none":
            continue
        seen_heights.add(height)
        formats.append(
            {
                "height": height,
                "width": f.get("width"),
                "ext": f.get("ext"),
                "fps": f.get("fps"),
                "filesize_mb": round(
                    (f.get("filesize") or f.get("filesize_approx") or 0)
                    / (1024 * 1024),
                    2,
                )
                or None,
                "format_id": f.get("format_id"),
            }
        )

    formats.sort(key=lambda x: x["height"] or 0, reverse=True)

    return {
        "title": info.get("title", ""),
        "duration": info.get("duration", 0),
        "video_id": info.get("id", ""),
        "is_short": (info.get("duration") or 0) <= 60,
        "formats": formats,
    }
