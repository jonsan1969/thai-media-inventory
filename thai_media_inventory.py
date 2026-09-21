#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Thai Media Inventory / Duplicate Scanner

Scans one or more media roots recursively, parses the archive filename standard,
extracts technical metadata with ffprobe, compares expected-vs-actual metadata,
uses an incremental SQLite cache, detects exact duplicates and likely alternate
versions of the same film, builds local Film Records / Title Aliases, and can optionally
match movie records against TMDb. Rename proposals are report-only; TMDb never rewrites
local titles or technical metadata.

Designed for Windows but works on macOS/Linux too.

Requirements:
    python -m pip install --user xlsxwriter

External:
    ffprobe / FFmpeg
    - Put ffprobe.exe next to this script, OR
    - have ffprobe available on PATH.

Typical usage:
    python thai_media_inventory.py D:\\ThaiMovies E:\\Archive\\Thai

Or place thai_media_inventory.ini next to the script and simply run:
    python thai_media_inventory.py
"""

from __future__ import annotations

import argparse
import configparser
import difflib
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable, Optional

VIDEO_EXTS = {
    ".mkv", ".mp4", ".m4v", ".mov", ".avi", ".wmv", ".ts", ".m2ts", ".mts",
    ".mpg", ".mpeg", ".vob", ".webm", ".flv", ".iso"
}

QUICK_HASH_CHUNK = 4 * 1024 * 1024  # 4 MiB from start + end
CACHE_SCHEMA_VERSION = 9

# Filename noise. Conservative on purpose. This remains useful as a fallback
# for non-standard filenames, but the primary parser below is structural.
NOISE_PATTERNS = [
    r"\b(?:2160p|1080p|1080i|720p|576p|480p|4k|uhd|hdr10\+?|hdr|dv|dolby[ ._-]?vision)\b",
    r"\b(?:web[ ._-]?dl|web[ ._-]?rip|webrip|bluray|blu[ ._-]?ray|bdrip|brrip|remux|hdtv|dvd(?:rip)?|vhs(?:rip)?)\b",
    r"\b(?:x264|x265|h\.?264|h\.?265|hevc|avc|av1|mpeg[ ._-]?2|vc[ ._-]?1)\b",
    r"\b(?:aac(?:2\.0|5\.1)?|ac3|eac3|ddp?|dts(?:hd)?|truehd|atmos|flac|opus|mp3)\b",
    r"\b(?:proper|repack|internal|limited|uncut|extended|remastered|restored|criterion)\b",
    r"\b(?:thai|th|eng|english|subbed|subs?|softsubs?|hardsubs?)\b",
]

YEAR_RE = re.compile(r"(?<!\d)(18\d{2}|19\d{2}|20\d{2}|24\d{2}|25\d{2})(?!\d)")
BRACKET_RE = re.compile(r"[\[\(\{][^\]\)\}]{1,45}[\]\)\}]")
MULTISPACE_RE = re.compile(r"\s+")
SEPARATOR_RE = re.compile(r"[._\-]+")
EPISODE_RE = re.compile(r"\b(?:s\d{1,2}e\d{1,3}|ep(?:isode)?[ ._-]?\d{1,4})\b", re.I)

# Structural filename parser.
#
# v2.2 deliberately treats the technical suffix separately from the film/year
# identity. A media file can therefore still be checked against ffprobe even
# when a year is missing (TV episodes, unknown-year archive files, placeholders
# such as 20XX). The prefix is greedy so the LAST recognized resolution token
# becomes the technical boundary.
FILENAME_TECH_BOUNDARY_RE = re.compile(
    r"^(?P<prefix>.+)\."
    r"(?P<resolution>2160p|1080p|1080i|720p|576p|480p|360p|4K|UHD)\."
    r"(?P<rest>.+)$",
    re.I,
)

PREFIX_YEAR_RE = re.compile(
    r"^(?P<title>.+)\.(?P<year>(?:18|19|20|24|25)\d{2})$",
    re.I,
)

PREFIX_PLACEHOLDER_YEAR_RE = re.compile(
    r"^(?P<title>.+)\.(?P<year>(?:18|19|20|24|25)(?:XX|xx))$",
    re.I,
)

RELEASE_TYPE_PATTERN = (
    r"WEB-DL|WEB\.DL|WEBRip|WEB-Rip|BluRay|Blu-Ray|BDRip|BRRip|REMUX|HDTV|"
    r"DVDRip|DVD|VHSRip|VHS-Rip"
)
VIDEO_TOKEN_PATTERN = r"H\.?264|H\.?265|x264|x265|AVC|HEVC|AV1"
CHANNEL_TOKEN_PATTERN = r"7\.1|6\.1|5\.1|4\.0|3\.0|2\.1|2\.0|1\.0|mono|stereo"
AUDIO_TOKEN_PATTERN = (
    r"DDP|DD\+|DD|EAC3|E-AC-3|AC3|AC-3|AAC|DTS-HD\.MA|DTS-HD|DTS|TRUEHD|"
    r"FLAC|OPUS|MP3|LPCM|PCM"
)
AUDIO_FEATURE_PATTERN = r"Atmos"
VARIANT_TOKEN_PATTERN = r"TH|v\d+"


@dataclass
class FilenameMetadata:
    title: str = ""
    year: str = ""
    resolution: str = ""
    source: str = ""
    release_type: str = ""
    variant: str = ""
    audio_codec: str = ""
    channels: str = ""
    audio_features: str = ""
    video_codec: str = ""
    standard: str = "UNPARSED"
    notes: str = ""


@dataclass
class MediaRecord:
    path: str
    root: str
    filename: str
    extension: str
    size_bytes: int
    size_gib: float
    modified: str
    duration_sec: Optional[float]
    duration_text: str
    container: str
    overall_bitrate_kbps: Optional[int]
    computed_bitrate_kbps: Optional[int]
    video_codec: str
    video_profile: str
    video_encoder: str
    width: Optional[int]
    height: Optional[int]
    resolution: str
    actual_resolution_class: str
    fps: Optional[float]
    pixel_format: str
    bit_depth: str
    hdr: str
    video_bitrate_kbps: Optional[int]
    audio_tracks: int
    audio_summary: str
    primary_audio_codec: str
    primary_audio_channels: str
    actual_audio_features: str
    max_audio_channels: Optional[int]
    has_51: bool
    subtitle_tracks: int
    subtitle_summary: str
    has_english_sub: bool
    title_tag: str

    # Parsed / expected filename metadata.
    filename_title: str
    normalized_title: str
    filename_year: str
    name_resolution: str
    streaming_source: str
    release_type: str
    filename_variant: str
    name_audio: str
    name_channels: str
    name_audio_features: str
    name_video: str
    filename_standard: str
    filename_parse_notes: str

    # Expected-vs-actual checks.
    resolution_match: str
    audio_match: str
    channels_match: str
    audio_features_match: str
    video_match: str
    metadata_check: str

    # Report-only rename proposal. No rename action exists in this version.
    rename_candidate: str
    rename_status: str
    rename_target: str

    quick_hash: str
    sha256: str
    ffprobe_ok: bool
    error: str


def human_duration(seconds: Optional[float]) -> str:
    if seconds is None:
        return ""
    s = int(round(seconds))
    h, rem = divmod(s, 3600)
    m, sec = divmod(rem, 60)
    return f"{h:d}:{m:02d}:{sec:02d}"


def to_float(v) -> Optional[float]:
    try:
        if v in (None, "", "N/A"):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def to_int(v) -> Optional[int]:
    try:
        if v in (None, "", "N/A"):
            return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


def parse_rate(rate: str) -> Optional[float]:
    if not rate or rate == "0/0":
        return None
    try:
        if "/" in rate:
            a, b = rate.split("/", 1)
            b = float(b)
            return float(a) / b if b else None
        return float(rate)
    except Exception:
        return None


def script_dir() -> Path:
    return Path(__file__).resolve().parent


def find_ffprobe(explicit: Optional[str] = None) -> str:
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    candidates.extend([
        script_dir() / "ffprobe.exe",
        script_dir() / "ffprobe",
    ])
    on_path = shutil.which("ffprobe")
    if on_path:
        candidates.append(Path(on_path))

    for p in candidates:
        if p and Path(p).exists():
            return str(Path(p))
    raise FileNotFoundError(
        "ffprobe hittades inte. Lägg ffprobe.exe i samma katalog som scriptet "
        "eller lägg FFmpeg på PATH."
    )


def _pretty_title(raw_title: str) -> str:
    s = unicodedata.normalize("NFKC", raw_title or "")
    s = re.sub(r"[._]+", " ", s)
    return MULTISPACE_RE.sub(" ", s).strip()


def _take_tail_token(text: str, pattern: str) -> tuple[str, str]:
    m = re.search(rf"(?:^|\.)(?P<token>{pattern})$", text, flags=re.I)
    if not m:
        return text, ""
    token = m.group("token")
    prefix = text[:m.start()].rstrip(".")
    return prefix, token


def _take_video_tail(text: str) -> tuple[str, str, str]:
    """Return (prefix, video token, trailing annotation).

    Normal filenames end directly in H.264/H.265/x264/etc.  A small number of
    archive files carry an intentional marker directly after the codec, e.g.
    ``H.264-kass``.  We still want to verify the codec and the preceding audio
    fields, but keep the filename in REVIEW because of the extra annotation.
    """
    text = (text or "").strip(".")
    if not text:
        return "", "", ""

    # Normal case first.
    prefix, video = _take_tail_token(text, VIDEO_TOKEN_PATTERN)
    if video:
        return prefix, video, ""

    # Allow a small trailing annotation after the codec. This covers archive
    # markers such as H.264-kass, Windows duplicate suffixes such as H264 (1),
    # and accidental dot-delimited leftovers such as H.264.1999.  The latter
    # must stay visible as a filename REVIEW, but should not destroy otherwise
    # valid audio/channel/video parsing.
    m = re.search(
        rf"(?:^|\.)(?P<video>{VIDEO_TOKEN_PATTERN})"
        rf"(?P<annotation>(?:[-_][^.]+| \(\d+\)|\.[^.]+))$",
        text,
        flags=re.I,
    )
    if not m:
        return text, "", ""
    return text[:m.start()].rstrip("."), m.group("video"), m.group("annotation")


def _canonical_variant_token(token: str) -> str:
    token = (token or "").strip()
    if token.upper() == "TH":
        return "TH"
    if re.fullmatch(r"v\d+", token, flags=re.I):
        return token.lower()
    return token


def _split_source_variant(source: str) -> tuple[str, str]:
    """Split trailing archive variant tokens from a source field.

    Example: NF.TH -> (NF, TH).  These markers distinguish alternate
    provider versions; they are not language/codec errors.
    """
    parts = [p for p in (source or "").split(".") if p]
    variants = []
    while parts and re.fullmatch(VARIANT_TOKEN_PATTERN, parts[-1], flags=re.I):
        variants.append(_canonical_variant_token(parts.pop()))
    variants.reverse()
    return ".".join(parts), ".".join(variants)


def _split_modifier_variants(modifiers: str) -> tuple[str, str]:
    """Return (known variant markers, unknown modifiers).

    TH and vN are intentional archive version markers and must not create
    REVIEW by themselves. Anything else remains visible for manual review.
    """
    if not modifiers:
        return "", ""
    variants, unknown = [], []
    for token in [p for p in modifiers.split(".") if p]:
        if re.fullmatch(VARIANT_TOKEN_PATTERN, token, flags=re.I):
            variants.append(_canonical_variant_token(token))
        else:
            unknown.append(token)
    return ".".join(variants), ".".join(unknown)


def _combine_variants(*values: str) -> str:
    out = []
    for value in values:
        for token in [p for p in (value or "").split(".") if p]:
            token = _canonical_variant_token(token)
            if token not in out:
                out.append(token)
    return ".".join(out)


def _embedded_resolution_tokens(title_area: str) -> list[str]:
    """Return scene resolution tokens accidentally left in the title area.

    The structural boundary intentionally uses the LAST resolution token.  This
    lets us parse malformed names such as ``Taxi.720p.1080p.MUBI...`` while
    still reporting the earlier 720p as a filename anomaly instead of silently
    treating it as part of the title.
    """
    if not title_area:
        return []
    return re.findall(
        r"(?:^|\.)(2160p|1080p|1080i|720p|576p|480p|360p|4K|UHD)(?=\.|$)",
        title_area,
        flags=re.I,
    )


def _remove_embedded_resolution_tokens(title_area: str) -> str:
    if not title_area:
        return title_area
    cleaned = re.sub(
        r"(?:^|\.)(?:2160p|1080p|1080i|720p|576p|480p|360p|4K|UHD)(?=\.|$)",
        "",
        title_area,
        flags=re.I,
    )
    return re.sub(r"\.{2,}", ".", cleaned).strip(".")


def _split_audio_channels(text: str) -> tuple[str, str, str, str]:
    """Split technical audio tail.

    Returns ``(audio, channels, features, modifiers)``.

    Accepted archive spellings include:
        DDP.5.1 / AAC.2.0
        DDP5.1  / AAC2.0
        DD+5.1              (legacy Dolby Digital Plus spelling)
        DDP5.1.Atmos
        TH.DD5.1            (legacy language prefix)
        v2.DDP5.1           (legacy/version modifier prefix)

    ``modifiers`` is returned to the caller. Known archive variants such as
    TH/vN are accepted; unknown modifiers remain REVIEW items.
    """
    text = (text or "").strip(".")
    if not text:
        return "", "", "", ""

    features = ""
    fm = re.search(rf"(?:^|\.)(?P<feature>{AUDIO_FEATURE_PATTERN})$", text, flags=re.I)
    if fm:
        features = fm.group("feature")
        text = text[:fm.start()].rstrip(".")

    # Separate scene-style channel token: DDP.5.1
    prefix, channels = _take_tail_token(text, CHANNEL_TOKEN_PATTERN)
    if channels:
        # The token before channels should end in the audio codec. Allow an
        # older prefix such as TH. or v2. without treating it as the codec.
        am = re.search(rf"(?:^|\.)(?P<audio>{AUDIO_TOKEN_PATTERN})$", prefix, flags=re.I)
        if am:
            modifiers = prefix[:am.start()].rstrip(".")
            return am.group("audio"), channels, features, modifiers

    # Compact codec+channels: DDP5.1 / DD+2.0, again allowing a prefix.
    m = re.search(
        rf"(?:^|\.)(?P<audio>{AUDIO_TOKEN_PATTERN})(?P<channels>{CHANNEL_TOKEN_PATTERN})$",
        text,
        flags=re.I,
    )
    if m:
        modifiers = text[:m.start()].rstrip(".")
        return m.group("audio"), m.group("channels"), features, modifiers

    return text, "", features, ""


def parse_filename_metadata(filename: str) -> FilenameMetadata:
    """Parse expected archive metadata without touching the media file.

    v2.7 keeps identity errors separate from technical parsing. Movie filenames
    still require a release year, while SxxEyy episode filenames are allowed to
    omit it. An omitted resolution token is intentional and means SD; ffprobe
    decides whether that implicit SD expectation is correct.
    """
    stem = unicodedata.normalize("NFKC", Path(filename).stem)
    notes = []

    m = FILENAME_TECH_BOUNDARY_RE.match(stem)
    if m:
        prefix = m.group("prefix")
        resolution = m.group("resolution")
        rest = m.group("rest")
    else:
        # Fallback for older filenames that omitted the resolution token, e.g.
        # Title.2015.NF.WEB-DL.DDP2.0.H.264.mkv.  Find release type first and
        # treat the immediately preceding token as source.
        fallback = re.match(
            rf"^(?P<prefix>.+)\.(?P<source>[^.]+)\."
            rf"(?P<release>{RELEASE_TYPE_PATTERN})(?:\.(?P<tech>.+))?$",
            stem,
            flags=re.I,
        )
        if not fallback:
            years = YEAR_RE.findall(stem)
            year = years[-1] if years else ""
            title_stem = stem
            if year and re.search(rf"\.{re.escape(year)}$", stem):
                title_stem = re.sub(rf"\.{re.escape(year)}$", "", stem)
            return FilenameMetadata(
                title=_pretty_title(title_stem),
                year=year,
                standard="UNPARSED",
                notes="No recognizable technical boundary found",
            )
        prefix = fallback.group("prefix")
        resolution = ""
        source = fallback.group("source") or ""
        source, source_variant = _split_source_variant(source)
        release_type = fallback.group("release") or ""
        tech = fallback.group("tech") or ""
        # No resolution token is valid for SD material in this archive.
        # Do not mark the filename itself for review here; compare_resolution()
        # will flag it only if ffprobe shows 720p or higher.

        ym = PREFIX_YEAR_RE.match(prefix)
        pym = PREFIX_PLACEHOLDER_YEAR_RE.match(prefix) if not ym else None
        if ym:
            raw_title = ym.group("title")
            year = ym.group("year")
        elif pym:
            raw_title = pym.group("title")
            year = pym.group("year").upper()
            notes.append(f"placeholder year: {year}")
        else:
            raw_title = prefix
            year = ""
            # TV episode filenames (SxxEyy) intentionally omit release year.
            # Movies/non-episode media still require one.
            if not EPISODE_RE.search(_pretty_title(prefix)):
                notes.append("release year missing")

        embedded_res = _embedded_resolution_tokens(raw_title)
        if embedded_res:
            notes.append(
                "additional resolution token in title area: " + ", ".join(embedded_res)
            )
            raw_title = _remove_embedded_resolution_tokens(raw_title)

        tech, video, annotation = _take_video_tail(tech)
        audio, channels, audio_features, modifiers = _split_audio_channels(tech)
        modifier_variant, unknown_modifiers = _split_modifier_variants(modifiers)
        variant = _combine_variants(source_variant, modifier_variant)

        if annotation:
            if annotation.startswith("."):
                notes.append(f"unexpected trailing token after video codec: {annotation[1:]}")
            else:
                notes.append(f"trailing annotation after video codec: {annotation}")
        if unknown_modifiers:
            notes.append(f"technical modifier unrecognized: {unknown_modifiers}")
        if not video:
            notes.append("video codec token missing/unrecognized")
        if not channels:
            notes.append("channel token missing/unrecognized")
        if not audio or not re.fullmatch(AUDIO_TOKEN_PATTERN, audio, flags=re.I):
            if audio:
                notes.append(f"audio token unrecognized: {audio}")
            else:
                notes.append("audio codec token missing")
            audio = audio or ""

        release_display = "WEB-DL" if release_type.upper() == "WEB.DL" else release_type
        standard = "OK" if not notes else "REVIEW"
        legacy_video = (video or "").lower() in {"x264", "x265", "h264", "h265"}
        legacy_audio = (audio or "").upper() == "DD+"
        if not notes and (legacy_video or legacy_audio):
            standard = "LEGACY"
        return FilenameMetadata(
            title=_pretty_title(raw_title),
            year=year,
            resolution=resolution,
            source=source,
            release_type=release_display,
            variant=variant,
            audio_codec=audio,
            channels=channels,
            audio_features=audio_features,
            video_codec=video,
            standard=standard,
            notes="; ".join(notes),
        )

    # Film/year identity is allowed to be incomplete while technical parsing
    # continues. A normal movie filename ends prefix with a real year.
    ym = PREFIX_YEAR_RE.match(prefix)
    pym = PREFIX_PLACEHOLDER_YEAR_RE.match(prefix) if not ym else None
    if ym:
        raw_title = ym.group("title")
        year = ym.group("year")
    elif pym:
        raw_title = pym.group("title")
        year = pym.group("year").upper()
        notes.append(f"placeholder year: {year}")
    else:
        raw_title = prefix
        year = ""
        # TV episode filenames (SxxEyy) intentionally omit release year.
        # Movies/non-episode media still require one.
        if not EPISODE_RE.search(_pretty_title(prefix)):
            notes.append("release year missing")

    embedded_res = _embedded_resolution_tokens(raw_title)
    if embedded_res:
        notes.append(
            "additional resolution token in title area: " + ", ".join(embedded_res)
        )
        raw_title = _remove_embedded_resolution_tokens(raw_title)

    # Locate source + release type; everything after it is the technical tail.
    release_re = re.compile(
        rf"^(?P<source>.+?)\.(?P<release>{RELEASE_TYPE_PATTERN})(?:\.(?P<tech>.+))?$",
        re.I,
    )
    rm = release_re.match(rest)
    if not rm:
        notes.append("source/release-type structure was not recognized")
        return FilenameMetadata(
            title=_pretty_title(raw_title),
            year=year,
            resolution=resolution,
            standard="REVIEW",
            notes="; ".join(notes),
        )

    source = rm.group("source") or ""
    source, source_variant = _split_source_variant(source)
    release_type = rm.group("release") or ""
    tech = rm.group("tech") or ""

    tech, video, annotation = _take_video_tail(tech)
    audio, channels, audio_features, modifiers = _split_audio_channels(tech)
    modifier_variant, unknown_modifiers = _split_modifier_variants(modifiers)
    variant = _combine_variants(source_variant, modifier_variant)

    if annotation:
        if annotation.startswith("."):
            notes.append(f"unexpected trailing token after video codec: {annotation[1:]}")
        else:
            notes.append(f"trailing annotation after video codec: {annotation}")
    if unknown_modifiers:
        notes.append(f"technical modifier unrecognized: {unknown_modifiers}")
    if not video:
        notes.append("video codec token missing/unrecognized")
    if not channels:
        notes.append("channel token missing/unrecognized")
    if not audio or not re.fullmatch(AUDIO_TOKEN_PATTERN, audio, flags=re.I):
        if audio:
            notes.append(f"audio token unrecognized: {audio}")
        else:
            notes.append("audio codec token missing")
        audio = audio or ""

    # Atmos is a real expectation, but ffprobe does not expose it consistently.
    # Parse it structurally here; the comparison layer will use REVIEW when it
    # cannot positively verify the feature.
    standard = "OK" if not notes else "REVIEW"
    legacy_video = (video or "").lower() in {"x264", "x265", "h264", "h265"}
    legacy_audio = (audio or "").upper() == "DD+"
    if not notes and (legacy_video or legacy_audio):
        standard = "LEGACY"

    release_display = "WEB-DL" if release_type.upper() == "WEB.DL" else release_type

    return FilenameMetadata(
        title=_pretty_title(raw_title),
        year=year,
        resolution=resolution,
        source=source,
        release_type=release_display,
        variant=variant,
        audio_codec=audio,
        channels=channels,
        audio_features=audio_features,
        video_codec=video,
        standard=standard,
        notes="; ".join(notes),
    )


def normalize_title_text(title: str) -> str:
    s = unicodedata.normalize("NFKC", title or "")
    s = BRACKET_RE.sub(" ", s)
    s = EPISODE_RE.sub(" ", s)
    s = SEPARATOR_RE.sub(" ", s)
    s = "".join(ch if (ch.isalnum() or ch.isspace()) else " " for ch in s)
    return MULTISPACE_RE.sub(" ", s).strip().casefold()


def normalize_title(filename: str) -> tuple[str, str]:
    parsed = parse_filename_metadata(filename)
    if parsed.standard != "UNPARSED":
        return normalize_title_text(parsed.title), parsed.year

    # Conservative fallback for filenames outside the structural standard.
    stem = Path(filename).stem
    stem = unicodedata.normalize("NFKC", stem)
    years = YEAR_RE.findall(stem)
    year = years[-1] if years else ""
    s = BRACKET_RE.sub(" ", stem)
    s = YEAR_RE.sub(" ", s)
    s = EPISODE_RE.sub(" ", s)
    for pat in NOISE_PATTERNS:
        s = re.sub(pat, " ", s, flags=re.I)
    s = SEPARATOR_RE.sub(" ", s)
    s = re.sub(r"\b(?:cd|disc|disk|part)[ ._-]?\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\b(?:r\d+)\b", " ", s, flags=re.I)
    s = re.sub(r"\b(?:proper|repack)\b", " ", s, flags=re.I)
    s = "".join(ch if (ch.isalnum() or ch.isspace()) else " " for ch in s)
    return MULTISPACE_RE.sub(" ", s).strip().casefold(), year


def canonical_video_codec(value: str) -> str:
    x = re.sub(r"[^a-z0-9]+", "", (value or "").lower())
    if x in {"h264", "x264", "avc", "avc1"}:
        return "h264"
    if x in {"h265", "x265", "hevc", "hev1", "hvc1"}:
        return "hevc"
    if x == "av1":
        return "av1"
    return x


def canonical_audio_codec(value: str) -> str:
    raw = (value or "").strip().lower()
    if raw in {"dd+", "ddplus"}:
        return "eac3"
    x = re.sub(r"[^a-z0-9]+", "", raw)
    mapping = {
        "ddp": "eac3", "eac3": "eac3",
        "dd": "ac3", "ac3": "ac3",
        "aac": "aac", "dts": "dts", "dtshd": "dts", "dtshdma": "dts",
        "truehd": "truehd", "flac": "flac", "opus": "opus", "mp3": "mp3",
        "pcm": "pcm", "lpcm": "pcm",
    }
    return mapping.get(x, x)


def display_audio_codec(codec_name: str) -> str:
    c = canonical_audio_codec(codec_name)
    return {
        "eac3": "DDP", "ac3": "DD", "aac": "AAC", "dts": "DTS",
        "truehd": "TRUEHD", "flac": "FLAC", "opus": "OPUS", "mp3": "MP3",
        "pcm": "PCM",
    }.get(c, (codec_name or "").upper())


def channel_token_to_count(token: str) -> Optional[int]:
    x = (token or "").strip().lower()
    if x == "mono":
        return 1
    if x == "stereo":
        return 2
    mapping = {
        "1.0": 1, "2.0": 2, "2.1": 3, "3.0": 3, "4.0": 4,
        "5.1": 6, "6.1": 7, "7.1": 8,
    }
    return mapping.get(x)


def channels_to_label(channels: Optional[int]) -> str:
    if not channels:
        return ""
    mapping = {1: "1.0", 2: "2.0", 3: "3ch", 4: "4.0", 6: "5.1", 7: "6.1", 8: "7.1"}
    return mapping.get(channels, f"{channels}ch")


def resolution_token_to_class(token: str) -> str:
    x = (token or "").strip().lower()
    if x in {"4k", "uhd"}:
        return "2160p"
    return x


def infer_resolution_class(width: Optional[int], height: Optional[int], field_order: str = "") -> str:
    if not width or not height:
        return ""
    w, h = int(width), int(height)
    # Width thresholds intentionally allow scope/cropped WEB-DLs such as 1920x800.
    if w >= 3000 or h >= 1600:
        base = "2160"
    elif w >= 1600 or h >= 900:
        base = "1080"
    elif w >= 1100 or h >= 650:
        base = "720"
    elif h >= 520:
        base = "576"
    elif h >= 430:
        base = "480"
    else:
        base = str(h)
    fo = (field_order or "").lower()
    interlaced = fo not in {"", "unknown", "progressive"} and any(x in fo for x in ("tt", "bb", "tb", "bt"))
    return f"{base}{'i' if interlaced else 'p'}"


def select_primary_audio(audios: list[dict]) -> Optional[dict]:
    if not audios:
        return None
    for a in audios:
        if to_int((a.get("disposition") or {}).get("default")) == 1:
            return a
    return audios[0]


def compare_resolution(name_resolution: str, actual_class: str) -> str:
    if not actual_class:
        return "REVIEW"
    if not name_resolution:
        # Archive convention: no resolution token means SD. Therefore 480/576
        # (and lower) is correct, while 720p+ must be flagged because an HD
        # resolution token should have been present in the filename.
        m = re.match(r"^(\d+)", actual_class)
        if not m:
            return "REVIEW"
        return "MISMATCH" if int(m.group(1)) >= 720 else "OK"
    expected = resolution_token_to_class(name_resolution)
    if expected == actual_class:
        return "OK"
    # If ffprobe cannot reliably expose interlace/progressive, don't punish p/i alone.
    if expected.rstrip("pi") == actual_class.rstrip("pi"):
        return "REVIEW"
    return "MISMATCH"


def compare_video(name_video: str, actual_video: str) -> str:
    if not name_video or not actual_video:
        return "REVIEW"
    expected = canonical_video_codec(name_video)
    actual = canonical_video_codec(actual_video)
    if expected != actual:
        return "MISMATCH"
    # Current naming standard is H.264/H.265. Older x264/x265 and H264/H265
    # spellings are equivalent codecs but remain visible as LEGACY.
    if (name_video or "").lower() in {"x264", "x265", "h264", "h265"}:
        return "LEGACY"
    return "OK"


def compare_audio_codec(name_audio: str, audios: list[dict], primary: Optional[dict]) -> str:
    if not name_audio or not audios:
        return "REVIEW"
    expected = canonical_audio_codec(name_audio)
    legacy = (name_audio or "").strip().upper() == "DD+"
    if primary and canonical_audio_codec(str(primary.get("codec_name") or "")) == expected:
        return "LEGACY" if legacy else "OK"
    if any(canonical_audio_codec(str(a.get("codec_name") or "")) == expected for a in audios):
        return "REVIEW"
    return "MISMATCH"


def detect_audio_features(stream: Optional[dict]) -> str:
    if not stream:
        return ""
    blob = json.dumps(stream, ensure_ascii=False).lower()
    if "atmos" in blob or re.search(r"\bjoc\b", blob):
        return "Atmos"
    return ""


def compare_audio_features(name_features: str, actual_features: str) -> str:
    if not name_features:
        return "OK"
    if name_features.lower() == "atmos":
        # Positive verification is strong. Absence is only REVIEW because many
        # ffprobe/FFmpeg versions expose E-AC-3 JOC simply as EAC3 5.1.
        return "OK" if (actual_features or "").lower() == "atmos" else "REVIEW"
    return "REVIEW"


def compare_audio_channels(name_channels: str, audios: list[dict], primary: Optional[dict]) -> str:
    expected = channel_token_to_count(name_channels)
    if expected is None or not audios:
        return "REVIEW"
    if primary and to_int(primary.get("channels")) == expected:
        return "OK"
    if any(to_int(a.get("channels")) == expected for a in audios):
        return "REVIEW"
    return "MISMATCH"


def combine_metadata_status(filename_standard: str, checks: list[str], ffprobe_ok: bool = True) -> str:
    if not ffprobe_ok:
        return "REVIEW"
    if "MISMATCH" in checks:
        return "MISMATCH"
    if filename_standard in {"UNPARSED", "REVIEW"} or "REVIEW" in checks:
        return "REVIEW"
    if filename_standard == "LEGACY" or "LEGACY" in checks:
        return "LEGACY"
    return "OK"


def build_x264_rename(filename: str, actual_video: str) -> str:
    if canonical_video_codec(actual_video) != "h264":
        return ""
    p = Path(filename)
    new_stem, n = re.subn(r"(?i)\.x264$", ".H.264", p.stem, count=1)
    if n != 1:
        return ""
    return new_stem + p.suffix


def quick_hash(path: Path, chunk_size: int = QUICK_HASH_CHUNK) -> str:
    size = path.stat().st_size
    h = hashlib.sha256()
    with path.open("rb") as f:
        start = f.read(chunk_size)
        h.update(start)
        if size > chunk_size:
            seek_to = max(0, size - chunk_size)
            f.seek(seek_to)
            h.update(f.read(chunk_size))
    h.update(str(size).encode("ascii"))
    return h.hexdigest()


def full_sha256(path: Path, block_size: int = 8 * 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            block = f.read(block_size)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def detect_hdr(video: dict) -> str:
    transfer = str(video.get("color_transfer") or "").lower()
    primaries = str(video.get("color_primaries") or "").lower()
    pix_fmt = str(video.get("pix_fmt") or "").lower()
    profile = str(video.get("profile") or "").lower()

    # Dolby Vision can show up in side data depending on ffprobe version.
    blob = json.dumps(video.get("side_data_list", []), ensure_ascii=False).lower()
    if "dolby vision" in blob or "dovi" in blob:
        return "Dolby Vision"
    if transfer in {"smpte2084", "pq"}:
        return "HDR10/PQ"
    if transfer in {"arib-std-b67", "hlg"}:
        return "HLG"
    if "10" in pix_fmt and primaries == "bt2020":
        return "10-bit / BT.2020"
    return ""


def infer_bit_depth(pix_fmt: str) -> str:
    p = (pix_fmt or "").lower()
    if any(x in p for x in ("p16", "16le", "16be")):
        return "16"
    if any(x in p for x in ("p12", "12le", "12be")):
        return "12"
    if any(x in p for x in ("p10", "10le", "10be")):
        return "10"
    if p:
        return "8"
    return ""


def probe_media(ffprobe: str, path: Path) -> dict:
    cmd = [
        ffprobe, "-v", "error",
        "-show_format", "-show_streams",
        "-print_format", "json",
        str(path)
    ]
    cp = subprocess.run(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace"
    )
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr.strip() or f"ffprobe return code {cp.returncode}")
    return json.loads(cp.stdout)


def stream_language(s: dict) -> str:
    tags = s.get("tags") or {}
    lang = str(tags.get("language") or tags.get("LANGUAGE") or "").strip()
    return lang.lower()


def audio_desc(s: dict) -> str:
    codec = str(s.get("codec_name") or "").upper()
    profile = str(s.get("profile") or "")
    channels = to_int(s.get("channels"))
    layout = str(s.get("channel_layout") or "")
    lang = stream_language(s)
    br = to_int(s.get("bit_rate"))
    parts = []
    if lang:
        parts.append(lang)
    if codec:
        parts.append(codec)
    if profile and profile.lower() not in {"unknown"}:
        parts.append(profile)
    if channels:
        parts.append(f"{channels}ch")
    elif layout:
        parts.append(layout)
    if br:
        parts.append(f"{round(br/1000)}kbps")
    return " ".join(parts)


def sub_desc(s: dict) -> str:
    codec = str(s.get("codec_name") or "").upper()
    lang = stream_language(s)
    title = str((s.get("tags") or {}).get("title") or "").strip()
    parts = [x for x in (lang, codec, title) if x]
    return " ".join(parts)


def is_english_language(lang: str, title: str = "") -> bool:
    x = f"{lang} {title}".lower()
    return any(token in x for token in ("eng", "english", "en "))


def make_record(path: Path, root: Path, ffdata: Optional[dict],
                qhash: str = "", sha256: str = "", error: str = "") -> MediaRecord:
    st = path.stat()
    parsed = parse_filename_metadata(path.name)
    norm_title = normalize_title_text(parsed.title) if parsed.title else normalize_title(path.name)[0]
    year = parsed.year or normalize_title(path.name)[1]

    def common_expected_fields():
        return dict(
            filename_title=parsed.title,
            normalized_title=norm_title,
            filename_year=year,
            name_resolution=parsed.resolution,
            streaming_source=parsed.source,
            release_type=parsed.release_type,
            filename_variant=parsed.variant,
            name_audio=parsed.audio_codec,
            name_channels=parsed.channels,
            name_audio_features=parsed.audio_features,
            name_video=parsed.video_codec,
            filename_standard=parsed.standard,
            filename_parse_notes=parsed.notes,
        )

    if not ffdata:
        checks = ["REVIEW", "REVIEW", "REVIEW", "REVIEW"]
        return MediaRecord(
            path=str(path), root=str(root), filename=path.name, extension=path.suffix.lower(),
            size_bytes=st.st_size, size_gib=round(st.st_size / (1024**3), 3),
            modified=datetime.fromtimestamp(st.st_mtime).isoformat(sep=" ", timespec="seconds"),
            duration_sec=None, duration_text="", container="", overall_bitrate_kbps=None,
            computed_bitrate_kbps=None, video_codec="", video_profile="", video_encoder="",
            width=None, height=None, resolution="", actual_resolution_class="", fps=None,
            pixel_format="", bit_depth="", hdr="", video_bitrate_kbps=None,
            audio_tracks=0, audio_summary="", primary_audio_codec="", primary_audio_channels="",
            actual_audio_features="", max_audio_channels=None, has_51=False,
            subtitle_tracks=0, subtitle_summary="", has_english_sub=False, title_tag="",
            **common_expected_fields(),
            resolution_match="REVIEW", audio_match="REVIEW", channels_match="REVIEW",
            audio_features_match="REVIEW" if parsed.audio_features else "OK", video_match="REVIEW",
            metadata_check=combine_metadata_status(parsed.standard, checks, ffprobe_ok=False),
            rename_candidate="", rename_status="", rename_target="",
            quick_hash=qhash, sha256=sha256, ffprobe_ok=False, error=error
        )

    fmt = ffdata.get("format") or {}
    streams = ffdata.get("streams") or []
    videos = [s for s in streams if s.get("codec_type") == "video"]
    audios = [s for s in streams if s.get("codec_type") == "audio"]
    subs = [s for s in streams if s.get("codec_type") == "subtitle"]
    video = videos[0] if videos else {}
    primary_audio = select_primary_audio(audios)

    duration = to_float(fmt.get("duration"))
    overall_br = to_int(fmt.get("bit_rate"))
    computed_br = None
    if duration and duration > 0:
        computed_br = round((st.st_size * 8 / duration) / 1000)

    width = to_int(video.get("width"))
    height = to_int(video.get("height"))
    resolution = f"{width}x{height}" if width and height else ""
    actual_res_class = infer_resolution_class(width, height, str(video.get("field_order") or ""))
    fps = parse_rate(str(video.get("avg_frame_rate") or video.get("r_frame_rate") or ""))

    audio_summaries = [audio_desc(a) for a in audios]
    audio_summaries = [x for x in audio_summaries if x]
    max_channels = max([to_int(a.get("channels")) or 0 for a in audios], default=0) or None
    primary_audio_codec = display_audio_codec(str((primary_audio or {}).get("codec_name") or ""))
    primary_channels_count = to_int((primary_audio or {}).get("channels"))
    primary_audio_channels = channels_to_label(primary_channels_count)
    actual_audio_features = detect_audio_features(primary_audio)

    subtitle_summaries = [sub_desc(s) for s in subs]
    subtitle_summaries = [x for x in subtitle_summaries if x]

    has_eng = False
    for sub in subs:
        tags = sub.get("tags") or {}
        if is_english_language(stream_language(sub), str(tags.get("title") or "")):
            has_eng = True
            break

    fmt_tags = fmt.get("tags") or {}
    title_tag = str(fmt_tags.get("title") or fmt_tags.get("TITLE") or "").strip()
    video_tags = video.get("tags") or {}
    video_encoder = str(
        video_tags.get("encoder") or video_tags.get("ENCODER") or
        video_tags.get("writing_library") or video_tags.get("WRITING_LIBRARY") or ""
    ).strip()

    actual_video = str(video.get("codec_name") or "").upper()
    resolution_match = compare_resolution(parsed.resolution, actual_res_class)
    audio_match = compare_audio_codec(parsed.audio_codec, audios, primary_audio)
    channels_match = compare_audio_channels(parsed.channels, audios, primary_audio)
    audio_features_match = compare_audio_features(parsed.audio_features, actual_audio_features)
    video_match = compare_video(parsed.video_codec, actual_video)
    metadata_check = combine_metadata_status(
        parsed.standard,
        [resolution_match, audio_match, channels_match, audio_features_match, video_match],
        ffprobe_ok=True,
    )

    rename_target = ""
    rename_candidate = ""
    rename_status = ""
    if parsed.video_codec.lower() == "x264":
        rename_target = build_x264_rename(path.name, actual_video)
        if rename_target:
            rename_candidate = "YES"
            other_checks = [resolution_match, audio_match, channels_match]
            rename_status = "SAFE" if all(x in {"OK", "LEGACY"} for x in other_checks) else "REVIEW"

    pix = str(video.get("pix_fmt") or "")
    return MediaRecord(
        path=str(path), root=str(root), filename=path.name, extension=path.suffix.lower(),
        size_bytes=st.st_size, size_gib=round(st.st_size / (1024**3), 3),
        modified=datetime.fromtimestamp(st.st_mtime).isoformat(sep=" ", timespec="seconds"),
        duration_sec=round(duration, 3) if duration is not None else None,
        duration_text=human_duration(duration),
        container=str(fmt.get("format_name") or ""),
        overall_bitrate_kbps=round(overall_br / 1000) if overall_br else None,
        computed_bitrate_kbps=computed_br,
        video_codec=actual_video,
        video_profile=str(video.get("profile") or ""),
        video_encoder=video_encoder,
        width=width, height=height, resolution=resolution, actual_resolution_class=actual_res_class,
        fps=round(fps, 3) if fps else None,
        pixel_format=pix,
        bit_depth=infer_bit_depth(pix),
        hdr=detect_hdr(video),
        video_bitrate_kbps=round((to_int(video.get("bit_rate")) or 0) / 1000) or None,
        audio_tracks=len(audios),
        audio_summary=" | ".join(audio_summaries),
        primary_audio_codec=primary_audio_codec,
        primary_audio_channels=primary_audio_channels,
        actual_audio_features=actual_audio_features,
        max_audio_channels=max_channels,
        has_51=bool(max_channels and max_channels >= 6),
        subtitle_tracks=len(subs),
        subtitle_summary=" | ".join(subtitle_summaries),
        has_english_sub=has_eng,
        title_tag=title_tag,
        **common_expected_fields(),
        resolution_match=resolution_match,
        audio_match=audio_match,
        channels_match=channels_match,
        audio_features_match=audio_features_match,
        video_match=video_match,
        metadata_check=metadata_check,
        rename_candidate=rename_candidate,
        rename_status=rename_status,
        rename_target=rename_target,
        quick_hash=qhash,
        sha256=sha256,
        ffprobe_ok=True,
        error=error
    )


class CacheDB:
    def __init__(self, path: Path):
        self.path = path
        self.con = sqlite3.connect(path)
        self.con.row_factory = sqlite3.Row
        self._init()

    def _init(self):
        cur = self.con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS files (
                path TEXT PRIMARY KEY,
                size_bytes INTEGER NOT NULL,
                mtime_ns INTEGER NOT NULL,
                record_json TEXT NOT NULL
            )
        """)

        # Parser/record changes must not silently reuse stale cached rows.
        # v2 wrote the schema marker but did not actually enforce it.
        row = cur.execute("SELECT value FROM meta WHERE key='schema_version'").fetchone()
        old_schema = str(row[0]) if row else ""
        if old_schema != str(CACHE_SCHEMA_VERSION):
            cur.execute("DELETE FROM files")

        cur.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('schema_version',?)",
                    (str(CACHE_SCHEMA_VERSION),))
        self.con.commit()

    def get(self, path: Path):
        st = path.stat()
        row = self.con.execute(
            "SELECT size_bytes, mtime_ns, record_json FROM files WHERE path=?",
            (str(path),)
        ).fetchone()
        if not row:
            return None
        if row["size_bytes"] != st.st_size or row["mtime_ns"] != st.st_mtime_ns:
            return None
        try:
            return MediaRecord(**json.loads(row["record_json"]))
        except Exception:
            return None

    def put(self, path: Path, rec: MediaRecord):
        st = path.stat()
        self.con.execute("""
            INSERT INTO files(path,size_bytes,mtime_ns,record_json)
            VALUES(?,?,?,?)
            ON CONFLICT(path) DO UPDATE SET
                size_bytes=excluded.size_bytes,
                mtime_ns=excluded.mtime_ns,
                record_json=excluded.record_json
        """, (str(path), st.st_size, st.st_mtime_ns, json.dumps(asdict(rec), ensure_ascii=False)))
        self.con.commit()

    def close(self):
        self.con.close()


def discover_files(roots: list[Path]) -> list[tuple[Path, Path]]:
    out = []
    for root in roots:
        root = root.resolve()
        if not root.exists():
            print(f"VARNING: sökväg finns inte: {root}", file=sys.stderr)
            continue
        if root.is_file():
            if root.suffix.lower() in VIDEO_EXTS:
                out.append((root, root.parent))
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            # Skip some common hidden/system dirs.
            dirnames[:] = [
                d for d in dirnames
                if d not in {"$RECYCLE.BIN", "System Volume Information"}
                and not d.startswith(".")
            ]
            for fn in filenames:
                p = Path(dirpath) / fn
                if p.suffix.lower() in VIDEO_EXTS:
                    out.append((p.resolve(), root))
    out.sort(key=lambda x: str(x[0]).casefold())
    return out


def apply_candidate_hashing(records: list[MediaRecord], hash_mode: str):
    if hash_mode == "none":
        return

    size_groups = defaultdict(list)
    for i, r in enumerate(records):
        size_groups[r.size_bytes].append(i)

    # Quick hash only where size collides, unless "all".
    targets = range(len(records)) if hash_mode == "all" else [
        i for inds in size_groups.values() if len(inds) > 1 for i in inds
    ]

    for n, i in enumerate(targets, 1):
        r = records[i]
        if r.quick_hash:
            continue
        try:
            r.quick_hash = quick_hash(Path(r.path))
        except Exception as e:
            r.error = (r.error + " | " if r.error else "") + f"quick hash: {e}"
        if n % 25 == 0:
            print(f"  quick hash {n} files...")

    if hash_mode == "all":
        full_targets = list(range(len(records)))
    else:
        qgroups = defaultdict(list)
        for i, r in enumerate(records):
            if r.quick_hash:
                qgroups[(r.size_bytes, r.quick_hash)].append(i)
        full_targets = [
            i for inds in qgroups.values() if len(inds) > 1 for i in inds
        ]

    for n, i in enumerate(full_targets, 1):
        r = records[i]
        if r.sha256:
            continue
        try:
            r.sha256 = full_sha256(Path(r.path))
        except Exception as e:
            r.error = (r.error + " | " if r.error else "") + f"sha256: {e}"
        if n % 10 == 0:
            print(f"  full SHA-256 {n} files...")


def exact_duplicate_rows(records: list[MediaRecord]):
    groups = defaultdict(list)
    for r in records:
        if r.sha256:
            groups[r.sha256].append(r)

    rows = []
    gid = 0
    for sha, group in sorted(groups.items()):
        if len(group) < 2:
            continue
        gid += 1
        for r in group:
            rows.append({
                "Duplicate Group": gid,
                "SHA256": sha,
                "Size GiB": r.size_gib,
                "Duration": r.duration_text,
                "Resolution": r.resolution,
                "Path": r.path,
            })
    return rows


def possible_same_film_groups(records: list[MediaRecord], duration_tolerance: int = 120):
    """
    Conservative heuristic:
      - same normalized filename title, at least 3 chars
      - multiple files
      - if durations exist, keep only pairs/groups whose span <= tolerance OR
        split by ~2 minute duration buckets.
    """
    base = defaultdict(list)
    for r in records:
        key = r.normalized_title.strip()
        if len(key) >= 3:
            base[key].append(r)

    out = []
    gid = 0
    for title, group in sorted(base.items()):
        if len(group) < 2:
            continue

        # Split by year when both sides have years and they differ.
        yg = defaultdict(list)
        for r in group:
            yg[r.filename_year or "?"].append(r)

        candidate_sets = []
        if len([k for k in yg if k != "?"]) > 1:
            # unknown-year records stay together; known years separate
            candidate_sets.extend(v for v in yg.values() if len(v) > 1)
        else:
            candidate_sets = [group]

        for g in candidate_sets:
            if len(g) < 2:
                continue

            durations = [r.duration_sec for r in g if r.duration_sec]
            if durations and max(durations) - min(durations) > duration_tolerance:
                # split into rough duration buckets
                buckets = defaultdict(list)
                for r in g:
                    if r.duration_sec:
                        b = int(round(r.duration_sec / duration_tolerance))
                    else:
                        b = -1
                    buckets[b].append(r)
                subgroups = [x for x in buckets.values() if len(x) > 1]
            else:
                subgroups = [g]

            for sg in subgroups:
                if len(sg) < 2:
                    continue
                gid += 1
                for r in sg:
                    out.append({
                        "Possible Group": gid,
                        "Normalized Title": title,
                        "Filename Year": r.filename_year,
                        "Source": r.streaming_source,
                        "Variant": r.filename_variant,
                        "Duration": r.duration_text,
                        "Duration sec": r.duration_sec,
                        "Resolution": r.resolution,
                        "Computed kbps": r.computed_bitrate_kbps,
                        "Video": r.video_codec,
                        "Audio Tracks": r.audio_tracks,
                        "Max Audio Ch": r.max_audio_channels,
                        "English Subs": "YES" if r.has_english_sub else "",
                        "Size GiB": r.size_gib,
                        "Path": r.path,
                        "Exact SHA256": r.sha256,
                    })
    return out


def best_capabilities(records: list[MediaRecord]):
    rows = []
    groups = defaultdict(list)
    for r in records:
        if len(r.normalized_title) >= 3:
            groups[r.normalized_title].append(r)

    for title, g in sorted(groups.items()):
        if len(g) < 2:
            continue
        max_h = max((r.height or 0 for r in g), default=0)
        max_channels = max((r.max_audio_channels or 0 for r in g), default=0)
        eng = any(r.has_english_sub for r in g)
        hdr = any(bool(r.hdr) for r in g)
        max_br = max((r.computed_bitrate_kbps or 0 for r in g), default=0)

        rows.append({
            "Normalized Title": title,
            "Versions": len(g),
            "Variants": ", ".join(sorted({r.filename_variant for r in g if r.filename_variant})),
            "Highest Height": max_h or "",
            "Highest Computed kbps": max_br or "",
            "Max Audio Ch": max_channels or "",
            "Any English Subs": "YES" if eng else "",
            "Any HDR": "YES" if hdr else "",
            "Files": "\n".join(r.path for r in g),
        })
    return rows


def metadata_review_rows(records: list[MediaRecord]):
    rows = []
    for r in records:
        if r.metadata_check == "OK":
            continue
        # Verified x264 legacy names have their own dedicated rename sheet.
        if r.metadata_check == "LEGACY" and r.rename_candidate == "YES":
            continue
        rows.append({
            "Status": r.metadata_check,
            "Filename Standard": r.filename_standard,
            "Filename": r.filename,
            "Filename Title": r.filename_title,
            "Year": r.filename_year,
            "Source": r.streaming_source,
            "Variant": r.filename_variant,
            "Name Resolution": r.name_resolution,
            "Actual Resolution": r.resolution,
            "Actual Res Class": r.actual_resolution_class,
            "Resolution Match": r.resolution_match,
            "Name Audio": r.name_audio,
            "Actual Audio": r.primary_audio_codec,
            "Audio Match": r.audio_match,
            "Name Channels": r.name_channels,
            "Actual Channels": r.primary_audio_channels,
            "Channels Match": r.channels_match,
            "Name Audio Feature": r.name_audio_features,
            "Actual Audio Feature": r.actual_audio_features,
            "Audio Feature Match": r.audio_features_match,
            "Name Video": r.name_video,
            "Actual Video": r.video_codec,
            "Video Match": r.video_match,
            "Parser Notes": r.filename_parse_notes,
            "Error": r.error,
            "Path": r.path,
        })
    return rows


def rename_candidate_rows(records: list[MediaRecord]):
    rows = []
    for r in records:
        if r.rename_candidate != "YES":
            continue
        rows.append({
            "Rename Status": r.rename_status,
            "Current Filename": r.filename,
            "Proposed Filename": r.rename_target,
            "Actual Video": r.video_codec,
            "Encoder Tag": r.video_encoder,
            "Metadata Check": r.metadata_check,
            "Path": r.path,
        })
    return rows




@dataclass
class LocalFilmGroup:
    local_id: str
    title: str
    normalized_title: str
    year: str
    runtime_min: Optional[float]
    files: list[str]
    aliases: list[str]


@dataclass
class TmdbMatchResult:
    status: str = "NOT_RUN"
    confidence: float = 0.0
    margin: float = 0.0
    tmdb_id: Optional[int] = None
    imdb_id: str = ""
    canonical_title: str = ""
    original_title: str = ""
    original_language: str = ""
    release_date: str = ""
    runtime_min: Optional[int] = None
    genres: str = ""
    director: str = ""
    production_companies: str = ""
    production_countries: str = ""
    notes: str = ""
    aliases: list[dict] | None = None
    candidates: list[dict] | None = None


@dataclass
class FilmRecordRow:
    film_record_id: str
    local_group_ids: str
    canonical_title: str
    original_thai_title: str
    year: str
    tmdb_id: str
    imdb_id: str
    match_status: str
    match_confidence: float
    runtime_min: Optional[int]
    genres: str
    director: str
    production_companies: str
    file_count: int
    aliases_count: int
    files: str


def is_tv_episode_record(rec: MediaRecord) -> bool:
    """TV episodes are valid archive files, but v3.0's external matcher is movie-only."""
    return bool(EPISODE_RE.search(rec.filename_title or rec.filename or ""))


def stable_local_id(normalized_title: str, year: str) -> str:
    raw = f"{normalized_title}|{year}".encode("utf-8", "replace")
    return "LOCAL-" + hashlib.sha1(raw).hexdigest()[:12].upper()


def build_local_film_groups(records: list[MediaRecord]) -> tuple[list[LocalFilmGroup], dict[str, str], dict[str, str]]:
    """Build provisional movie groups from local filename identity.

    Returns (groups, path->local_group_id, path->skip_reason). TV episodes and
    genuinely unparsed files are deliberately excluded from TMDb movie matching.
    """
    grouped: dict[tuple[str, str], list[MediaRecord]] = defaultdict(list)
    path_to_local: dict[str, str] = {}
    skipped: dict[str, str] = {}

    for rec in records:
        if is_tv_episode_record(rec):
            skipped[rec.path] = "TV_EPISODE"
            continue
        if rec.filename_standard == "UNPARSED" or not rec.normalized_title.strip():
            skipped[rec.path] = "UNPARSED"
            continue
        key = (rec.normalized_title.strip(), rec.filename_year or "")
        grouped[key].append(rec)

    out: list[LocalFilmGroup] = []
    for (norm, year), group in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        local_id = stable_local_id(norm, year)
        titles = [r.filename_title.strip() for r in group if r.filename_title.strip()]
        # Prefer the most frequent local title, then the shortest clean variant.
        counts = defaultdict(int)
        for t in titles:
            counts[t] += 1
        if counts:
            title = sorted(counts, key=lambda t: (-counts[t], len(t), t.casefold()))[0]
        else:
            title = norm

        alias_set = {t for t in titles if t}
        for r in group:
            if r.title_tag and len(r.title_tag.strip()) >= 2:
                alias_set.add(r.title_tag.strip())

        durations = sorted(r.duration_sec / 60.0 for r in group if r.duration_sec and r.duration_sec > 0)
        runtime = None
        if durations:
            mid = len(durations) // 2
            runtime = durations[mid] if len(durations) % 2 else (durations[mid - 1] + durations[mid]) / 2.0

        files = [r.path for r in group]
        for r in group:
            path_to_local[r.path] = local_id

        out.append(LocalFilmGroup(
            local_id=local_id,
            title=title,
            normalized_title=norm,
            year=year,
            runtime_min=round(runtime, 1) if runtime is not None else None,
            files=files,
            aliases=sorted(alias_set, key=str.casefold),
        ))
    return out, path_to_local, skipped


def tmdb_norm(text: str) -> str:
    s = unicodedata.normalize("NFKC", text or "").casefold()
    s = "".join(ch if ch.isalnum() else " " for ch in s)
    return MULTISPACE_RE.sub(" ", s).strip()


def title_similarity(a: str, b: str) -> float:
    aa, bb = tmdb_norm(a), tmdb_norm(b)
    if not aa or not bb:
        return 0.0
    if aa == bb:
        return 100.0
    return round(difflib.SequenceMatcher(None, aa, bb).ratio() * 100.0, 1)


def extract_year(date_text: str) -> Optional[int]:
    m = re.match(r"^(\d{4})", str(date_text or ""))
    return int(m.group(1)) if m else None


def year_component(local_year: str, release_date: str) -> float:
    if not local_year or not local_year.isdigit():
        return 50.0
    y = int(local_year)
    cy = extract_year(release_date)
    if cy is None:
        return 35.0
    diff = abs(y - cy)
    return {0: 100.0, 1: 80.0, 2: 50.0}.get(diff, 0.0)


def runtime_component(local_runtime: Optional[float], candidate_runtime: Optional[int]) -> float:
    if local_runtime is None or not candidate_runtime:
        return 50.0
    diff = abs(local_runtime - candidate_runtime)
    if diff <= 2:
        return 100.0
    if diff <= 5:
        return 85.0
    if diff <= 10:
        return 60.0
    if diff <= 20:
        return 25.0
    return 0.0


def candidate_aliases(details: dict) -> list[dict]:
    rows: list[dict] = []
    seen: set[tuple[str, str, str]] = set()

    def add(title: str, source: str, country: str = ""):
        t = str(title or "").strip()
        if not t:
            return
        k = (tmdb_norm(t), source, country or "")
        if k in seen:
            return
        seen.add(k)
        rows.append({"title": t, "source": source, "country": country or ""})

    add(details.get("title"), "TMDb title")
    add(details.get("original_title"), "TMDb original title")
    for item in (details.get("alternative_titles") or {}).get("titles", []) or []:
        add(item.get("title"), "TMDb alternative title", item.get("iso_3166_1") or "")
    return rows


def score_tmdb_candidate(group: LocalFilmGroup, details: dict) -> dict:
    aliases = candidate_aliases(details)
    sims = [(title_similarity(group.title, a["title"]), a["title"]) for a in aliases]
    best_title_score, best_alias = max(sims, default=(0.0, ""), key=lambda x: x[0])
    yc = year_component(group.year, details.get("release_date") or "")
    lang = str(details.get("original_language") or "").lower()
    lang_score = 100.0 if lang == "th" else 20.0
    countries = [str(x.get("iso_3166_1") or "").upper() for x in details.get("production_countries", []) or []]
    country_score = 100.0 if "TH" in countries else (35.0 if not countries else 10.0)
    rc = runtime_component(group.runtime_min, details.get("runtime"))

    total = (
        best_title_score * 0.55
        + yc * 0.15
        + lang_score * 0.10
        + country_score * 0.10
        + rc * 0.10
    )
    return {
        "tmdb_id": details.get("id"),
        "score": round(total, 1),
        "title_score": best_title_score,
        "best_alias": best_alias,
        "year_score": yc,
        "language_score": lang_score,
        "country_score": country_score,
        "runtime_score": rc,
        "details": details,
    }


class TmdbHttpCache:
    def __init__(self, path: Path, ttl_days: int = 30):
        self.path = path
        self.ttl_days = max(1, int(ttl_days))
        self.con = sqlite3.connect(path)
        self.con.row_factory = sqlite3.Row
        self.con.execute("""
            CREATE TABLE IF NOT EXISTS responses (
                cache_key TEXT PRIMARY KEY,
                fetched_at TEXT NOT NULL,
                response_json TEXT NOT NULL
            )
        """)
        self.con.commit()

    def get(self, key: str) -> Optional[dict]:
        row = self.con.execute(
            "SELECT fetched_at, response_json FROM responses WHERE cache_key=?", (key,)
        ).fetchone()
        if not row:
            return None
        try:
            fetched = datetime.fromisoformat(row["fetched_at"])
            if datetime.now() - fetched > timedelta(days=self.ttl_days):
                return None
            return json.loads(row["response_json"])
        except Exception:
            return None

    def put(self, key: str, value: dict):
        self.con.execute(
            "INSERT OR REPLACE INTO responses(cache_key,fetched_at,response_json) VALUES(?,?,?)",
            (key, datetime.now().isoformat(timespec="seconds"), json.dumps(value, ensure_ascii=False)),
        )
        self.con.commit()

    def close(self):
        self.con.close()


class TmdbClient:
    BASE = "https://api.themoviedb.org/3"

    def __init__(self, access_token: str = "", api_key: str = "", language: str = "en-US",
                 cache_path: Optional[Path] = None, cache_ttl_days: int = 30,
                 timeout: int = 20, max_candidates: int = 3,
                 auto_match_score: float = 90.0, auto_match_margin: float = 8.0):
        self.access_token = access_token.strip()
        self.api_key = api_key.strip()
        self.language = language.strip() or "en-US"
        self.timeout = max(5, int(timeout))
        self.max_candidates = max(1, min(5, int(max_candidates)))
        self.auto_match_score = float(auto_match_score)
        self.auto_match_margin = float(auto_match_margin)
        self.cache = TmdbHttpCache(cache_path, cache_ttl_days) if cache_path else None

    @property
    def configured(self) -> bool:
        return bool(self.access_token or self.api_key)

    def close(self):
        if self.cache:
            self.cache.close()

    def _get(self, path: str, params: dict) -> dict:
        clean = {k: v for k, v in params.items() if v not in (None, "")}
        if self.api_key:
            clean["api_key"] = self.api_key
        query = urllib.parse.urlencode(clean)
        url = f"{self.BASE}{path}" + (f"?{query}" if query else "")
        # Do not leak credentials into the persistent cache key.
        cache_params = {k: v for k, v in clean.items() if k != "api_key"}
        cache_key = path + "?" + urllib.parse.urlencode(sorted(cache_params.items()))
        if self.cache:
            hit = self.cache.get(cache_key)
            if hit is not None:
                return hit

        headers = {"Accept": "application/json", "User-Agent": "ThaiMediaInventory/3.0"}
        if self.access_token:
            headers["Authorization"] = f"Bearer {self.access_token}"

        last_error = None
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                if self.cache:
                    self.cache.put(cache_key, data)
                return data
            except urllib.error.HTTPError as e:
                last_error = e
                if e.code == 429 and attempt < 2:
                    retry_after = to_int(e.headers.get("Retry-After")) or (2 + attempt * 2)
                    time.sleep(min(retry_after, 10))
                    continue
                body = e.read().decode("utf-8", "replace")[:500]
                raise RuntimeError(f"TMDb HTTP {e.code}: {body or e.reason}") from e
            except urllib.error.URLError as e:
                last_error = e
                if attempt < 2:
                    time.sleep(1 + attempt)
                    continue
                raise RuntimeError(f"TMDb network error: {e}") from e
        raise RuntimeError(f"TMDb request failed: {last_error}")

    def search_movie(self, title: str, year: str = "") -> list[dict]:
        params = {"query": title, "include_adult": "false", "language": self.language}
        if year and year.isdigit():
            params["year"] = year
        data = self._get("/search/movie", params)
        results = data.get("results", []) or []
        if not results and year:
            params.pop("year", None)
            data = self._get("/search/movie", params)
            results = data.get("results", []) or []
        return results

    def movie_details(self, tmdb_id: int) -> dict:
        return self._get(
            f"/movie/{int(tmdb_id)}",
            {
                "language": self.language,
                "append_to_response": "alternative_titles,external_ids,credits",
            },
        )

    def match_group(self, group: LocalFilmGroup) -> TmdbMatchResult:
        try:
            search = self.search_movie(group.title, group.year)
            if not search:
                return TmdbMatchResult(status="NO_MATCH", notes="TMDb search returned no candidates", candidates=[])

            prelim = []
            for idx, c in enumerate(search[:10]):
                ts = max(title_similarity(group.title, c.get("title") or ""),
                         title_similarity(group.title, c.get("original_title") or ""))
                ys = year_component(group.year, c.get("release_date") or "")
                prelim.append((ts * 0.8 + ys * 0.2, idx, c))
            prelim.sort(reverse=True, key=lambda x: (x[0], -x[1]))

            scored = []
            for _, _, c in prelim[:self.max_candidates]:
                cid = c.get("id")
                if not cid:
                    continue
                details = self.movie_details(int(cid))
                scored.append(score_tmdb_candidate(group, details))
            scored.sort(key=lambda x: x["score"], reverse=True)
            if not scored:
                return TmdbMatchResult(status="NO_MATCH", notes="TMDb candidates could not be enriched", candidates=[])

            top = scored[0]
            second = scored[1]["score"] if len(scored) > 1 else 0.0
            margin = round(top["score"] - second, 1)
            details = top["details"]
            high = (
                top["score"] >= self.auto_match_score
                and margin >= self.auto_match_margin
                and top["title_score"] >= 88.0
                and (not group.year or top["year_score"] >= 80.0)
            )
            status = "HIGH_CONFIDENCE" if high else "REVIEW"

            credits = details.get("credits") or {}
            directors = []
            for crew in credits.get("crew", []) or []:
                if str(crew.get("job") or "").casefold() == "director":
                    name = str(crew.get("name") or "").strip()
                    if name and name not in directors:
                        directors.append(name)

            external_ids = details.get("external_ids") or {}
            genres = ", ".join(str(x.get("name") or "") for x in details.get("genres", []) or [] if x.get("name"))
            companies = ", ".join(str(x.get("name") or "") for x in details.get("production_companies", []) or [] if x.get("name"))
            countries = ", ".join(str(x.get("iso_3166_1") or x.get("name") or "") for x in details.get("production_countries", []) or [] if x)

            cand_rows = []
            for rank, c in enumerate(scored, 1):
                d = c["details"]
                cand_rows.append({
                    "rank": rank,
                    "tmdb_id": d.get("id"),
                    "title": d.get("title") or "",
                    "original_title": d.get("original_title") or "",
                    "release_date": d.get("release_date") or "",
                    "original_language": d.get("original_language") or "",
                    "runtime": d.get("runtime"),
                    "score": c["score"],
                    "title_score": c["title_score"],
                    "best_alias": c["best_alias"],
                    "year_score": c["year_score"],
                    "runtime_score": c["runtime_score"],
                })

            return TmdbMatchResult(
                status=status,
                confidence=top["score"],
                margin=margin,
                tmdb_id=details.get("id"),
                imdb_id=str(external_ids.get("imdb_id") or ""),
                canonical_title=str(details.get("title") or ""),
                original_title=str(details.get("original_title") or ""),
                original_language=str(details.get("original_language") or ""),
                release_date=str(details.get("release_date") or ""),
                runtime_min=to_int(details.get("runtime")),
                genres=genres,
                director=", ".join(directors),
                production_companies=companies,
                production_countries=countries,
                aliases=candidate_aliases(details),
                candidates=cand_rows,
                notes=(f"best alias: {top['best_alias']}; margin: {margin:.1f}"),
            )
        except Exception as e:
            return TmdbMatchResult(status="ERROR", notes=str(e), candidates=[])


def build_film_layer(records: list[MediaRecord], tmdb_client: Optional[TmdbClient] = None):
    """Return Film Record rows, Title Alias rows, API-review rows and file enrichment."""
    groups, path_to_local, skipped = build_local_film_groups(records)
    matches: dict[str, TmdbMatchResult] = {}

    for idx, group in enumerate(groups, 1):
        if tmdb_client and tmdb_client.configured:
            print(f"TMDb [{idx}/{len(groups)}] {group.title} ({group.year or '?'})")
            matches[group.local_id] = tmdb_client.match_group(group)
        else:
            matches[group.local_id] = TmdbMatchResult(status="NOT_RUN")

    # Merge only independently high-confidence local groups that resolve to the same TMDb movie.
    bucket: dict[str, list[LocalFilmGroup]] = defaultdict(list)
    for g in groups:
        m = matches[g.local_id]
        if m.status == "HIGH_CONFIDENCE" and m.tmdb_id:
            key = f"TMDB-{m.tmdb_id}"
        else:
            key = g.local_id
        bucket[key].append(g)

    film_rows: list[dict] = []
    alias_rows: list[dict] = []
    api_review_rows: list[dict] = []
    file_enrichment: dict[str, dict] = {}

    for film_key, members in sorted(bucket.items(), key=lambda x: x[0]):
        primary = members[0]
        pm = matches[primary.local_id]
        accepted = film_key.startswith("TMDB-") and pm.status == "HIGH_CONFIDENCE"
        film_id = film_key

        local_aliases: set[str] = set()
        all_files: list[str] = []
        for g in members:
            local_aliases.update(g.aliases)
            all_files.extend(g.files)

        canonical = pm.canonical_title if accepted else primary.title
        original_thai = pm.original_title if accepted and pm.original_language == "th" else ""
        record_year = str(extract_year(pm.release_date) or primary.year or "") if accepted else primary.year
        alias_external = pm.aliases or [] if accepted else []
        all_alias_norms = {tmdb_norm(x) for x in local_aliases}

        for a in sorted(local_aliases, key=str.casefold):
            alias_rows.append({
                "Film Record ID": film_id,
                "TMDb ID": pm.tmdb_id if accepted else "",
                "Alias": a,
                "Alias Source": "Local filename/title tag",
                "Country": "",
                "Status": "LOCAL",
            })
        if accepted:
            for a in alias_external:
                if tmdb_norm(a["title"]) in all_alias_norms:
                    continue
                alias_rows.append({
                    "Film Record ID": film_id,
                    "TMDb ID": pm.tmdb_id or "",
                    "Alias": a["title"],
                    "Alias Source": a["source"],
                    "Country": a.get("country", ""),
                    "Status": "TMDB",
                })
                all_alias_norms.add(tmdb_norm(a["title"]))

        film_rows.append({
            "Film Record ID": film_id,
            "Local Group IDs": ", ".join(g.local_id for g in members),
            "Canonical Title": canonical,
            "Original Thai Title": original_thai,
            "Year": record_year,
            "TMDb ID": pm.tmdb_id if accepted else "",
            "IMDb ID": pm.imdb_id if accepted else "",
            "Match Status": pm.status if len(members) == 1 else "HIGH_CONFIDENCE / MERGED",
            "Match Confidence": pm.confidence if pm.status != "NOT_RUN" else "",
            "Match Margin": pm.margin if pm.status != "NOT_RUN" else "",
            "Runtime min": pm.runtime_min if accepted else primary.runtime_min,
            "Genres": pm.genres if accepted else "",
            "Director": pm.director if accepted else "",
            "Production Companies": pm.production_companies if accepted else "",
            "Production Countries": pm.production_countries if accepted else "",
            "Files": len(all_files),
            "Aliases": len(all_alias_norms),
            "Paths": "\n".join(all_files),
        })

        for g in members:
            gm = matches[g.local_id]
            for path in g.files:
                file_enrichment[path] = {
                    "film_record_id": film_id,
                    "canonical_title": canonical,
                    "original_thai_title": original_thai,
                    "tmdb_id": pm.tmdb_id if accepted else "",
                    "imdb_id": pm.imdb_id if accepted else "",
                    "tmdb_match_status": gm.status,
                    "tmdb_match_confidence": gm.confidence if gm.status != "NOT_RUN" else "",
                }

    for path, reason in skipped.items():
        file_enrichment[path] = {
            "film_record_id": "",
            "canonical_title": "",
            "original_thai_title": "",
            "tmdb_id": "",
            "imdb_id": "",
            "tmdb_match_status": f"SKIPPED_{reason}",
            "tmdb_match_confidence": "",
        }

    # Review shows local group plus candidate ranking, but never auto-accepts uncertain rows.
    for g in groups:
        m = matches[g.local_id]
        if m.status not in {"REVIEW", "NO_MATCH", "ERROR"}:
            continue
        candidates = m.candidates or []
        if not candidates:
            api_review_rows.append({
                "Local Group ID": g.local_id,
                "Filename Title": g.title,
                "Year": g.year,
                "Local Runtime min": g.runtime_min,
                "Status": m.status,
                "Candidate Rank": "",
                "TMDb ID": m.tmdb_id or "",
                "Candidate Title": m.canonical_title,
                "Original Title": m.original_title,
                "Release Date": m.release_date,
                "Original Language": m.original_language,
                "Candidate Runtime": m.runtime_min or "",
                "Score": m.confidence or "",
                "Title Score": "",
                "Best Alias": "",
                "Year Score": "",
                "Runtime Score": "",
                "Notes": m.notes,
            })
        else:
            for c in candidates:
                api_review_rows.append({
                    "Local Group ID": g.local_id,
                    "Filename Title": g.title,
                    "Year": g.year,
                    "Local Runtime min": g.runtime_min,
                    "Status": m.status,
                    "Candidate Rank": c["rank"],
                    "TMDb ID": c["tmdb_id"] or "",
                    "Candidate Title": c["title"],
                    "Original Title": c["original_title"],
                    "Release Date": c["release_date"],
                    "Original Language": c["original_language"],
                    "Candidate Runtime": c["runtime"] or "",
                    "Score": c["score"],
                    "Title Score": c["title_score"],
                    "Best Alias": c["best_alias"],
                    "Year Score": c["year_score"],
                    "Runtime Score": c["runtime_score"],
                    "Notes": m.notes if c["rank"] == 1 else "",
                })

    return film_rows, alias_rows, api_review_rows, file_enrichment


def write_excel(output: Path, records: list[MediaRecord], exact_rows, possible_rows, capability_rows,
                review_rows=None, rename_rows=None, film_rows=None, alias_rows=None,
                api_review_rows=None, file_enrichment=None):
    try:
        import xlsxwriter
    except ImportError:
        raise RuntimeError(
            "xlsxwriter saknas. Installera med: python -m pip install --user xlsxwriter"
        )

    review_rows = review_rows if review_rows is not None else metadata_review_rows(records)
    rename_rows = rename_rows if rename_rows is not None else rename_candidate_rows(records)
    film_rows = film_rows or []
    alias_rows = alias_rows or []
    api_review_rows = api_review_rows or []
    file_enrichment = file_enrichment or {}

    wb = xlsxwriter.Workbook(str(output))
    wb.set_properties({
        "title": "Thai Media Inventory",
        "subject": "Technical media inventory, filename verification and duplicate scan",
        "comments": "Generated by thai_media_inventory.py",
    })

    header_fmt = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#1F4E78",
        "border": 1, "border_color": "#1F4E78", "align": "center",
        "valign": "vcenter", "text_wrap": True
    })
    cell_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "valign": "vcenter"})
    wrap_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "valign": "vcenter", "text_wrap": True})
    center_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "align": "center", "valign": "vcenter"})
    num_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "num_format": "0.000", "valign": "vcenter"})
    ok_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "align": "center", "bg_color": "#E2F0D9"})
    legacy_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "align": "center", "bg_color": "#FCE4D6"})
    review_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "align": "center", "bg_color": "#FFF2CC"})
    mismatch_fmt = wb.add_format({"border": 1, "border_color": "#D6E2EA", "align": "center", "bg_color": "#F4CCCC"})

    status_keys = {
        "filename_standard", "resolution_match", "audio_match", "channels_match",
        "audio_features_match", "video_match", "metadata_check", "rename_status", "tmdb_match_status"
    }

    inventory_cols = [
        ("Filename", "filename"), ("Path", "path"),
        ("Film Record ID", "film_record_id"), ("Canonical Title", "canonical_title"),
        ("Original Thai Title", "original_thai_title"), ("TMDb ID", "tmdb_id"),
        ("IMDb ID", "imdb_id"), ("TMDb Match", "tmdb_match_status"),
        ("Match Confidence", "tmdb_match_confidence"),
        ("Filename Title", "filename_title"), ("Year", "filename_year"),
        ("Streaming Source", "streaming_source"), ("Release Type", "release_type"),
        ("Variant", "filename_variant"),
        ("Name Resolution", "name_resolution"), ("Actual Resolution", "resolution"),
        ("Actual Res Class", "actual_resolution_class"), ("Resolution Match", "resolution_match"),
        ("Name Audio", "name_audio"), ("Actual Audio", "primary_audio_codec"), ("Audio Match", "audio_match"),
        ("Name Channels", "name_channels"), ("Actual Channels", "primary_audio_channels"),
        ("Channels Match", "channels_match"),
        ("Name Audio Feature", "name_audio_features"), ("Actual Audio Feature", "actual_audio_features"),
        ("Audio Feature Match", "audio_features_match"),
        ("Name Video", "name_video"), ("Actual Video", "video_codec"), ("Video Match", "video_match"),
        ("Encoder Tag", "video_encoder"),
        ("Filename Standard", "filename_standard"), ("Metadata Check", "metadata_check"),
        ("Parser Notes", "filename_parse_notes"),
        ("Rename Candidate", "rename_candidate"), ("Rename Status", "rename_status"), ("Rename Target", "rename_target"),
        ("Size GiB", "size_gib"), ("Duration", "duration_text"), ("Container", "container"),
        ("Computed kbps", "computed_bitrate_kbps"), ("Overall kbps", "overall_bitrate_kbps"),
        ("Profile", "video_profile"), ("FPS", "fps"), ("Pixel Format", "pixel_format"),
        ("Bit Depth", "bit_depth"), ("HDR", "hdr"), ("Video kbps", "video_bitrate_kbps"),
        ("Audio Tracks", "audio_tracks"), ("Audio", "audio_summary"),
        ("Max Audio Ch", "max_audio_channels"), ("5.1+", "has_51"),
        ("Subtitle Tracks", "subtitle_tracks"), ("Subtitles", "subtitle_summary"),
        ("English Subs", "has_english_sub"), ("Title Tag", "title_tag"),
        ("Normalized Title", "normalized_title"),
        ("Quick Hash", "quick_hash"), ("SHA256", "sha256"),
        ("Modified", "modified"), ("ffprobe OK", "ffprobe_ok"), ("Error", "error"),
    ]

    ws = wb.add_worksheet("Inventory")
    ws.freeze_panes(1, 3)
    if records:
        ws.autofilter(0, 0, len(records), len(inventory_cols)-1)
    ws.set_row(0, 34)

    widths = {
        "Filename": 48, "Path": 90, "Film Record ID": 20, "Canonical Title": 34,
        "Original Thai Title": 30, "TMDb ID": 10, "IMDb ID": 14, "TMDb Match": 18,
        "Match Confidence": 14, "Filename Title": 34, "Year": 8,
        "Streaming Source": 14, "Release Type": 12, "Variant": 12,
        "Name Resolution": 13, "Actual Resolution": 14, "Actual Res Class": 14, "Resolution Match": 15,
        "Name Audio": 12, "Actual Audio": 12, "Audio Match": 12,
        "Name Channels": 13, "Actual Channels": 14, "Channels Match": 14,
        "Name Audio Feature": 16, "Actual Audio Feature": 16, "Audio Feature Match": 16,
        "Name Video": 12, "Actual Video": 12, "Video Match": 12, "Encoder Tag": 24,
        "Filename Standard": 16, "Metadata Check": 15, "Parser Notes": 42,
        "Rename Candidate": 16, "Rename Status": 14, "Rename Target": 48,
        "Size GiB": 10, "Duration": 10, "Container": 16,
        "Computed kbps": 13, "Overall kbps": 12, "Profile": 16,
        "FPS": 9, "Pixel Format": 14, "Bit Depth": 9, "HDR": 14,
        "Video kbps": 11, "Audio Tracks": 11, "Audio": 55, "Max Audio Ch": 12,
        "5.1+": 8, "Subtitle Tracks": 12, "Subtitles": 55, "English Subs": 12,
        "Title Tag": 28, "Normalized Title": 36,
        "Quick Hash": 20, "SHA256": 20, "Modified": 19, "ffprobe OK": 10, "Error": 40,
    }

    def status_format(value: str):
        return {
            "OK": ok_fmt,
            "SAFE": ok_fmt,
            "LEGACY": legacy_fmt,
            "REVIEW": review_fmt,
            "UNPARSED": review_fmt,
            "MISMATCH": mismatch_fmt,
            "HIGH_CONFIDENCE": ok_fmt,
            "HIGH_CONFIDENCE / MERGED": ok_fmt,
            "NOT_RUN": center_fmt,
            "NO_MATCH": review_fmt,
            "ERROR": mismatch_fmt,
        }.get(str(value).upper(), center_fmt)

    for c, (label, key) in enumerate(inventory_cols):
        ws.write(0, c, label, header_fmt)
        ws.set_column(c, c, widths.get(label, 14))

    for r_idx, rec in enumerate(records, 1):
        d = asdict(rec)
        d.update(file_enrichment.get(rec.path, {}))
        for c, (label, key) in enumerate(inventory_cols):
            v = d.get(key, "")
            fmt = cell_fmt
            if key in {"audio_summary", "subtitle_summary", "path", "error", "filename_parse_notes"}:
                fmt = wrap_fmt
            elif key in status_keys:
                fmt = status_format(v)
            elif key in {
                "has_51", "has_english_sub", "ffprobe_ok", "resolution", "actual_resolution_class",
                "video_codec", "hdr", "rename_candidate"
            }:
                fmt = center_fmt
            if isinstance(v, bool):
                v = "YES" if v else ""
            if key == "size_gib" and isinstance(v, (int, float)):
                fmt = num_fmt
            ws.write(r_idx, c, v if v is not None else "", fmt)

    def simple_sheet(name: str, rows: list[dict], widths_override=None):
        sws = wb.add_worksheet(name)
        if not rows:
            sws.write(0, 0, "No rows", header_fmt)
            return
        headers = list(rows[0].keys())
        sws.freeze_panes(1, 0)
        sws.autofilter(0, 0, len(rows), len(headers)-1)
        sws.set_row(0, 34)
        for c, h in enumerate(headers):
            sws.write(0, c, h, header_fmt)
            width = (widths_override or {}).get(h, 18)
            sws.set_column(c, c, width)
        for rr, row in enumerate(rows, 1):
            for c, h in enumerate(headers):
                v = row.get(h, "")
                if h in {"Path", "Paths", "Files", "Parser Notes", "Error", "Notes"}:
                    fmt = wrap_fmt
                elif h in {
                    "Status", "Filename Standard", "Resolution Match", "Audio Match", "Channels Match",
                    "Audio Feature Match", "Video Match", "Metadata Check", "Rename Status", "Match Status", "Status"
                }:
                    fmt = status_format(v)
                else:
                    fmt = cell_fmt
                sws.write(rr, c, v if v is not None else "", fmt)

    simple_sheet("Metadata Review", review_rows, {
        "Filename": 48, "Filename Title": 34, "Path": 90, "Parser Notes": 42, "Error": 40,
        "Actual Resolution": 14, "Actual Res Class": 14,
    })
    simple_sheet("Rename Candidates", rename_rows, {
        "Current Filename": 52, "Proposed Filename": 52, "Encoder Tag": 28, "Path": 90,
    })
    simple_sheet("Exact Duplicates", exact_rows, {"Path": 90, "SHA256": 22})
    simple_sheet("Possible Same Film", possible_rows, {"Path": 90, "Normalized Title": 38, "Variant": 12, "Exact SHA256": 22})
    simple_sheet("Multi-Version Summary", capability_rows, {"Normalized Title": 38, "Variants": 18, "Files": 100})
    simple_sheet("Film Records", film_rows, {
        "Film Record ID": 20, "Local Group IDs": 38, "Canonical Title": 36, "Original Thai Title": 32,
        "Director": 28, "Production Companies": 50, "Production Countries": 22, "Paths": 100,
    })
    simple_sheet("Title Aliases", alias_rows, {
        "Film Record ID": 20, "Alias": 42, "Alias Source": 26, "Country": 10, "Status": 12,
    })
    simple_sheet("TMDb Review", api_review_rows, {
        "Local Group ID": 20, "Filename Title": 38, "Candidate Title": 38, "Original Title": 34,
        "Best Alias": 38, "Notes": 55,
    })

    sws = wb.add_worksheet("Summary")
    summary = [
        ("Generated", datetime.now().isoformat(sep=" ", timespec="seconds")),
        ("Files scanned", len(records)),
        ("ffprobe errors", sum(1 for r in records if not r.ffprobe_ok)),
        ("Filename OK", sum(1 for r in records if r.filename_standard == "OK")),
        ("Filename legacy", sum(1 for r in records if r.filename_standard == "LEGACY")),
        ("Filename review/unparsed", sum(1 for r in records if r.filename_standard in {"REVIEW", "UNPARSED"})),
        ("Metadata OK", sum(1 for r in records if r.metadata_check == "OK")),
        ("Metadata legacy", sum(1 for r in records if r.metadata_check == "LEGACY")),
        ("Metadata mismatch", sum(1 for r in records if r.metadata_check == "MISMATCH")),
        ("Metadata review", sum(1 for r in records if r.metadata_check == "REVIEW")),
        ("x264 rename candidates", len(rename_rows)),
        ("Film Records", len(film_rows)),
        ("Title Alias rows", len(alias_rows)),
        ("TMDb review rows", len(api_review_rows)),
        ("TMDb high-confidence files", sum(1 for v in file_enrichment.values() if v.get("tmdb_match_status") == "HIGH_CONFIDENCE")),
        ("Files with English subs", sum(1 for r in records if r.has_english_sub)),
        ("Files with 5.1+ audio", sum(1 for r in records if r.has_51)),
        ("Files actual 1080-class+", sum(1 for r in records if (r.width or 0) >= 1600 or (r.height or 0) >= 900)),
        ("Files actual 2160-class+", sum(1 for r in records if (r.width or 0) >= 3000 or (r.height or 0) >= 1600)),
        ("Exact duplicate rows", len(exact_rows)),
        ("Possible same-film rows", len(possible_rows)),
        ("Multi-version title groups", len(capability_rows)),
    ]
    sws.set_column(0, 0, 30)
    sws.set_column(1, 1, 24)
    for r, (k, v) in enumerate(summary):
        sws.write(r, 0, k, header_fmt if r == 0 else cell_fmt)
        sws.write(r, 1, v, header_fmt if r == 0 else cell_fmt)

    wb.close()


def read_ini(path: Path):
    cfg = configparser.ConfigParser()
    if not path.exists():
        return {}
    cfg.read(path, encoding="utf-8-sig")
    sec = cfg["scanner"] if cfg.has_section("scanner") else {}
    roots = []
    if sec:
        raw = sec.get("roots", "")
        for line in raw.splitlines():
            line = line.strip().strip('"')
            if line:
                roots.append(line)
    tmdb = cfg["tmdb"] if cfg.has_section("tmdb") else {}
    return {
        "roots": roots,
        "output": sec.get("output", "") if sec else "",
        "ffprobe": sec.get("ffprobe", "") if sec else "",
        "hash_mode": sec.get("hash_mode", "") if sec else "",
        "tmdb_enabled": str(tmdb.get("enabled", "false")).strip().lower() in {"1", "true", "yes", "on"} if tmdb else False,
        "tmdb_access_token": tmdb.get("access_token", "") if tmdb else "",
        "tmdb_api_key": tmdb.get("api_key", "") if tmdb else "",
        "tmdb_language": tmdb.get("language", "en-US") if tmdb else "en-US",
        "tmdb_cache": tmdb.get("cache", "") if tmdb else "",
        "tmdb_cache_ttl_days": tmdb.get("cache_ttl_days", "30") if tmdb else "30",
        "tmdb_timeout": tmdb.get("timeout", "20") if tmdb else "20",
        "tmdb_max_candidates": tmdb.get("max_candidates", "3") if tmdb else "3",
        "tmdb_auto_match_score": tmdb.get("auto_match_score", "90") if tmdb else "90",
        "tmdb_auto_match_margin": tmdb.get("auto_match_margin", "8") if tmdb else "8",
    }


def parse_args():
    p = argparse.ArgumentParser(description="Scan Thai media files and build an Excel inventory.")
    p.add_argument("roots", nargs="*", help="One or more root folders/files to scan.")
    p.add_argument("--config", default=str(script_dir() / "thai_media_inventory.ini"),
                   help="INI config path. Default: next to script.")
    p.add_argument("--output", default="", help="Output XLSX path.")
    p.add_argument("--cache", default=str(script_dir() / ".thai_media_inventory.sqlite"),
                   help="SQLite cache path.")
    p.add_argument("--ffprobe", default="", help="Explicit ffprobe path.")
    p.add_argument("--hash", choices=["candidates", "all", "none"], default="",
                   help="Duplicate hashing mode. Default: candidates.")
    p.add_argument("--refresh", action="store_true",
                   help="Ignore cached ffprobe metadata and rescan.")
    p.add_argument("--tmdb", action="store_true",
                   help="Enable TMDb matching for movie Film Records (requires credentials).")
    p.add_argument("--no-tmdb", action="store_true",
                   help="Disable TMDb matching even if enabled in INI.")
    return p.parse_args()


def main():
    args = parse_args()
    ini = read_ini(Path(args.config))

    roots_raw = args.roots or ini.get("roots") or []
    if not roots_raw:
        print(
            "Inga media-rötter angivna.\n"
            "Ange dem på kommandoraden eller i thai_media_inventory.ini.\n"
            'Exempel: python thai_media_inventory.py "D:\\ThaiMovies" "E:\\ThaiArchive"',
            file=sys.stderr
        )
        return 2

    roots = [Path(os.path.expandvars(os.path.expanduser(x))) for x in roots_raw]
    ffprobe = find_ffprobe(args.ffprobe or ini.get("ffprobe") or None)
    hash_mode = args.hash or ini.get("hash_mode") or "candidates"

    tmdb_enabled = (args.tmdb or bool(ini.get("tmdb_enabled"))) and not args.no_tmdb
    tmdb_access_token = (os.environ.get("TMDB_ACCESS_TOKEN") or ini.get("tmdb_access_token") or "").strip()
    tmdb_api_key = (os.environ.get("TMDB_API_KEY") or ini.get("tmdb_api_key") or "").strip()
    tmdb_language = ini.get("tmdb_language") or "en-US"
    tmdb_cache_raw = ini.get("tmdb_cache") or ""
    tmdb_cache_path = Path(os.path.expandvars(os.path.expanduser(tmdb_cache_raw))) if tmdb_cache_raw else (script_dir() / ".thai_media_tmdb.sqlite")

    if args.output:
        output = Path(args.output)
    elif ini.get("output"):
        output = Path(os.path.expandvars(os.path.expanduser(ini["output"])))
    else:
        output = script_dir() / f"Thai_Media_Inventory_{datetime.now():%Y-%m-%d}.xlsx"

    files = discover_files(roots)
    print(f"Found {len(files)} media files.")
    print(f"ffprobe: {ffprobe}")
    print(f"Cache: {args.cache}")
    print(f"Output: {output}")
    print(f"Hash mode: {hash_mode}")
    if tmdb_enabled:
        auth_state = "configured" if (tmdb_access_token or tmdb_api_key) else "MISSING CREDENTIALS"
        print(f"TMDb: enabled ({auth_state}), language={tmdb_language}, cache={tmdb_cache_path}")
    else:
        print("TMDb: disabled (local Film Records / aliases will still be built)")

    cache = CacheDB(Path(args.cache))
    records: list[MediaRecord] = []
    cache_hits = 0

    try:
        for idx, (path, root) in enumerate(files, 1):
            rec = None if args.refresh else cache.get(path)
            if rec is not None:
                cache_hits += 1
            else:
                try:
                    data = probe_media(ffprobe, path)
                    rec = make_record(path, root, data)
                except Exception as e:
                    rec = make_record(path, root, None, error=str(e))
                cache.put(path, rec)

            records.append(rec)
            if idx % 25 == 0 or idx == len(files):
                print(f"[{idx}/{len(files)}] scanned  (cache hits: {cache_hits})")
    finally:
        cache.close()

    print("Duplicate hashing...")
    apply_candidate_hashing(records, hash_mode)

    # Store hashes back to cache.
    cache = CacheDB(Path(args.cache))
    try:
        for rec in records:
            cache.put(Path(rec.path), rec)
    finally:
        cache.close()

    exact = exact_duplicate_rows(records)
    possible = possible_same_film_groups(records)
    capabilities = best_capabilities(records)
    reviews = metadata_review_rows(records)
    renames = rename_candidate_rows(records)

    tmdb_client = None
    try:
        if tmdb_enabled and (tmdb_access_token or tmdb_api_key):
            tmdb_client = TmdbClient(
                access_token=tmdb_access_token,
                api_key=tmdb_api_key,
                language=tmdb_language,
                cache_path=tmdb_cache_path,
                cache_ttl_days=to_int(ini.get("tmdb_cache_ttl_days")) or 30,
                timeout=to_int(ini.get("tmdb_timeout")) or 20,
                max_candidates=to_int(ini.get("tmdb_max_candidates")) or 3,
                auto_match_score=to_float(ini.get("tmdb_auto_match_score")) or 90.0,
                auto_match_margin=to_float(ini.get("tmdb_auto_match_margin")) or 8.0,
            )
        elif tmdb_enabled:
            print("WARNING: TMDb enabled but no TMDB_ACCESS_TOKEN / TMDB_API_KEY or INI credential was provided.", file=sys.stderr)

        film_rows, alias_rows, api_reviews, file_enrichment = build_film_layer(records, tmdb_client)
    finally:
        if tmdb_client:
            tmdb_client.close()

    write_excel(
        output, records, exact, possible, capabilities, reviews, renames,
        film_rows, alias_rows, api_reviews, file_enrichment
    )

    print()
    print("DONE")
    print(f"Files: {len(records)}")
    print(f"ffprobe errors: {sum(1 for r in records if not r.ffprobe_ok)}")
    print(f"Exact duplicate rows: {len(exact)}")
    print(f"Possible same-film rows: {len(possible)}")
    print(f"Multi-version groups: {len(capabilities)}")
    print(f"Metadata review rows: {len(reviews)}")
    print(f"x264 rename candidates: {len(renames)}")
    print(f"Film Records: {len(film_rows)}")
    print(f"Title Alias rows: {len(alias_rows)}")
    print(f"TMDb review rows: {len(api_reviews)}")
    print(f"Excel: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())