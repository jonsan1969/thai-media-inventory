THAI MEDIA INVENTORY / FILM RECORDS / TMDB MATCHER v3.0
======================================================

v3.0 builds on the validated v2.7 technical scanner. The filename parser,
ffprobe verification, duplicate detection and report-only rename logic remain
in place. This version adds a separate archive identity layer:

    Media File -> Local Film Group -> Film Record -> Title Aliases

NEW SHEETS
----------
Film Records
    One provisional archive record per locally identified movie. If two local
    title groups independently receive the same HIGH_CONFIDENCE TMDb ID, they
    are merged into one Film Record while retaining all local aliases/files.

Title Aliases
    Local filename/title-tag aliases are always retained. After a safe TMDb
    match, TMDb original and alternative titles are added as external aliases.
    External data never replaces local titles.

TMDb Review
    Only uncertain / no-match / API-error cases. Candidate score components
    are shown so questionable matches can be reviewed instead of silently
    accepted.

Inventory additions
    Film Record ID
    Canonical Title
    Original Thai Title
    TMDb ID
    IMDb ID
    TMDb Match
    Match Confidence

LOCAL FILM RECORDS WORK WITHOUT TMDB
------------------------------------
TMDb is disabled by default. The scanner still groups versions by local title
identity and creates Film Records + local aliases. TV episodes (SxxEyy) are
valid files but are deliberately skipped by the v3.0 movie matcher.

TMDB SETUP
----------
Recommended authentication is the TMDb API Read Access Token.

Windows cmd example:
    set TMDB_ACCESS_TOKEN=YOUR_TOKEN_HERE
    python thai_media_inventory_v3_0.py

PowerShell example:
    $env:TMDB_ACCESS_TOKEN="YOUR_TOKEN_HERE"
    python thai_media_inventory_v3_0.py

Then set in thai_media_inventory_v3_0.ini:
    [tmdb]
    enabled = true

A v3 API key can also be supplied through TMDB_API_KEY or api_key= in the INI.
Environment variables take precedence over INI credentials.

TMDB MATCHING MODEL
-------------------
Search input:
    parsed local title + year (year is omitted when unknown)

The best few candidates are enriched with movie details plus:
    alternative titles
    external IDs (including IMDb when present)
    credits

Scoring currently weighs:
    title / alias similarity     55%
    release year                15%
    original language Thai      10%
    production country Thailand 10%
    runtime                     10%

Default HIGH_CONFIDENCE requirements:
    total score >= 90
    lead over runner-up >= 8
    title/alias similarity >= 88
    if local year exists, year must be exact or +/- 1

Anything below those conditions remains REVIEW. No API result renames files,
changes parsed metadata, or overwrites local titles.

TMDB CACHE
----------
TMDb HTTP responses are stored separately in:
    .thai_media_tmdb.sqlite

This is intentionally separate from:
    .thai_media_inventory.sqlite

You can therefore delete the technical ffprobe cache for a fresh media scan
without forcing hundreds of repeated TMDb requests. Default TMDb cache TTL is
30 days.

COMMAND LINE
------------
Enable TMDb for one run:
    python thai_media_inventory_v3_0.py --tmdb

Disable TMDb even if enabled in INI:
    python thai_media_inventory_v3_0.py --no-tmdb

Technical full refresh remains:
    python thai_media_inventory_v3_0.py --refresh

SAFETY
------
The script remains report-only:
    - no deletion
    - no moving
    - no automatic renaming
    - no API data overwrites local metadata

v3.0 is the first Film Record/TMDb implementation. The next logical layer is
manual acceptance/override persistence for REVIEW matches and importing known
archive aliases/IDs (for example from the Five Star master catalog).

TMDB ATTRIBUTION
----------------
This product uses the TMDB API but is not endorsed or certified by TMDB.

Project page:
https://github.com/jonsan1969/thai-media-inventory
