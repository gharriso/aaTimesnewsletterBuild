# AATimes Newsletter Builder

Automatically generates the weekly AATimes newsletter `.docx` by scraping event and meeting data from [aatimes.org.au](https://aatimes.org.au).

## Requirements

- Python 3.x (standard library only — no pip installs needed)
- `sampleNewsletter.docx` — the template file (included in this repo)

## Usage

Run once each week, any time before Monday:

```
python3 build_newsletter.py
```

The script will:

1. Calculate the upcoming Monday's date automatically
2. Scrape all events from aatimes.org.au/events/
3. Scrape new and changed meetings from aatimes.org.au/meetings/
4. Download flyer images for events that have them (up to 6)
5. Build a formatted two-page events section followed by a flyer image page
6. Update the date in the newsletter header
7. Write output to `aatimesYYYYMMDD.docx`

No manual editing is required.

## Output

- **Pages 1–2**: Upcoming events table, covering roughly two A4 pages
  - Month banner rows when the month changes
  - Each event shows date, time, title, description, and venue
  - Events with a flyer image show "See Next Page" in red
- **New Meetings** and **Recently Changed Meetings** sections follow the events
- **Page 3+**: Flyer images arranged in a 3-column grid

## Files

| File | Purpose |
|------|---------|
| `build_newsletter.py` | Main script — run this each week |
| `sampleNewsletter.docx` | Template — provides all styles, fonts, and headers |
| `aatimesYYYYMMDD.docx` | Generated output (not committed) |

## Tuning

One constant at the top of `build_newsletter.py` controls how many events are included:

```python
MAX_EVENT_LINES = 220   # approx 2 pages of 2-column A4 at 9pt
```

Increase this to include more events, decrease it to include fewer.

## How it works

The script modifies `sampleNewsletter.docx` (a zip of XML files) in memory:

- Replaces `word/document.xml` with a freshly built events table
- Updates `word/header2.xml` and `word/header3.xml` with the new date
- Adds downloaded flyer images to `word/media/` and registers them in the relationships file
- Writes everything out as a new zip with the dated filename

The document has two sections:
- **Section 1** (pages 1–2): Two-column layout for the events table
- **Section 2** (page 3+): Single-column layout for flyer images as floating anchors
