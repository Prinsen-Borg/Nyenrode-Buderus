# DESIGN.md: Visual rules for the Buderus case

The case is about a fictional company and is presented at Nyenrode, so this repository uses a **neutral, sober style**. It deliberately does not use any employer's brand (no Moba colours or fonts) and does not imitate the real Buderus or Bosch branding.

## 1. Colour tokens

```css
:root {
  --ink: #18242f;      /* headings */
  --body: #3d4b57;     /* body text */
  --muted: #6b7a86;    /* captions, meta */
  --line: #dde3e8;     /* hairlines, table rules */
  --canvas: #f4f6f8;   /* page background */
  --surface: #ffffff;  /* cards */
  --navy: #1f3a5f;     /* primary: header band, table heads, numbers */
  --navy-2: #2d5487;   /* eyebrows, links */
  --tint: #e9eff6;     /* key-message callouts */
  --flag: #b8741a;     /* open points, things that shift */
  --flag-bg: #fbf1e3;
  --ok: #2f7a4f;       /* status: done */
  --ok-bg: #e6f3eb;
}
```

- Navy carries structure; amber (`--flag`) marks only what the group still has to decide or what changes. Green and red appear only in status pills.
- Light only. Set `color-scheme: light` and an explicit body background.

## 2. Type

- System sans: `"Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif`. No web fonts, so pages work offline and in mail attachments.
- Headings 650 weight; body 16px / 1.6.
- Section rhythm: small uppercase eyebrow (`§ number · topic`), then a heading that states the conclusion, then content.

## 3. Components

| Component | Rule |
|---|---|
| Card | White, 1px `--line` border, 10px radius, no shadow |
| Table | Navy header row, hairline rows; becomes stacked cards under 760px (use `data-l` labels) |
| Callout "key" | `--tint` background for the central message or core tension |
| Callout "flag" | `--flag-bg` for a weakness or tension we must address |
| Open point | Starts with **Open:** in `--flag` colour |
| Decisions | Numbered list with navy circles; max. five |

## 4. Slides

- 16:9, white background, navy titles that state the conclusion, max. ~6 bullets per slide.
- One message per slide; the answer (the five decisions) comes first.
- Use the same palette. If Nyenrode provides a template, that template wins.

## 5. Checks before sharing

- No horizontal scroll on a 400px-wide phone.
- Every figure labelled as case, public source or assumption.
- No employer branding, no real people.
