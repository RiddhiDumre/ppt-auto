# BombayDC Presentation Automation — Master Design System & Layout Rules

This document serves as the complete, authoritative specification for generating, styling, and re-rendering BombayDC pitch decks and sales presentations. All presentation plans (`*_plan.json`), automation engines (`generate_ppt.js`), and LLM agents MUST strictly adhere to these rules without exception.

---

## 1. Canvas & Layout Dimensions (16:9 Widescreen)

All measurements are based on a 16:9 HD presentation canvas.

| Dimension / Coordinate | Points (pt) | Inches (in) | Purpose & Usage |
| :--- | :--- | :--- | :--- |
| **Canvas Width** | 960 pt | `10.0 in` | Standard widescreen canvas width |
| **Canvas Height** | 540 pt | `5.625 in` | Standard widescreen canvas height |
| **Far Left Margin (`xIndex`)** | 15.36 pt | `0.16 in` | Grid alignment line for slide numbers & horizontal divider lines |
| **Primary Content X (`contentX`)** | 191.04 pt | `1.99 in` | Left margin for slide titles, taglines, & column 1 content |
| **Secondary Content X (`secondaryX`)** | 571.2 pt | `5.95 in` | Left margin for column 2 / description area in row lists |
| **Vertical Divider Line X (`dividerX`)** | 561.6 pt | `5.85 in` | Center vertical rule coordinate for 2-column layouts |
| **Cover Margin X (`coverX`)** | 48 pt | `0.50 in` | Left margin for cover slide title & subtitle |
| **Header Banner Top Y** | 0 pt | `0.0 in` | Top status bar / header image Y coordinate |
| **Header Banner Height** | 54.72 pt | `0.57 in` | Top status bar / header image height |

---

## 2. Fixed Header & Alignment Rules

> [!IMPORTANT]
> **CRITICAL RULE (Fixed Horizontal Divider Line & Content Y Position)**
> The horizontal header divider line MUST ALWAYS be placed at **`y: 1.35 in`** and the content start Y position MUST ALWAYS be **`y: 1.60 in`** (`contentY`), regardless of whether a tagline/subline is present or absent on the slide. The line and content area must NEVER jump up or down between slides!

### Visual Grid Architecture

```
┌──────────────────────────────────────────────────────────────────────────┐
│ [0.16"] Slide Index (y: 0.57", fs: 18.8pt)   [1.99"] SLIDE TITLE (y: 0.57", fs: 18.8pt, BOLD, UPPERCASE) │
│                                             [1.99"] Tagline (y: 0.92", fs: 8.55pt, sentence case)     │
├──────────────────────────────────────────────────────────────────────────┤  <-- Line Y = 1.35" FIXED
│                                                                          │  <-- Content Y = 1.60" FIXED
│                     SLIDE CONTENT REGION (y: 1.60" to 5.25")             │
└──────────────────────────────────────────────────────────────────────────┘
```

### Exact Coordinates & Specifications
- **Slide Index Number**: Placed at `x: 0.16 in`, `y: 0.57 in`, `w: 1.68 in`, `h: 0.35 in`, `fontSize: 18.8pt`, `color: colors.body`, `fontFace: Inter`, `bold: false`, `valign: top`. Displays slide section numbers (e.g. `1.1`, `1.2`, ..., `1.10`).
- **Slide Title**: Placed at `x: 1.99 in`, `y: 0.57 in`, `w: 7.85 in`, `h: 0.35 in`, `fontSize: 18.8pt`, `color: colors.title`, `fontFace: Inter`, `bold: true`, `valign: top`. **MUST ALWAYS BE UPPERCASE AND BOLD**.
- **Slide Tagline (if present)**: Placed at `x: 1.99 in`, `y: 0.92 in`, `w: 7.65 in`, `h: 0.35 in`, `fontSize: 8.55pt`, `color: colors.body`, `fontFace: Inter`, `bold: false`, `lineSpacing: "140%"`, `valign: top`, **Sentence Case**.
- **Horizontal Divider Rule Line**: Placed at `x: 0.16 in`, `y: 1.35 in` (FIXED), `w: 9.68 in`, `h: 0 in`, `lineWidth: 0.5pt`, `transparency: 40%`.
- **Content Start Position (`contentY`)**: **`1.60 in`** (FIXED on all content slides).

---

## 3. Master Typography Hierarchy & Casing Rules

| Element | Font Family | Size | Weight | Color | Casing | Line Spacing | Notes |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Cover Title** | Inter | `29.0 pt` | Bold | Theme Title (`#ECE9E4` / `#1D1D1F`) | **UPPERCASE** | 110% | Top margin `y: 2.1 in` |
| **Cover Subtitle** | Inter | `8.55 pt` | Regular | Theme Body (`#FFFFFF` / `#6A6A6B`) | Sentence Case | 150% | Placed directly below cover title |
| **Section Starter Num** | Inter | `72.0 pt` | Bold | Theme Accent Green | N/A | 100% | Giant section index |
| **Section Starter Title**| Inter | `24.0 pt` | Bold | Theme Title | **UPPERCASE** | 120% | Large section title |
| **Slide Index Num** | Inter | `18.8 pt` | Regular | Theme Body | N/A | 100% | Left aligned at `x: 0.16 in` |
| **Slide Title** | Inter | `18.8 pt` | **Bold** | Theme Title | **UPPERCASE** | 110% | Left aligned at `x: 1.99 in` |
| **Slide Tagline** | Inter | `8.55 pt` | Regular | Theme Body | Sentence Case | 140% | Sub-header text below title |
| **Header (Non-Bullet)**| Inter | `9.5 pt` | **Bold** | Accent Green (`#034E48` / `#4DB89A`) | **UPPERCASE** | 125%–130% | Any non-bullet single-line text |
| **Table / Section Label**| Inter | `7.5 pt` | **Bold** | Theme Body | **UPPERCASE** (charSpacing: 1.5) | 100% | Column headers (e.g. CATEGORY / DESCRIPTION) |
| **Body Bullet / Text** | Inter | `8.55 pt` | Regular | Theme Body (`#6A6A6B` / `#B4B4B4`) | Sentence Case | **145%** | Standard bullet body text |
| **Primary Narrative Text**| Inter | `9.5 pt` | Regular | Theme Title / Body | Sentence Case | 145% | Headline/narrative emphasis text |
| **Card Title (Grid)** | Inter | `8.0 pt` | **Bold** | `#FAFAF7` | **UPPERCASE** | 120% | Center aligned inside dark green pill |
| **Card Body (Grid)** | Inter | `8.0 pt` | Regular | Theme Body | Sentence Case | 140% | Text inside grid cards |
| **Timeline Num / Time**| Inter | `13.5 pt` | **Bold** | Accent Green | Mixed / Capitalized | 100% | Timeline metric / duration |
| **BEAM Acronym Letter** | Inter | `48.0 pt` | **Bold** | Accent Green | **UPPERCASE** | 100% | Giant acronym letter (B, E, A, M) |

### Category Header Formatting Rule
ONLY explicit short uppercase category titles (e.g., `TRANSACTIONS WORK`, `THEN`, `NOW`, `WHAT THIS FEELS LIKE FOR THE USER`) or lines ending with a colon `:` should be rendered in **Accent Green Bold** (`#034E48` for light theme, `#4DB89A` for dark theme). Regular text lines, body paragraphs, standalone sentences, and taglines MUST be rendered in standard title or body colors (`#1D1D1F`, `#ECE9E4`, `#6A6A6B`, `#B4B4B4`) with regular font weight and casing.

### Bullet Spacing & Indentation Rule
- Bullet points MUST use a tight single-space gap (`"• "`) without excessive PowerPoint default paragraph indentation.
- Standalone sentences that are not part of bullet lists MUST NOT be auto-bulleted.

---

## 4. Theme Color Palettes

The automation engine dynamically applies themes to maintain visual rhythm across presentation decks.

### Dark Theme (`theme: "dark"`)
* **Background (`bg`)**: `#1D1D1F` (Apple Dark Grey)
* **Title Text (`title`)**: `#ECE9E4` (Warm Off-white)
* **Body Text (`body`)**: `#B4B4B4` (Muted Light Grey)
* **Accent Color (`accent`)**: `#4DB89A` (Bright Teal Green for dark contrast)
* **Line Color (`line`)**: `#FFFFFF` (White with 40% opacity)

### Light Theme (`theme: "light"`)
* **Background (`bg`)**: `#ECE9E4` (Warm Light Cream / Linen)
* **Title Text (`title`)**: `#1D1D1F` (Charcoal / Near Black)
* **Body Text (`body`)**: `#6A6A6B` (Mid Grey)
* **Accent Color (`accent`)**: `#034E48` (BombayDC Deep Forest Green)
* **Line Color (`line`)**: `#B4B4B4` (Grey with 40% opacity)

### Green Accent Theme (`theme: "green"`)
* **Background (`bg`)**: `#034E48` (BombayDC Signature Deep Forest Green)
* **Title Text (`title`)**: `#ECE9E4` (Off-white)
* **Body Text (`body`)**: `#B4B4B4` (Muted Grey)
* **Accent Color (`accent`)**: `#4DB89A` (Bright Teal Green)
* **Line Color (`line`)**: `#3E8D86` (Teal Line)

---

## 5. Layout Selection & Mathematical Specifications

### 5.1 TWO_COLUMN Layout
- **Side-by-Side Comparison (Divider Mode)**: Used for direct comparisons (e.g., "Then vs Now", "Problem vs Solution") or multi-column lists (`slide.divider === true`).
  - **Left Column**: `x: 1.99 in`, `y: 1.60 in`, `w: 3.56 in`, `h: 3.75 in`, `fontSize: 9.5pt`, `lineSpacing: "145%"`.
  - **Vertical Center Divider Line**: `x: 5.85 in`, `y: 1.60 in`, `w: 0 in`, `h: 3.70 in`, `lineWidth: 0.5pt`.
  - **Right Column**: `x: 5.95 in`, `y: 1.60 in`, `w: 3.70 in`, `h: 3.75 in`, `fontSize: 9.5pt`, `lineSpacing: "145%"`.
- **Single Column Narrative**: Left headline across top (`x: 1.99 in`, `y: 1.70 in`, `w: 7.65 in`, `h: 0.65 in`), horizontal divider line, then right body below (`y: 2.65 in`).

### 5.2 ROW_LIST Layout
- **Column Header Labels**: `CATEGORY` (`x: 1.99 in`) and `DESCRIPTION` (`x: 5.95 in`), `y: 1.60 in`, `fontSize: 7.5pt`, **UPPERCASE**, `bold: true`, `charSpacing: 1.5`.
- **Row Y Calculation**: `rowH = Math.min(0.95, (5.25 - 1.90) / rowCount)`.
- **Category Titles**: `x: 1.99 in`, `w: 3.4 in`, `fontSize: 9.5pt`, `bold: true`, `color: colors.accent`, **UPPERCASE**.
- **Description Body**: `x: 5.95 in`, `w: 3.7 in`, `fontSize: 8.55pt`, `lineSpacing: "145%"`.

### 5.3 TIMELINE_LIST Layout
- **Row Height Restriction**: `rowH = Math.max(0.68, 3.6 / rowCount)` (Enforces minimum 0.68" height to prevent text overlap or bottom overflow).
- **Column 1 (Metric / Duration)**: `x: 1.99 in`, `w: 1.3 in`, `fontSize: 13.5pt`, `bold: true`, `color: colors.accent`.
- **Column 2 (Phase Title)**: `x: 3.4 in`, `w: 2.1 in`, `fontSize: 9.5pt`, `bold: true`, `color: colors.title`.
- **Column 3 (Description / Bullets)**: `x: 5.6 in`, `w: 4.1 in`, `fontSize: 8.55pt`, `lineSpacing: "145%"`.
- **Single-String Title Fallback**: When `row.title` has no newline, title spans left area (`x: 1.99 in`, `w: 3.4 in`), `fontSize: 9.5pt`, `bold: true`, `color: colors.accent`.
- **Bottom Outro Text**: Placed at `y = rowStartY + rowCount * rowH + 0.12 in`, `fontSize: 8.55pt`, `bold: true`, `color: colors.accent`.

### 5.4 CARD_GRID Layout (Cap: Max 2–3 Per Deck)
> [!WARNING]
> Do NOT use `CARD_GRID` for more than **2 to 3 slides maximum** in any deck. Reserve `CARD_GRID` for critical capability or solution slides and render remaining qualifying lists using alternative layouts (`THREE_COLUMN`, `ACRONYM_GRID`, `TIMELINE_LIST`, `TWO_COLUMN`, `ROW_LIST`).

- **Card Dimensions**:
  - **Adaptive Card Header Pill Height**: For short 1–2 word titles (e.g., `ACQUIRE`, `ACTIVATE`, `ENGAGE`), use compact header pill height (`0.35 in`). For longer multi-line titles, use standard height (`0.65 in`).
- **Card Pill Fill**: `#034E48` (Light Theme) or `#4DB89A` (Dark Theme).
- **Card Title**: `fontSize: 8.0pt`, `bold: true`, `color: #FAFAF7`, **UPPERCASE**, centered inside pill.
- **Card Body**: `fontSize: 8.0pt`, `color: colors.body`, `lineSpacing: "140%"`.
- **Footer Outro Text**: `fontSize: 8.55pt`, `color: colors.title` (Ensures high legibility against dark backgrounds).

### 5.5 THREE_COLUMN Layout (Default for 3 Offerings / Pillars)
- **Use Case**: ALWAYS use `THREE_COLUMN` layout for slides presenting 3 core offerings, team structures, or response pillars (e.g., "Our response: three simple BFSI offerings").
- **Grid Math**: 3 columns, `gap: 0.35 in`, `colW = (10.0 - 0.16 - 1.99 - 0.70) / 3 = 2.38 in`.
- **Column X Coordinates**: Column 1 (`x: 1.99 in`), Column 2 (`x: 4.72 in`), Column 3 (`x: 7.45 in`).
- **Headers**: `y: 1.65 in`, `w: 2.38 in`, `h: 0.35 in`, `fontSize: 9.5pt`, `bold: true`, `color: colors.accent`, **UPPERCASE**.
- **Body**: `y: 2.03 in`, `w: 2.38 in`, `fontSize: 8.55pt`, `lineSpacing: "145%"`.

### 5.6 ACRONYM_GRID Layout (BEAM Framework & 4-Item Lists)
- **Use Case**: 4-item list layouts requiring prominent letters/numbers (e.g. B-E-A-M framework).
- **Grid Math**: 2x2 grid. Column width `w: 3.5 in`. Row gap `0.25 in`.
- **Big Acronym Letter**: `fontSize: 48.0pt`, `bold: true`, `color: colors.accent`.
- **Title**: `fontSize: 9.5pt`, `bold: true`, `color: colors.accent`, **UPPERCASE**.
- **Underline Divider**: Horizontal rule beneath title.
- **Body**: `fontSize: 8.0pt`, `lineSpacing: "140%"`.

---

## 6. Overlapping Prevention & Defensive Layout Rules

1. **Text Overlap Prevention**:
   - For `SINGLE COLUMN NARRATIVE`, if top headline/text is long (e.g. multi-line body text), scale font size down to `9.5pt`, push the horizontal divider down dynamically, and use standard column bounds.
   - For `TIMELINE_LIST`, strictly cap timeline row heights (`Math.max(0.68, 3.6 / rowCount)`) so text and bottom taglines are never pushed off the slide canvas.
2. **Text-Only Slides**: Do NOT use `TWO_COLUMN` layouts for text-only slides unless it is an explicit comparison (e.g. "Then vs Now") or `divider === true`. Use `SINGLE COLUMN NARRATIVE` or body text layouts instead.
3. **Dynamic Theme Rhythm**: Alternately assign `theme: "dark"` to key problem statements, strategic shifts, team structure, and core frameworks to break visual monotony and create an engaging presentation flow.

---

## 7. General Presentation Structure Rules

1. **No Index / Agenda Slides**: Never generate a "Today's Presentation" or Index slide. The cover slide MUST lead directly into content slide 1.1.
2. **Cover Slide**: Displays title (`29pt` Bold UPPERCASE) and subtitle (`8.55pt` Regular Sentence Case) matching the reference plans. Do NOT add extraneous shapes or text.
3. **Closing Slide**: Every presentation deck MUST end with the green `LET'S BUILD WHAT'S NEXT` contact slide (`xEnd` of dividing line = `9.84 in`).
4. **Slide Section Index Numbering**: Display section slide index numbers (`1.1`, `1.2`, ..., `1.10`) on the far left column (`x: 0.16 in`, `y: 0.57 in`) of all content slides.
