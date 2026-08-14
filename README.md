# BombayDC Presentation Automation Engine (`ppt-auto`)

Automated pitch deck styling, standardization, and zero-collision layout engine for BombayDC presentation decks.

---

## 🚀 How to Get 100% Identical Output on Any Machine

If you are cloning this repository to another laptop or running it on a different machine, follow these instructions to ensure the presentations generate and display **identically**.

### 1. Prerequisites for Users (Human Setup)

#### A. Install the "Inter" Font Family (CRITICAL)
The presentation layout engine calculates exact text bounding boxes based on the **Inter** typography metrics:
- **Font Face**: Inter (Regular), Inter Medium, Inter Bold
- If **Inter** is not installed in your operating system, PowerPoint will fall back to **Calibri** or **Arial**, causing text to wrap differently and look misaligned on screen.
- **How to Install**:
  1. Download the Inter font family free from [Google Fonts (Inter)](https://fonts.google.com/specimen/Inter) or [rsms.me/inter](https://rsms.me/inter/).
  2. Unzip and install `Inter-Regular.ttf`, `Inter-Medium.ttf`, and `Inter-Bold.ttf` for all users on Windows / macOS.
  3. Restart PowerPoint / your editor.

#### B. Install Node.js
- Ensure **Node.js (v18 or newer)** is installed (`node -v`).

---

### 2. Quick Start & Installation

```bash
# 1. Clone the repository
git clone https://github.com/RiddhiDumre/ppt-auto.git
cd ppt-auto

# 2. Install dependencies
npm install

# 3. Build all 11 standardized presentation decks
npm run build
# OR: node build_standardized_decks.js

# 4. Audit collisions and clipping
npm run audit
# OR: node audit_both_dirs.js

# 5. Verify header positions and bounds
npm test
# OR: node test_header_positions.js
```

---

## 🤖 Instructions for AI Assistants & LLMs

When developing, refactoring, or generating presentations in this repository, **ALWAYS** follow these strict invariants:

### 1. Locked Chrome & Header Coordinates (Never Move)
- **Slide Section Index Number**: `x: 0.16 in`, `y: 0.57 in`, `fontSize: 18.8pt`, `fontFace: "Inter Medium"`.
- **Slide Title**: `x: 1.99 in`, `y: 0.57 in`, `w: 7.85 in`, `h: 0.76 in`, `fontSize: 18.8pt`, uppercase `Inter Medium`.
- **Fixed Header Divider Line**: **LOCKED at `y: 1.35 in`** (`x: 0.16 in`, `w: 9.68 in`, `line: { color: "C8C8CC", width: 0.5 }`).
  - *Rule*: The line must NEVER move up or down between slides regardless of subtitle presence.
- **Content Area Start Y (`contentStartY`)**: Fixed at **`y: 1.60 in`** (or `1.95 in` if a 2-line title forces the subline below the divider).

### 2. Thick Card Header Box Design
- **Card Pill / Box Height**: **`h = 0.50 in`** (`CARD_PILL_H = 0.50`).
- **Pill Fill Color**: `#034E48` (Light Theme) or `#1C1C1E` (Dark Theme).
- **Pill Text**: `10.5pt`, `bold: true`, `fontFace: "Inter Medium"`, `color: "FFFFFF"`, `align: "center"`, `valign: "middle"`.
- **Card Body Text Y Offset**: Starts at `textY = pillY + 0.58 in` (`CARD_TEXT_OFFSET = 0.58`).
- **Row 2 Card Pill Y**: `r2Y = Math.min(3.45, Math.max(3.20, r1Y + CARD_PILL_H + maxR1TextH + 0.22))`.
- **Bottom Summary / Footnotes**: Positioned dynamically with `bottomNoteY = Math.min(4.85, Math.max(4.00, r2Y + CARD_TEXT_OFFSET + maxR2TextH + 0.14))` to guarantee zero collision.

### 3. Typography & Color Tokens
| Element | Light Theme | Dark Theme | Size / Style |
|---|---|---|---|
| **Slide Index Number** | `#1A1A1A` | `#FFFFFF` | `18.8pt` Medium |
| **Slide Title** | `#1A1A1A` | `#FFFFFF` | `18.8pt` Medium, UPPERCASE |
| **Subtitle / Subline** | `#5A5A5E` | `#D0D0D4` | `9.5pt` Regular |
| **Category Header** | `#034E48` | `#4DB89A` | `10.5pt` Bold |
| **Card Header Box** | `#034E48` | `#1C1C1E` | Box `h: 0.50in`, Text `10.5pt` White |
| **Body Text / Bullets** | `#2C2C2E` | `#A0A0A6` | `8.5pt` - `9.0pt` Regular |
| **Header Divider Line** | `#B0B0B4` | `#555558` | `y: 1.35in`, `width: 0.5pt` |
| **Body Divider Line** | `#D4D4D8` | `#38383C` | `width: 0.4pt` |

### 4. Bounding Box & Overlap Prevention Rules
- **Zero Collision Guarantee**: Every shape must maintain at least `0.10 in` vertical clearance (`minGap = 0.10`).
- **Bottom Slide Margin Floor**: Content must never extend past `y: 5.08 in`.
- **Downward-Only Collision Flow**: `resolveVerticalCollision` must only push elements downward; never pull elements upward.

---

## 📁 Repository Structure

```
├── build_standardized_decks.js    # Master PPT standardization and generation engine
├── audit_both_dirs.js             # Automated bounding-box collision detection auditor
├── test_header_positions.js       # Header invariant and boundary verification tests
├── deck_bg_mappings.json          # Slide background image mappings
├── cover_bg_map.json              # Cover background image mappings
├── BDC Deck (Copy).pptx           # Master closing slide reference deck
├── package.json                   # Project metadata and dependencies
├── .agents/
│   └── AGENTS.md                  # Strict AI pair-programming design rules
└── SALES DECKS/
    └── new MD/                    # Output directory for styled presentation decks
```

---

## 🔍 Verification & QA Checklist

Before committing or delivering any deck updates, run:

1. `node build_standardized_decks.js` -> Rebuilds all 11 decks in both output directories.
2. `node audit_both_dirs.js` -> Must report **`0 Total Collisions in 11 decks`**.
3. `node test_header_positions.js` -> Must report **`✅ PASS`** on Header Invariants and Content Bounds.
