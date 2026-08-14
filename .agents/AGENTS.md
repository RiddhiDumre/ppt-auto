# Presentation Design Rules (BombayDC)

When generating or formatting presentation plans (`*_plan.json`) and updating the engine (`generate_ppt.js`), ALWAYS follow these strict design rules requested by the user:

## 1. Content & Alignment Guidelines
- **Fixed Header Line & Content Y Position**: The horizontal header divider line MUST ALWAYS be placed at **`y: 1.35 in`** and content start Y position MUST ALWAYS be **`y: 1.60 in`** (`contentY`), regardless of whether a tagline/subline is present or absent on the slide. The line must NEVER jump up or down between slides!
- **Slide Title Casing & Alignment**: All slide titles MUST be formatted as **UPPERCASE, BOLD (18.8pt)** placed at `x: 1.99 in`, `y: 0.57 in`, `w: 7.85 in`.
- **Slide Index Number**: Placed at `x: 0.16 in`, `y: 0.57 in`, `fontSize: 18.8pt`, matching title top Y alignment.
- **Header Formatting**: ONLY explicit short uppercase category titles (e.g. `TRANSACTIONS WORK`, `THEN`, `NOW`) or lines ending with colons should be formatted as uppercase bold in accent green color. Regular text lines, standalone sentences, and taglines MUST use standard body/title colors with normal casing.
- **Bullet Spacing**: Bullet points must use a tight single-space gap (`"• "`) without excessive PowerPoint default paragraph indentation. Standalone sentences that are not part of bullet lists must NOT be auto-bulleted.
- **Overlapping Prevention**: Text must NEVER overlap lines, shapes, or other text blocks.
  - For `SINGLE COLUMN NARRATIVE`, if the top text is long (e.g. multi-line bullets), it must be scaled down to standard body size (9.5pt), the horizontal divider must be dynamically pushed down, and standard column formatting must be used.
  - For `TIMELINE_LIST`, restrict the timeline row heights (`Math.max(0.68, 3.6 / rowCount)`) to ensure text and bottom taglines do not get pushed off the frame.

## 2. Layout Selection Rules & Visual Diversity
- **Card Grid Cap (Max 2–3 per deck)**: Do NOT use `CARD_GRID` (the dark green rectangular card design) for more than **2 to 3 slides maximum** in any deck. If a deck has more qualifying lists, reserve `CARD_GRID` for the 2–3 most critical capability/solution slides and render the rest using alternative layouts (`ACRONYM_GRID`, `THREE_COLUMN`, `TIMELINE_LIST`, `TWO_COLUMN`, `ROW_LIST`).
- **3-Column Offering / Response Slides**: For slides presenting 3 core offerings or response pillars (e.g., "Our response: three simple BFSI offerings"), ALWAYS use `THREE_COLUMN` layout with green bold headers and bullet points (`• `) instead of card grids.
- **Text-Only Slides**: Do NOT use `TWO_COLUMN` layouts for text-only slides unless it is an explicit comparison (e.g., "Then vs Now") or has `divider: true`. Otherwise, use `SINGLE COLUMN NARRATIVE` or standard body layouts.
- **Three-Column / Team Layouts**: Use `THREE_COLUMN` layout for team structure, multi-role, or 3-pillar slides. Combine with `theme: "dark"` for strong visual contrast (green small-caps headers with white text bullets).
- **BEAM Framework / 4-item lists**: Use `ACRONYM_GRID` for 4-item list layouts where huge letters/numbers are needed. Be explicit in the JSON `layout: "ACRONYM_GRID"`.
- **Dynamic Theme Rhythm**: Alternately assign `theme: "dark"` to key problem statements, strategic shifts, team structure, and core frameworks to break monotony and create a sleek presentation flow.

## 3. General Deck Structure
- **No Index/Agenda Slides**: Never generate a "Today's Presentation" or Index slide. The cover slide should be immediately followed by the content.
- **Cover Slide**: The cover slide should have exactly the title and subtitle as provided in the reference plans. Do not insert extra elements.
- **Closing Slide**: Every deck must end with the green `LET'S BUILD WHAT'S NEXT` contact slide (`xEnd` of dividing line = `9.84 in`).
- **Numbering**: Display section slide index numbers (`1.1`, `1.2`, ..., `1.10`) on the far left column (`x: 0.16 in`, `y: 0.57 in`) of all content slides.

## 4. Strict Zero-Overlap & Bounding Box Collision Prevention Rule
- **Zero Element Overlap**: Text boxes, headers, subtitles, card pills, lines, and images MUST NEVER overlap each other under any circumstances across all slides and presentation decks.
- **Dynamic Content Flow & Collision Resolution**:
  - **2-Line Titles**: If a title wraps to 2 lines, do NOT display the tagline/subtitle above the divider line. Relocate the tagline to `y: 1.60 in` below the divider line as the first content element, and push all subsequent content down by `+0.35 in`.
  - **Fixed Header Divider**: The horizontal line divider MUST ALWAYS stay anchored at `y: 1.35 in`.
  - **Minimum Vertical Gap**: Maintain a minimum vertical gap of `0.10 in` between adjacent shapes in the same column.
  - **Slide Margin Protection**: Clamp all content shapes so their bottom edge NEVER exceeds `y: 5.08 in` to prevent text cropping or slide edge cut-offs.

## 5. Thick Card Header Box Guidelines
- **Card Pill / Header Box Height**: MUST ALWAYS be **`h = 0.50 in`** (`CARD_PILL_H = 0.50`) with substantial solid visual weight.
- **Card Pill Text Styling**: Formatted as **10.5pt Bold (`Inter Medium`)**, text vertically centered with `valign: "middle"` and `align: "center"` in `#FFFFFF` on green (`#034E48`) or dark (`#1C1C1E`) boxes.
- **Text Starting Offset**: Card body text MUST start at `pillY + 0.58 in` (`CARD_TEXT_OFFSET = 0.58`) to leave clean whitespace below the header box.
- **Row 2 Positioning**: Row 2 card pills must be placed below Row 1 text at `r2Y = Math.min(3.45, Math.max(3.20, r1Y + CARD_PILL_H + maxR1TextH + 0.22))`.
- **Bottom Footnote Clearance**: Bottom notes on card slides must be anchored below Row 2 text at `bottomNoteY = Math.min(4.85, Math.max(4.00, r2Y + CARD_TEXT_OFFSET + maxR2TextH + 0.14))`.

## 6. Cross-Machine Reproducibility & Fonts
- **Required Font**: All slides MUST use the **`Inter`** font family (`Inter`, `Inter Medium`, `Inter Bold`). Any machine generating or viewing decks must have the Inter font family installed to avoid font fallback differences.
- **Relative Path Resolution**: Never hardcode user profiles or machine-specific directories; always use `__dirname` and portable relative paths.

