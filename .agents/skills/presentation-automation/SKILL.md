---
name: presentation-automation
description: Automates generating, formatting, and re-rendering BombayDC styled pitch decks and sales presentations from JSON plans.
---

# Presentation Automation Skill (BombayDC Standard)

This skill provides guidelines, layout rules, mathematical specifications, and instructions for generating, styling, and updating BombayDC presentation decks.

## Core Rules & Design System Guidelines
1. **Fixed Header Line & Content Y Position**: The horizontal header divider line MUST ALWAYS be placed at **`y: 1.35 in`** and content start Y position MUST ALWAYS be **`y: 1.60 in`** (`contentY`), regardless of whether a tagline/subline is present or absent on the slide.
2. **Slide Title & Index Formatting**: Slide titles MUST be **UPPERCASE, BOLD (18.8pt)** placed at `x: 1.99 in`, `y: 0.57 in`. Slide index numbers (`1.1`, `1.2`, ..., `1.10`) placed at `x: 0.16 in`, `y: 0.57 in`.
3. **Header Formatting**: Any single line of text that is NOT a bullet point should be treated as a header. It must be formatted as **uppercase, bold, and in accent green color** (`#034E48` in light mode, `#4DB89A` in dark mode).
4. **Bullet Spacing**: Bullet points must use a tight single-space gap (`"• "`) without PowerPoint default paragraph indentation.
5. **Card Grid Cap (Max 2–3 per deck)**: Do NOT use `CARD_GRID` (the dark green rectangular card design) for more than **2 to 3 slides maximum** in any deck. If a deck has more qualifying lists, reserve `CARD_GRID` for the 2–3 most critical capability/solution slides and render the rest using alternative layouts (`ACRONYM_GRID`, `THREE_COLUMN`, `TIMELINE_LIST`, `TWO_COLUMN`, `ROW_LIST`).
6. **Adaptive Card Height for Short Terms**: For card grid slides with short 1–2 word titles (e.g., `ACQUIRE`, `ACTIVATE`, `ENGAGE`), use compact header pill heights (`0.35"`) instead of tall bulky boxes.
7. **3-Column Offering / Response Slides**: For slides presenting 3 core offerings or response pillars (e.g., "Our response: three simple BFSI offerings"), ALWAYS use `THREE_COLUMN` layout with green bold headers and bullet points (`• `).
8. **Closing Slide**: Every presentation deck must end with the green **`LET'S BUILD WHAT'S NEXT`** closing slide featuring contact details and COO initial branding (`xEnd` of dividing line = `9.84 in`).

## Master References & Rules
- `BOMBAYDC_DESIGN_SYSTEM_RULES.md`: Authoritative design system rulebook containing complete mathematical formulas, typography tokens, grid coordinates, and color palettes.
- `.agents/AGENTS.md`: Central design rule sheet for LLM behavior.
- `generate_ppt.js`: Core pptxgenjs compilation engine enforcing layouts, themes, typography, and card height calculations.
