"use strict";

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const L = require("./layout_constants");

const BASE = __dirname;
const RIDDHI_MEDIA = path.join(BASE, "riddhi_unpacked/ppt/media");

// Helpers to get image paths
const M = n => path.join(RIDDHI_MEDIA, n);
const ASSETS = {
    cover: M("image-1-1.png"),
    coverLogo: M("image-1-2.png"),
    hdrDark: M("image-2-1.png"),
    hdrLight: M("image-3-1.png"),
    hdrGreen: M("image-10-1.png"),
    closingHeadshot: M("image-10-2.png")
};

function loadJson(filepath) {
    if (!fs.existsSync(filepath)) return null;
    return JSON.parse(fs.readFileSync(filepath, 'utf8'));
}

async function generatePPT(plan, outputPath) {
    if (!plan) {
        console.log("Reading slide_plan.json...");
        plan = loadJson(path.join(__dirname, 'slide_plan.json'));
    }
    if (!plan) throw new Error("Slide plan not found!");

    let pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.title = plan.proposalName || "Client Proposal";

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // HELPERS
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    function addHeader(s, theme) {
        let img = ASSETS.hdrLight;
        if (theme === "dark") img = ASSETS.hdrDark;
        if (theme === "green") img = ASSETS.hdrGreen;

        if (fs.existsSync(img)) {
            s.addImage({ path: img, x: L.LEFT_MARGIN, y: L.HDR_Y, w: L.HDR_W, h: L.HDR_H });
        }
    }

    function addRule(s, y, theme, xStart = L.LEFT_MARGIN, xEnd = L.WIDTH - L.LEFT_MARGIN) {
        const t = L.THEMES[theme];
        s.addShape(pres.shapes.LINE, {
            x: xStart, y, w: xEnd - xStart, h: 0,
            line: { color: t.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
        });
    }

    function addVRule(s, x, yStart, len, theme) {
        const t = L.THEMES[theme];
        s.addShape(pres.shapes.LINE, {
            x, y: yStart, w: 0, h: len,
            line: { color: t.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
        });
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // SLIDE 1 — COVER (Dark theme)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    {
        const s = pres.addSlide();
        s.background = { color: L.THEMES.dark.bg };

        if (plan.coverBgImage && fs.existsSync(plan.coverBgImage)) {
            s.addImage({ path: plan.coverBgImage, x: 0, y: 0, w: L.WIDTH, h: 5.63 });
        } else if (fs.existsSync(ASSETS.cover)) {
            s.addImage({ path: ASSETS.cover, x: 0, y: 0, w: L.WIDTH, h: 5.63 });
        }
        if (fs.existsSync(ASSETS.coverLogo)) {
            s.addImage({ path: ASSETS.coverLogo, x: L.COVER_X, y: 0.5, w: 1.89, h: 0.21 });
        }

        const proposalName = plan.proposalName || plan.presentationTitle || (plan.coverSlide && plan.coverSlide.title) || "BOMBAYDC PLATFORMS";
        const coverTitleLines = Math.max(1, Math.ceil(proposalName.length / 32));
        const coverTitleH = parseFloat((coverTitleLines * 0.45).toFixed(2));
        const coverTitleY = 1.60;

        // LAYER B — DYNAMIC COVER TITLE FLOW
        s.addText(proposalName.toUpperCase(), {
            x: L.COVER_X, y: coverTitleY, w: 8.8, h: coverTitleH,
            fontSize: 28, color: L.THEMES.dark.title, fontFace: L.FONT_TITLE, bold: true,
            charSpacing: -0.5, lineSpacing: "115%", valign: "top", margin: 0
        });

        const subtitle = plan.proposalSubtitle || "A strategic layout and experience design presentation.";
        const coverSubY = parseFloat((coverTitleY + coverTitleH + 0.08).toFixed(2));
        s.addText(subtitle, {
            x: L.COVER_X, y: coverSubY, w: 7.5, h: 0.45,
            fontSize: 8.55, color: "ECE9E4", fontFace: L.FONT_BODY, lineSpacing: "145%", valign: "top", margin: 0
        });

        const coverLineY = parseFloat((coverSubY + 0.50).toFixed(2));
        s.addShape(pres.shapes.LINE, {
            x: L.COVER_X, y: coverLineY, w: 9.0, h: 0,
            line: { color: "FFFFFF", width: 0.5, transparency: 70 }
        });

        const infoY = parseFloat((coverLineY + 0.15).toFixed(2));
        const createdBy = (plan.proposalCreatedBy || "BOMBAYDC").toUpperCase();
        const createdFor = (plan.proposalCreatedFor || "ENTERPRISE CLIENT PARTNERSHIP").toUpperCase();
        const proposalDate = plan.proposalDate || "";

        let infoText = `CREATED FOR: ${createdFor}\nCREATED BY: ${createdBy}`;
        if (proposalDate) {
            infoText += `\nDATE: ${proposalDate}`;
        }
        s.addText(infoText, {
            x: L.COVER_X, y: infoY, w: 7.5, h: 0.8,
            fontSize: 8.55, color: "B4B4B4", fontFace: L.FONT_BODY, lineSpacing: "150%", valign: "top", margin: 0
        });

        // LAYER A — FIXED MASTER FRAME FOOTER (Anchored! Never moves!)
        s.addText("CONFIDENTIAL AND PROPRIETARY | © BombayDC. This material is intended solely for your internal use.", {
            x: L.COVER_X, y: 4.96, w: 4.5, h: 0.35, fontSize: 4.27, color: "FFFFFF", fontFace: L.FONT_BODY, margin: 0
        });

        s.addText("bombaydc.com", {
            x: 8.0, y: 4.96, w: 1.5, h: 0.25,
            fontSize: 9.0, color: L.THEMES.dark.title, fontFace: L.FONT_BODY, align: "right", margin: 0
        });
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // SLIDE 2 — INDEX / TABLE OF CONTENTS (Light theme)
    // Rendered ONLY if there are more than 2 logical sections
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const numSections = (plan.dynamicSlides || []).filter(s => s && s.layout === "SECTION_STARTER").length;
    if (numSections > 2) {
        const s = pres.addSlide();
        s.background = { color: L.THEMES.light.bg };
        addHeader(s, "light");

        s.addText("TODAY'S PRESENTATION", {
            x: 1.99, y: 0.57, w: 7.85, h: 0.35,
            fontSize: 18.8, color: L.THEMES.light.title, fontFace: L.FONT_TITLE, bold: true, valign: "top", margin: 0
        });

        // Fixed divider line at 1.35 in
        s.addShape(pres.shapes.LINE, {
            x: 0.16, y: 1.35, w: 9.68, h: 0,
            line: { color: L.THEMES.light.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
        });

        // Find all section starter titles
        const sections = plan.dynamicSlides
            .filter(slide => slide && slide.layout === "SECTION_STARTER")
            .map((slide, idx) => ({ num: `${idx + 1}.0`, title: slide.title }));

        sections.forEach((item, idx) => {
            const rowY = 1.60 + idx * 0.75;
            
            s.addText(item.num, {
                x: 1.99, y: rowY + 0.05, w: 0.8, h: 0.35,
                fontSize: 11.5, color: L.THEMES.light.accent, fontFace: L.FONT_TITLE, bold: true, valign: "top", margin: 0
            });
            s.addText((item.title || "").toUpperCase(), {
                x: 2.85, y: rowY + 0.05, w: 6.8, h: 0.35,
                fontSize: 11.5, color: L.THEMES.light.title, fontFace: L.FONT_TITLE, bold: true, valign: "top", margin: 0, fit: 'shrink'
            });

            s.addShape(pres.shapes.LINE, {
                x: 1.99, y: rowY + 0.52, w: 7.69, h: 0,
                line: { color: L.THEMES.light.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
            });
        });
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // DYNAMIC SLIDES
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    if (plan.dynamicSlides && Array.isArray(plan.dynamicSlides)) {

        // ── Pre-processing ────────────────────────────────────────────
        // 0. Filter out redundant pre-closing CTA slides and empty slides
        plan.dynamicSlides = plan.dynamicSlides.filter(s => {
            if (!s) return false;
            if (s.layout === "SECTION_STARTER") return Boolean(s.title);
            
            const titleL = (s.title || "").toLowerCase();
            const leftL = (s.leftText || "").toLowerCase();
            const rightL = (s.rightText || "").toLowerCase();
            const bodyL = (s.body || "").toLowerCase();

            // Filter out redundant pre-closing CTA / teaser slides
            const isRedundantCTA = titleL.includes("where is digital value leaking") ||
                                   titleL.includes("where is value leaking") ||
                                   titleL.includes("talk to us") ||
                                   titleL.includes("opportunity to work together") ||
                                   titleL.includes("next steps") ||
                                   leftL.includes("highest-value journeys to fix first") ||
                                   rightL.includes("www.bombaydc.com") ||
                                   bodyL.includes("www.bombaydc.com");

            if (isRedundantCTA) return false;

            const hasData = (s.leftText && s.leftText.trim()) ||
                            (s.rightText && s.rightText.trim()) ||
                            (s.body && s.body.trim()) ||
                            (s.headline && s.headline.trim()) ||
                            (s.rows && s.rows.length > 0) ||
                            (s.cards && s.cards.length > 0) ||
                            (s.items && s.items.length > 0) ||
                            (s.cols && s.cols.length > 0) ||
                            (s.columns && s.columns.length > 0) ||
                            (s.quads && Object.keys(s.quads).length > 0) ||
                            (s.steps && s.steps.length > 0) ||
                            (s.needs && s.needs.length > 0) ||
                            (s.reasons && s.reasons.length > 0) ||
                            (s.cases && s.cases.length > 0) ||
                            (s.clients && s.clients.length > 0) ||
                            (s.metrics && s.metrics.length > 0) ||
                            (s.bios && s.bios.length > 0) ||
                            (s.phases && s.phases.length > 0) ||
                            (s.image || s.img);
            return Boolean(hasData);
        });

        // 1. Remove section starters from short decks (<=2 sections)
        const starterCount = plan.dynamicSlides.filter(s => s.layout === "SECTION_STARTER").length;
        if (starterCount <= 2) {
            plan.dynamicSlides = plan.dynamicSlides.filter(s => s.layout !== "SECTION_STARTER");
        }

        // 2. Auto-assign dark theme: every 3rd content slide gets dark
        let contentSlideIdx = 0;
        let cardGridCount = plan.dynamicSlides.filter(s => s.layout === "CARD_GRID").length;
        plan.dynamicSlides.forEach(slide => {
            if (slide.layout !== "SECTION_STARTER") {
                if (!slide.theme) {
                    if (contentSlideIdx === 0) slide.theme = "light";
                    else if (contentSlideIdx % 4 === 2) slide.theme = "dark";
                    else slide.theme = "light";
                }
                contentSlideIdx++;
            }

            // AUTO-MAP LAYOUTS BASED ON USER FEEDBACK
            if (slide.layout === "TWO_COLUMN" && slide.leftText && slide.rightText) {
                const titleL = (slide.title || "").toLowerCase();
                const isClosingSlide = titleL.includes("where is") || titleL.includes("leaking") || titleL.includes("talk to us") || titleL.includes("let's build") || titleL.includes("opportunity");
                
                if (!isClosingSlide) {
                    const leftParts = slide.leftText.split(/\n\s*\n/);
                    const rightParts = slide.rightText.split(/\n\s*\n/);
                    if (leftParts.length === 2 && rightParts.length === 2) {
                        // Check that parts actually have titles and bodies, not floating bullet points
                        const validParts = [...leftParts, ...rightParts].every(p => p.includes("\n") && p.split("\n").length >= 2);
                        if (validParts) {
                            slide.layout = "ROW_LIST";
                            slide.rows = [];
                            [...leftParts, ...rightParts].forEach(part => {
                                const lines = part.split("\n");
                                slide.rows.push({
                                    title: lines[0].trim(),
                                    body: lines.slice(1).join("\n").trim()
                                });
                            });
                        }
                    }
                }
            }

            // Dynamic theme rhythm for visual contrast across slides
            if (!slide.theme) {
                const titleLower = (slide.title || "").toLowerCase();
                if (titleLower.includes("shift") || titleLower.includes("problem") || titleLower.includes("cause") || titleLower.includes("friction") || titleLower.includes("underperform")) {
                    slide.theme = "dark";
                }
            }

            // Auto-promote 3-column / Team slides to THREE_COLUMN
            const titleLower = (slide.title || "").toLowerCase();
            if (titleLower.includes("team") || titleLower.includes("structure") || titleLower.includes("roles")) {
                if (slide.layout !== "TIMELINE_LIST") {
                    slide.layout = "THREE_COLUMN";
                    slide.theme = "dark";
                    if (!slide.cols && slide.rows) {
                        slide.cols = slide.rows;
                    }
                }
            }

            // Auto-promote lists to CARD_GRID up to maximum 3 slides per deck
            if (slide.layout === "ROW_LIST" && slide.rows) {
                const titleL = (slide.title || "").toLowerCase();
                const isCardCandidate = titleL.includes("capability") || titleL.includes("capabilities") ||
                                       titleL.includes("solution") || titleL.includes("solutions") ||
                                       titleL.includes("category") || titleL.includes("categories") ||
                                       titleL.includes("offering") || titleL.includes("offerings") ||
                                       titleL.includes("opportunity") || titleL.includes("opportunities") ||
                                       titleL.includes("integrated") || titleL.includes("ecosystem") ||
                                       titleL.includes("outcomes") || titleL.includes("feature") ||
                                       titleL.includes("modules");

                // If 4 items, promote to ACRONYM_GRID (BEAM style)
                if (slide.rows.length === 4) {
                    slide.layout = "ACRONYM_GRID";
                }
                // If 3 items with newlines, promote to TIMELINE_LIST; if offerings/response, promote to THREE_COLUMN
                else if (slide.rows.length === 3) {
                    const hasNewlines = slide.rows.some(r => r.title && r.title.includes("\n"));
                    if (hasNewlines || (slide.title && slide.title.toLowerCase().includes("validated product"))) {
                        slide.layout = "TIMELINE_LIST";
                    } else if (titleL.includes("response") || titleL.includes("offering") || titleL.includes("offerings") || titleL.includes("three") || titleL.includes("pillars") || titleL.includes("simple")) {
                        slide.layout = "THREE_COLUMN";
                        slide.cols = slide.rows;
                    } else if (isCardCandidate && cardGridCount < 3) {
                        slide.layout = "CARD_GRID";
                        slide.cards = slide.rows;
                        cardGridCount++;
                    }
                }
                // If 5 to 10 items or card candidate, promote to CARD_GRID up to max 2-3 per deck
                else if ((isCardCandidate || (slide.rows.length >= 5 && slide.rows.length <= 10)) && cardGridCount < 3) {
                    slide.layout = "CARD_GRID";
                    slide.cards = slide.rows;
                    cardGridCount++;
                }
            }
        });

        // Pre-compute hierarchical slide numbers
        // If starters were removed, start section at 1 so slides show 1.1, 1.2, not 0.1
        const hasStarters = plan.dynamicSlides.some(s => s.layout === "SECTION_STARTER");
        let sectionCount = hasStarters ? 0 : 1;
        let slideInSection = 0;
        const slideNums = plan.dynamicSlides.map(slide => {
            if (slide.layout === "SECTION_STARTER") {
                sectionCount++;
                slideInSection = 0;
                return String(sectionCount);
            } else {
                slideInSection++;
                // If only one section (no starters), show sequential: 1, 2, 3...
                return hasStarters ? `${sectionCount}.${slideInSection}` : String(slideInSection);
            }
        });

        plan.dynamicSlides.forEach((slide, idx) => {
            const s = pres.addSlide();
            // SECTION_STARTER is always green theme
            const theme = slide.layout === "SECTION_STARTER" ? "green" : (slide.theme || "light");
            const colors = L.THEMES[theme];

            s.background = { color: colors.bg };
            if (slide.bg_image && fs.existsSync(slide.bg_image)) {
                s.addImage({ path: slide.bg_image, x: 0, y: 0, w: L.WIDTH, h: 5.625 });
            }
            addHeader(s, theme);

            if (slide.layout === "SECTION_STARTER") {
                // Section starter slide render with native 16:9 un-stretched cover cropped photo background
                const secNum = slide.number || slideNums[idx];
                const secIdx = parseInt(secNum, 10) || (idx + 1);
                const bgImgName = `sec${Math.min(secIdx, 3)}_bg.png`;
                const bgImgPath = path.resolve(BASE, "bg_images", bgImgName);

                if (fs.existsSync(bgImgPath)) {
                    // Native PowerPoint slide background image (16:9 1920x1080 cover cropped with dark overlay)
                    s.background = { path: bgImgPath };
                } else {
                    s.background = { color: "034E48" };
                }

                // Section starter top status bar
                addHeader(s, "green");

                // Section starter number (01, 02, 03)
                const doubleDigit = String(secNum).padStart(2, "0");
                s.addText(doubleDigit, {
                    x: L.PRIMARY_X, y: L.STARTER_NUM_Y, w: 6.8, h: 1.1,
                    fontSize: L.STARTER_NUM_FS, color: "ECE9E4", fontFace: L.FONT_TITLE, bold: true, valign: "bottom", margin: 0
                });

                s.addShape(pres.shapes.LINE, {
                    x: L.PRIMARY_X, y: L.STARTER_LINE_Y, w: L.STARTER_LINE_W, h: 0,
                    line: { color: "3E8D86", width: L.LINE_WIDTH }
                });

                const titleText = (slide.title || "UNTITLED SECTION").toUpperCase();
                s.addText(titleText, {
                    x: L.PRIMARY_X, y: L.STARTER_TITLE_Y, w: 6.8, h: 1.8,
                    fontSize: L.STARTER_TITLE_FS, color: "ECE9E4", fontFace: L.FONT_TITLE, bold: true, valign: "top", margin: 0, fit: 'shrink'
                });
            }
            else {
                // ── EXACT PIXEL-PERFECT MATCH TO MANUAL_REDESIGN_BOMBAYDC_RMZ ──
                // QA-verified measurements (2026-08-04) against Manual RMZ reference:
                //   Slide num:  x:0.16, y:0.65, w:1.68, h:0.25, fs:18pt, color:body
                //   Title:      x:1.99, y:0.57, w:7.85, h:0.35, fs:20pt, UPPERCASE, color:title
                //   Tagline:    x:1.99, y:0.92, w:7.65, h:0.35, fs:8.55pt, color:body (only when present)
                //   Divider:    x:0.16, y:0.89 (no tagline) / y:1.35 (with tagline), w:9.68
                //   ContentY:   dividerY + 0.25

                s.addText(slideNums[idx], {
                    x: 0.16, y: 0.57, w: 1.68, h: 0.35,
                    fontSize: 18.8, color: colors.body, fontFace: L.FONT_BODY, valign: "top", margin: 0
                });

                const isTimelineTeamTagline = slide.layout === "TIMELINE_LIST" && slide.tagline && slide.tagline.toLowerCase().includes("team structure");
                const hasTagline = Boolean(slide.tagline && !isTimelineTeamTagline);

                const slideTitleText = (slide.title || "UNTITLED SLIDE").toUpperCase();
                s.addText(slideTitleText, {
                    x: 1.99, y: 0.57, w: 7.85, h: hasTagline ? 0.35 : 0.68,
                    fontSize: 18.8, color: colors.title, fontFace: L.FONT_TITLE, bold: true,
                    lineSpacing: "115%", valign: "top", margin: 0, fit: 'shrink'
                });

                if (hasTagline) {
                    s.addText(slide.tagline, {
                        x: 1.99, y: 0.92, w: 7.65, h: 0.35,
                        fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY,
                        lineSpacing: "140%", valign: "top", margin: 0, fit: 'shrink'
                    });
                }

                // FIXED DIVIDER Y POSITION (1.35) & CONTENT Y POSITION (1.60)
                // Guaranteed consistent horizontal rule and content Y start across ALL slides
                const dividerY = 1.35;
                s.addShape(pres.shapes.LINE, {
                    x: 0.16, y: dividerY, w: 9.68, h: 0,
                    line: { color: colors.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
                });

                // contentY is fixed at 1.60 across all content slides
                const contentY = 1.60;

                // Helper to format text: only explicit uppercase category headers get green bold, plain sentences stay normal text
                const formatColText = (txt, defaultColor) => {
                    if (!txt) return "";
                    const lines = txt.split("\n").map(l => l.trim()).filter(l => l.length > 0);
                    
                    return lines.map((trimmed, idx) => {
                        const isBullet = /^[•\-\*]/.test(trimmed);
                        const rawText = trimmed.replace(/^[•\-\*]\s*/, '');
                        const isAllCaps = rawText.toUpperCase() === rawText && /[A-Z]/.test(rawText) && rawText.length >= 2 && rawText.length <= 50;
                        const endsWithColon = rawText.endsWith(":");

                        // Only explicit uppercase category headers (e.g. "TRANSACTIONS WORK", "THEN", "NOW") are green bold
                        const isHeader = !isBullet && (isAllCaps || endsWithColon);

                        if (isHeader) {
                            return {
                                text: rawText.toUpperCase(),
                                options: {
                                    color: colors.accent,
                                    bold: true,
                                    breakLine: true,
                                    paraSpaceBefore: idx === 0 ? 0 : 10,
                                    paraSpaceAfter: 4
                                }
                            };
                        }

                        // Preserves original text formatting (bulleted vs plain paragraph)
                        const displayText = isBullet ? "• " + rawText : rawText;

                        return {
                            text: displayText,
                            options: {
                                color: defaultColor,
                                breakLine: true,
                                paraSpaceAfter: 8
                            }
                        };
                    });
                };

                if (slide.layout === "TWO_COLUMN" || slide.layout === "SINGLE_COLUMN") {
                    let imgPath = slide.image || slide.img;
                    if (imgPath && !path.isAbsolute(imgPath)) {
                        imgPath = path.join(BASE, imgPath);
                    }
                    
                    const hasLeft = Boolean(slide.leftText && slide.leftText.trim());
                    const hasRight = Boolean(slide.rightText && slide.rightText.trim());

                    if (imgPath && fs.existsSync(imgPath)) {
                        // Image on left, narrative text below
                        const fullText = (slide.leftText ? slide.leftText + "\n" : "") + (slide.rightText || slide.body || "");
                        s.addImage({ path: imgPath, x: L.PRIMARY_X, y: contentY, w: 3.56, h: 1.9 });
                        s.addText(formatColText(fullText, colors.body), {
                            x: L.PRIMARY_X, y: contentY + 2.0, w: 7.65, h: 1.75,
                            fontSize: 8.55, fontFace: L.FONT_BODY,
                            valign: "top", margin: 0, lineSpacing: "145%"
                        });
                    } else if (hasLeft && hasRight) {
                        // 2 Distinct Columns Side-by-Side (Matches Manual_Redesign_BOMBAYDC for RMZ reference)
                        s.addText(formatColText(slide.leftText, colors.title), {
                            x: L.PRIMARY_X, y: contentY, w: 3.56, h: 3.75,
                            fontSize: 8.55, fontFace: L.FONT_BODY,
                            valign: "top", margin: 0, lineSpacing: "145%"
                        });

                        if (slide.divider !== false) {
                            addVRule(s, L.DIVIDER_X, contentY, 3.7, theme);
                        }

                        s.addText(formatColText(slide.rightText, colors.body), {
                            x: L.SECONDARY_X, y: contentY, w: 3.70, h: 3.75,
                            fontSize: 8.55, fontFace: L.FONT_BODY,
                            valign: "top", margin: 0, lineSpacing: "145%"
                        });
                    } else {
                        // Single Column Narrative Layout (when only leftText or body exists)
                        const fullText = slide.leftText || slide.rightText || slide.body || slide.headline || "";
                        s.addText(formatColText(fullText, colors.body), {
                            x: L.PRIMARY_X, y: contentY + 0.1, w: 7.65, h: 3.6,
                            fontSize: 8.55, fontFace: L.FONT_BODY,
                            lineSpacing: "145%", valign: "top", margin: 0
                        });
                    }
                }
                else if (slide.layout === "ROW_LIST") {
                    // ── ROW_LIST: Pixel-perfect match to build_rmz_full.js rowList() ──
                    const rows = slide.rows || [];
                    const rowCount = rows.length;
                    const headerH = 0.25;
                    const rowStartY = contentY + headerH + 0.05;
                    const availH = 5.25 - rowStartY;
                    const rowH = Math.min(0.95, availH / Math.max(rowCount, 1));

                    // Column headers — 7.5pt with letter-spacing
                    s.addText("CATEGORY", {
                        x: L.PRIMARY_X, y: contentY, w: 3.2, h: headerH,
                        fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, charSpacing: 1.5, valign: "middle", margin: 0
                    });
                    s.addText("DESCRIPTION", {
                        x: L.SECONDARY_X, y: contentY, w: 3.7, h: headerH,
                        fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, charSpacing: 1.5, valign: "middle", margin: 0
                    });
                    // Separator below header row
                    s.addShape(pres.shapes.LINE, {
                        x: L.PRIMARY_X, y: contentY + headerH, w: L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X, h: 0,
                        line: { color: colors.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
                    });

                    rows.forEach((row, rIdx) => {
                        const y = rowStartY + rIdx * rowH;

                        // Category title — 9.5pt bold accent, wide enough to prevent wrapping
                        s.addText(row.title || "", {
                            x: L.PRIMARY_X, y: y + 0.06, w: 3.4, h: rowH - 0.12,
                            fontSize: 9.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true,
                            lineSpacing: "130%", valign: "middle", fit: 'shrink', margin: 0
                        });

                        // Body description — Secondary column
                        s.addText(row.body || "", {
                            x: L.SECONDARY_X, y: y + 0.06, w: 3.7, h: rowH - 0.12,
                            fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY,
                            lineSpacing: "145%", valign: "middle", fit: 'shrink', margin: 0
                        });

                        // Horizontal separator between rows
                        s.addShape(pres.shapes.LINE, {
                            x: L.PRIMARY_X, y: y + rowH, w: L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X, h: 0,
                            line: { color: colors.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
                        });
                    });
                }
                else if (slide.layout === "TIMELINE_LIST") {
                    // ── TIMELINE_LIST: Matches Screenshot "From idea to validated product" ──
                    const rows = slide.rows || [];
                    const rowCount = rows.length;
                    const totalAvailH = 3.6;
                    // Adjusted row height for better vertical distribution
                    const rowH = Math.max(0.68, totalAvailH / Math.max(rowCount, 1));
                    const rowStartY = contentY + 0.1;

                    rows.forEach((row, rIdx) => {
                        const y = rowStartY + rIdx * rowH;
                        
                        let col1Text = row.title || "";
                        let col2Text = "";
                        let singleTitleSpan = false;
                        if (row.title && row.title.includes("\n")) {
                            const parts = row.title.split("\n");
                            col1Text = parts[0].trim();
                            col2Text = parts.slice(1).join(" ").trim();
                        } else {
                            singleTitleSpan = true;
                        }

                        if (singleTitleSpan) {
                            s.addText(col1Text, {
                                x: L.PRIMARY_X, y: y + 0.05, w: 3.4, h: rowH - 0.1,
                                fontSize: 9.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true,
                                lineSpacing: "130%", valign: "middle", fit: 'shrink', margin: 0
                            });
                        } else {
                            s.addText(col1Text, {
                                x: L.PRIMARY_X, y: y + 0.05, w: 1.3, h: rowH - 0.1,
                                fontSize: 13.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true,
                                valign: "middle", fit: 'shrink', margin: 0
                            });

                            s.addText(col2Text, {
                                x: L.PRIMARY_X + 1.4, y: y + 0.05, w: 2.1, h: rowH - 0.1,
                                fontSize: 9.5, color: colors.title, fontFace: L.FONT_TITLE, bold: true,
                                lineSpacing: "125%", valign: "middle", fit: 'shrink', margin: 0
                            });
                        }

                        // Col 3: Body description — starts at 3.6" from left
                        s.addText(row.body || "", {
                            x: L.PRIMARY_X + 3.6, y: y + 0.05, w: 4.1, h: rowH - 0.1,
                            fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY,
                            lineSpacing: "145%", valign: "middle", fit: 'shrink', margin: 0
                        });

                        // Horizontal separator between rows
                        s.addShape(pres.shapes.LINE, {
                            x: L.PRIMARY_X, y: y + rowH, w: L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X, h: 0,
                            line: { color: colors.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
                        });
                    });

                    // Tagline at the bottom
                    const tlBottomY = rowStartY + rowCount * rowH + 0.12;
                    const tlTagline = slide.outroText || (slide.tagline && slide.tagline.toLowerCase().includes("team structure") ? slide.tagline : "");
                    if (tlTagline && tlBottomY < 5.2) {
                        s.addText(tlTagline, {
                            x: L.PRIMARY_X, y: tlBottomY, w: 7.65, h: 0.35,
                            fontSize: 8.55, color: colors.accent, fontFace: L.FONT_BODY, bold: true,
                            valign: "top", margin: 0, fit: 'shrink'
                        });
                    }
                }
                else if (slide.layout === "GRID_LAYOUT") {
                    // ── GRID_LAYOUT: Exact match to Delivery Methodology reference ──
                    const q = slide.quads || {};
                    const startX = L.PRIMARY_X;
                    const colW = 1.80;
                    const colGap = 0.12;
                    const totalW = colW + colGap;

                    const gridItems = [
                        { label: "SITUATION",      text: q.situation    || "", num: "1" },
                        { label: "CONSTRAINT",     text: q.constraint   || "", num: "2" },
                        { label: "INTERVENTION",   text: q.intervention || "", num: "3" },
                        { label: "WHY IT MATTERS", text: q.whyItMatters || "", num: "4" }
                    ];

                    gridItems.forEach((item, gi) => {
                        const cx = startX + gi * totalW;

                        // Huge Number (matches Delivery Methodology reference)
                        s.addText(item.num, {
                            x: cx, y: contentY, w: colW, h: 0.85,
                            fontSize: 65, color: colors.accent, fontFace: L.FONT_TITLE, valign: "bottom", margin: 0
                        });

                        // Category Title — bold accent
                        s.addText(item.label, {
                            x: cx, y: contentY + 0.90, w: colW, h: 0.35,
                            fontSize: 9.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true, valign: "top", fit: 'shrink', margin: 0
                        });

                        // Thin underline below title
                        s.addShape(pres.shapes.LINE, {
                            x: cx, y: contentY + 1.30, w: colW * 0.85, h: 0,
                            line: { color: colors.line, width: L.LINE_WIDTH, transparency: L.LINE_TRANSPARENCY }
                        });

                        // Body text with formatted bullets
                        s.addText(formatColText(item.text, colors.body), {
                            x: cx, y: contentY + 1.38, w: colW, h: 2.2,
                            fontSize: 8.0, fontFace: L.FONT_BODY,
                            lineSpacing: "140%", valign: "top", fit: 'shrink', margin: 0
                        });
                    });
                }
                else if (slide.layout === "PROCESS_FLOW") {
                    // ── PROCESS_FLOW: Horizontal boxes with arrows (flowchart style) ──
                    const steps = slide.steps || slide.rows || [];
                    const stepCount = Math.min(steps.length, 5);
                    const boxW = (L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X - (stepCount - 1) * 0.28) / stepCount;
                    const boxH = 2.0;
                    const boxY = contentY + 0.4;

                    steps.slice(0, stepCount).forEach((step, si) => {
                        const bx = L.PRIMARY_X + si * (boxW + 0.28);

                        // Filled accent rectangle box
                        const isAlt = si % 2 === 1;
                        const boxFill = isAlt ? (theme === "dark" ? "1a3632" : "E0EDE9") : colors.accent;
                        const textCol = isAlt ? colors.accent : colors.bg;

                        s.addShape(pres.shapes.RECTANGLE, {
                            x: bx, y: boxY, w: boxW, h: boxH,
                            fill: { color: boxFill }, line: { color: colors.accent, width: 1.0 }
                        });

                        // Step number (top-left inside box)
                        s.addText(String(si + 1).padStart(2, "0"), {
                            x: bx + 0.1, y: boxY + 0.1, w: 0.4, h: 0.25,
                            fontSize: 7, color: isAlt ? colors.body : "FFFFFF60", fontFace: L.FONT_BODY, bold: true
                        });

                        // Step title
                        s.addText(step.title || "", {
                            x: bx + 0.1, y: boxY + 0.42, w: boxW - 0.2, h: 0.45,
                            fontSize: 9.5, color: textCol, fontFace: L.FONT_TITLE,
                            lineSpacing: "115%", valign: "top", fit: 'shrink'
                        });

                        // Step body
                        s.addText(step.body || "", {
                            x: bx + 0.1, y: boxY + 0.98, w: boxW - 0.2, h: boxH - 1.1,
                            fontSize: 8.0, color: textCol, fontFace: L.FONT_BODY,
                            lineSpacing: "140%", valign: "top", fit: 'shrink'
                        });
                    });

                    // Optional footer label below boxes
                    if (slide.footerText) {
                        s.addText(slide.footerText, {
                            x: L.PRIMARY_X, y: boxY + boxH + 0.2, w: L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X, h: 0.3,
                            fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, lineSpacing: "140%"
                        });
                    }
                }
                else if (slide.layout === "CARD_GRID") {
                    // ── CARD_GRID: Wider boxes & Dark Grey fill on dark background ──
                    const cards = slide.cards || slide.items || slide.rows || [];
                    let startY = contentY + 0.1;

                    // Intro tagline above grid if present
                    if (slide.introText || slide.subhead) {
                        s.addText(slide.introText || slide.subhead, {
                            x: L.PRIMARY_X, y: startY, w: 7.65, h: 0.3,
                            fontSize: 9.5, color: colors.title, fontFace: L.FONT_BODY, valign: "top", margin: 0
                        });
                        startY += 0.4;
                    }

                    const cardCount = cards.length;
                    // Support 3, 4, and 5 column layouts dynamically
                    let cols = 3;
                    if (cardCount === 4) cols = 4;
                    else if (cardCount >= 5) cols = 5;

                    const rows = Math.ceil(cardCount / cols);
                    const hasBody = cards.some(c => typeof c === "object" && c.body);

                    // Standardized gaps and card widths for 3, 4, and 5 column layouts
                    let gapX = 0.25;
                    if (cols === 4) gapX = 0.18;
                    else if (cols === 5) gapX = 0.12;

                    const cardW = (L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X - (cols - 1) * gapX) / cols;
                    
                    // STANDARDIZED TALLER CARD BOX DIMENSIONS (0.65" HEADER PILL + 0.65" BODY BOX)
                    const cardH = 0.65; // Taller header pill matching "AI-era opportunities in BFSI"
                    const descH = hasBody ? 0.65 : 0;
                    const rowUnitH = cardH + (hasBody ? 0.08 + descH + 0.20 : 0.20);

                    cards.forEach((card, ci) => {
                        const colIdx = ci % cols;
                        const rowIdx = Math.floor(ci / cols);

                        const cx = L.PRIMARY_X + colIdx * (cardW + gapX);
                        const cy = startY + rowIdx * rowUnitH;

                        const cardTitle = typeof card === "string" ? card : (card.title || card.label || "");
                        const cardBody = typeof card === "object" ? card.body : "";

                        // Box Fill: Dark Grey (#1C1C1E) on dark theme, Dark Forest Green (#034E48) on light theme
                        const cardFill = theme === "dark" ? "1C1C1E" : "034E48";

                        s.addShape(pres.shapes.RECTANGLE, {
                            x: cx, y: cy, w: cardW, h: cardH,
                            fill: { color: cardFill }, line: { width: 0 }
                        });

                        s.addText(cardTitle.toUpperCase(), {
                            x: cx + 0.08, y: cy + 0.04, w: cardW - 0.16, h: cardH - 0.08,
                            fontSize: cols >= 5 ? 7.8 : 8.5, color: "FAFAF7", fontFace: L.FONT_TITLE, bold: true,
                            align: "center", valign: "middle", lineSpacing: "120%", fit: 'shrink'
                        });

                        // Description below if provided — with ZERO overlap
                        if (cardBody) {
                            s.addText(cardBody, {
                                x: cx, y: cy + cardH + 0.08, w: cardW, h: descH,
                                fontSize: cols >= 5 ? 8.0 : 8.55, color: colors.body, fontFace: L.FONT_BODY,
                                lineSpacing: "130%", align: "center", valign: "middle", margin: 0, fit: 'shrink'
                            });
                        }
                    });

                    // Footer text below grid
                    if (slide.footerText || slide.outroText) {
                        const totalGridH = rows * rowUnitH;
                        s.addText(slide.footerText || slide.outroText, {
                            x: L.PRIMARY_X, y: startY + totalGridH + 0.10, w: 7.65, h: 0.4,
                            fontSize: 8.55, color: colors.title, fontFace: L.FONT_BODY,
                            lineSpacing: "140%", valign: "top", fit: 'shrink', margin: 0
                        });
                    }
                }
                else if (slide.layout === "THREE_COLUMN") {
                    // ── THREE_COLUMN: Matches Screenshot 1 ("Project Team") ──
                    const cols = slide.cols || slide.columns || slide.items || slide.rows || [];
                    const count = Math.min(cols.length || 3, 3);
                    const gap = 0.35;
                    const colW = (L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X - (count - 1) * gap) / count;
                    const startY = contentY + 0.1;

                    cols.slice(0, count).forEach((col, ci) => {
                        const cx = L.PRIMARY_X + ci * (colW + gap);
                        const colTitle = typeof col === "string" ? "" : (col.title || col.label || col.header || "");
                        const colBody = typeof col === "string" ? col : (col.body || col.text || col.content || "");

                        // Column Header in bold green accent — 9.5pt to match ROW_LIST category title
                        if (colTitle) {
                            s.addText(colTitle.toUpperCase(), {
                                x: cx, y: startY, w: colW, h: 0.35,
                                fontSize: 9.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true,
                                lineSpacing: "120%", valign: "top", margin: 0, fit: 'shrink'
                            });
                        }

                        // Column body text formatted with bullets
                        const bodyY = colTitle ? startY + 0.38 : startY;
                        s.addText(formatColText(colBody, colors.body), {
                            x: cx, y: bodyY, w: colW, h: 3.5 - (bodyY - contentY),
                            fontSize: 8.55, fontFace: L.FONT_BODY,
                            lineSpacing: "145%", valign: "top", margin: 0, fit: 'shrink'
                        });
                    });
                }
                else if (slide.layout === "FRAMEWORK_GRID" || slide.layout === "ACRONYM_GRID") {
                    // ── ACRONYM_GRID: Matches Screenshot 1 ("BEAM Framework") ──
                    const items = slide.items || slide.cols || slide.rows || [];
                    const count = Math.min(items.length, 4);

                    const gap = 0.25;
                    const colW = (L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X - (count - 1) * gap) / count;
                    const startY = contentY + 0.1;

                    items.slice(0, count).forEach((item, ii) => {
                        const cx = L.PRIMARY_X + ii * (colW + gap);
                        const letter = item.letter || item.num || String(ii + 1);

                        // Huge letter (B, E, A, M) in dark green accent
                        s.addText(letter, {
                            x: cx, y: startY, w: colW, h: 0.9,
                            fontSize: 48, color: colors.accent, fontFace: L.FONT_TITLE,
                            align: "left", valign: "top", margin: 0
                        });

                        // Title below letter
                        s.addText(item.title || "", {
                            x: cx, y: startY + 0.95, w: colW, h: 0.25,
                            fontSize: 9.0, color: colors.accent, fontFace: L.FONT_TITLE, bold: true,
                            valign: "top", margin: 0
                        });

                        // Thin underline
                        s.addShape(pres.shapes.LINE, {
                            x: cx, y: startY + 1.25, w: colW * 0.8, h: 0,
                            line: { color: colors.line, width: L.LINE_WIDTH }
                        });

                        // Description with formatted bullets
                        s.addText(formatColText(item.body || item.text || "", colors.body), {
                            x: cx, y: startY + 1.35, w: colW, h: 1.5,
                            fontSize: 8.0, fontFace: L.FONT_BODY,
                            lineSpacing: "140%", valign: "top", fit: 'shrink', margin: 0
                        });
                    });

                    // Bottom section (e.g. Illustrative use cases with bullets)
                    if (slide.useCases || slide.bulletPoints || slide.bottomSection) {
                        const useCasesY = startY + 2.8;
                        const useCaseTitle = slide.useCaseTitle || "Illustrative use cases";

                        s.addText(useCaseTitle, {
                            x: L.PRIMARY_X, y: useCasesY, w: 7.65, h: 0.25,
                            fontSize: 9.0, color: colors.accent, fontFace: L.FONT_TITLE, bold: true, margin: 0
                        });

                        const bullets = slide.useCases || slide.bulletPoints || [];
                        const bulletText = Array.isArray(bullets) ? bullets.map(b => "• " + b).join("\n") : bullets;

                        s.addText(bulletText, {
                            x: L.PRIMARY_X, y: useCasesY + 0.3, w: 7.65, h: 1.2,
                            fontSize: 8.0, color: colors.body, fontFace: L.FONT_BODY,
                            lineSpacing: "145%", valign: "top", fit: 'shrink', margin: 0
                        });
                    }
                }
                else if (slide.layout === "BRIDGE_FLOW") {
                    // Col 1 label & items
                    s.addText("PLATFORM NOW NEEDS TO BE...", {
                        x: L.PRIMARY_X, y: contentY, w: 2.3, h: 0.2,
                        fontSize: 7.5, color: colors.accent, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, margin: 0
                    });
                    const needs = slide.needs || [];
                    needs.forEach((n, idx) => {
                        const itemY = contentY + 0.35 + idx * 0.48;
                        s.addText(n, {
                            x: L.PRIMARY_X, y: itemY, w: 2.3, h: 0.35,
                            fontSize: 8.55, color: colors.title, fontFace: L.FONT_BODY, valign: "top", margin: 0
                        });
                        addRule(s, itemY + 0.4, theme, L.PRIMARY_X, L.PRIMARY_X + 2.3);
                    });

                    // Col 2
                    s.addText("THIS MEANS...", {
                        x: 5.0, y: contentY, w: 1.95, h: 0.2,
                        fontSize: 7.5, color: colors.accent, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, margin: 0
                    });
                    s.addText(slide.transitionText || "", {
                        x: 5.0, y: contentY + 0.35, w: 1.95, h: 2.0,
                        fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY,
                        lineSpacing: "140%", valign: "top", margin: 0
                    });

                    // Col 3
                    s.addText("THEREFORE...", {
                        x: 7.7, y: contentY, w: 2.0, h: 0.2,
                        fontSize: 7.5, color: colors.accent, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, margin: 0
                    });
                    s.addText(slide.outcomeText || "", {
                        x: 7.7, y: contentY + 0.35, w: 2.0, h: 1.8,
                        fontSize: 8.55, color: colors.title, fontFace: L.FONT_TITLE,
                        lineSpacing: "140%", valign: "top", margin: 0
                    });
                }
                else if (slide.layout === "SHIFT_LAYOUT") {
                    // Headers
                    s.addText("THEN", {
                        x: L.PRIMARY_X, y: contentY, w: 3.5, h: 0.25,
                        fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.5, margin: 0
                    });
                    addRule(s, contentY + 0.3, theme, L.PRIMARY_X, L.PRIMARY_X + 3.5);

                    s.addText("NOW", {
                        x: L.SECONDARY_X, y: contentY, w: 3.7, h: 0.25,
                        fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.5, margin: 0
                    });
                    addRule(s, contentY + 0.3, theme, L.SECONDARY_X, L.SECONDARY_X + 3.7);

                    const thenItems = slide.thenItems || [];
                    const nowItems = slide.nowItems || [];
                    for (let i = 0; i < 5; i++) {
                        const rowY = contentY + 0.35 + i * 0.58;
                        addRule(s, rowY + 0.52, theme, L.PRIMARY_X, L.PRIMARY_X + 3.5);
                        addRule(s, rowY + 0.52, theme, L.SECONDARY_X, L.SECONDARY_X + 3.7);

                        s.addText(thenItems[i] || "", {
                            x: L.PRIMARY_X, y: rowY + 0.1, w: 3.5, h: 0.35,
                            fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top", margin: 0
                        });
                        s.addText(nowItems[i] || "", {
                            x: L.SECONDARY_X, y: rowY + 0.1, w: 3.7, h: 0.35,
                            fontSize: 8.55, color: colors.accent, fontFace: L.FONT_BODY, valign: "top", margin: 0
                        });
                    }
                }
                else if (slide.layout === "OPERATING_MODEL") {
                    const phases = slide.phases || [];
                    phases.forEach((p, i) => {
                        const y = contentY + i * 0.95;
                        addRule(s, y + 0.85, theme, L.PRIMARY_X, L.WIDTH - L.LEFT_MARGIN);

                        s.addText(p.duration || "", {
                            x: L.PRIMARY_X, y: y + 0.16, w: 1.2, h: 0.25,
                            fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top"
                        });
                        s.addText(p.title || "", {
                            x: L.PRIMARY_X + 1.3, y: y + 0.16, w: 2.3, h: 0.3,
                            fontSize: 9.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true, valign: "top"
                        });
                        s.addText(p.desc || "", {
                            x: L.SECONDARY_X, y: y + 0.16, w: 3.7, h: 0.7,
                            fontSize: 8.55, color: colors.title, fontFace: L.FONT_BODY, valign: "top",
                            lineSpacing: "140%"
                        });
                    });

                    // Footer team structure row
                    const footerY = contentY + 2.9;
                    addRule(s, footerY, theme, L.PRIMARY_X, L.WIDTH - L.LEFT_MARGIN);
                    s.addText("NEW TEAM STRUCTURE:", {
                        x: L.PRIMARY_X, y: footerY + 0.15, w: 1.68, h: 0.22,
                        fontSize: 7.5, color: colors.body, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, valign: "top"
                    });
                    s.addText(slide.footerText || "", {
                        x: L.PRIMARY_X + 1.7, y: footerY + 0.15, w: 6.1, h: 0.22,
                        fontSize: 8.55, color: colors.title, fontFace: L.FONT_BODY, valign: "top"
                    });
                }
                else if (slide.layout === "WHY_BOMBAYDC") {
                    const reasons = slide.reasons || [];
                    reasons.forEach((r, i) => {
                        const y = contentY + i * 0.72;
                        addRule(s, y + 0.68, theme, L.PRIMARY_X, L.DIVIDER_X - 0.1);

                        s.addText(String(i + 1), {
                            x: L.PRIMARY_X, y: y + 0.15, w: 0.35, h: 0.35,
                            fontSize: 8.55, color: colors.accent, fontFace: L.FONT_BODY, bold: true, valign: "top"
                        });
                        s.addText(r, {
                            x: L.PRIMARY_X + 0.45, y: y + 0.15, w: 3.2, h: 0.45,
                            fontSize: 8.55, color: colors.title, fontFace: L.FONT_BODY, valign: "top"
                        });
                    });

                    // Trust Signals Box
                    s.addShape(pres.shapes.RECTANGLE, {
                        x: L.SECONDARY_X, y: contentY, w: 3.65, h: 3.55,
                        fill: { color: "034E48" }, line: { width: 0 }
                    });
                    s.addText((slide.trustTitle || "TRUST SIGNALS").toUpperCase(), {
                        x: L.SECONDARY_X + 0.25, y: contentY + 0.2, w: 3.15, h: 0.25,
                        fontSize: 6.41, color: "ECE9E4", fontFace: L.FONT_BODY, valign: "top"
                    });
                    const trustItems = slide.trustItems || [];
                    trustItems.forEach((t, i) => {
                        s.addText(`• ${t}`, {
                            x: L.SECONDARY_X + 0.25, y: contentY + 0.55 + i * 0.52, w: 3.15, h: 0.3,
                            fontSize: 8.55, color: "ECE9E4", fontFace: L.FONT_BODY, valign: "top"
                        });
                    });
                }
                else if (slide.layout.startsWith("CREDS_") && slide.layout.endsWith("_EXACT")) {
                    // ━━━ EXACT CREDS SLIDE: Copy images directly from BDC_Fidelity_Test.pptx ━━━
                    // Map layout ID to fidelity slide number
                    const FIDELITY_MAP = {
                        "CREDS_THE_SHIELD_EXACT": 2,
                        "CREDS_DIAGNOSTICS_EXACT": 3,
                        "CREDS_CAPABILITIES_EXACT": 4,
                        "CREDS_METRICS_EXACT": 5,
                        "CREDS_BDC_POV_EXACT": 6,
                        "CREDS_TEAM_EXACT": 7,
                        "CREDS_CASE_STUDIES_EXACT": 8,
                        "CREDS_CLIENTS_EXACT": 9,
                    };
                    const fidSlideNum = FIDELITY_MAP[slide.layout];
                    if (fidSlideNum !== undefined) {
                        const fidDataPath = path.join(BASE, "fidelity_exact_data.json");
                        const fidMediaDir = path.join(BASE, "fidelity_unpacked/ppt/media");
                        if (fs.existsSync(fidDataPath) && fs.existsSync(fidMediaDir)) {
                            const fidData = JSON.parse(fs.readFileSync(fidDataPath, 'utf8'));
                            const slideData = fidData[`slide${fidSlideNum}`];
                            if (slideData) {
                                // Set background
                                s.background = { color: slideData.bg || "1D1D1F" };
                                let elementsToRender = slideData.elements;
                                if (fidSlideNum === 7 || fidSlideNum === 8) {
                                    elementsToRender = elementsToRender.filter(el => {
                                        if (fidSlideNum === 7) {
                                            if (el.type === 'img' && el.imgName === 'image-7-1.png') return true;
                                            if (el.type === 'text' && (el.text === '5' || el.text === 'GLOBAL PRODUCT EXPERTS')) return true;
                                            return false;
                                        }
                                        if (fidSlideNum === 8) {
                                            if (el.type === 'img' && el.imgName === 'image-8-1.png') return true;
                                            if (el.type === 'text' && (el.text === '6' || el.text === 'SELECTED CASES')) return true;
                                            return false;
                                        }
                                        return true;
                                    });
                                }

                                // Place base exact elements
                                elementsToRender.forEach(el => {
                                    if (el.type === "img") {
                                        const imgPath = path.join(fidMediaDir, el.imgName);
                                        if (fs.existsSync(imgPath)) {
                                            s.addImage({
                                                path: imgPath,
                                                x: el.x, y: el.y, w: el.w, h: el.h
                                            });
                                        }
                                    } else if (el.type === "text") {
                                        s.addText(el.text, {
                                            x: el.x, y: el.y, w: el.w, h: el.h + 0.1,
                                            fontSize: el.fontSize,
                                            color: el.color,
                                            fontFace: L.FONT_BODY,
                                            bold: el.bold || false,
                                            align: el.align || "left",
                                            valign: "top",
                                            wrap: true,
                                        });
                                    } else if (el.type === "rect") {
                                        s.addShape(pres.shapes.RECTANGLE, {
                                            x: el.x, y: el.y, w: el.w, h: el.h,
                                            fill: { color: el.fill },
                                            line: { width: 0 }
                                        });
                                    } else if (el.type === "line") {
                                        const alpha = el.alpha !== undefined ? Math.round(el.alpha * 100) : 100;
                                        s.addShape(pres.shapes.LINE, {
                                            x: el.x, y: el.y, w: el.w, h: 0,
                                            line: { color: el.color || "3E3E40", width: 0.5, transparency: 100 - alpha }
                                        });
                                    }
                                });

                                // Draw exact dynamic overlays
                                if (fidSlideNum === 7) {
                                    const exactBios = [
                                        { name: "Ankur Rander", role: "Founder & CEO", img: "media/image-7-2.png" },
                                        { name: "Siddesh Pednekar", role: "Partner & COO", img: "media/image-10-2.png" },
                                        { name: "Jorge Andres Sierra", role: "Product Design Director", img: "media/image-7-3.png" },
                                        { name: "Alex Solod", role: "Product Design Director", img: "" },
                                        { name: "Sugesh Sugathan", role: "Associate Director, Design", img: "" }
                                    ];
                                    const bios = slide.bios || exactBios;
                                    bios.forEach((bio, i) => {
                                        const col = Math.floor(i / 3);
                                        const row = i % 3;
                                        const rectX = 2.0 + col * 3.07;
                                        const rectY = 1.46 + row * 1.20;
                                        const x = 2.16 + col * 3.07;
                                        const y = 1.61 + row * 1.20;
                                        s.addShape(pres.shapes.RECTANGLE, { x: rectX, y: rectY, w: 2.92, h: 1.04, fill: { color: "1D1D1F" }, line: { width: 0 } });

                                        if (bio.img) {
                                            if (bio.img.startsWith('data:image/')) {
                                                s.addImage({ data: bio.img, x, y, w: 0.73, h: 0.73, rounding: true });
                                            } else {
                                                const imgPath = path.join(BASE, bio.img);
                                                if (fs.existsSync(imgPath)) {
                                                    s.addImage({ path: imgPath, x, y, w: 0.73, h: 0.73, rounding: true });
                                                }
                                            }
                                        }
                                        s.addText(bio.name || "", { x: x + 0.85, y: y + 0.1, w: 2.0, h: 0.25, fontSize: 7.52, color: "FFFFFF", fontFace: L.FONT_BODY, bold: true, valign: "top" });
                                        s.addText(bio.role || "", { x: x + 0.85, y: y + 0.35, w: 2.0, h: 0.2, fontSize: 5.64, color: "FFFFFF", fontFace: L.FONT_BODY, valign: "top" });
                                    });
                                } else if (fidSlideNum === 8) {
                                    const exactCases = [
                                        { title: "Kotak Netbanking", desc: "22% increase in engagement" },
                                        { title: "AU Bank (Mobile App)", desc: "Reimagining Mobile banking for a universal bank" },
                                        { title: "TCX Platform", desc: "Simplifying enterprise connectivity at scale." },
                                        { title: "Kokuyo Camlin", desc: "2,00,000 artists joined the online artists community we envi" },
                                        { title: "Godrej Enterprises", desc: "₹189 B worth of business represented through the website." },
                                        { title: "StayVista", desc: "97% increase in conversion rate from checkout to payment." },
                                        { title: "AU Bank (Website)", desc: "80% increase in leads" },
                                        { title: "Threadspan", desc: "A real-time network monitoring platform." }
                                    ];
                                    const cases = slide.cases || exactCases;
                                    cases.forEach((c, i) => {
                                        const col = i % 4;
                                        const row = Math.floor(i / 4);
                                        const x = 2.00 + col * 1.85;
                                        const y = 1.38 + row * 1.72;

                                        if (c.img && c.img.startsWith('data:image/')) {
                                            s.addImage({ data: c.img, x, y, w: 1.69, h: 1.07 });
                                        } else if (c.img) {
                                            const imgPath = path.join(BASE, c.img);
                                            if (fs.existsSync(imgPath)) {
                                                s.addImage({ path: imgPath, x, y, w: 1.69, h: 1.07 });
                                            }
                                        } else {
                                            const imgName = `image-8-${(i % 8) + 2}.png`;
                                            const imgPath = path.join(fidMediaDir, imgName);
                                            if (fs.existsSync(imgPath)) {
                                                s.addImage({ path: imgPath, x, y, w: 1.69, h: 1.07 });
                                            }
                                        }

                                        s.addText(c.title || "", { x, y: y + 1.15, w: 1.69, h: 0.25, fontSize: 7.52, color: "1D1D1F", fontFace: L.FONT_BODY, bold: true, valign: "top" });
                                        s.addText(c.desc || "", { x, y: y + 1.4, w: 1.69, h: 0.5, fontSize: 5.64, color: "6A6A6B", fontFace: L.FONT_BODY, valign: "top" });
                                    });
                                }
                            }
                        }
                    }
                }
                else if (slide.layout === "CREDS_THE_SHIELD") {
                    s.addShape(pres.shapes.RECTANGLE, {
                        x: L.PRIMARY_X, y: contentY, w: 3.65, h: 3.55,
                        fill: { color: "034E48" }, line: { width: 0 }
                    });
                    const shieldLogo = path.join(BASE, "fidelity_unpacked/ppt/media/image-2-2.png");
                    if (fs.existsSync(shieldLogo)) {
                        s.addImage({ path: shieldLogo, x: L.PRIMARY_X + 0.8, y: contentY + 0.8, w: 2.0, h: 1.8 });
                    }
                    const points = slide.shieldPoints || ["10 Years building Digital Products", "80 Specialists in Strategy, Design & Engineering", "85% Client Repeat Rate"];
                    points.forEach((p, i) => {
                        s.addText(p, {
                            x: L.SECONDARY_X, y: contentY + i * 1.1, w: 3.8, h: 0.8,
                            fontSize: 10.5, color: colors.title, fontFace: L.FONT_BODY, valign: "top",
                            lineSpacing: "140%"
                        });
                    });
                }
                else if (slide.layout === "CREDS_DIAGNOSTICS") {
                    s.addShape(pres.shapes.RECTANGLE, {
                        x: L.PRIMARY_X, y: contentY, w: 3.65, h: 3.55,
                        fill: { color: theme === "dark" ? "000000" : "ECE9E4" }, line: { color: colors.line, width: 1 }
                    });
                    s.addText((slide.diagnosticTitleLeft || "THE PROBLEM").toUpperCase(), {
                        x: L.PRIMARY_X + 0.25, y: contentY + 0.2, w: 3.15, h: 0.3,
                        fontSize: 7.5, color: colors.accent, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, valign: "top"
                    });
                    s.addText(slide.diagnosticTextLeft || "", {
                        x: L.PRIMARY_X + 0.25, y: contentY + 0.7, w: 3.15, h: 2.65,
                        fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top", lineSpacing: "150%"
                    });
                    s.addShape(pres.shapes.RECTANGLE, {
                        x: L.SECONDARY_X, y: contentY, w: 3.65, h: 3.55,
                        fill: { color: "034E48" }, line: { width: 0 }
                    });
                    s.addText((slide.diagnosticTitleRight || "THE RECOMMENDATION").toUpperCase(), {
                        x: L.SECONDARY_X + 0.25, y: contentY + 0.2, w: 3.15, h: 0.3,
                        fontSize: 7.5, color: "ECE9E4", fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, valign: "top"
                    });
                    s.addText(slide.diagnosticTextRight || "", {
                        x: L.SECONDARY_X + 0.25, y: contentY + 0.7, w: 3.15, h: 2.65,
                        fontSize: 8.55, color: "ECE9E4", fontFace: L.FONT_BODY, valign: "top", lineSpacing: "150%"
                    });
                }
                else if (slide.layout === "CREDS_CAPABILITIES") {
                    const caps = slide.capabilities || [];
                    caps.forEach((cap, i) => {
                        const y = contentY + i * 1.15;
                        addRule(s, y + 1.05, theme, L.PRIMARY_X, L.WIDTH - L.LEFT_MARGIN);
                        s.addText((cap.label || "").toUpperCase(), { x: L.PRIMARY_X, y: y + 0.15, w: 2.0, h: 0.25, fontSize: 7.5, color: colors.accent, fontFace: L.FONT_BODY, bold: true, charSpacing: 1.2, valign: "top" });
                        s.addText(cap.title || "", { x: L.PRIMARY_X + 2.2, y: y + 0.15, w: 2.2, h: 0.25, fontSize: 9.5, color: colors.title, fontFace: L.FONT_TITLE, bold: true, valign: "top" });
                        s.addText(cap.desc || "", { x: L.PRIMARY_X + 4.6, y: y + 0.15, w: 3.2, h: 0.8, fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top", lineSpacing: "140%" });
                    });
                }
                else if (slide.layout === "CREDS_METRICS") {
                    const metrics = slide.metrics || [];
                    metrics.forEach((m, i) => {
                        const x = L.PRIMARY_X + i * 2.8;
                        s.addText(m.num || "", { x, y: contentY, w: 2.5, h: 0.8, fontSize: 28.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true, valign: "top" });
                        s.addText((m.title || "").toUpperCase(), { x, y: contentY + 0.9, w: 2.5, h: 0.25, fontSize: 7.5, color: colors.title, fontFace: L.FONT_TITLE, bold: true, charSpacing: 1.2, valign: "top" });
                        s.addText(m.desc || "", { x, y: contentY + 1.25, w: 2.5, h: 1.5, fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top", lineSpacing: "140%" });
                    });
                }
                else if (slide.layout === "CREDS_TEAM") {
                    const bios = slide.bios || [];
                    bios.forEach((bio, i) => {
                        const x = L.PRIMARY_X + i * 2.8;
                        const avatarPath = path.join(BASE, bio.img || "");
                        if (fs.existsSync(avatarPath)) {
                            s.addImage({ path: avatarPath, x, y: contentY, w: 1.2, h: 1.2 });
                        }
                        s.addText(bio.name || "", { x, y: contentY + 1.35, w: 2.6, h: 0.25, fontSize: 10.5, color: colors.title, fontFace: L.FONT_TITLE, bold: true, valign: "top" });
                        s.addText(bio.role || "", { x, y: contentY + 1.65, w: 2.6, h: 0.2, fontSize: 8.55, color: colors.accent, fontFace: L.FONT_BODY, valign: "top" });
                    });
                }
                else if (slide.layout === "CREDS_CASE_STUDIES") {
                    const cases = slide.cases || [];
                    cases.forEach((c, i) => {
                        const x = L.PRIMARY_X + i * 3.9;
                        s.addShape(pres.shapes.RECTANGLE, { x, y: contentY, w: 3.6, h: 3.5, fill: { color: theme === "dark" ? "1C1C1E" : "E1DDD8" }, line: { color: colors.line, width: 1 } });
                        s.addText(c.title || "", { x: x + 0.25, y: contentY + 0.25, w: 3.1, h: 0.4, fontSize: 10.5, color: colors.accent, fontFace: L.FONT_TITLE, bold: true, valign: "top" });
                        s.addText(c.desc || "", { x: x + 0.25, y: contentY + 0.8, w: 3.1, h: 2.2, fontSize: 8.55, color: colors.body, fontFace: L.FONT_BODY, valign: "top", lineSpacing: "150%" });
                    });
                }
                else if (slide.layout === "CREDS_CLIENTS") {
                    const clients = slide.clients || [];
                    clients.forEach((c, i) => {
                        const col = i % 3;
                        const row = Math.floor(i / 3);
                        const x = L.PRIMARY_X + col * 2.6;
                        const y = contentY + row * 1.2;
                        s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.4, h: 1.0, fill: { color: theme === "dark" ? "2C2C2E" : "E8E5E0" }, line: { color: colors.line, width: 1 } });
                        s.addText(c, { x: x + 0.1, y: y + 0.3, w: 2.3, h: 0.4, fontSize: 9.0, color: colors.title, fontFace: L.FONT_TITLE, align: "center", valign: "middle" });
                    });
                }
                else if (slide.layout === "CREDS_BDC_POV") {
                    const contentY = L.TITLE_DIVIDER_Y + 0.2;
                    const mainW = L.WIDTH - L.LEFT_MARGIN - L.PRIMARY_X;
                    s.addShape(pres.shapes.RECTANGLE, { x: L.PRIMARY_X, y: contentY, w: mainW, h: 3.4, fill: { color: "023D38" }, line: { color: "3E8D86", width: 1 } });
                    s.addText((slide.povTitle || "BOMBAYDC POINT OF VIEW").toUpperCase(), {
                        x: L.PRIMARY_X + 0.3, y: contentY + 0.3, w: mainW - 0.6, h: 0.3,
                        fontSize: 7.5, color: "3E8D86", fontFace: L.FONT_BODY, valign: "top"
                    });
                    s.addText(slide.povText || "", {
                        x: L.PRIMARY_X + 0.3, y: contentY + 0.8, w: mainW - 0.6, h: 2.2,
                        fontSize: 12.0, color: "ECE9E4", fontFace: L.FONT_BODY, valign: "top", lineSpacing: "150%"
                    });
                }
                else {
                    // Fallback for any unknown layout — prevents blank empty slides
                    const fallbackText = slide.body || slide.leftText || slide.rightText || slide.text || slide.tagline || "";
                    if (fallbackText) {
                        s.addText(formatColText(fallbackText, colors.body), {
                            x: L.PRIMARY_X, y: contentY + 0.1, w: 7.65, h: 3.5,
                            fontSize: 9.5, fontFace: L.FONT_BODY,
                            lineSpacing: "140%", valign: "top", margin: 0, fit: 'shrink'
                        });
                    }
                }
            }
        });
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // CLOSING SLIDE (Green theme - only if not already in dynamicSlides)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const lastSlide = plan.dynamicSlides ? plan.dynamicSlides[plan.dynamicSlides.length - 1] : null;
    const lastTitle = lastSlide ? (lastSlide.title || "").toLowerCase() : "";
    const alreadyHasClosing = lastTitle.includes("let's build") || lastTitle.includes("build what's next");

    if (!alreadyHasClosing) {
        const s = pres.addSlide();
        s.background = { color: L.THEMES.green.bg };
        addHeader(s, "green");

        s.addText("LET'S BUILD\nWHAT'S NEXT", {
            x: 2.0, y: 0.57, w: 7.84, h: 0.82,
            fontSize: 28.5, color: L.THEMES.green.title, fontFace: L.FONT_TITLE, lineSpacing: "110%", valign: "top"
        });

        s.addText("Explore our work, sectors, and point of view at www.bombaydc.com", {
            x: 2.0, y: 1.52, w: 7.84, h: 0.22,
            fontSize: 8.55, color: L.THEMES.green.title, fontFace: L.FONT_BODY, valign: "top"
        });

        addRule(s, 2.47, "green", 2.0, 9.84);

        s.addShape(pres.shapes.RECTANGLE, {
            x: 2.0, y: 2.62, w: 3.69, h: 1.67,
            fill: { color: "000000" }, line: { width: 0 }
        });

        if (fs.existsSync(ASSETS.closingHeadshot)) {
            s.addImage({ path: ASSETS.closingHeadshot, x: 2.31, y: 2.93, w: 1.04, h: 1.04 });
        } else {
            // Draw clean contact initial circle if image asset not present
            s.addShape(pres.shapes.OVAL, {
                x: 2.31, y: 2.93, w: 1.04, h: 1.04,
                fill: { color: "034E48" }, line: { color: "3E8D86", width: 1 }
            });
            s.addText("SP", {
                x: 2.31, y: 2.93, w: 1.04, h: 1.04,
                fontSize: 16, color: "ECE9E4", fontFace: L.FONT_TITLE, bold: true,
                align: "center", valign: "middle"
            });
        }

        s.addText("Siddesh Pednekar", {
            x: 3.52, y: 2.93, w: 3.0, h: 0.25,
            fontSize: 14.25, color: "FFFFFF", fontFace: L.FONT_BODY, valign: "top"
        });
        s.addText("Partner & COO", {
            x: 3.52, y: 3.23, w: 1.86, h: 0.15,
            fontSize: 8.55, color: "FFFFFF", fontFace: L.FONT_BODY, valign: "top"
        });
        s.addText("sid@bombaydc.com", {
            x: 3.52, y: 3.54, w: 1.86, h: 0.15,
            fontSize: 7.13, color: "FFFFFF", fontFace: L.FONT_BODY, valign: "top"
        });
        s.addText("9819981354", {
            x: 3.52, y: 3.74, w: 1.86, h: 0.15,
            fontSize: 7.13, color: "FFFFFF", fontFace: L.FONT_BODY, valign: "top"
        });

        s.addText("CONFIDENTIAL AND PROPRIETARY | © BombayDC.\nThis material is intended solely for your internal use and any use of this material without specific permission of BombayDC is strictly prohibited. All rights reserved.", {
            x: 2.0, y: 4.95, w: 7.5, h: 0.35, fontSize: 4.27, color: "FFFFFF", fontFace: L.FONT_BODY, lineSpacing: "140%", valign: "top"
        });
    }

    outputPath = outputPath || path.join(BASE, "Generated_AI_Proposal.pptx");
    console.log("Saving presentation to:", outputPath);
    await pres.writeFile({ fileName: outputPath });
    console.log("Successfully compiled BOMBAYDC deck!");
}

if (require.main === module) {
    generatePPT().catch(console.error);
}

module.exports = { generatePPT };
