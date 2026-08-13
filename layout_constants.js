"use strict";

const LayoutConstants = {
    // Screen aspect ratio: 16:9
    WIDTH: 10.0,
    HEIGHT: 5.625,

    // Alignments
    LEFT_MARGIN: 0.16,
    PRIMARY_X: 1.99,
    SECONDARY_X: 5.95,
    DIVIDER_X: 5.85,
    COVER_X: 0.5,

    // Header dimensions
    HDR_Y: 0.0,
    HDR_W: 9.69,
    HDR_H: 0.57,

    // Title / Subtitle positioning
    TITLE_Y: 0.57,
    TITLE_H: 0.25,
    TAGLINE_Y: 0.92,
    TAGLINE_H: 0.35,
    TITLE_DIVIDER_Y: 1.35,

    // Section starter layout
    STARTER_NUM_Y: 2.1,
    STARTER_NUM_FS: 68,
    STARTER_LINE_Y: 3.3,
    STARTER_LINE_W: 6.86,
    STARTER_TITLE_Y: 3.5,
    STARTER_TITLE_FS: 23,

    // Row layout configurations
    ROW_SPACING_DEFAULT: 0.72,
    ROW_SPACING_3ROW: 0.9,
    ROW_TEXT_H: 0.45,

    // Lines & dividers
    LINE_WIDTH: 0.5,
    LINE_TRANSPARENCY: 40, // 40% transparency in PPTX (CSS opacity = 0.6)

    // Palette definition - BombayDC Design System
    THEMES: {
        dark: {
            bg: "1D1D1F",
            title: "ECE9E4",
            body: "B4B4B4",
            accent: "4DB89A", // Light readable green/teal
            line: "FFFFFF",
            lineOpacity: 0.2
        },
        light: {
            bg: "ECE9E4",
            title: "1D1D1F",
            body: "6A6A6B",
            accent: "034E48", // Dark forest green
            line: "B4B4B4",
            lineOpacity: 0.6
        },
        green: {
            bg: "034E48", // Brand green
            title: "ECE9E4",
            body: "B4B4B4",
            accent: "4DB89A",
            line: "3E8D86",
            lineOpacity: 0.6
        }
    },

    // Fonts
    FONT_BODY: "Inter",
    FONT_TITLE: "Inter Medium"
};

// Export for Node.js or expose to window in the browser
if (typeof module !== "undefined" && module.exports) {
    module.exports = LayoutConstants;
} else {
    window.LayoutConstants = LayoutConstants;
}
