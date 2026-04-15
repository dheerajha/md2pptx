/**
 * themes.js — Color palettes and typography for each theme
 */

const PALETTES = {
  midnight: {
    primary: "1E2761",
    secondary: "CADCFC",
    accent: "4FC3F7",
    dark: "0D1B4B",
    light: "EEF4FF",
    text: "1A1A2E",
    white: "FFFFFF",
    chartColors: ["4FC3F7", "CADCFC", "7B8CDE", "A8D8EA", "1E2761", "6C63FF"],
  },
  ocean: {
    primary: "065A82",
    secondary: "1C7293",
    accent: "68C3E8",
    dark: "021527",
    light: "E8F5FA",
    text: "1A2F3A",
    white: "FFFFFF",
    chartColors: ["068BC7", "1C7293", "68C3E8", "A8D8EA", "065A82", "0D3B5E"],
  },
  teal: {
    primary: "028090",
    secondary: "00A896",
    accent: "02C39A",
    dark: "013A40",
    light: "E0F5F3",
    text: "012E33",
    white: "FFFFFF",
    chartColors: ["02C39A", "00A896", "028090", "05668D", "028A7B", "78C8C8"],
  },
  forest: {
    primary: "2C5F2D",
    secondary: "97BC62",
    accent: "5B9B2F",
    dark: "1A3A1B",
    light: "EDF5E1",
    text: "1A2E1A",
    white: "FFFFFF",
    chartColors: ["5B9B2F", "97BC62", "2C5F2D", "A8C686", "4A8B23", "C8DDA0"],
  },
  coral: {
    primary: "F96167",
    secondary: "2F3C7E",
    accent: "F9E795",
    dark: "1A2050",
    light: "FFF5F5",
    text: "1A0A0A",
    white: "FFFFFF",
    chartColors: ["F96167", "F9E795", "2F3C7E", "F47C80", "6B74B5", "FAD02E"],
  },
  charcoal: {
    primary: "36454F",
    secondary: "607D8B",
    accent: "00BCD4",
    dark: "1A2227",
    light: "F5F7F8",
    text: "212121",
    white: "FFFFFF",
    chartColors: ["00BCD4", "607D8B", "36454F", "80CBC4", "455A64", "B0BEC5"],
  },
};

const FONTS = {
  heading: "Calibri",
  body: "Calibri",
};

module.exports = { PALETTES, FONTS };
