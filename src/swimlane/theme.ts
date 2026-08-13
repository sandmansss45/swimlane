import { createTheme, ITheme } from '@fluentui/react';

// Single source of truth for the app's brand colors - every Fluent
// component (buttons, dropdowns, tabs, message bars) picks these up
// automatically via ThemeProvider, so hand-styled pieces (ShapeNode,
// SwimlaneCanvas, the picker cards) just need to match these by eye
// rather than every component needing its own override.
export const BRAND = {
  navy: '#0b2148',
  blue: '#1f4fa3',
  blueLight: '#e6f0ff',
  approval: '#146c43',
  greenLight: '#e0f5ea',
  decision: '#c96a1f',
  document: '#5b6478',
  riskHigh: '#d92d20',
  riskMedium: '#dc6803',
  riskLow: '#17803d',
  textPrimary: '#16233d',
  textSecondary: '#5b6b84',
  border: '#dde3ea',
  surface: '#ffffff',
  canvas: '#eaf2fd',
  // Lifted from the real QLE logo (dark navy background, brighter
  // orbit-ring blue, small emerald accent dot) - used only on the
  // sign-in screen. The rest of the app stays on the light BRAND
  // palette above: this dark bg would hurt readability across the
  // dense swimlane grid, and the logo's green would visually collide
  // with green's actual meaning there (Approval shapes).
  darkBg: '#0a1420',
  darkSurface: '#111d33',
  brightBlue: '#3f6fe0',
  accentGreen: '#4ade80'
};

export const swimlaneTheme: ITheme = createTheme({
  defaultFontStyle: { fontFamily: 'inherit' },
  // Fluent's own default (2px) reads as boxy/dated next to the rest of
  // the app's hand-styled cards and buttons, which all use 8-16px radii -
  // this brings every Fluent-rendered button, input, and dropdown in line
  // with that same softer look instead of standing out as "the generic
  // component library one".
  effects: { roundedCorner2: '6px' },
  palette: {
    themePrimary: BRAND.blue,
    themeLighterAlt: '#f2f7fd',
    themeLighter: '#cddef4',
    themeLight: '#a5c2e8',
    themeTertiary: '#5c8bcd',
    themeSecondary: '#2f66b6',
    themeDarkAlt: '#1c479390',
    themeDark: '#173c7c',
    themeDarker: '#112c5c',
    neutralLighterAlt: BRAND.canvas,
    neutralLighter: '#eef1f6',
    neutralLight: '#e2e7ee',
    neutralQuaternaryAlt: '#d3d9e2',
    neutralQuaternary: '#c9d0da',
    neutralTertiaryAlt: '#bac2cd',
    neutralTertiary: BRAND.textSecondary,
    neutralSecondary: '#3c4a63',
    neutralPrimaryAlt: '#273656',
    neutralPrimary: BRAND.textPrimary,
    neutralDark: BRAND.navy,
    black: '#0a1526',
    white: '#ffffff'
  }
});

// Scoped to just the sign-in screen (nested ThemeProvider) so its dark,
// bright-blue button doesn't bleed into the rest of the app's lighter
// palette above. isInverted:true flips Fluent's text/background
// semantic colors for a dark surface automatically.
export const signInTheme: ITheme = createTheme({
  isInverted: true,
  defaultFontStyle: { fontFamily: 'inherit' },
  effects: { roundedCorner2: '6px' },
  palette: {
    themePrimary: BRAND.brightBlue,
    themeLighterAlt: '#0a1420',
    themeLighter: '#152a4f',
    themeLight: '#274e9e',
    themeTertiary: '#3f6fe0',
    themeSecondary: '#5a86e8',
    themeDarkAlt: '#5a86e8',
    themeDark: '#7ea0ed',
    themeDarker: '#a3bcf2',
    neutralLighterAlt: BRAND.darkSurface,
    neutralLighter: '#16233d',
    neutralLight: '#1c2c4a',
    neutralQuaternaryAlt: '#233457',
    neutralQuaternary: '#2a3c62',
    neutralTertiaryAlt: '#5b6b84',
    neutralTertiary: '#9fb0cc',
    neutralSecondary: '#cdd7e8',
    neutralPrimaryAlt: '#e2e8f4',
    neutralPrimary: '#ffffff',
    neutralDark: '#f2f5fa',
    black: '#ffffff',
    white: BRAND.darkBg
  }
});
