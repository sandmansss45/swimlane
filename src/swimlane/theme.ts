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
  decision: '#c96a1f',
  textPrimary: '#16233d',
  textSecondary: '#5b6b84',
  border: '#dde3ea',
  surface: '#ffffff',
  canvas: '#f5f7fb'
};

export const swimlaneTheme: ITheme = createTheme({
  defaultFontStyle: { fontFamily: 'inherit' },
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
    neutralLighterAlt: '#f5f7fb',
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
