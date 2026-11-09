import * as React from 'react';

// ─── BIAT IT Brand Palette ────────────────────────────────────────────────────
export const BIAT_NAVY   = '#1B3669';
export const BIAT_TEAL   = '#00B4D8';
export const BIAT_WHITE  = '#FFFFFF';
export const BIAT_BG     = '#F2F5F9';
export const BIAT_BORDER = '#D0DCE8';

// ─── Shared layout styles (inline, no SCSS dependency) ───────────────────────
export const biatContainerStyle: React.CSSProperties = {
  fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
  background: BIAT_BG,
  padding: '0 0 28px 0',
};

export const biatCardStyle: React.CSSProperties = {
  background: BIAT_WHITE,
  borderRadius: 0,
  padding: '22px 28px',
  boxShadow: '0 4px 24px rgba(27, 54, 105, 0.10)',
};

export const biatToolbarStyle: React.CSSProperties = {
  paddingBottom: 16,
  marginBottom: 4,
  borderBottom: `1px solid ${BIAT_BORDER}`,
};

// ─── Fluent UI button style overrides ────────────────────────────────────────
// fontWeight must be a string in Fluent UI's IStyle / IRawStyle
// Use with: <PrimaryButton styles={biatPrimaryBtn} .../>
export const biatPrimaryBtn = {
  root:         { background: BIAT_NAVY, borderColor: BIAT_NAVY, borderRadius: 5, minWidth: 120 },
  rootHovered:  { background: '#2A4D8E', borderColor: '#2A4D8E' },
  rootPressed:  { background: '#0D2450', borderColor: '#0D2450' },
  rootDisabled: { background: '#9BA8BF', borderColor: '#9BA8BF' },
  label:        { color: BIAT_WHITE, fontWeight: '600' as const },
};

// Use with: <DefaultButton styles={biatDefaultBtn} .../>
export const biatDefaultBtn = {
  root:        { borderColor: BIAT_NAVY, color: BIAT_NAVY, borderRadius: 5, background: BIAT_WHITE },
  rootHovered: { background: '#EEF2F8', borderColor: '#2A4D8E', color: '#2A4D8E' },
  rootPressed: { background: '#DDE6F0', borderColor: BIAT_NAVY },
  label:       { color: BIAT_NAVY, fontWeight: '500' as const },
};

// Use with: <DefaultButton styles={biatDangerBtn} .../> for destructive actions
export const biatDangerBtn = {
  root:        { borderColor: '#C0392B', color: '#C0392B', borderRadius: 5, background: BIAT_WHITE },
  rootHovered: { background: '#FDF1F0', borderColor: '#A93226', color: '#A93226' },
  label:       { color: '#C0392B', fontWeight: '500' as const },
};

// ─── BiatHeader component ─────────────────────────────────────────────────────

interface IBiatHeaderProps {
  /** Main page title displayed next to the logo */
  title: string;
  /** Optional secondary line below the title */
  subtitle?: string;
}

import biatLogo from '../assets/biat-logo.png';

const BiatHeader: React.FC<IBiatHeaderProps> = ({ title, subtitle }) => (
  <div
    style={{
      background: BIAT_WHITE,
      padding: '0 28px',
      display: 'flex',
      alignItems: 'center',
      gap: 22,
      borderRadius: 0,
      boxShadow: '0 2px 8px rgba(27, 54, 105, 0.10)',
      borderBottom: `3px solid ${BIAT_TEAL}`,
      minHeight: 70,
    }}
  >
    {/* ── Logo image ───────────────────────────────────────────────────────── */}
    <img
      src={biatLogo}
      alt="BIAT Innovation & Technology"
      style={{ height: 46, flexShrink: 0 }}
    />

    {/* ── Divider ─────────────────────────────────────────────────────────── */}
    <div
      style={{
        width: 1,
        height: 38,
        background: BIAT_BORDER,
        flexShrink: 0,
        margin: '0 4px',
      }}
    />

    {/* ── Page title ──────────────────────────────────────────────────────── */}
    <div>
      <div
        style={{
          color: BIAT_NAVY,
          fontSize: 18,
          fontWeight: 700,
          letterSpacing: 0.2,
          fontFamily: "'Segoe UI', sans-serif",
          lineHeight: 1.2,
        }}
      >
        {title}
      </div>
      {subtitle && (
        <div
          style={{
            color: '#6B7A99',
            fontSize: 12,
            marginTop: 3,
            fontFamily: "'Segoe UI', sans-serif",
          }}
        >
          {subtitle}
        </div>
      )}
    </div>
  </div>
);

export default BiatHeader;
