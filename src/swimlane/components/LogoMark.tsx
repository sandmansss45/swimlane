import * as React from 'react';

// Hand-recreated approximation of the real QLE logo (three overlapping
// orbit rings + a small accent dot) since only a pasted chat image was
// available, not the actual logo file. Swap the <svg> body below for an
// <img src="/qle-logo.svg"> (or similar) once the real asset file is
// provided - see AUTH_SETUP.md-adjacent conversation history for context.
const LogoMark: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <ellipse cx="50" cy="50" rx="34" ry="16" stroke="#3f6fe0" strokeWidth="5" transform="rotate(0 50 50)" />
    <ellipse cx="50" cy="50" rx="34" ry="16" stroke="#3f6fe0" strokeWidth="5" transform="rotate(60 50 50)" />
    <ellipse cx="50" cy="50" rx="34" ry="16" stroke="#3f6fe0" strokeWidth="5" transform="rotate(120 50 50)" />
    <circle cx="62" cy="58" r="6" fill="#4ade80" />
  </svg>
);

export default LogoMark;
