/** @jsxImportSource react */
// "XLSX Editor by EigenPal" lockup, after docx-editor/examples/shared/BrandLogo.tsx at a0825f1f
// and pptx-editor/examples/vite/src/BrandLogo.tsx at 5b6ead87 (EigenPal). DOCX Editor's icon is
// a document page and PPTX Editor's is a slide. This one is a worksheet in the same family
// style: the same outline, paper, and label badge, with a cell grid under a header row, one
// selected cell, and a `.XLSX` label in green (#107C41) instead of blue (#185ABD) or orange
// (#C43E1C). The same artwork is in
// `.github/assets/readme-logo-*.svg`; `examples/vite/public/icon.svg` is a simpler 64 px version.
//
// Differences, each on purpose:
// - The link goes to eigenpal.com and shows the product name only, as in PPTX Editor. Its
//   `title` says "XLSX Editor by EigenPal", and its `aria-label` also says that it opens in a
//   new tab.
// - The EigenPal mark is not in the lockup, as in PPTX Editor, so `EigenpalLogo` is not here.
// - `BrandLogo` takes no `name` prop: the demo's text is English and stays out of
//   `packages/i18n`, so the name is the literal "XLSX Editor".
// The JSX import source pragma on the first line is added, as every React file here has it.

function XlsxIcon() {
  return (
    <svg
      viewBox="0 0 100 84"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
      style={{ width: '30px', height: 'auto', flexShrink: 0 }}
    >
      <path
        d="M2 2H98V60H2V2Z"
        fill="#FAF9F8"
        stroke="#201F1E"
        strokeWidth="4"
        strokeLinejoin="round"
      />
      <rect x="4" y="4" width="92" height="12" fill="#DFF0E5" />
      <rect x="4" y="16" width="92" height="2" fill="#201F1E" opacity="0.35" />
      <rect x="4" y="30" width="92" height="2" fill="#201F1E" opacity="0.15" />
      <rect x="4" y="44" width="92" height="2" fill="#201F1E" opacity="0.15" />
      <rect x="34" y="4" width="2" height="54" fill="#201F1E" opacity="0.15" />
      <rect x="66" y="4" width="2" height="54" fill="#201F1E" opacity="0.15" />
      <rect x="11" y="8" width="16" height="4" fill="#201F1E" opacity="0.35" />
      <rect x="43" y="8" width="16" height="4" fill="#201F1E" opacity="0.35" />
      <rect x="75" y="8" width="14" height="4" fill="#201F1E" opacity="0.35" />
      <rect x="11" y="23" width="18" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="11" y="37" width="14" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="11" y="51" width="16" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="47" y="23" width="14" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="49" y="51" width="12" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="77" y="23" width="14" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="73" y="37" width="18" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="79" y="51" width="12" height="3" fill="#201F1E" opacity="0.15" />
      <rect x="36" y="32" width="30" height="12" fill="#107C41" opacity="0.2" />
      <rect
        x="36"
        y="32"
        width="30"
        height="12"
        stroke="#107C41"
        strokeWidth="3"
        strokeLinejoin="round"
      />
      <rect x="28" y="66" width="44" height="16" fill="#107C41" stroke="#201F1E" strokeWidth="3" />
      <text
        x="50"
        y="78"
        textAnchor="middle"
        fill="#FAF9F8"
        fontSize="11"
        fontWeight="900"
        fontFamily="system-ui, sans-serif"
      >
        .XLSX
      </text>
    </svg>
  );
}

/** The lockup: the worksheet icon and the product name. */
export function BrandLogo() {
  return (
    <a
      className="demo-brand"
      href="https://eigenpal.com"
      target="_blank"
      rel="noopener noreferrer"
      title="XLSX Editor by EigenPal"
      aria-label="XLSX Editor by EigenPal (opens in a new tab)"
    >
      <XlsxIcon />
      <span className="demo-brand__name">XLSX Editor</span>
    </a>
  );
}
