/** @jsxImportSource vue */
// Inline SVG icons, as AGENTS.md "Icons" asks: Material Symbols path data (Apache-2.0), 24 by
// 960 view box. The Vue twin of the React adapter's `icons.tsx`, with the same paths.

const ADD = 'M440-440H200v-80h240v-240h80v240h240v80H520v240h-80v-240Z';
const REMOVE = 'M200-440v-80h560v80H200Z';

function Icon(props: { readonly d: string }) {
  return (
    <svg viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
      <path d={props.d} />
    </svg>
  );
}

export const ZoomInIcon = () => <Icon d={ADD} />;
export const ZoomOutIcon = () => <Icon d={REMOVE} />;
export const AddIcon = () => <Icon d={ADD} />;
