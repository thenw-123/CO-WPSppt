export type PptDslVersion = '2.0';

export type SlideType =
  | 'cover'
  | 'agenda'
  | 'section'
  | 'content'
  | 'left-image'
  | 'argument'
  | 'comparison'
  | 'timeline'
  | 'data'
  | 'swot'
  | 'custom';

export type ElementType = 'text' | 'bullets' | 'image' | 'chart' | 'table';
export type ElementRole =
  | 'title'
  | 'subtitle'
  | 'body'
  | 'claim'
  | 'evidence'
  | 'insight'
  | 'caption'
  | 'quote'
  | 'footer'
  | 'decorative';

export interface PptDsl {
  dslVersion: PptDslVersion;
  meta: {
    title: string;
    subtitle?: string;
    audience?: string;
    purpose?: string;
    language?: string;
    author?: string;
    savePath?: string;
    tags?: string[];
  };
  theme?: {
    preset?: 'business-blue' | 'clean-light' | 'dark-executive' | 'academic' | 'custom';
    template?: { id?: string; path?: string; variant?: string };
    fonts?: { title?: string; body?: string; mono?: string };
    colors?: Record<string, string>;
    motion?: DslMotion;
  };
  defaults?: {
    layout?: 'auto' | string;
    imagePlacement?: 'auto' | 'right' | 'bottom' | 'fullBleed';
    chartEngine?: 'auto' | 'png' | 'com';
    assetFetch?: Record<string, unknown>;
  };
  slides: DslSlide[];
}

export interface DslCardChrome {
  surface?: 'none' | 'muted' | 'elevated' | 'accent';
  background?: string;
  borderColor?: string;
  padding?: number;
  cornerRadius?: number;
}

export interface DslSlide {
  id?: string;
  type: SlideType;
  message?: string;
  title?: string;
  layout?: {
    name?:
      | 'auto'
      | 'title'
      | 'title-content'
      | 'section'
      | 'two-column'
      | 'three-column'
      | 'image-right'
      | 'left-image'
      | 'image-bottom'
      | 'full-bleed-image'
      | 'comparison-matrix'
      | 'timeline-horizontal'
      | 'claim-with-evidence'
      | 'swot-grid'
      | 'data-insight'
      | 'grid'
      | 'comparison'
      | 'custom';
    density?: 'compact' | 'normal' | 'spacious';
    templateSlot?: string;
    /** Grid column count when using `grid` layout (2–4 typical). */
    columns?: number;
    /** Gap between grid cells (slide points). */
    gap?: number;
    /** Default card chrome for grid / three-column; `element.style.card` overrides. */
    card?: DslCardChrome;
  };
  style?: DslStyle;
  elements: DslElement[];
  notes?: string;
  motion?: DslMotion;
  constraints?: { maxBullets?: number; maxTextLines?: number; mustFitOneSlide?: boolean };
}

export type DslElement =
  | { id?: string; type: 'text'; role?: ElementRole; text: string; style?: DslStyle; layoutHint?: DslLayoutHint }
  | {
      id?: string;
      type: 'bullets';
      role?: ElementRole;
      items?: Array<string | DslBulletItem>;
      content?: { items: Array<string | DslBulletItem> };
      style?: DslStyle;
      layoutHint?: DslLayoutHint;
    }
  | { id?: string; type: 'image'; role?: ElementRole; source: DslImageSource; alt?: string; style?: DslStyle; layoutHint?: DslLayoutHint }
  | { id?: string; type: 'chart'; role?: ElementRole; chart: DslChart; alt?: string; style?: DslStyle; layoutHint?: DslLayoutHint }
  | { id?: string; type: 'table'; role?: ElementRole; table: DslTable; style?: DslStyle; layoutHint?: DslLayoutHint };

export interface DslBulletItem {
  text: string;
  level?: 1 | 2 | 3;
  evidence?: string;
}

export interface DslImageSource {
  kind: 'local' | 'remote' | 'commons' | 'generated' | 'template';
  path?: string;
  url?: string;
  title?: string;
  prompt?: string;
  assetId?: string;
  credit?: string;
  policy?: 'default' | 'trustedCdn' | 'strict';
}

export interface DslChart {
  kind: 'bar' | 'line' | 'donut' | 'pie' | 'area' | 'scatter';
  title?: string;
  data: {
    labels: string[];
    series: Array<{ name?: string; values: number[] }>;
  };
  insight?: string;
  render?: { engine?: 'auto' | 'png' | 'com'; fallback?: 'text' | 'png' | 'none' };
}

export interface DslTable {
  title?: string;
  columns: string[];
  rows: Array<Array<string | number | boolean | null>>;
  highlight?: { rows?: number[]; columns?: number[] };
}

export interface DslMotion {
  transition?: 'none' | 'fade' | 'push' | 'wipe' | 'cut' | 'uncover' | 'split' | 'cover' | 'random' | 'blinds' | 'dissolve';
  duration?: number;
  build?: 'none' | 'byElement' | 'byParagraph';
  effect?: 'fade' | 'appear' | 'fly' | 'float' | 'wipe' | 'zoom';
}

export interface DslStyle {
  variant?: 'default' | 'primary' | 'secondary' | 'accent' | 'muted' | 'danger' | 'success';
  background?: string;
  color?: string;
  fontSize?: number;
  minFontSize?: number;
  weight?: 'regular' | 'medium' | 'bold';
  imageFit?: 'contain' | 'fill';
  card?: DslCardChrome;
}

export interface DslLayoutHint {
  placement?: 'auto' | 'left' | 'right' | 'top' | 'bottom' | 'center' | 'background' | 'fullBleed';
  emphasis?: 'low' | 'normal' | 'high';
  span?: 'auto' | 'full' | 'half' | 'third';
}

