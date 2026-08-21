export interface PptxSceneSize {
  height: number;
  width: number;
}

export interface PptxSceneTransform extends PptxSceneSize {
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  rotation?: number;
  x: number;
  y: number;
}

export interface PptxSceneCoordinateSpace extends PptxSceneSize {
  x: number;
  y: number;
}

export interface PptxSceneGroupTransform extends PptxSceneTransform {
  childSpace: PptxSceneCoordinateSpace;
}

export interface PptxSceneAuthoredElement {
  fillColor?: string;
  geometry?: 'ellipse' | 'rect' | 'roundRect';
  hidden?: boolean;
  lineColor?: string;
  lineWidth?: number;
  transform?: PptxSceneTransform;
}

export interface PptxSceneResolvedElement {
  hidden: boolean;
  transform?: PptxSceneTransform;
}

export interface PptxScenePlaceholder {
  hasCustomPrompt?: boolean;
  index?: number;
  orientation?: 'horizontal' | 'vertical';
  prompt?: string;
  role: 'layout-definition' | 'master-definition' | 'slide-instance';
  size?: 'full' | 'half' | 'quarter';
  sourceKey?: string;
  type?: string;
}

export interface PptxSceneElementBase {
  authored: PptxSceneAuthoredElement;
  description?: string;
  key: string;
  name?: string;
  placeholder?: PptxScenePlaceholder;
  resolved: PptxSceneResolvedElement;
  title?: string;
}

export interface PptxSceneTextBodyProperties {
  anchor?: 'bottom' | 'center' | 'distributed' | 'justified' | 'top';
  autoFit?: 'none' | 'shape' | 'text';
  vertical?: boolean;
  wrap?: boolean;
}

export interface PptxSceneRunProperties {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  italic?: boolean;
  language?: string;
}

export interface PptxSceneTextRun {
  key: string;
  preserveSpace?: boolean;
  properties?: PptxSceneRunProperties;
  text: string;
  type: 'run';
}

export interface PptxSceneTextField {
  fieldType: string;
  key: string;
  properties?: PptxSceneRunProperties;
  text: string;
  type: 'field';
}

export interface PptxSceneTextBreak {
  key: string;
  properties?: PptxSceneRunProperties;
  type: 'break';
}

export type PptxSceneTextNode =
  PptxSceneTextBreak | PptxSceneTextField | PptxSceneTextRun;

export interface PptxSceneParagraphProperties {
  alignment?: 'center' | 'distributed' | 'justify' | 'left' | 'right';
  level?: number;
}

export interface PptxSceneParagraph {
  children: PptxSceneTextNode[];
  endProperties?: PptxSceneRunProperties;
  key: string;
  properties?: PptxSceneParagraphProperties;
}

export interface PptxSceneTextBody {
  body: PptxSceneTextBodyProperties;
  paragraphs: PptxSceneParagraph[];
}

export interface PptxSceneTextElement extends PptxSceneElementBase {
  text: PptxSceneTextBody;
  type: 'text';
}

export interface PptxSceneShapeElement extends PptxSceneElementBase {
  type: 'shape';
}

export interface PptxSceneImageCrop {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export interface PptxSceneImageElement extends PptxSceneElementBase {
  crop?: PptxSceneImageCrop;
  mediaKey?: string;
  type: 'image';
}

export type PptxSceneChartType =
  'barChart' | 'doughnutChart' | 'lineChart' | 'pieChart';

export interface PptxSceneChartSeries {
  categories: string[];
  color?: string;
  key: string;
  name: string;
  values: number[];
}

export interface PptxSceneChartElement extends PptxSceneElementBase {
  barDirection?: 'bar' | 'col';
  chartType: PptxSceneChartType;
  grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
  holeSize?: number;
  marker?: boolean;
  series: PptxSceneChartSeries[];
  type: 'chart';
}

export interface PptxSceneTableBorder {
  color: string;
  style?: 'dashed' | 'dotted' | 'solid';
  width: number;
}

export interface PptxSceneTableBorders {
  bottom?: PptxSceneTableBorder;
  left?: PptxSceneTableBorder;
  right?: PptxSceneTableBorder;
  top?: PptxSceneTableBorder;
}

export interface PptxSceneTableCell {
  borders?: PptxSceneTableBorders;
  colSpan?: number;
  fillColor?: string;
  hMerge?: boolean;
  rowSpan?: number;
  text: PptxSceneTextBody;
  vMerge?: boolean;
}

export interface PptxSceneTableRow {
  cells: PptxSceneTableCell[];
  height: number;
}

export interface PptxSceneTableElement extends PptxSceneElementBase {
  columns: number[];
  rows: PptxSceneTableRow[];
  type: 'table';
}

export interface PptxSceneGroupElement extends Omit<
  PptxSceneElementBase,
  'authored' | 'resolved'
> {
  authored: Omit<PptxSceneAuthoredElement, 'transform'> & {
    transform?: PptxSceneGroupTransform;
  };
  elements: PptxSceneElement[];
  resolved: Omit<PptxSceneResolvedElement, 'transform'> & {
    transform?: PptxSceneGroupTransform;
  };
  type: 'group';
}

export interface PptxSceneUnsupportedElement extends PptxSceneElementBase {
  feature: string;
  previewText?: string;
  type: 'unsupported';
}

export type PptxSceneElement =
  | PptxSceneChartElement
  | PptxSceneGroupElement
  | PptxSceneImageElement
  | PptxSceneShapeElement
  | PptxSceneTableElement
  | PptxSceneTextElement
  | PptxSceneUnsupportedElement;

export interface PptxSceneMedia {
  data: Uint8Array;
  key: string;
  mimeType: 'image/jpeg' | 'image/png';
}

export interface PptxSceneTheme {
  key: string;
  name?: string;
}

export interface PptxSceneMaster {
  elements: PptxSceneElement[];
  key: string;
  name?: string;
  themeKey: string;
}

export interface PptxSceneLayout {
  elements: PptxSceneElement[];
  key: string;
  masterKey: string;
  name?: string;
}

export interface PptxSceneSlide {
  backgroundColor?: string;
  elements: PptxSceneElement[];
  hidden?: boolean;
  key: string;
  layoutKey?: string;
  name?: string;
}

export interface PptxSceneDocument {
  layouts: PptxSceneLayout[];
  masters: PptxSceneMaster[];
  media: PptxSceneMedia[];
  schemaVersion: 2;
  size: PptxSceneSize;
  slides: PptxSceneSlide[];
  themes: PptxSceneTheme[];
}

export type PptxSceneValidationCode =
  | 'duplicate-public-key'
  | 'invalid-hierarchy-reference'
  | 'invalid-numeric-value'
  | 'invalid-office-text-escape'
  | 'resource-limit-exceeded'
  | 'invalid-scene-document'
  | 'unsupported-feature'
  | 'unsupported-schema-version';

export interface PptxSceneValidationIssue {
  code: PptxSceneValidationCode;
  message: string;
  path: string;
}

export interface PptxSceneValidationOptions {
  profile?: 'create-native-v1' | 'create-text-v1' | 'scene';
}

export interface PptxSceneValidationResult {
  issues: PptxSceneValidationIssue[];
  valid: boolean;
}
