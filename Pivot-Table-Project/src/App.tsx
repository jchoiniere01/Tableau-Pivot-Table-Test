import { useEffect, useMemo, useRef, useState } from 'react';

declare const tableau: any;
declare const XLSX: any;

type PreviewRow = Record<string, any>;

type FieldRole = 'dimension' | 'measure';

type FieldMeta = {
  fieldName: string;
  caption: string;
  dataType?: string;
  role: FieldRole;
};

type TreeNode = {
  id: string;
  label: string;
  level: number;
  dimensionField: string;
  pathValues: string[];
  measures: Record<string, number>;
  children: TreeNode[];
};

type VisibleNode = {
  id: string;
  label: string;
  level: number;
  dimensionField: string;
  pathValues: string[];
  measures: Record<string, number>;
  hasChildren: boolean;
};

type DateRangeState = {
  field: string;
  start: string;
  end: string;
  preset: string;
};

type SimpleFilterState = {
  field: string;
  values: string[];
};

type MeasureFormat = {
  formatType: 'number' | 'currency' | 'percent';
  decimals: number;
  prefix: string;
  suffix: string;
  useThousandsSeparator: boolean;
  alignment: 'left' | 'center' | 'right';
};

type LayoutSettings = {
  hierarchyWidth: number;
  defaultMeasureWidth: number;
  measureWidths: Record<string, number>;
};

type GroupedFieldSection = {
  label: string;
  fields: FieldMeta[];
};

type FilterGroupDefinition = {
  label: string;
  field: string;
  options: string[];
};

type InfoCalculation = {
  id: string;
  label: string;
  worksheet: string;
  field: string;
  aggregation: 'MAX' | 'MIN' | 'SUM' | 'AVG' | 'COUNT';
  format: 'text' | 'date' | 'number' | 'currency';
};

type InfoResult = {
  id: string;
  label: string;
  value: string;
};

type WorksheetFieldsMap = Record<string, FieldMeta[]>;
type FieldRoleOverrideMap = Record<string, Record<string, FieldRole>>;
type WorksheetFieldExpandMap = Record<string, boolean>;

const DEFAULT_MEASURE_FORMAT: MeasureFormat = {
  formatType: 'number',
  decimals: 0,
  prefix: '',
  suffix: '',
  useThousandsSeparator: true,
  alignment: 'right'
};

const DEFAULT_LAYOUT: LayoutSettings = {
  hierarchyWidth: 320,
  defaultMeasureWidth: 140,
  measureWidths: {}
};

const MIN_HIERARCHY_WIDTH = 180;
const MIN_MEASURE_WIDTH = 90;

const DATE_PRESETS = [
  'Custom Date Range',
  'Today',
  'Yesterday',
  'Last 7 Days',
  'Last 30 Days',
  'This Month',
  'Last Month',
  'This Quarter',
  'This Year'
];

function parseNumber(value: any): number {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

function normalizeTypeName(dataType?: string): string {
  return String(dataType || '').toLowerCase();
}

function looksLikeDateField(fieldName: string): boolean {
  const lower = fieldName.toLowerCase();
  return (
    lower.includes('date') ||
    lower.includes('year') ||
    lower.includes('month') ||
    lower.includes('quarter') ||
    lower.includes('week') ||
    lower.includes('day') ||
    lower.includes('timestamp') ||
    lower.includes('time')
  );
}

function looksNumericType(dataType?: string): boolean {
  const lower = normalizeTypeName(dataType);
  return (
    lower.includes('int') ||
    lower.includes('float') ||
    lower.includes('double') ||
    lower.includes('number') ||
    lower.includes('numeric') ||
    lower.includes('decimal') ||
    lower.includes('real')
  );
}

function looksDateType(dataType?: string): boolean {
  const lower = normalizeTypeName(dataType);
  return lower.includes('date') || lower.includes('time');
}

function inferFieldRole(column: any): FieldRole {
  const fieldName = String(column?.fieldName || '');
  const dataType = String(column?.dataType || column?.datatype || '');

  if (looksNumericType(dataType)) return 'measure';
  if (looksDateType(dataType) || looksLikeDateField(fieldName)) return 'dimension';

  const lower = fieldName.toLowerCase();
  if (
    lower.includes('sales') ||
    lower.includes('revenue') ||
    lower.includes('profit') ||
    lower.includes('amount') ||
    lower.includes('qty') ||
    lower.includes('quantity') ||
    lower.includes('cost') ||
    lower.includes('price') ||
    lower.includes('total') ||
    lower.includes('margin') ||
    lower.includes('rate') ||
    lower.includes('percent') ||
    lower.includes('score') ||
    lower.includes('count')
  ) {
    return 'measure';
  }

  return 'dimension';
}

function buildFieldMeta(columns: any[]): FieldMeta[] {
  return columns.map((col) => {
    const fieldName = String(col?.fieldName || '');
    const caption = col?.caption || col?.alias || fieldName;
    return {
      fieldName,
      caption,
      dataType: String(col?.dataType || col?.datatype || ''),
      role: inferFieldRole(col)
    };
  });
}

function createNode(
  id: string,
  label: string,
  level: number,
  dimensionField: string,
  pathValues: string[],
  measureFields: string[]
): TreeNode {
  const measures: Record<string, number> = {};
  measureFields.forEach((field) => {
    measures[field] = 0;
  });

  return {
    id,
    label,
    level,
    dimensionField,
    pathValues,
    measures,
    children: []
  };
}

function sanitizeSheetName(name: string) {
  return (name || 'Sheet1').replace(/[\\/?*[\]:]/g, '').slice(0, 31) || 'Sheet1';
}

function toDateInputValue(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, '0');
  const day = `${date.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function startOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function startOfQuarter(date: Date): Date {
  const quarterStartMonth = Math.floor(date.getMonth() / 3) * 3;
  return new Date(date.getFullYear(), quarterStartMonth, 1);
}

function getPresetDateRange(preset: string): { start: string; end: string } {
  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (preset === 'Today') {
    return { start: toDateInputValue(todayOnly), end: toDateInputValue(todayOnly) };
  }

  if (preset === 'Yesterday') {
    const yesterday = new Date(todayOnly);
    yesterday.setDate(yesterday.getDate() - 1);
    return { start: toDateInputValue(yesterday), end: toDateInputValue(yesterday) };
  }

  if (preset === 'Last 7 Days') {
    const start = new Date(todayOnly);
    start.setDate(start.getDate() - 6);
    return { start: toDateInputValue(start), end: toDateInputValue(todayOnly) };
  }

  if (preset === 'Last 30 Days') {
    const start = new Date(todayOnly);
    start.setDate(start.getDate() - 29);
    return { start: toDateInputValue(start), end: toDateInputValue(todayOnly) };
  }

  if (preset === 'This Month') {
    return {
      start: toDateInputValue(startOfMonth(todayOnly)),
      end: toDateInputValue(endOfMonth(todayOnly))
    };
  }

  if (preset === 'Last Month') {
    const lastMonth = new Date(todayOnly.getFullYear(), todayOnly.getMonth() - 1, 1);
    return {
      start: toDateInputValue(startOfMonth(lastMonth)),
      end: toDateInputValue(endOfMonth(lastMonth))
    };
  }

  if (preset === 'This Quarter') {
    const start = startOfQuarter(todayOnly);
    const end = new Date(start.getFullYear(), start.getMonth() + 3, 0);
    return { start: toDateInputValue(start), end: toDateInputValue(end) };
  }

  if (preset === 'This Year') {
    return {
      start: `${todayOnly.getFullYear()}-01-01`,
      end: `${todayOnly.getFullYear()}-12-31`
    };
  }

  return { start: '', end: '' };
}

function getFieldGroupLabel(field: FieldMeta): string {
  const name = field.fieldName.toLowerCase();
  const caption = field.caption.toLowerCase();

  if (
    name.includes('date') ||
    name.includes('month') ||
    name.includes('quarter') ||
    name.includes('year') ||
    caption.includes('date') ||
    caption.includes('month') ||
    caption.includes('quarter') ||
    caption.includes('year')
  ) return 'TIME';

  if (
    name.includes('region') ||
    name.includes('country') ||
    name.includes('city') ||
    name.includes('state') ||
    name.includes('territory') ||
    caption.includes('region') ||
    caption.includes('country') ||
    caption.includes('city') ||
    caption.includes('state')
  ) return 'GEOGRAPHY';

  if (
    name.includes('product') ||
    name.includes('sku') ||
    name.includes('item') ||
    name.includes('category') ||
    caption.includes('product') ||
    caption.includes('category')
  ) return 'PRODUCT';

  if (
    name.includes('sales') ||
    name.includes('channel') ||
    name.includes('revenue') ||
    name.includes('profit') ||
    caption.includes('sales') ||
    caption.includes('channel')
  ) return field.role === 'dimension' ? 'SALES' : 'MEASURES';

  if (
    name.includes('customer') ||
    name.includes('segment') ||
    name.includes('account') ||
    caption.includes('customer') ||
    caption.includes('segment')
  ) return 'CUSTOMER';

  return field.role === 'measure' ? 'MEASURES' : 'OTHER';
}

function groupFields(fields: FieldMeta[]): GroupedFieldSection[] {
  const map = new Map<string, FieldMeta[]>();

  fields.forEach((field) => {
    const label = getFieldGroupLabel(field);
    if (!map.has(label)) map.set(label, []);
    map.get(label)!.push(field);
  });

  return Array.from(map.entries()).map(([label, groupFields]) => ({
    label,
    fields: groupFields.sort((a, b) => a.caption.localeCompare(b.caption))
  }));
}

function getLevelStyles(level: number, hasChildren: boolean) {
  if (level === 0) return { hierarchyWeight: 700, measureWeight: 700 };
  if (level === 1) return { hierarchyWeight: hasChildren ? 600 : 500, measureWeight: hasChildren ? 600 : 500 };
  return { hierarchyWeight: hasChildren ? 500 : 400, measureWeight: hasChildren ? 500 : 400 };
}

function safeJsonParse<T>(raw: string | undefined, fallback: T): T {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

function formatInfoValue(value: any, format: InfoCalculation['format']): string {
  if (value === null || value === undefined || value === '') return '—';

  if (format === 'date') {
    const d = new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return d.toLocaleDateString();
  }

  if (format === 'currency') {
    const n = parseNumber(value);
    return new Intl.NumberFormat(undefined, {
      style: 'currency',
      currency: 'USD',
      maximumFractionDigits: 2
    }).format(n);
  }

  if (format === 'number') {
    const n = parseNumber(value);
    return new Intl.NumberFormat(undefined, {
      maximumFractionDigits: 2
    }).format(n);
  }

  return String(value);
}

function evaluateInfoCalculation(calculation: InfoCalculation, rows: PreviewRow[]): string {
  const scopedRows = rows.filter((row) => row.__sourceWorksheet === calculation.worksheet);

  if (calculation.aggregation === 'COUNT') {
    return formatInfoValue(scopedRows.length, calculation.format);
  }

  const values = scopedRows
    .map((row) => row[calculation.field])
    .filter((value) => value !== null && value !== undefined && value !== '');

  if (!values.length) return '—';

  if (calculation.aggregation === 'SUM') {
    const total = values.reduce((sum, value) => sum + parseNumber(value), 0);
    return formatInfoValue(total, calculation.format);
  }

  if (calculation.aggregation === 'AVG') {
    const total = values.reduce((sum, value) => sum + parseNumber(value), 0);
    return formatInfoValue(total / values.length, calculation.format);
  }

  if (calculation.aggregation === 'MAX') {
    const isDate = calculation.format === 'date' || looksLikeDateField(calculation.field);
    if (isDate) {
      const dateValues = values
        .map((value) => new Date(value))
        .filter((d) => !isNaN(d.getTime()));
      if (!dateValues.length) return '—';
      const maxDate = new Date(Math.max(...dateValues.map((d) => d.getTime())));
      return formatInfoValue(maxDate.toISOString(), calculation.format);
    }

    const maxValue = values.reduce((max, value) => {
      const numeric = parseNumber(value);
      return numeric > max ? numeric : max;
    }, Number.NEGATIVE_INFINITY);

    return formatInfoValue(maxValue, calculation.format);
  }

  if (calculation.aggregation === 'MIN') {
    const isDate = calculation.format === 'date' || looksLikeDateField(calculation.field);
    if (isDate) {
      const dateValues = values
        .map((value) => new Date(value))
        .filter((d) => !isNaN(d.getTime()));
      if (!dateValues.length) return '—';
      const minDate = new Date(Math.min(...dateValues.map((d) => d.getTime())));
      return formatInfoValue(minDate.toISOString(), calculation.format);
    }

    const minValue = values.reduce((min, value) => {
      const numeric = parseNumber(value);
      return numeric < min ? numeric : min;
    }, Number.POSITIVE_INFINITY);

    return formatInfoValue(minValue, calculation.format);
  }

  return '—';
}

function getEffectiveFieldRole(
  worksheetName: string,
  field: FieldMeta,
  overrides: FieldRoleOverrideMap
): FieldRole {
  return overrides[worksheetName]?.[field.fieldName] ?? field.role;
}

export default function App() {
  const [message, setMessage] = useState('Starting extension...');
  const [viewMode, setViewMode] = useState<'pivot' | 'straight'>('straight');
  const [isSelectionsPanelOpen, setIsSelectionsPanelOpen] = useState(true);
  const [isInfoOpen, setIsInfoOpen] = useState(false);

  const [worksheets, setWorksheets] = useState<string[]>([]);
  const [selectedSources, setSelectedSources] = useState<string[]>([]);
  const [dataSourceSearch, setDataSourceSearch] = useState('');

  const [worksheetFields, setWorksheetFields] = useState<WorksheetFieldsMap>({});
  const [fieldRoleOverrides, setFieldRoleOverrides] = useState<FieldRoleOverrideMap>({});
  const [expandedWorksheetFields, setExpandedWorksheetFields] = useState<WorksheetFieldExpandMap>({});

  const [availableFields, setAvailableFields] = useState<FieldMeta[]>([]);
  const [dimensionFields, setDimensionFields] = useState<string[]>([]);
  const [measureFields, setMeasureFields] = useState<string[]>([]);
  const [selectedDimensions, setSelectedDimensions] = useState<string[]>([]);
  const [selectedMeasures, setSelectedMeasures] = useState<string[]>([]);

  const [dateRange, setDateRange] = useState<DateRangeState>({
    field: '',
    start: '',
    end: '',
    preset: 'Custom Date Range'
  });

  const [filters, setFilters] = useState<SimpleFilterState[]>([]);
  const [openNestedFilterGroups, setOpenNestedFilterGroups] = useState<Record<string, boolean>>({});

  const [allRows, setAllRows] = useState<PreviewRow[]>([]);
  const [treeRoots, setTreeRoots] = useState<TreeNode[]>([]);
  const [expandedNodes, setExpandedNodes] = useState<Record<string, boolean>>({});

  const [layoutSettings, setLayoutSettings] = useState<LayoutSettings>(DEFAULT_LAYOUT);
  const [isResizing, setIsResizing] = useState(false);

  const [infoCalculations, setInfoCalculations] = useState<InfoCalculation[]>([]);

  const [sourceMenuOpen, setSourceMenuOpen] = useState(false);
  const [dateMenuOpen, setDateMenuOpen] = useState(false);
  const [dimensionMenuOpen, setDimensionMenuOpen] = useState(false);
  const [measureMenuOpen, setMeasureMenuOpen] = useState(false);
  const [filterMenuOpen, setFilterMenuOpen] = useState(false);

  const sourceMenuRef = useRef<HTMLDivElement | null>(null);
  const dateMenuRef = useRef<HTMLDivElement | null>(null);
  const dimensionMenuRef = useRef<HTMLDivElement | null>(null);
  const measureMenuRef = useRef<HTMLDivElement | null>(null);
  const filterMenuRef = useRef<HTMLDivElement | null>(null);

  const unregisterHandlersRef = useRef<(() => void)[]>([]);
  const refreshTimerRef = useRef<number | null>(null);
  const initialLoadCompleteRef = useRef(false);

  function clearEventListeners() {
    unregisterHandlersRef.current.forEach((unregister) => {
      try {
        unregister();
      } catch {
      }
    });
    unregisterHandlersRef.current = [];
  }

  function loadInfoCalculationsFromSettings() {
    const raw = tableau.extensions.settings.get('infoCalculations');
    const parsed = safeJsonParse<InfoCalculation[]>(raw, []);
    setInfoCalculations(parsed);
  }

  async function configure() {
    const popupUrl = `${window.location.origin}/configure.html`;
    const payload = JSON.stringify({
      worksheets,
      fields: availableFields
    });

    try {
      await tableau.extensions.ui.displayDialogAsync(popupUrl, payload, {
        height: 700,
        width: 900
      });
      loadInfoCalculationsFromSettings();
    } catch (error: any) {
      if (error?.errorCode === tableau.ErrorCodes.DialogClosedByUser) {
        return;
      }
      setMessage(`Configure error: ${error?.message || error}`);
    }
  }

  function getDisplayLabel(field: string): string {
    return availableFields.find((f) => f.fieldName === field)?.caption || field;
  }

  function getDisplayLabelFromWorksheetField(field: FieldMeta): string {
    return field.caption || field.fieldName;
  }

  function getMeasureFormat(_field: string): MeasureFormat {
    return DEFAULT_MEASURE_FORMAT;
  }

  function getSavedMeasureWidth(field: string): number {
    return layoutSettings.measureWidths[field] || layoutSettings.defaultMeasureWidth;
  }

  function formatMeasureValue(field: string, value: number): string {
    const fmt = getMeasureFormat(field);
    let numericValue = value;

    if (fmt.formatType === 'percent') numericValue = value * 100;

    const formatted = new Intl.NumberFormat(undefined, {
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: fmt.useThousandsSeparator
    }).format(numericValue);

    const percentSuffix = fmt.formatType === 'percent' ? '%' : '';
    const currencyPrefix = fmt.formatType === 'currency' && !fmt.prefix ? '$' : '';

    return `${fmt.prefix || currencyPrefix}${formatted}${percentSuffix}${fmt.suffix}`;
  }

  function buildTree(rows: PreviewRow[], dims: string[], measures: string[]): TreeNode[] {
    const roots: TreeNode[] = [];
    const nodeMap = new Map<string, TreeNode>();

    if (measures.length === 0) return [];

    if (dims.length === 0) {
      const root = createNode('__grand_total__', 'Grand Total', 0, '', [], measures);
      rows.forEach((row) => {
        measures.forEach((measureField) => {
          root.measures[measureField] += parseNumber(row[measureField]);
        });
      });
      return [root];
    }

    rows.forEach((row) => {
      let parentChildren = roots;
      let pathValues: string[] = [];

      dims.forEach((dimensionField, level) => {
        const label = String(row[dimensionField] ?? '');
        pathValues = [...pathValues, label];
        const nodeId = `${dimensionField}::${pathValues.join(' || ')}`;

        let node = nodeMap.get(nodeId);
        if (!node) {
          node = createNode(nodeId, label, level, dimensionField, [...pathValues], measures);
          nodeMap.set(nodeId, node);
          parentChildren.push(node);
        }

        measures.forEach((measureField) => {
          node!.measures[measureField] += parseNumber(row[measureField]);
        });

        parentChildren = node.children;
      });
    });

    function sortNodes(nodes: TreeNode[]) {
      nodes.sort((a, b) => a.label.localeCompare(b.label));
      nodes.forEach((node) => sortNodes(node.children));
    }

    sortNodes(roots);
    return roots;
  }

  function flattenVisibleNodes(nodes: TreeNode[], expanded: Record<string, boolean>): VisibleNode[] {
    const result: VisibleNode[] = [];

    function visit(nodeList: TreeNode[]) {
      nodeList.forEach((node) => {
        result.push({
          id: node.id,
          label: node.label,
          level: node.level,
          dimensionField: node.dimensionField,
          pathValues: node.pathValues,
          measures: node.measures,
          hasChildren: node.children.length > 0
        });

        if (expanded[node.id]) visit(node.children);
      });
    }

    visit(nodes);
    return result;
  }

  function flattenAllNodes(nodes: TreeNode[]): VisibleNode[] {
    const result: VisibleNode[] = [];

    function visit(nodeList: TreeNode[]) {
      nodeList.forEach((node) => {
        result.push({
          id: node.id,
          label: node.label,
          level: node.level,
          dimensionField: node.dimensionField,
          pathValues: node.pathValues,
          measures: node.measures,
          hasChildren: node.children.length > 0
        });
        if (node.children.length > 0) visit(node.children);
      });
    }

    visit(nodes);
    return result;
  }

  function toggleNode(nodeId: string) {
    setExpandedNodes((prev) => ({
      ...prev,
      [nodeId]: !prev[nodeId]
    }));
  }

  function toggleSource(name: string) {
    setSelectedSources((prev) => {
      const isSelected = prev.includes(name);
      if (isSelected) {
        if (prev.length === 1) return prev;
        return prev.filter((s) => s !== name);
      }
      return [...prev, name];
    });
  }

  function selectAllSources() {
    if (worksheets.length > 0) {
      setSelectedSources([...worksheets]);
    }
  }

  function resetSourcesToDefault() {
    setSelectedSources(worksheets.length > 0 ? [worksheets[0]] : []);
  }

  function toggleWorksheetFieldExpansion(worksheetName: string) {
    setExpandedWorksheetFields((prev) => ({
      ...prev,
      [worksheetName]: !prev[worksheetName]
    }));
  }

  function setFieldOverride(worksheetName: string, fieldName: string, role: FieldRole) {
    setFieldRoleOverrides((prev) => ({
      ...prev,
      [worksheetName]: {
        ...(prev[worksheetName] || {}),
        [fieldName]: role
      }
    }));
  }

  function toggleDimension(field: string) {
    setSelectedDimensions((prev) =>
      prev.includes(field) ? prev.filter((f) => f !== field) : [...prev, field]
    );
  }

  function toggleMeasure(field: string) {
    setSelectedMeasures((prev) =>
      prev.includes(field) ? prev.filter((f) => f !== field) : [...prev, field]
    );
  }

  function clearAllSelections() {
    const defaultDims = dimensionFields.slice(0, 2);
    const defaultMeasures = measureFields.slice(0, 2);

    setSelectedSources(worksheets.length > 0 ? [worksheets[0]] : []);
    setSelectedDimensions(defaultDims);
    setSelectedMeasures(defaultMeasures);
    setDateRange({ field: '', start: '', end: '', preset: 'Custom Date Range' });
    setFilters([]);
    setOpenNestedFilterGroups({});
    setDataSourceSearch('');
    setFieldRoleOverrides({});
    setExpandedWorksheetFields({});
    setMessage('');
  }

  function getDateFields() {
    return availableFields.filter(
      (field) => looksDateType(field.dataType) || looksLikeDateField(field.fieldName)
    );
  }

  function getFieldDistinctValuesFromRows(rows: PreviewRow[], fieldName: string): string[] {
    const values = new Set<string>();
    rows.forEach((row) => {
      const value = String(row[fieldName] ?? '');
      if (value !== '') values.add(value);
    });
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }

  function getFilterDefinitions(): FilterGroupDefinition[] {
    const dimensionCandidates = availableFields.filter(
      (field) =>
        field.role === 'dimension' &&
        !looksDateType(field.dataType) &&
        !looksLikeDateField(field.fieldName)
    );

    return dimensionCandidates
      .map((field) => ({
        label: field.caption,
        field: field.fieldName,
        options: getFieldDistinctValuesFromRows(allRows, field.fieldName).slice(0, 25)
      }))
      .filter((group) => group.options.length > 0)
      .sort((a, b) => a.label.localeCompare(b.label));
  }

  function isFilterValueSelected(field: string, value: string): boolean {
    const filter = filters.find((item) => item.field === field);
    return !!filter?.values.includes(value);
  }

  function toggleFilterValue(field: string, value: string) {
    setFilters((prev) => {
      const existing = prev.find((item) => item.field === field);

      if (!existing) return [...prev, { field, values: [value] }];

      const nextValues = existing.values.includes(value)
        ? existing.values.filter((v) => v !== value)
        : [...existing.values, value];

      const remaining = prev.filter((item) => item.field !== field);

      if (nextValues.length === 0) return remaining;
      return [...remaining, { field, values: nextValues }];
    });
  }

  function removeFilter(field: string) {
    setFilters((prev) => prev.filter((item) => item.field !== field));
  }

  function toggleNestedFilterGroup(field: string) {
    setOpenNestedFilterGroups((prev) => ({
      ...prev,
      [field]: !prev[field]
    }));
  }

  function applyDatePreset(preset: string) {
    if (preset === 'Custom Date Range') {
      setDateRange((prev) => ({ ...prev, preset }));
      return;
    }

    const { start, end } = getPresetDateRange(preset);
    setDateRange((prev) => ({ ...prev, preset, start, end }));
  }

  function debouncedRefresh() {
    if (refreshTimerRef.current) window.clearTimeout(refreshTimerRef.current);

    refreshTimerRef.current = window.setTimeout(() => {
      if (selectedSources.length > 0) loadAllSelectedWorksheets(selectedSources, true);
    }, 250);
  }

  async function registerWorksheetListeners(worksheetNames: string[]) {
    clearEventListeners();

    const dashboard = tableau.extensions.dashboardContent.dashboard;

    worksheetNames.forEach((worksheetName) => {
      const worksheet = dashboard.worksheets.find((w: any) => w.name === worksheetName);
      if (!worksheet) return;

      try {
        const filterHandler = worksheet.addEventListener(
          tableau.TableauEventType.FilterChanged,
          () => debouncedRefresh()
        );
        unregisterHandlersRef.current.push(filterHandler);
      } catch {
      }

      try {
        const summaryHandler = worksheet.addEventListener(
          tableau.TableauEventType.SummaryDataChanged,
          () => debouncedRefresh()
        );
        unregisterHandlersRef.current.push(summaryHandler);
      } catch {
      }
    });
  }

  async function loadWorksheetRows(worksheet: any): Promise<{ rows: PreviewRow[]; fields: FieldMeta[] }> {
    const reader = await worksheet.getSummaryDataReaderAsync();

    try {
      const dataTable = await reader.getAllPagesAsync();
      const fields = buildFieldMeta(dataTable.columns);

      const rows: PreviewRow[] = dataTable.data.map((row: any[]) => {
        const obj: PreviewRow = { __sourceWorksheet: worksheet.name };
        dataTable.columns.forEach((col: any, index: number) => {
          obj[col.fieldName] = row[index]?.value ?? row[index]?.formattedValue ?? '';
        });
        return obj;
      });

      return { rows, fields };
    } finally {
      await reader.releaseAsync();
    }
  }

  function rebuildFieldsFromSelectedSources(
    sourceNames: string[],
    currentWorksheetFields: WorksheetFieldsMap,
    overrides: FieldRoleOverrideMap,
    mergedRowsArg?: PreviewRow[]
  ) {
    const mergedFieldMap = new Map<string, FieldMeta>();

    sourceNames.forEach((worksheetName) => {
      const fields = currentWorksheetFields[worksheetName] || [];
      fields.forEach((field) => {
        const effectiveRole = getEffectiveFieldRole(worksheetName, field, overrides);
        const existing = mergedFieldMap.get(field.fieldName);

        if (!existing) {
          mergedFieldMap.set(field.fieldName, {
            ...field,
            role: effectiveRole
          });
        } else if (existing.role !== effectiveRole) {
          mergedFieldMap.set(field.fieldName, {
            ...existing,
            role: field.role
          });
        }
      });
    });

    const mergedFields = Array.from(mergedFieldMap.values()).sort((a, b) =>
      a.caption.localeCompare(b.caption)
    );

    const dims = mergedFields.filter((f) => f.role === 'dimension').map((f) => f.fieldName);
    const measures = mergedFields.filter((f) => f.role === 'measure').map((f) => f.fieldName);
    const dateFields = mergedFields.filter(
      (f) => looksDateType(f.dataType) || looksLikeDateField(f.fieldName)
    );

    const rowsToUse = mergedRowsArg ?? allRows;

    setAvailableFields(mergedFields);
    setDimensionFields(dims);
    setMeasureFields(measures);

    setSelectedDimensions((prev) => {
      const valid = prev.filter((field) => dims.includes(field));
      return valid.length > 0 ? valid : dims.slice(0, 2);
    });

    setSelectedMeasures((prev) => {
      const valid = prev.filter((field) => measures.includes(field));
      return valid.length > 0 ? valid : measures.slice(0, 2);
    });

    setFilters((prev) =>
      prev
        .filter((filter) => dims.includes(filter.field))
        .map((filter) => ({
          ...filter,
          values: filter.values.filter((value) =>
            getFieldDistinctValuesFromRows(rowsToUse, filter.field).includes(value)
          )
        }))
        .filter((filter) => filter.values.length > 0)
    );

    setDateRange((prev) => {
      const nextField =
        prev.field && dims.includes(prev.field)
          ? prev.field
          : dateFields[0]?.fieldName || '';

      if (!nextField) {
        return { field: '', start: '', end: '', preset: 'Custom Date Range' };
      }

      return { ...prev, field: nextField };
    });
  }

  async function loadAllSelectedWorksheets(sourceNames: string[], isRefresh = false) {
    if (!sourceNames.length) {
      setAllRows([]);
      setAvailableFields([]);
      setMessage('Select at least one data source.');
      return;
    }

    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const selectedWorksheets = dashboard.worksheets.filter((w: any) => sourceNames.includes(w.name));

    if (!isRefresh) setMessage('Loading data...');

    const results = await Promise.all(
      selectedWorksheets.map((worksheet: any) => loadWorksheetRows(worksheet))
    );

    const mergedRows = results.flatMap((r) => r.rows);

    const nextWorksheetFields: WorksheetFieldsMap = {};
    selectedWorksheets.forEach((worksheet: any, index: number) => {
      nextWorksheetFields[worksheet.name] = results[index].fields;
    });

    setAllRows(mergedRows);
    setWorksheetFields((prev) => {
      const combined = { ...prev, ...nextWorksheetFields };
      rebuildFieldsFromSelectedSources(sourceNames, combined, fieldRoleOverrides, mergedRows);
      return combined;
    });

    setExpandedWorksheetFields((prev) => {
      const next = { ...prev };
      sourceNames.forEach((name) => {
        if (next[name] === undefined) next[name] = false;
      });
      return next;
    });

    await registerWorksheetListeners(sourceNames);
    setMessage('');
  }

  useEffect(() => {
    async function init() {
      try {
        await tableau.extensions.initializeAsync({ configure });

        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheetNames = dashboard.worksheets.map((w: any) => w.name);

        setWorksheets(worksheetNames);

        const initialSources = worksheetNames.length > 0 ? worksheetNames.slice(0, 1) : [];
        setSelectedSources(initialSources);

        if (initialSources.length > 0) {
          await loadAllSelectedWorksheets(initialSources);
        } else {
          setMessage('No worksheets were found on this dashboard.');
        }

        loadInfoCalculationsFromSettings();
        initialLoadCompleteRef.current = true;
      } catch (error: any) {
        setMessage(`Error: ${error?.message || error}`);
      }
    }

    init();

    return () => {
      clearEventListeners();
      if (refreshTimerRef.current) window.clearTimeout(refreshTimerRef.current);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
  }, []);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;

      if (sourceMenuRef.current && !sourceMenuRef.current.contains(target)) setSourceMenuOpen(false);
      if (dateMenuRef.current && !dateMenuRef.current.contains(target)) setDateMenuOpen(false);
      if (dimensionMenuRef.current && !dimensionMenuRef.current.contains(target)) setDimensionMenuOpen(false);
      if (measureMenuRef.current && !measureMenuRef.current.contains(target)) setMeasureMenuOpen(false);
      if (filterMenuRef.current && !filterMenuRef.current.contains(target)) setFilterMenuOpen(false);
    }

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (!initialLoadCompleteRef.current) return;

    if (selectedSources.length > 0) {
      loadAllSelectedWorksheets(selectedSources);
    } else {
      setAllRows([]);
      setAvailableFields([]);
    }
  }, [selectedSources]);

  useEffect(() => {
    if (selectedSources.length > 0 && Object.keys(worksheetFields).length > 0) {
      rebuildFieldsFromSelectedSources(selectedSources, worksheetFields, fieldRoleOverrides);
    }
  }, [fieldRoleOverrides]);

  const filteredRows = useMemo(() => {
    return allRows.filter((row) => {
      if (dateRange.field && (dateRange.start || dateRange.end)) {
        const rawValue = row[dateRange.field];
        const rowDate = new Date(rawValue);
        if (isNaN(rowDate.getTime())) return false;

        if (dateRange.start) {
          const startDate = new Date(dateRange.start);
          if (rowDate < startDate) return false;
        }

        if (dateRange.end) {
          const endDate = new Date(dateRange.end);
          endDate.setHours(23, 59, 59, 999);
          if (rowDate > endDate) return false;
        }
      }

      for (const filter of filters) {
        const value = String(row[filter.field] ?? '');
        if (!filter.values.includes(value)) return false;
      }

      return true;
    });
  }, [allRows, dateRange, filters]);

  useEffect(() => {
    const roots = buildTree(filteredRows, selectedDimensions, selectedMeasures);
    setTreeRoots(roots);

    setExpandedNodes((prev) => {
      const next = { ...prev };
      roots.forEach((node) => {
        if (next[node.id] === undefined) next[node.id] = true;
      });
      return next;
    });
  }, [filteredRows, selectedDimensions, selectedMeasures]);

  const visibleNodes = useMemo(
    () => flattenVisibleNodes(treeRoots, expandedNodes),
    [treeRoots, expandedNodes]
  );

  const hierarchyWidth = Math.max(MIN_HIERARCHY_WIDTH, layoutSettings.hierarchyWidth);

  const measureColumnWidths = useMemo(() => {
    const result: Record<string, number> = {};
    selectedMeasures.forEach((field) => {
      result[field] = Math.max(MIN_MEASURE_WIDTH, getSavedMeasureWidth(field));
    });
    return result;
  }, [selectedMeasures, layoutSettings]);

  const groupedDimensions = useMemo(
    () => groupFields(availableFields.filter((field) => field.role === 'dimension')),
    [availableFields]
  );

  const groupedMeasures = useMemo(
    () => groupFields(availableFields.filter((field) => field.role === 'measure')),
    [availableFields]
  );

  const filterDefinitions = useMemo(
    () => getFilterDefinitions(),
    [availableFields, allRows]
  );

  const infoResults = useMemo<InfoResult[]>(() => {
    return infoCalculations.map((calc) => ({
      id: calc.id,
      label: calc.label,
      value: evaluateInfoCalculation(calc, allRows)
    }));
  }, [infoCalculations, allRows]);

  const filteredWorksheets = useMemo(() => {
    const search = dataSourceSearch.trim().toLowerCase();
    if (!search) return worksheets;
    return worksheets.filter((name) => name.toLowerCase().includes(search));
  }, [worksheets, dataSourceSearch]);

  function startResize(
    event: React.MouseEvent<HTMLDivElement>,
    type: 'hierarchy' | 'measure',
    field?: string
  ) {
    event.preventDefault();
    event.stopPropagation();

    const startX = event.clientX;
    const startWidth =
      type === 'hierarchy'
        ? layoutSettings.hierarchyWidth
        : getSavedMeasureWidth(field || '');

    setIsResizing(true);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    function onMouseMove(moveEvent: MouseEvent) {
      const deltaX = moveEvent.clientX - startX;
      const proposedWidth = startWidth + deltaX;

      setLayoutSettings((prev) => {
        if (type === 'hierarchy') {
          return {
            ...prev,
            hierarchyWidth: Math.max(MIN_HIERARCHY_WIDTH, proposedWidth)
          };
        }

        if (!field) return prev;

        return {
          ...prev,
          measureWidths: {
            ...prev.measureWidths,
            [field]: Math.max(MIN_MEASURE_WIDTH, proposedWidth)
          }
        };
      });
    }

    function onMouseUp() {
      setIsResizing(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    }

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  function buildStraightExportRows() {
    return filteredRows.map((row) => {
      const exportRow: Record<string, any> = {};

      selectedDimensions.forEach((field) => {
        exportRow[getDisplayLabel(field)] = row[field] ?? '';
      });

      selectedMeasures.forEach((field) => {
        exportRow[getDisplayLabel(field)] = parseNumber(row[field]);
      });

      return exportRow;
    });
  }

  function buildPivotExportRowsFromNodes(nodesToExport: VisibleNode[]) {
    return nodesToExport.map((node) => {
      const exportRow: Record<string, any> = {};

      selectedDimensions.forEach((field, index) => {
        exportRow[getDisplayLabel(field)] = node.pathValues[index] ?? '';
      });

      selectedMeasures.forEach((field) => {
        exportRow[getDisplayLabel(field)] = node.measures[field] ?? 0;
      });

      return exportRow;
    });
  }

  function exportRowsToExcel(rows: Record<string, any>[]) {
    if (typeof XLSX === 'undefined') {
      setMessage('Excel export library is not available.');
      return;
    }

    if (!rows.length) {
      setMessage('There is no data to export.');
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeSheetName('Report Builder'));
    XLSX.writeFile(workbook, 'report-builder-export.xlsx');
    setMessage('');
  }

  function exportCurrentView() {
    if (selectedMeasures.length === 0) {
      setMessage('Please select at least one metric.');
      return;
    }

    if (viewMode === 'pivot') {
      exportRowsToExcel(buildPivotExportRowsFromNodes(flattenAllNodes(treeRoots)));
    } else {
      exportRowsToExcel(buildStraightExportRows());
    }
  }

  function renderToolbarDropdown(
    label: string,
    icon: string,
    menuOpen: boolean,
    setMenuOpen: (value: boolean | ((prev: boolean) => boolean)) => void,
    menuRef: React.RefObject<HTMLDivElement | null>,
    content: React.ReactNode,
    accentColor: string
  ) {
    return (
      <div ref={menuRef} style={{ position: 'relative' }}>
        <button
          type="button"
          onClick={() => setMenuOpen((prev) => !prev)}
          style={{
            height: '34px',
            minWidth: '128px',
            padding: '0 10px',
            border: '1px solid #d9dfe8',
            borderRadius: '8px',
            background: '#f7f9fc',
            color: '#1f2937',
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '8px',
            fontSize: '12px',
            fontWeight: 600,
            cursor: 'pointer',
            boxShadow: menuOpen ? '0 2px 6px rgba(0,0,0,0.06)' : 'none'
          }}
        >
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
            <span style={{ color: accentColor, fontSize: '13px', lineHeight: 1 }}>{icon}</span>
            <span>{label}</span>
          </span>
          <span aria-hidden="true" style={{ color: '#6b7280', fontSize: '10px', lineHeight: 1 }}>▼</span>
        </button>

        {menuOpen && (
          <div
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              marginTop: '6px',
              minWidth: '420px',
              maxWidth: '760px',
              maxHeight: '560px',
              overflowY: 'auto',
              background: '#ffffff',
              border: '1px solid #e6ebf2',
              borderRadius: '18px',
              boxShadow: '0 16px 36px rgba(15,23,42,0.14)',
              zIndex: 50,
              padding: '18px'
            }}
          >
            {content}
          </div>
        )}
      </div>
    );
  }

  function renderSourceOverrideButtons(worksheetName: string, field: FieldMeta) {
    const effectiveRole = getEffectiveFieldRole(worksheetName, field, fieldRoleOverrides);
    const isDimension = effectiveRole === 'dimension';
    const isMeasure = effectiveRole === 'measure';

    return (
      <div style={{ display: 'flex', gap: '10px' }}>
        <button
          type="button"
          onClick={() => setFieldOverride(worksheetName, field.fieldName, 'dimension')}
          style={{
            minWidth: '132px',
            height: '32px',
            border: '2px solid #222',
            background: isDimension ? '#eef6ff' : '#ffffff',
            color: '#111827',
            fontSize: '12px',
            fontWeight: 600,
            cursor: 'pointer'
          }}
        >
          Dimension
        </button>

        <button
          type="button"
          onClick={() => setFieldOverride(worksheetName, field.fieldName, 'measure')}
          style={{
            minWidth: '86px',
            height: '32px',
            border: '2px solid #222',
            background: isMeasure ? '#fff1e8' : '#ffffff',
            color: '#111827',
            fontSize: '12px',
            fontWeight: 600,
            cursor: 'pointer'
          }}
        >
          Metric
        </button>
      </div>
    );
  }

  function renderSourceMenu() {
    return (
      <div>
        <div style={{ fontSize: '15px', fontWeight: 700, color: '#111827', marginBottom: '10px' }}>
          Select Data Sources
        </div>

        <div style={{ fontSize: '12px', color: '#6b7280', marginBottom: '12px' }}>
          {selectedSources.length} of {worksheets.length} source(s) selected
        </div>

        <div style={{ marginBottom: '12px' }}>
          <input
            type="text"
            value={dataSourceSearch}
            onChange={(e) => setDataSourceSearch(e.target.value)}
            placeholder="Search data sources..."
            style={{
              width: '100%',
              height: '38px',
              border: '1px solid #d7dfea',
              borderRadius: '10px',
              padding: '0 12px',
              background: '#f8fbff',
              fontSize: '13px'
            }}
          />
        </div>

        <div
          style={{
            display: 'flex',
            gap: '8px',
            flexWrap: 'wrap',
            marginBottom: '14px'
          }}
        >
          <button
            type="button"
            onClick={selectAllSources}
            style={{
              height: '32px',
              padding: '0 12px',
              border: '1px solid #d8e2ee',
              borderRadius: '10px',
              background: '#f4f8ff',
              color: '#1d4ed8',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer'
            }}
          >
            Select All
          </button>

          <button
            type="button"
            onClick={resetSourcesToDefault}
            style={{
              height: '32px',
              padding: '0 12px',
              border: '1px solid #d8e2ee',
              borderRadius: '10px',
              background: '#fff7ed',
              color: '#c2410c',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer'
            }}
          >
            Reset to Default
          </button>

          <button
            type="button"
            onClick={() => {
              setDataSourceSearch('');
              setSourceMenuOpen(false);
            }}
            style={{
              height: '32px',
              padding: '0 12px',
              border: '1px solid #d8e2ee',
              borderRadius: '10px',
              background: '#f8fafc',
              color: '#475569',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer'
            }}
          >
            Close
          </button>
        </div>

        <div style={{ display: 'grid', gap: '12px' }}>
          {filteredWorksheets.length === 0 && (
            <div
              style={{
                border: '1px dashed #d8e2ee',
                borderRadius: '12px',
                padding: '16px',
                fontSize: '13px',
                color: '#6b7280',
                background: '#fbfdff'
              }}
            >
              No matching data sources found.
            </div>
          )}

          {filteredWorksheets.map((name) => {
            const selected = selectedSources.includes(name);
            const expanded = !!expandedWorksheetFields[name];
            const worksheetFieldList = worksheetFields[name] || [];

            const description =
              name.toLowerCase().includes('marketing')
                ? 'Campaign and engagement metrics'
                : name.toLowerCase().includes('customer')
                  ? 'User demographics and behavior'
                  : name.toLowerCase().includes('product')
                    ? 'Usage and performance data'
                    : 'Transaction and revenue data';

            return (
              <div
                key={name}
                style={{
                  border: '1px solid #d8e2ee',
                  borderRadius: '14px',
                  background: selected ? '#f8fbff' : '#ffffff',
                  overflow: 'hidden'
                }}
              >
                <button
                  type="button"
                  onClick={() => toggleSource(name)}
                  style={{
                    width: '100%',
                    border: 'none',
                    padding: '14px 16px',
                    textAlign: 'left',
                    background: selected ? '#eef6ff' : '#f8fbff',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'flex-start',
                    gap: '12px'
                  }}
                >
                  <div style={{ color: '#2c6cff', width: '14px', marginTop: '2px', fontWeight: 700 }}>
                    {selected ? '✓' : ''}
                  </div>

                  <div style={{ minWidth: 0, flex: 1 }}>
                    <div
                      style={{
                        fontSize: '16px',
                        fontWeight: 600,
                        color: '#111827',
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}
                    >
                      {name}
                    </div>
                    <div style={{ fontSize: '13px', color: '#4b5563', marginTop: '2px' }}>{description}</div>
                  </div>
                </button>

                {selected && worksheetFieldList.length > 0 && (
                  <div style={{ padding: '0 16px 16px' }}>
                    <button
                      type="button"
                      onClick={() => toggleWorksheetFieldExpansion(name)}
                      style={{
                        width: '100%',
                        height: '28px',
                        border: '2px solid #222',
                        background: '#fff',
                        color: '#111827',
                        textAlign: 'left',
                        padding: '0 10px',
                        fontSize: '12px',
                        cursor: 'pointer',
                        marginBottom: expanded ? '0' : '6px'
                      }}
                    >
                      Show Fields
                      <span style={{ float: 'right' }}>{expanded ? '▲' : '▼'}</span>
                    </button>

                    {expanded && (
                      <div
                        style={{
                          border: '2px solid #222',
                          borderTop: 'none',
                          background: '#fff',
                          padding: '14px'
                        }}
                      >
                        <div style={{ display: 'grid', gap: '12px' }}>
                          {worksheetFieldList.map((field) => (
                            <div
                              key={`${name}-${field.fieldName}`}
                              style={{
                                display: 'grid',
                                gridTemplateColumns: '1fr auto',
                                gap: '16px',
                                alignItems: 'center'
                              }}
                            >
                              <div style={{ fontSize: '13px', color: '#111827' }}>
                                {getDisplayLabelFromWorksheetField(field)}
                              </div>
                              {renderSourceOverrideButtons(name, field)}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  function renderDateMenu() {
    const dateFields = getDateFields();

    return (
      <div>
        <div style={{ fontSize: '15px', fontWeight: 700, color: '#111827', marginBottom: '12px' }}>
          Select Time Period
        </div>

        <div style={{ marginBottom: '14px' }}>
          <select
            value={dateRange.field}
            onChange={(e) => setDateRange((prev) => ({ ...prev, field: e.target.value }))}
            style={{
              width: '100%',
              height: '40px',
              border: '1px solid #d7dfea',
              borderRadius: '10px',
              padding: '0 12px',
              background: '#f8fbff'
            }}
          >
            <option value="">Select date field</option>
            {dateFields.map((field) => (
              <option key={field.fieldName} value={field.fieldName}>
                {field.caption}
              </option>
            ))}
          </select>
        </div>

        <div
          style={{
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: '10px',
            marginBottom: '18px'
          }}
        >
          {DATE_PRESETS.map((preset) => {
            const selected = dateRange.preset === preset;
            return (
              <button
                key={preset}
                type="button"
                onClick={() => applyDatePreset(preset)}
                style={{
                  height: '38px',
                  border: 'none',
                  borderRadius: '10px',
                  background: selected ? 'linear-gradient(90deg, #c23af5 0%, #7c4dff 100%)' : '#f4f7fb',
                  color: selected ? '#fff' : '#374151',
                  fontWeight: selected ? 700 : 600,
                  fontSize: '13px',
                  cursor: 'pointer',
                  textAlign: 'left',
                  padding: '0 12px'
                }}
              >
                {preset}
              </button>
            );
          })}
        </div>

        <div style={{ borderTop: '1px solid #e5e7eb', paddingTop: '14px' }}>
          <div
            style={{
              background: '#fafbfd',
              border: '1px solid #eef2f7',
              borderRadius: '16px',
              padding: '16px'
            }}
          >
            <div style={{ fontSize: '15px', fontWeight: 700, marginBottom: '12px', color: '#111827' }}>
              {new Date().toLocaleString(undefined, { month: 'long', year: 'numeric' })}
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
              <label style={{ display: 'grid', gap: '6px' }}>
                <span style={{ fontSize: '12px', color: '#6b7280', fontWeight: 600 }}>Start</span>
                <input
                  type="date"
                  value={dateRange.start}
                  onChange={(e) =>
                    setDateRange((prev) => ({
                      ...prev,
                      preset: 'Custom Date Range',
                      start: e.target.value
                    }))
                  }
                  style={{
                    height: '40px',
                    border: '1px solid #d7dfea',
                    borderRadius: '10px',
                    padding: '0 10px',
                    background: '#fff'
                  }}
                />
              </label>

              <label style={{ display: 'grid', gap: '6px' }}>
                <span style={{ fontSize: '12px', color: '#6b7280', fontWeight: 600 }}>End</span>
                <input
                  type="date"
                  value={dateRange.end}
                  onChange={(e) =>
                    setDateRange((prev) => ({
                      ...prev,
                      preset: 'Custom Date Range',
                      end: e.target.value
                    }))
                  }
                  style={{
                    height: '40px',
                    border: '1px solid #d7dfea',
                    borderRadius: '10px',
                    padding: '0 10px',
                    background: '#fff'
                  }}
                />
              </label>
            </div>

            <div style={{ textAlign: 'center', fontSize: '13px', color: '#4b5563' }}>
              {dateRange.start || '...'} - {dateRange.end || '...'}
            </div>
          </div>
        </div>
      </div>
    );
  }

  function renderFieldChip(
    label: string,
    subtitle: string | null,
    selected: boolean,
    onClick: () => void,
    selectedColor: string
  ) {
    return (
      <button
        type="button"
        onClick={onClick}
        style={{
          border: 'none',
          borderRadius: '12px',
          padding: subtitle ? '12px 14px' : '10px 14px',
          background: selected ? selectedColor : '#f3f6fa',
          color: selected ? '#fff' : '#374151',
          minHeight: subtitle ? '64px' : '40px',
          textAlign: 'left',
          cursor: 'pointer',
          fontWeight: 600,
          boxShadow: selected ? 'inset 0 -1px 0 rgba(255,255,255,0.18)' : 'none'
        }}
      >
        <div style={{ fontSize: '13px', lineHeight: 1.2 }}>{label}</div>
        {subtitle && (
          <div
            style={{
              fontSize: '12px',
              lineHeight: 1.2,
              marginTop: '6px',
              color: selected ? 'rgba(255,255,255,0.88)' : '#6b7280',
              fontWeight: 500
            }}
          >
            {subtitle}
          </div>
        )}
      </button>
    );
  }

  function renderDimensionMenu() {
    return (
      <div>
        <div style={{ fontSize: '15px', fontWeight: 700, color: '#111827', marginBottom: '14px' }}>
          Select Dimensions
        </div>

        <div style={{ display: 'grid', gap: '14px' }}>
          {groupedDimensions.map((section) => (
            <div key={section.label}>
              <div style={{ fontSize: '12px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                {section.label}
              </div>

              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
                  gap: '10px'
                }}
              >
                {section.fields.map((field) => (
                  <div key={field.fieldName}>
                    {renderFieldChip(
                      field.caption,
                      null,
                      selectedDimensions.includes(field.fieldName),
                      () => toggleDimension(field.fieldName),
                      '#08cf49'
                    )}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  function renderMeasureMenu() {
    return (
      <div>
        <div style={{ fontSize: '15px', fontWeight: 700, color: '#111827', marginBottom: '14px' }}>
          Select Measures
        </div>

        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
            gap: '10px'
          }}
        >
          {groupedMeasures.flatMap((section) =>
            section.fields.map((field) => (
              <div key={field.fieldName}>
                {renderFieldChip(
                  field.caption,
                  field.dataType
                    ? looksNumericType(field.dataType)
                      ? normalizeTypeName(field.dataType).includes('percent')
                        ? 'Percentage'
                        : 'Number'
                      : field.dataType
                    : 'Number',
                  selectedMeasures.includes(field.fieldName),
                  () => toggleMeasure(field.fieldName),
                  '#ff6a00'
                )}
              </div>
            ))
          )}
        </div>
      </div>
    );
  }

  function renderFilterValueDropdown(group: FilterGroupDefinition) {
    const open = !!openNestedFilterGroups[group.field];
    const selectedCount = filters.find((f) => f.field === group.field)?.values.length || 0;

    return (
      <div
        key={group.field}
        style={{
          border: '1px solid #e7ecf2',
          borderRadius: '12px',
          padding: '10px 12px',
          background: '#fbfcfe'
        }}
      >
        <button
          type="button"
          onClick={() => toggleNestedFilterGroup(group.field)}
          style={{
            width: '100%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            color: '#1f2937',
            fontWeight: 600,
            cursor: 'pointer'
          }}
        >
          <span>{group.label}{selectedCount ? ` (${selectedCount})` : ''}</span>
          <span>{open ? '▲' : '▼'}</span>
        </button>

        {open && (
          <div style={{ marginTop: '10px', display: 'grid', gap: '8px' }}>
            {group.options.map((option) => {
              const selected = isFilterValueSelected(group.field, option);
              return (
                <label
                  key={option}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    padding: '8px 10px',
                    borderRadius: '10px',
                    background: selected ? '#eef0ff' : '#f5f7fb',
                    cursor: 'pointer'
                  }}
                >
                  <input
                    type="checkbox"
                    checked={selected}
                    onChange={() => toggleFilterValue(group.field, option)}
                  />
                  <span style={{ fontSize: '13px', color: '#374151', fontWeight: 600 }}>{option}</span>
                </label>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  function renderFilterMenu() {
    return (
      <div>
        <div style={{ fontSize: '15px', fontWeight: 700, color: '#111827', marginBottom: '14px' }}>
          Select Filters
        </div>

        <div style={{ fontSize: '12px', color: '#6b7280', marginBottom: '14px' }}>
          All filterable fields from the selected data source(s).
        </div>

        <div style={{ display: 'grid', gap: '16px' }}>
          {filterDefinitions.map((group) => (
            <div key={group.field}>
              <div style={{ fontSize: '13px', fontWeight: 700, color: '#374151', marginBottom: '8px' }}>
                {group.label}
              </div>

              {group.options.length > 3 ? (
                renderFilterValueDropdown(group)
              ) : (
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
                    gap: '10px'
                  }}
                >
                  {group.options.map((option) => (
                    <div key={option}>
                      {renderFieldChip(
                        option,
                        null,
                        isFilterValueSelected(group.field, option),
                        () => toggleFilterValue(group.field, option),
                        '#5f5cf1'
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}

          {filterDefinitions.length === 0 && (
            <div style={{ color: '#6b7280', fontSize: '13px' }}>
              No filterable fields are available.
            </div>
          )}
        </div>
      </div>
    );
  }

  function renderSelectionPill(text: string, color = '#065f73', onRemove?: () => void) {
    return (
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '8px',
          minHeight: '28px',
          padding: '0 8px 0 12px',
          borderRadius: '999px',
          background: color,
          color: '#fff',
          fontSize: '12px',
          fontWeight: 700,
          marginBottom: '8px'
        }}
      >
        <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {text}
        </span>

        {onRemove && (
          <button
            type="button"
            onClick={onRemove}
            style={{
              width: '18px',
              height: '18px',
              minWidth: '18px',
              borderRadius: '999px',
              border: 'none',
              background: color,
              color: '#fff',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: 0,
              lineHeight: 1
            }}
            aria-label={`Remove ${text}`}
          >
            ×
          </button>
        )}
      </div>
    );
  }

  function renderPivotTable() {
    return (
      <div style={{ width: '100%', overflow: 'auto' }}>
        <table
          style={{
            borderCollapse: 'collapse',
            width: 'max-content',
            minWidth: '100%',
            tableLayout: 'fixed'
          }}
        >
          <thead>
            <tr>
              <th
                style={{
                  position: 'relative',
                  background: '#075b67',
                  color: '#fff',
                  border: '1px solid #dce4ec',
                  padding: '8px 10px',
                  textAlign: 'left',
                  width: `${hierarchyWidth}px`,
                  boxSizing: 'border-box',
                  fontSize: '12px'
                }}
              >
                <div style={{ paddingRight: '12px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {selectedDimensions[0] ? getDisplayLabel(selectedDimensions[0]) : 'Hierarchy'}
                </div>
                <div
                  onMouseDown={(e) => startResize(e, 'hierarchy')}
                  style={{
                    position: 'absolute',
                    top: 0,
                    right: 0,
                    width: '8px',
                    height: '100%',
                    cursor: 'col-resize'
                  }}
                />
              </th>

              {selectedMeasures.map((field) => {
                const width = measureColumnWidths[field];
                return (
                  <th
                    key={field}
                    style={{
                      position: 'relative',
                      background: '#075b67',
                      color: '#fff',
                      border: '1px solid #dce4ec',
                      padding: '8px 10px',
                      textAlign: 'right',
                      width: `${width}px`,
                      boxSizing: 'border-box',
                      fontSize: '12px'
                    }}
                  >
                    <div style={{ paddingRight: '12px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {getDisplayLabel(field)}
                    </div>
                    <div
                      onMouseDown={(e) => startResize(e, 'measure', field)}
                      style={{
                        position: 'absolute',
                        top: 0,
                        right: 0,
                        width: '8px',
                        height: '100%',
                        cursor: 'col-resize'
                      }}
                    />
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {visibleNodes.map((node, index) => {
              const styles = getLevelStyles(node.level, node.hasChildren);

              return (
                <tr key={node.id}>
                  <td
                    style={{
                      border: '1px solid #e8edf3',
                      background: index % 2 === 0 ? '#ffffff' : '#f7fafc',
                      padding: '8px 10px',
                      width: `${hierarchyWidth}px`,
                      boxSizing: 'border-box',
                      fontSize: '12px'
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', paddingLeft: `${node.level * 18}px`, overflow: 'hidden' }}>
                      {node.hasChildren ? (
                        <button
                          type="button"
                          onClick={() => toggleNode(node.id)}
                          style={{
                            width: '18px',
                            height: '18px',
                            minWidth: '18px',
                            marginRight: '8px',
                            border: '1px solid #9aa5b1',
                            borderRadius: '4px',
                            background: '#fff',
                            color: '#334155',
                            fontWeight: 700,
                            cursor: 'pointer'
                          }}
                        >
                          {expandedNodes[node.id] ? '-' : '+'}
                        </button>
                      ) : (
                        <span style={{ width: '26px', minWidth: '26px' }} />
                      )}

                      <span
                        style={{
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          fontWeight: styles.hierarchyWeight
                        }}
                      >
                        {node.label}
                      </span>
                    </div>
                  </td>

                  {selectedMeasures.map((field) => {
                    const width = measureColumnWidths[field];
                    return (
                      <td
                        key={field}
                        style={{
                          border: '1px solid #e8edf3',
                          background: index % 2 === 0 ? '#ffffff' : '#f7fafc',
                          padding: '8px 10px',
                          textAlign: 'right',
                          width: `${width}px`,
                          boxSizing: 'border-box',
                          whiteSpace: 'nowrap',
                          fontSize: '12px',
                          fontWeight: styles.measureWeight
                        }}
                      >
                        {formatMeasureValue(field, node.measures[field] ?? 0)}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }

  function renderStraightTable() {
    return (
      <div style={{ width: '100%', overflow: 'auto' }}>
        <table
          style={{
            borderCollapse: 'collapse',
            width: 'max-content',
            minWidth: '100%',
            tableLayout: 'fixed'
          }}
        >
          <thead>
            <tr>
              {selectedDimensions.map((field) => (
                <th
                  key={field}
                  style={{
                    background: '#075b67',
                    color: '#fff',
                    border: '1px solid #dce4ec',
                    padding: '8px 10px',
                    textAlign: 'left',
                    fontSize: '12px'
                  }}
                >
                  {getDisplayLabel(field)}
                </th>
              ))}
              {selectedMeasures.map((field) => (
                <th
                  key={field}
                  style={{
                    background: '#075b67',
                    color: '#fff',
                    border: '1px solid #dce4ec',
                    padding: '8px 10px',
                    textAlign: 'right',
                    fontSize: '12px'
                  }}
                >
                  {getDisplayLabel(field)}
                </th>
              ))}
            </tr>
          </thead>

          <tbody>
            {filteredRows.map((row, index) => (
              <tr key={index}>
                {selectedDimensions.map((field) => (
                  <td
                    key={field}
                    style={{
                      border: '1px solid #e8edf3',
                      background: index % 2 === 0 ? '#ffffff' : '#f7fafc',
                      padding: '8px 10px',
                      fontSize: '12px'
                    }}
                  >
                    {row[field] ?? ''}
                  </td>
                ))}

                {selectedMeasures.map((field) => (
                  <td
                    key={field}
                    style={{
                      border: '1px solid #e8edf3',
                      background: index % 2 === 0 ? '#ffffff' : '#f7fafc',
                      padding: '8px 10px',
                      textAlign: 'right',
                      whiteSpace: 'nowrap',
                      fontSize: '12px'
                    }}
                  >
                    {formatMeasureValue(field, parseNumber(row[field]))}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  return (
    <div
      style={{
        width: '100%',
        minHeight: '100vh',
        background: '#ffffff',
        fontFamily: 'Arial, sans-serif',
        color: '#111827',
        boxSizing: 'border-box',
        cursor: isResizing ? 'col-resize' : 'default'
      }}
    >
      <div style={{ padding: '12px 18px 0' }}>
        <div
          style={{
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            gap: '16px',
            marginBottom: '8px'
          }}
        >
          <div>
            <div style={{ fontSize: '18px', fontWeight: 700, color: '#0f172a', marginBottom: '4px' }}>
              Report Builder
            </div>
            <div style={{ fontSize: '12px', color: '#6b7280' }}>
              Configure and export your custom reports
            </div>
          </div>

          <div style={{ display: 'grid', gap: '8px' }}>
            <button
              type="button"
              onClick={exportCurrentView}
              style={{
                height: '38px',
                padding: '0 14px',
                border: 'none',
                borderRadius: '8px',
                background: 'linear-gradient(180deg, #24c5c0 0%, #16b3ad 100%)',
                color: '#fff',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              Export Report ▼
            </button>

            <button
              type="button"
              onClick={() => setIsInfoOpen((prev) => !prev)}
              style={{
                height: '34px',
                padding: '0 14px',
                border: '1px solid #d6dee7',
                borderRadius: '8px',
                background: isInfoOpen ? '#eaf6f5' : '#f7f9fc',
                color: '#0f172a',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              Information
            </button>

            <button
              type="button"
              onClick={configure}
              style={{
                height: '34px',
                padding: '0 14px',
                border: '1px solid #d6dee7',
                borderRadius: '8px',
                background: '#f7f9fc',
                color: '#0f172a',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              Configure
            </button>
          </div>
        </div>

        {isInfoOpen && (
          <div
            style={{
              marginBottom: '10px',
              border: '1px solid #e3e8ef',
              borderRadius: '14px',
              background: '#fbfdff',
              padding: '14px'
            }}
          >
            <div style={{ fontSize: '13px', fontWeight: 700, color: '#0f172a', marginBottom: '10px' }}>
              Information Panel
            </div>

            {infoResults.length > 0 ? (
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
                  gap: '10px'
                }}
              >
                {infoResults.map((item) => (
                  <div
                    key={item.id}
                    style={{
                      border: '1px solid #e6ebf2',
                      borderRadius: '12px',
                      background: '#ffffff',
                      padding: '12px'
                    }}
                  >
                    <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 700, marginBottom: '6px' }}>
                      {item.label}
                    </div>
                    <div style={{ fontSize: '18px', color: '#0f172a', fontWeight: 700 }}>
                      {item.value}
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ fontSize: '12px', color: '#6b7280' }}>
                No information calculations have been configured yet.
              </div>
            )}
          </div>
        )}

        <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', marginBottom: '10px', alignItems: 'center' }}>
          {renderToolbarDropdown(
            `Data Source(s)${selectedSources.length ? ` (${selectedSources.length})` : ''}`,
            '🛢️',
            sourceMenuOpen,
            setSourceMenuOpen,
            sourceMenuRef,
            renderSourceMenu(),
            '#2563eb'
          )}

          {renderToolbarDropdown(
            'Date Range',
            '🗓️',
            dateMenuOpen,
            setDateMenuOpen,
            dateMenuRef,
            renderDateMenu(),
            '#a855f7'
          )}

          {renderToolbarDropdown(
            `Dimensions${selectedDimensions.length ? ` (${selectedDimensions.length})` : ''}`,
            '📊',
            dimensionMenuOpen,
            setDimensionMenuOpen,
            dimensionMenuRef,
            renderDimensionMenu(),
            '#10b981'
          )}

          {renderToolbarDropdown(
            `Metrics${selectedMeasures.length ? ` (${selectedMeasures.length})` : ''}`,
            '↗',
            measureMenuOpen,
            setMeasureMenuOpen,
            measureMenuRef,
            renderMeasureMenu(),
            '#f97316'
          )}

          {renderToolbarDropdown(
            `Filters${filters.length ? ` (${filters.length})` : ''}`,
            '⏃',
            filterMenuOpen,
            setFilterMenuOpen,
            filterMenuRef,
            renderFilterMenu(),
            '#7c3aed'
          )}
        </div>
      </div>

      {message && (
        <div style={{ margin: '0 18px 10px', color: '#b00020', fontSize: '13px' }}>
          {message}
        </div>
      )}

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: isSelectionsPanelOpen ? '150px 1fr' : '32px 1fr',
          gap: '0',
          borderTop: '1px solid #e5e7eb',
          minHeight: 'calc(100vh - 108px)'
        }}
      >
        <aside
          style={{
            borderRight: '1px solid #e5e7eb',
            background: '#ffffff',
            overflow: 'hidden'
          }}
        >
          {isSelectionsPanelOpen ? (
            <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  padding: '12px 8px 10px 10px',
                  borderBottom: '1px solid #eef2f7'
                }}
              >
                <div style={{ fontSize: '12px', fontWeight: 700, color: '#374151' }}>
                  Current Selections
                </div>

                <button
                  type="button"
                  onClick={() => setIsSelectionsPanelOpen(false)}
                  style={{
                    width: '22px',
                    height: '22px',
                    borderRadius: '6px',
                    border: '1px solid #e1e7ef',
                    background: '#f7f9fc',
                    cursor: 'pointer',
                    color: '#64748b',
                    fontSize: '12px',
                    fontWeight: 700
                  }}
                  aria-label="Collapse current selections"
                >
                  ‹
                </button>
              </div>

              <div style={{ padding: '10px' }}>
                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Data Source(s):
                  </div>
                  {selectedSources.length > 0
                    ? selectedSources.map((source) => (
                        <div key={source}>
                          {renderSelectionPill(
                            source,
                            '#065f73',
                            selectedSources.length > 1 ? () => toggleSource(source) : undefined
                          )}
                        </div>
                      ))
                    : <div style={{ fontSize: '12px', color: '#94a3b8' }}>None selected</div>}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Date Range:
                  </div>
                  {(dateRange.start || dateRange.end)
                    ? renderSelectionPill(`${dateRange.start || '...'} - ${dateRange.end || '...'}`, '#065f73')
                    : <div style={{ fontSize: '12px', color: '#94a3b8' }}>Not set</div>}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Active Filters:
                  </div>
                  {filters.length > 0
                    ? filters.map((filter) => (
                        <div key={filter.field}>
                          {renderSelectionPill(
                            `${getDisplayLabel(filter.field)}: ${filter.values.join(', ')}`,
                            '#065f73',
                            () => removeFilter(filter.field)
                          )}
                        </div>
                      ))
                    : <div style={{ fontSize: '12px', color: '#94a3b8', marginBottom: '12px' }}>None applied</div>}
                </div>

                <button
                  type="button"
                  onClick={clearAllSelections}
                  style={{
                    width: '100%',
                    height: '32px',
                    border: '1px solid #d8e2ee',
                    borderRadius: '8px',
                    background: '#fff7ed',
                    color: '#c2410c',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer'
                  }}
                >
                  Clear All
                </button>
              </div>
            </div>
          ) : (
            <div style={{ paddingTop: '10px', display: 'flex', justifyContent: 'center' }}>
              <button
                type="button"
                onClick={() => setIsSelectionsPanelOpen(true)}
                style={{
                  width: '22px',
                  height: '22px',
                  borderRadius: '6px',
                  border: '1px solid #e1e7ef',
                  background: '#f7f9fc',
                  cursor: 'pointer',
                  color: '#64748b',
                  fontSize: '12px',
                  fontWeight: 700
                }}
                aria-label="Expand current selections"
              >
                ›
              </button>
            </div>
          )}
        </aside>

        <main style={{ padding: '10px 12px 14px', minWidth: 0 }}>
          <div style={{ display: 'flex', gap: '8px', marginBottom: '10px' }}>
            <button
              type="button"
              onClick={() => setViewMode('straight')}
              style={{
                height: '32px',
                padding: '0 12px',
                borderRadius: '8px',
                border: viewMode === 'straight' ? 'none' : '1px solid #d6dee7',
                background: viewMode === 'straight' ? '#075b67' : '#f8fafc',
                color: viewMode === 'straight' ? '#fff' : '#334155',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              Straight Table
            </button>

            <button
              type="button"
              onClick={() => setViewMode('pivot')}
              style={{
                height: '32px',
                padding: '0 12px',
                borderRadius: '8px',
                border: viewMode === 'pivot' ? 'none' : '1px solid #d6dee7',
                background: viewMode === 'pivot' ? '#075b67' : '#f8fafc',
                color: viewMode === 'pivot' ? '#fff' : '#334155',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              Pivot Table
            </button>
          </div>

          <div style={{ minWidth: 0 }}>
            {viewMode === 'pivot' ? renderPivotTable() : renderStraightTable()}
          </div>
        </main>
      </div>
    </div>
  );
}