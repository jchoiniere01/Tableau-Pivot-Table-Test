// app.tsx
// Legacy-compatible Tableau build: use getSummaryDataAsync only

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import exportIcon from './assets/icons/export-icon.png';
import { Grid } from 'react-window';
import {
  DndContext,
  PointerSensor,
  KeyboardSensor,
  closestCenter,
  useSensor,
  useSensors
} from '@dnd-kit/core';
import {
  SortableContext,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
  sortableKeyboardCoordinates
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';

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

type PivotNode = {
  id: string;
  label: string;
  level: number;
  dimensionIndex: number;
  hasChildren: boolean;
  parentId?: string;
  measures: Record<string, number>;
  children: PivotNode[];
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

type SortableSelectionPillProps = {
  id: string;
  label: string;
  index: number;
  total: number;
  onMoveUp: () => void;
  onMoveDown: () => void;
  onEdit?: () => void;
  onRemove?: () => void;
};

type SelectionPillProps = {
  text: string;
  color?: string;
  options?: {
    onRemove?: () => void;
    onMoveUp?: () => void;
    onMoveDown?: () => void;
    onEdit?: () => void;
    disableMoveUp?: boolean;
    disableMoveDown?: boolean;
  };
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

const MAX_PIVOT_ROWS = 25000;
const STRAIGHT_ROW_HEIGHT = 36;
const STRAIGHT_HEADER_HEIGHT = 38;

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

function SelectionPill({
  text,
  color = '#065f73',
  options
}: SelectionPillProps) {
  const {
    onRemove,
    onMoveUp,
    onMoveDown,
    onEdit,
    disableMoveUp,
    disableMoveDown
  } = options || {};

  const hasActions = !!onMoveUp || !!onMoveDown || !!onEdit || !!onRemove;

  const actionButtonStyle: React.CSSProperties = {
    width: '18px',
    height: '18px',
    minWidth: '18px',
    borderRadius: '999px',
    border: 'none',
    background: 'rgba(255,255,255,0.18)',
    color: '#fff',
    fontSize: '11px',
    fontWeight: 700,
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 0,
    lineHeight: 1,
    opacity: 1
  };

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
        fontWeight: 700
      }}
    >
      <span
        style={{
          paddingLeft: '15px',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis'
        }}
      >
        {text}
      </span>

      {hasActions && (
        <div
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '4px',
            flexShrink: 0
          }}
        >
          {onMoveUp && (
            <button
              type="button"
              onClick={onMoveUp}
              disabled={disableMoveUp}
              style={{
                ...actionButtonStyle,
                opacity: disableMoveUp ? 0.45 : 1,
                cursor: disableMoveUp ? 'default' : 'pointer'
              }}
              aria-label={`Move ${text} up`}
            >
              ↑
            </button>
          )}

          {onMoveDown && (
            <button
              type="button"
              onClick={onMoveDown}
              disabled={disableMoveDown}
              style={{
                ...actionButtonStyle,
                opacity: disableMoveDown ? 0.45 : 1,
                cursor: disableMoveDown ? 'default' : 'pointer'
              }}
              aria-label={`Move ${text} down`}
            >
              ↓
            </button>
          )}

          {onEdit && (
            <button
              type="button"
              onClick={onEdit}
              style={actionButtonStyle}
              aria-label={`Edit label for ${text}`}
            >
              ✎
            </button>
          )}

          {onRemove && (
            <button
              type="button"
              onClick={onRemove}
              style={actionButtonStyle}
              aria-label={`Remove ${text}`}
            >
              ×
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function SortableSelectionPill({
  id,
  label,
  index,
  total,
  onMoveUp,
  onMoveDown,
  onEdit,
  onRemove
}: SortableSelectionPillProps) {
  const {
    attributes,
    listeners,
    setNodeRef,
    setActivatorNodeRef,
    transform,
    transition,
    isDragging
  } = useSortable({ id });

  return (
    <div
      ref={setNodeRef}
      style={{
        transform: CSS.Transform.toString(transform),
        transition,
        position: 'relative',
        marginBottom: '8px',
        opacity: isDragging ? 0.92 : 1,
        zIndex: isDragging ? 2 : 'auto'
      }}
    >
      <button
        type="button"
        ref={setActivatorNodeRef}
        {...attributes}
        {...listeners}
        title="Drag to reorder"
        aria-label={`Drag to reorder ${label}`}
        style={{
          position: 'absolute',
          left: '8px',
          top: '50%',
          transform: 'translateY(-50%)',
          width: '16px',
          height: '16px',
          border: 'none',
          background: 'transparent',
          color: '#ffffff',
          cursor: isDragging ? 'grabbing' : 'grab',
          fontSize: '12px',
          lineHeight: 1,
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: 0,
          userSelect: 'none',
          zIndex: 3
        }}
      >
        ⋮⋮
      </button>

      <div>
        <SelectionPill
          text={label}
          color="#065f73"
          options={{
            onMoveUp,
            onMoveDown,
            onEdit,
            onRemove,
            disableMoveUp: index === 0,
            disableMoveDown: index === total - 1
          }}
        />
      </div>
    </div>
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

function dataTableToRows(dataTable: any, worksheetName: string): PreviewRow[] {
  return dataTable.data.map((row: any[]) => {
    const obj: PreviewRow = { __sourceWorksheet: worksheetName };

    dataTable.columns.forEach((col: any, index: number) => {
      obj[col.fieldName] = row[index]?.value ?? row[index]?.formattedValue ?? '';
    });

    return obj;
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

  return Array.from(map.entries()).map(([label, group]) => ({
    label,
    fields: group.sort((a, b) => a.caption.localeCompare(b.caption))
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

  const [fieldLabelOverrides, setFieldLabelOverrides] = useState<Record<string, string>>({});
  const [editingFieldLabel, setEditingFieldLabel] = useState<string | null>(null);
  const [editingFieldValue, setEditingFieldValue] = useState('');
  const [editingFieldFormat, setEditingFieldFormat] = useState('number');
  const [fieldFormatOverrides, setFieldFormatOverrides] = useState<Record<string, string>>({});

  const [worksheetFields, setWorksheetFields] = useState<WorksheetFieldsMap>({});
  const [fieldRoleOverrides, setFieldRoleOverrides] = useState<FieldRoleOverrideMap>({});
  const [expandedWorksheetFields, setExpandedWorksheetFields] = useState<WorksheetFieldExpandMap>({});

  const [sortConfig, setSortConfig] = useState<{
    field: string;
    direction: 'asc' | 'desc';
  } | null>(null);

  const [headerFilterMenu, setHeaderFilterMenu] = useState<{
    field: string;
    anchorRect: DOMRect | null;
  } | null>(null);

  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [columnFilterSearch, setColumnFilterSearch] = useState('');

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

  const fieldMapRef = useRef<WorksheetFieldsMap>({});

  const [fieldFormats, setFieldFormats] = useState<Record<string, string>>({});

  const [loadedRows, setLoadedRows] = useState<PreviewRow[]>([]);
  const [totalAvailableRows, setTotalAvailableRows] = useState(0);
  const [isPartialData, setIsPartialData] = useState(false);
  const [isLoadingData, setIsLoadingData] = useState(false);

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
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const exportMenuRef = useRef<HTMLDivElement | null>(null);

  const sourceMenuRef = useRef<HTMLDivElement | null>(null);
  const dateMenuRef = useRef<HTMLDivElement | null>(null);
  const dimensionMenuRef = useRef<HTMLDivElement | null>(null);
  const measureMenuRef = useRef<HTMLDivElement | null>(null);
  const filterMenuRef = useRef<HTMLDivElement | null>(null);
  const straightTableViewportRef = useRef<HTMLDivElement | null>(null);
  const straightTableHeaderRef = useRef<HTMLDivElement | null>(null);
  const straightTableGridRef = useRef<any>(null);
  const straightTableHeaderContentRef = useRef<HTMLDivElement | null>(null);
  const handleStraightTableGridScroll: React.UIEventHandler<HTMLDivElement> = useCallback(
    () => {
      const body =
        ((straightTableGridRef.current as any)?.element ??
          (straightTableGridRef.current as any)?._outerRef) as HTMLDivElement | null;

      if (!body) return;

      const left = body.scrollLeft;

      if (straightTableHeaderRef.current && straightTableHeaderRef.current.scrollLeft !== left) {
        straightTableHeaderRef.current.scrollLeft = left;
      }
    },
    []
  );
  
  const [straightTableViewportSize, setStraightTableViewportSize] = useState({
    width: 0,
    height: 0
  });

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 6
      }
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates
    })
  );

  function handleDimensionDragEnd(event: any) {
    const { active, over } = event;

    if (!over || active.id === over.id) return;

    setSelectedDimensions((prev) => {
      const oldIndex = prev.indexOf(String(active.id));
      const newIndex = prev.indexOf(String(over.id));

      if (oldIndex === -1 || newIndex === -1) return prev;

      return arrayMove(prev, oldIndex, newIndex);
    });
  }

  function handleMeasureDragEnd(event: any) {
    const { active, over } = event;

    if (!over || active.id === over.id) return;

    setSelectedMeasures((prev) => {
      const oldIndex = prev.indexOf(String(active.id));
      const newIndex = prev.indexOf(String(over.id));

      if (oldIndex === -1 || newIndex === -1) return prev;

      return arrayMove(prev, oldIndex, newIndex);
    });
  }

  const unregisterHandlersRef = useRef<(() => void)[]>([]);
  const refreshTimerRef = useRef<number | null>(null);
  const initialLoadCompleteRef = useRef(false);

  function getDisplayLabel(field: string): string {
    const override = fieldLabelOverrides[field]?.trim();
    if (override) return override;
    return availableFields.find((f) => f.fieldName === field)?.caption || field;
  }

  const straightTableColumns = useMemo(
    () => [
      ...selectedDimensions.map((field) => ({
        key: field,
        label: getDisplayLabel(field),
        width: 180,
        kind: 'dimension' as const
      })),
      ...selectedMeasures.map((field) => ({
        key: field,
        label: getDisplayLabel(field),
        width: 140,
        kind: 'measure' as const
      }))
    ],
    [selectedDimensions, selectedMeasures, fieldLabelOverrides, availableFields]
  );

  const getStraightTableColumnWidth = (index: number) =>
    straightTableColumns[index]?.width ?? 140;

  const straightTableTotalWidth = useMemo(() => {
    return straightTableColumns.reduce(
      (_, __, index) => _ + getStraightTableColumnWidth(index),
      0
    );
  }, [straightTableColumns, layoutSettings, selectedDimensions, selectedMeasures]);

  async function loadWorksheetSummaryLegacy(
    worksheet: any
  ): Promise<{ rows: PreviewRow[]; fields: FieldMeta[]; totalRowCount: number }> {
    const table = await worksheet.getSummaryDataAsync({ maxRows: 0 });
    const fields = buildFieldMeta(table.columns);
    const rows = dataTableToRows(table, worksheet.name);
    const totalRowCount = rows.length;

    return { rows, fields, totalRowCount };
  }

  async function loadAllWorksheetRowsForExport(
    worksheet: any
  ): Promise<{ rows: PreviewRow[]; fields: FieldMeta[]; totalRowCount: number }> {
    const table = await worksheet.getSummaryDataAsync({ maxRows: 0 });
    const fields = buildFieldMeta(table.columns);
    const rows = dataTableToRows(table, worksheet.name);
    const totalRowCount = rows.length;

    return { rows, fields, totalRowCount };
  }

  function buildStraightExportRowsFromRows(rows: any[]) {
    return rows.map((row) => {
      const exportRow: Record<string, string | number> = {};

      for (const field of selectedDimensions) {
        exportRow[getDisplayLabel(field)] = row[field] ?? '';
      }

      for (const field of selectedMeasures) {
        exportRow[getDisplayLabel(field)] = parseNumber(row[field]);
      }

      return exportRow;
    });
  }

  function buildPivotTreeFromRows(rows: any[]): PivotNode[] {
    if (selectedDimensions.length === 0) {
      return [];
    }

    function buildLevel(
      levelRows: any[],
      dimensionIndex: number,
      parentId?: string,
      parentPath = ''
    ): PivotNode[] {
      const field = selectedDimensions[dimensionIndex];
      if (!field) return [];

      const grouped = new Map<string, any[]>();

      for (const row of levelRows) {
        const rawValue = row[field];
        const label =
          rawValue === null || rawValue === undefined || rawValue === ''
            ? '(Blank)'
            : String(rawValue);

        const bucket = grouped.get(label);
        if (bucket) {
          bucket.push(row);
        } else {
          grouped.set(label, [row]);
        }
      }

      const nodes = Array.from(grouped.entries()).map(([label, groupRows]) => {
        const nodeId = parentPath ? `${parentPath}__${field}:${label}` : `${field}:${label}`;

        const measures: Record<string, number> = {};
        for (const measure of selectedMeasures) {
          let total = 0;
          for (const row of groupRows) {
            const value = parseNumber(row[measure]);
            total += Number.isFinite(value) ? value : 0;
          }
          measures[measure] = total;
        }

        const children =
          dimensionIndex < selectedDimensions.length - 1
            ? buildLevel(groupRows, dimensionIndex + 1, nodeId, nodeId)
            : [];

        return {
          id: nodeId,
          label,
          level: dimensionIndex,
          dimensionIndex,
          hasChildren: children.length > 0,
          parentId,
          measures,
          children
        };
      });

      nodes.sort((a, b) => a.label.localeCompare(b.label));
      return nodes;
    }

    return buildLevel(rows, 0);
  }

  function flattenAllPivotNodes(nodes: PivotNode[]): PivotNode[] {
    const result: PivotNode[] = [];

    function walk(items: PivotNode[]) {
      for (const node of items) {
        result.push(node);
        if (node.children.length > 0) {
          walk(node.children);
        }
      }
    }

    walk(nodes);
    return result;
  }

  function buildPivotExportRowsFromNodes(nodes: PivotNode[]) {
    const hierarchyLabel =
      selectedDimensions.length > 0
        ? selectedDimensions.map(getDisplayLabel).join(' / ')
        : 'Hierarchy';

    return nodes.map((node) => {
      const row: Record<string, string | number> = {
        [hierarchyLabel]: `${'  '.repeat(node.level)}${node.label}`
      };

      for (const field of selectedMeasures) {
        row[getDisplayLabel(field)] = node.measures[field] ?? 0;
      }

      return row;
    });
  }

  async function getExportRowsForCurrentSelection(): Promise<{
    rows: PreviewRow[];
    fields: FieldMeta[];
    totalRowCount: number;
  }> {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const selectedWorksheets = dashboard.worksheets.filter((w: any) =>
      selectedSources.includes(w.name)
    );

    const results = await Promise.all(
      selectedWorksheets.map((worksheet: any) =>
        loadAllWorksheetRowsForExport(worksheet)
      )
    );

    return {
      rows: results.flatMap((result) => result.rows),
      fields: results.flatMap((result) => result.fields),
      totalRowCount: results.reduce(
        (sum, result) => sum + result.totalRowCount,
        0
      )
    };
  }

  async function handleExportStraightTable() {
    const { rows } = await getExportRowsForCurrentSelection();
    const filteredExportRows = applyAllActiveFilters(rows);
    const exportRows = buildStraightExportRowsFromRows(filteredExportRows);
    exportRowsToExcel(exportRows);
  }

  async function handleExportPivotRawData() {
    const { rows } = await getExportRowsForCurrentSelection();
    const filteredExportRows = applyAllActiveFilters(rows);
    const exportRows = buildStraightExportRowsFromRows(filteredExportRows);
    exportRowsToExcel(exportRows);
  }

  async function handleExportPivotExpanded() {
    const { rows } = await getExportRowsForCurrentSelection();
    const filteredExportRows = applyAllActiveFilters(rows);
    const tree = buildPivotTreeFromRows(filteredExportRows);
    const allNodes = flattenAllPivotNodes(tree);
    const exportRows = buildPivotExportRowsFromNodes(allNodes);
    exportRowsToExcel(exportRows);
  }

  async function exportCurrentView() {
    if (selectedMeasures.length === 0) {
      setMessage('Please select at least one metric.');
      return;
    }

    if (viewMode === 'straight') {
      await handleExportStraightTable();
    } else {
      await handleExportPivotExpanded();
    }
  }

  function clearEventListeners() {
    unregisterHandlersRef.current.forEach((unregister) => {
      try {
        unregister();
      } catch {
        //
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

  function isHeaderFilterValueSelected(field: string, value: string): boolean {
    return !!columnFilters[field]?.includes(value);
  }

  function toggleHeaderFilterValue(field: string, value: string) {
    setColumnFilters((prev) => {
      const current = prev[field] || [];
      const exists = current.includes(value);

      const nextValues = exists
        ? current.filter((v) => v !== value)
        : [...current, value];

      const next: Record<string, string[]> = { ...prev };
      if (nextValues.length === 0) {
        delete next[field];
      } else {
        next[field] = nextValues;
      }

      return next;
    });
  }

  function clearHeaderFilter(field: string) {
    setColumnFilters((prev) => {
      const next: Record<string, string[]> = { ...prev };
      delete next[field];
      return next;
    });
  }

  function getHeaderFilterOptions(field: string): string[] {
    return Array.from(
      new Set(
        loadedRows.map((row) => {
          const value = row?.[field];
          return value == null || value === '' ? '(Blank)' : String(value);
        })
      )
    ).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
    );
  }

  useEffect(() => {
    const element = straightTableViewportRef.current;
    if (!element) return;

    const updateSize = () => {
      const rect = element.getBoundingClientRect();
      setStraightTableViewportSize({
        width: Math.max(320, Math.floor(rect.width)),
        height: Math.max(240, Math.floor(rect.height))
      });
    };

    updateSize();

    const observer = new ResizeObserver(() => {
      updateSize();
    });

    observer.observe(element);
    window.addEventListener('resize', updateSize);

    return () => {
      observer.disconnect();
      window.removeEventListener('resize', updateSize);
    };
  }, []);

  function getDisplayLabelFromWorksheetField(field: FieldMeta): string {
    const override = fieldLabelOverrides[field.fieldName]?.trim();
    if (override) return override;
    return field.caption || field.fieldName;
  }

  function getMeasureFormat(_field: string): MeasureFormat {
    return DEFAULT_MEASURE_FORMAT;
  }

  function getSavedMeasureWidth(field: string): number {
    return layoutSettings.measureWidths[field] || layoutSettings.defaultMeasureWidth;
  }

  function formatMeasureValue(field: string, value: number) {
    const format = fieldFormatOverrides[field] || 'number';

    switch (format) {
      case 'currency':
        return value.toLocaleString(undefined, {
          style: 'currency',
          currency: 'USD',
          maximumFractionDigits: 2
        });
      case 'percent':
        return `${(value * 100).toFixed(1)}%`;
      case 'integer':
        return value.toLocaleString(undefined, {
          maximumFractionDigits: 0
        });
      case 'decimal-1':
        return value.toLocaleString(undefined, {
          minimumFractionDigits: 1,
          maximumFractionDigits: 1
        });
      case 'decimal-2':
        return value.toLocaleString(undefined, {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        });
      case 'decimal-3':
        return value.toLocaleString(undefined, {
          minimumFractionDigits: 3,
          maximumFractionDigits: 3
        });
      case 'number':
      default:
        return value.toLocaleString();
    }
  }

  function getFieldFormat(field: string) {
    return fieldFormats[field] || 'number';
  }

  function isMeasureField(field: string) {
    return selectedMeasures.includes(field);
  }

  function toggleSort(field: string) {
    setSortConfig((prev) => {
      if (!prev || prev.field !== field) {
        return { field, direction: 'asc' };
      }

      return {
        field,
        direction: prev.direction === 'asc' ? 'desc' : 'asc'
      };
    });
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

  function moveArrayItem<T>(items: T[], fromIndex: number, toIndex: number): T[] {
    if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= items.length || toIndex >= items.length) {
      return items;
    }

    const next = [...items];
    const [moved] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, moved);
    return next;
  }

  function moveDimension(field: string, direction: 'up' | 'down') {
    setSelectedDimensions((prev) => {
      const index = prev.indexOf(field);
      if (index === -1) return prev;
      const nextIndex = direction === 'up' ? index - 1 : index + 1;
      return moveArrayItem(prev, index, nextIndex);
    });
  }

  function moveMeasure(field: string, direction: 'up' | 'down') {
    setSelectedMeasures((prev) => {
      const index = prev.indexOf(field);
      if (index === -1) return prev;
      const nextIndex = direction === 'up' ? index - 1 : index + 1;
      return moveArrayItem(prev, index, nextIndex);
    });
  }

  function startEditingFieldLabel(field: string) {
    setEditingFieldLabel(field);
    setEditingFieldValue(getDisplayLabel(field));
    setEditingFieldFormat(getFieldFormat(field));
  }

  function saveFieldLabel(field: string) {
    const trimmed = editingFieldValue.trim();
    const original = availableFields.find((f) => f.fieldName === field)?.caption || field;

    setFieldLabelOverrides((prev) => {
      const next = { ...prev };
      if (!trimmed || trimmed === original) {
        delete next[field];
      } else {
        next[field] = trimmed;
      }
      return next;
    });

    if (isMeasureField(field)) {
      setFieldFormatOverrides((prev) => {
        const next = { ...prev };

        if (!editingFieldFormat || editingFieldFormat === 'number') {
          delete next[field];
        } else {
          next[field] = editingFieldFormat;
        }

        return next;
      });
    }

    setEditingFieldLabel(null);
    setEditingFieldValue('');
    setEditingFieldFormat('number');
  }

  function cancelFieldLabelEdit() {
    setEditingFieldLabel(null);
    setEditingFieldValue('');
    setEditingFieldFormat('number');
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
    setColumnFilters({});
    setHeaderFilterMenu(null);
    setColumnFilterSearch('');
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
        !selectedDimensions.includes(field.fieldName) &&
        !looksDateType(field.dataType) &&
        !looksLikeDateField(field.fieldName)
    );

    return dimensionCandidates
      .map((field) => ({
        label: field.caption,
        field: field.fieldName,
        options: getFieldDistinctValuesFromRows(loadedRows, field.fieldName).slice(0, 25)
      }))
      .filter((group) => group.options.length > 0);
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
        //
      }

      if (tableau.TableauEventType?.SummaryDataChanged) {
        try {
          const summaryHandler = worksheet.addEventListener(
            tableau.TableauEventType.SummaryDataChanged,
            () => debouncedRefresh()
          );
          unregisterHandlersRef.current.push(summaryHandler);
        } catch {
          //
        }
      }
    });
  }


  
  function rebuildFieldsFromSelectedSources(
    sourceNames: string[],
    currentWorksheetFields: WorksheetFieldsMap,
    overrides: FieldRoleOverrideMap
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
            role: effectiveRole
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

    setFilters((prev) => prev.filter((filter) => dims.includes(filter.field)));

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
        setLoadedRows([]);
        setAvailableFields([]);
        setTotalAvailableRows(0);
        setIsPartialData(false);
        setMessage('Select at least one data source.');
        return;
      }

      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const selectedWorksheets = dashboard.worksheets.filter((w: any) =>
        sourceNames.includes(w.name)
      );

      if (!isRefresh) setMessage('Loading data...');
      setIsLoadingData(true);

      try {
        const results = await Promise.all(
          selectedWorksheets.map((worksheet: any) => loadWorksheetSummaryLegacy(worksheet))
        );

        const mergedRows = results.flatMap((r) => r.rows);
        const totalRows = results.reduce((sum, r) => sum + r.totalRowCount, 0);

        const nextWorksheetFields: WorksheetFieldsMap = {};
        selectedWorksheets.forEach((worksheet: any, index: number) => {
          nextWorksheetFields[worksheet.name] = results[index].fields;
        });

        fieldMapRef.current = nextWorksheetFields;

        setLoadedRows(mergedRows);
        setTotalAvailableRows(totalRows);
        setIsPartialData(false);

        setWorksheetFields((prev) => {
          const combined = { ...prev, ...nextWorksheetFields };
          rebuildFieldsFromSelectedSources(sourceNames, combined, fieldRoleOverrides);
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
      } catch (error: any) {
        setMessage(`Load error: ${error?.message || error}`);
      } finally {
        setIsLoadingData(false);
      }
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
    setIsPartialData(loadedRows.length < totalAvailableRows);
  }, [loadedRows, totalAvailableRows]);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;

      if (sourceMenuRef.current && !sourceMenuRef.current.contains(target)) setSourceMenuOpen(false);
      if (dateMenuRef.current && !dateMenuRef.current.contains(target)) setDateMenuOpen(false);
      if (dimensionMenuRef.current && !dimensionMenuRef.current.contains(target)) setDimensionMenuOpen(false);
      if (measureMenuRef.current && !measureMenuRef.current.contains(target)) setMeasureMenuOpen(false);
      if (filterMenuRef.current && !filterMenuRef.current.contains(target)) setFilterMenuOpen(false);
      if (exportMenuRef.current && !exportMenuRef.current.contains(target)) {setExportMenuOpen(false);}
    }

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (!initialLoadCompleteRef.current) return;

    if (selectedSources.length > 0) {
      loadAllSelectedWorksheets(selectedSources);
    } else {
      setLoadedRows([]);
      setAvailableFields([]);
      setTotalAvailableRows(0);
      setIsPartialData(false);
    }
  }, [selectedSources]);

  useEffect(() => {
    if (selectedSources.length > 0 && Object.keys(worksheetFields).length > 0) {
      rebuildFieldsFromSelectedSources(selectedSources, worksheetFields, fieldRoleOverrides);
    }
  }, [fieldRoleOverrides]);

  function applyAllActiveFilters(rows: PreviewRow[]): PreviewRow[] {
    return rows.filter((row) => {
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

      for (const [field, selectedValues] of Object.entries(columnFilters)) {
        if (!selectedValues.length) continue;
        const rawValue = row[field];
        const value = rawValue == null || rawValue === '' ? '(Blank)' : String(rawValue);
        if (!selectedValues.includes(value)) return false;
      }

      return true;
    });
  }

  const filteredRows = useMemo(() => {
    return applyAllActiveFilters(loadedRows);
  }, [loadedRows, dateRange, filters, columnFilters]);

  const visibleRowCount = filteredRows.length;
  const loadedRowCount = loadedRows.length;
  const availableRowCount = totalAvailableRows;

  const hasActiveFilters =
    (dateRange.field && (dateRange.start || dateRange.end)) ||
    filters.length > 0 ||
    Object.values(columnFilters).some((values) => values.length > 0);

  const sortedStraightRows = useMemo(() => {
    if (!sortConfig) return filteredRows;

    const { field, direction } = sortConfig;
    const multiplier = direction === 'asc' ? 1 : -1;
    const isMeasure = selectedMeasures.includes(field);

    return [...filteredRows].sort((a, b) => {
      if (isMeasure) {
        const aNum = parseNumber(a[field]);
        const bNum = parseNumber(b[field]);
        return (aNum - bNum) * multiplier;
      }

      const aVal = String(a[field] ?? '').toLowerCase();
      const bVal = String(b[field] ?? '').toLowerCase();
      return aVal.localeCompare(bVal, undefined, { numeric: true, sensitivity: 'base' }) * multiplier;
    });
  }, [filteredRows, sortConfig, selectedMeasures]);

  const canUsePivot = filteredRows.length <= MAX_PIVOT_ROWS && !isPartialData;

  useEffect(() => {
    if (!canUsePivot) {
      setTreeRoots([]);
      return;
    }

    const roots = buildTree(filteredRows, selectedDimensions, selectedMeasures);
    setTreeRoots(roots);

    setExpandedNodes((prev) => {
      const next = { ...prev };
      roots.forEach((node) => {
        if (next[node.id] === undefined) next[node.id] = true;
      });
      return next;
    });
  }, [canUsePivot, filteredRows, selectedDimensions, selectedMeasures]);

  const pivotTreeRoots = useMemo<PivotNode[]>(() => {
    if (selectedDimensions.length === 0) {
      return [];
    }

    function buildLevel(
      rows: any[],
      dimensionIndex: number,
      parentId?: string,
      parentPath = ''
    ): PivotNode[] {
      const field = selectedDimensions[dimensionIndex];
      if (!field) {
        return [];
      }

      const grouped = new Map<string, any[]>();

      for (const row of rows) {
        const rawValue = row[field];
        const label =
          rawValue === null || rawValue === undefined || rawValue === ''
            ? '(Blank)'
            : String(rawValue);

        const bucket = grouped.get(label);
        if (bucket) {
          bucket.push(row);
        } else {
          grouped.set(label, [row]);
        }
      }

      const nodes = Array.from(grouped.entries()).map(([label, groupRows]) => {
        const nodeId = parentPath
          ? `${parentPath}__${field}:${label}`
          : `${field}:${label}`;

        const measures: Record<string, number> = {};
        for (const measure of selectedMeasures) {
          let total = 0;
          for (const row of groupRows) {
            const value = parseNumber(row[measure]);
            total += Number.isFinite(value) ? value : 0;
          }
          measures[measure] = total;
        }

        const children =
          dimensionIndex < selectedDimensions.length - 1
            ? buildLevel(groupRows, dimensionIndex + 1, nodeId, nodeId)
            : [];

        return {
          id: nodeId,
          label,
          level: dimensionIndex,
          dimensionIndex,
          hasChildren: children.length > 0,
          parentId,
          measures,
          children
        };
      });

      nodes.sort((a, b) => a.label.localeCompare(b.label));
      return nodes;
    }

    return buildLevel(filteredRows, 0);
  }, [filteredRows, selectedDimensions, selectedMeasures]);

  const visibleNodes = useMemo<PivotNode[]>(() => {
    const result: PivotNode[] = [];

    function walk(nodes: PivotNode[]) {
      for (const node of nodes) {
        result.push(node);

        if (node.hasChildren && expandedNodes[node.id]) {
          walk(node.children);
        }
      }
    }

    walk(pivotTreeRoots);
    return result;
  }, [pivotTreeRoots, expandedNodes]);

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
    [availableFields, loadedRows, selectedDimensions, filters]
  );

  const infoResults = useMemo<InfoResult[]>(() => {
    return infoCalculations.map((calc) => ({
      id: calc.id,
      label: calc.label,
      value: evaluateInfoCalculation(calc, loadedRows)
    }));
  }, [infoCalculations, loadedRows]);

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

  const [filterGroupSearch, setFilterGroupSearch] = useState<Record<string, string>>({});

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
            minWidth: '86px',
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
    const search = filterGroupSearch[group.field] || '';

    const visibleOptions = group.options.filter((option) =>
      option.toLowerCase().includes(search.trim().toLowerCase())
    );

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
            <input
              type="text"
              value={search}
              onChange={(e) =>
                setFilterGroupSearch((prev) => ({
                  ...prev,
                  [group.field]: e.target.value
                }))
              }
              onKeyDown={(e) => e.stopPropagation()}
              placeholder="Search values..."
              style={{
                width: '100%',
                padding: '8px 10px',
                border: '1px solid #d7deea',
                borderRadius: '8px',
                fontSize: '13px',
                background: '#ffffff',
                color: '#374151',
                boxSizing: 'border-box'
              }}
            />

            <div
              style={{
                maxHeight: '240px',
                overflowY: 'auto',
                display: 'grid',
                gap: '8px',
                paddingRight: '4px'
              }}
            >
              {visibleOptions.map((option) => {
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
                    <span style={{ fontSize: '13px', color: '#374151', fontWeight: 600 }}>
                      {option}
                    </span>
                  </label>
                );
              })}

              {visibleOptions.length === 0 && (
                <div style={{ fontSize: '12px', color: '#6b7280', padding: '4px 2px' }}>
                  No matching values.
                </div>
              )}
            </div>
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
              No additional filter groups are available.
            </div>
          )}
        </div>
      </div>
    );
  }

  

  function renderPivotTable() {
    return (
      <div style={{ width: '100%', height: '100%', overflow: 'auto' }}>
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

  type StraightGridCellExtraProps = {
    rows: PreviewRow[];
    columns: Array<{
      key: string;
      label: string;
      width: number;
      kind: 'dimension' | 'measure';
    }>;
    formatMeasureValue: (field: string, value: number) => string;
    parseNumber: (value: any) => number;
  };

  type StraightGridCellProps = {
    ariaAttributes?: React.HTMLAttributes<HTMLDivElement>;
    columnIndex: number;
    rowIndex: number;
    style: React.CSSProperties;
  } & StraightGridCellExtraProps;

  const StraightGridCell = ({
    ariaAttributes,
    columnIndex,
    rowIndex,
    style,
    rows,
    columns,
    formatMeasureValue,
    parseNumber
  }: StraightGridCellProps) => {
    const row = rows[rowIndex];
    const column = columns[columnIndex];

    if (!row || !column) return null;

    const isEven = rowIndex % 2 === 0;
    const field = column.key;
    const isMeasure = column.kind === 'measure';

    return (
      <div
        {...ariaAttributes}
        style={{
          ...style,
          display: 'flex',
          alignItems: 'center',
          justifyContent: isMeasure ? 'flex-end' : 'flex-start',
          padding: '8px 10px',
          fontSize: '12px',
          background: isEven ? '#ffffff' : '#f7fafc',
          borderRight: '1px solid #e8edf3',
          borderBottom: '1px solid #e8edf3',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          boxSizing: 'border-box'
        }}
      >
        {isMeasure
          ? formatMeasureValue(field, parseNumber(row[field]))
          : row[field] ?? ''}
      </div>
    );
  };

  const straightGridCellProps: StraightGridCellExtraProps = {
    rows: sortedStraightRows,
    columns: straightTableColumns,
    formatMeasureValue,
    parseNumber
  };

  const gridWidth = straightTableViewportSize.width;
  const gridHeight = Math.max(
    120,
    straightTableViewportSize.height - STRAIGHT_HEADER_HEIGHT
  );

  useEffect(() => {
    const viewport = straightTableViewportRef.current;
    if (!viewport || !straightTableHeaderContentRef.current) return;

    const element =
      ((straightTableGridRef.current as any)?.element ??
        (straightTableGridRef.current as any)?._outerRef ??
        viewport.querySelector('[role="grid"], [data-testid="virtuoso-scroller"], div[style*="overflow"]')) as HTMLDivElement | null;

    if (!element) return;

    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (straightTableHeaderContentRef.current) {
          straightTableHeaderContentRef.current.style.transform =
            `translateX(${-element.scrollLeft}px)`;
        }
      });
    });
  }, [loadedRows.length, gridWidth, gridHeight, straightTableColumns.length]);

  function renderStraightTable() {
    
    return (
      <div
        ref={straightTableViewportRef}
        style={{
          width: '100%',
          height: '100%',
          overflow: 'hidden',
          minWidth: 0,
          minHeight: 0
        }}
      >
        <div
          style={{
            width: '100%',
            height: '100%',
            overflow: 'hidden',
            //border: '1px solid #dce4ec',
            display: 'flex',
            flexDirection: 'column',
            minHeight: 0,
            minWidth: 0,
            background: '#ffffff',
            marginRight: '10px'
          }}
        >
          <div
            ref={straightTableHeaderRef}
              style={{
              overflowX: 'hidden',
              overflowY: 'hidden',
              borderBottom: '1px solid #dce4ec',
              background: '#075b67',
              flex: '0 0 auto',
              scrollbarWidth: 'none',
              marginRight: '20px'
            }}
          >
            <div
              ref={straightTableHeaderContentRef}
              style={{
                width: straightTableTotalWidth,
                minWidth: straightTableTotalWidth,
                display: 'flex',
                minHeight: STRAIGHT_HEADER_HEIGHT,
                color: '#fff',
                marginRight: '10px'
              }}
            >
              {straightTableColumns.map((column, columnIndex) => {
                const field = column.key;
                const resolvedWidth = getStraightTableColumnWidth(columnIndex);

                if (column.kind === 'dimension') {
                  return (
                    <div
                      key={field}
                      style={{
                        width: resolvedWidth,
                        minWidth: resolvedWidth,
                        maxWidth: resolvedWidth,
                        padding: '8px 10px',
                        borderRight: '1px solid #dce4ec',
                        fontSize: '12px',
                        boxSizing: 'border-box'
                      }}
                    >
                      <div
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          gap: '8px'
                        }}
                      >
                        <button
                          type="button"
                          onClick={() => toggleSort(field)}
                          style={{
                            flex: 1,
                            minWidth: 0,
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            background: 'transparent',
                            border: 'none',
                            color: '#fff',
                            fontSize: '12px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            padding: 0
                          }}
                        >
                          <span
                            style={{
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              whiteSpace: 'nowrap'
                            }}
                          >
                            {column.label}
                          </span>
                          <span>
                            {sortConfig?.field === field
                              ? sortConfig.direction === 'asc'
                                ? '▲'
                                : '▼'
                              : '↕'}
                          </span>
                        </button>

                        <button
                          type="button"
                          onClick={(e) => {
                            const rect = (
                              e.currentTarget as HTMLButtonElement
                            ).getBoundingClientRect();
                            setColumnFilterSearch('');
                            setHeaderFilterMenu({ field, anchorRect: rect });
                          }}
                          style={{
                            width: '20px',
                            height: '20px',
                            border: 'none',
                            background: 'transparent',
                            cursor: 'pointer',
                            padding: 0,
                            color: '#ffffff',
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            flexShrink: 0
                          }}
                          aria-label={`Filter ${column.label}`}
                          title={`Filter ${column.label}`}
                        >
                          <svg
                            xmlns="http://www.w3.org/2000/svg"
                            width="14"
                            height="14"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            aria-hidden="true"
                          >
                            <circle cx="11" cy="11" r="7" />
                            <path d="M20 20l-3.5-3.5" />
                          </svg>
                        </button>
                      </div>
                    </div>
                  );
                }

                return (
                  <div
                    key={field}
                    style={{
                      width: resolvedWidth,
                      minWidth: resolvedWidth,
                      maxWidth: resolvedWidth,
                      padding: '8px 10px',
                      borderRight: '1px solid #dce4ec',
                      fontSize: '12px',
                      boxSizing: 'border-box'
                    }}
                  >
                    <button
                      type="button"
                      onClick={() => toggleSort(field)}
                      style={{
                        width: '100%',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        background: 'transparent',
                        border: 'none',
                        color: '#fff',
                        fontSize: '12px',
                        fontWeight: 700,
                        cursor: 'pointer',
                        padding: 0
                      }}
                    >
                      <span
                        style={{
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap'
                        }}
                      >
                        {column.label}
                      </span>
                      <span>
                        {sortConfig?.field === field
                          ? sortConfig.direction === 'asc'
                            ? '▲'
                            : '▼'
                          : '↕'}
                      </span>
                    </button>
                  </div>
                );
              })}
            </div>
          </div>

          <div
            style={{
              flex: '1 1 auto',
              minHeight: 0,
              minWidth: 0
            }}
          >
            {gridWidth > 0 && gridHeight > 0 && straightTableColumns.length > 0 ? (
              <Grid
                gridRef={straightTableGridRef}
                cellComponent={StraightGridCell}
                cellProps={straightGridCellProps as any}
                columnCount={straightTableColumns.length}
                columnWidth={getStraightTableColumnWidth}
                rowCount={sortedStraightRows.length}
                rowHeight={() => STRAIGHT_ROW_HEIGHT}
                defaultHeight={gridHeight}
                defaultWidth={gridWidth}
                overscanCount={8}
                onScroll={handleStraightTableGridScroll}
                style={{
                  width: gridWidth,
                  height: gridHeight
                }}
              />
            ) : null}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        width: '100%',
        height: '100vh',
        overflow: 'hidden',
        background: '#ffffff',
        fontFamily: 'Arial, sans-serif',
        color: '#111827',
        boxSizing: 'border-box',
        cursor: isResizing ? 'col-resize' : 'default',
        display: 'flex',
        flexDirection: 'column'
      }}
    >
      {editingFieldLabel && (
        <div
          onClick={cancelFieldLabelEdit}
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(15, 23, 42, 0.36)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '24px',
            zIndex: 300
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-labelledby="edit-header-title"
            style={{
              width: 'min(420px, 100%)',
              border: '1px solid #d8e2ee',
              borderRadius: '16px',
              background: '#ffffff',
              boxShadow: '0 24px 60px rgba(15, 23, 42, 0.20)',
              padding: '16px'
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '12px',
                marginBottom: '14px'
              }}
            >
              <div id="edit-header-title" style={{ fontSize: '14px', fontWeight: 700, color: '#0f172a' }}>
                Edit Header Title
              </div>

              <button
                type="button"
                onClick={cancelFieldLabelEdit}
                style={{
                  width: '28px',
                  height: '28px',
                  border: '1px solid #d8e2ee',
                  borderRadius: '8px',
                  background: '#f8fafc',
                  color: '#475569',
                  fontSize: '14px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}
                aria-label="Close edit dialog"
              >
                ×
              </button>
            </div>

            <div style={{ fontSize: '12px', color: '#64748b', marginBottom: '8px' }}>
              Editing: <span style={{ fontWeight: 700, color: '#0f172a' }}>{getDisplayLabel(editingFieldLabel)}</span>
            </div>

            <input
              type="text"
              value={editingFieldValue}
              onChange={(e) => setEditingFieldValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  saveFieldLabel(editingFieldLabel);
                }
                if (e.key === 'Escape') {
                  cancelFieldLabelEdit();
                }
              }}
              style={{
                width: '100%',
                height: '36px',
                border: '1px solid #d7dfea',
                borderRadius: '8px',
                padding: '0 10px',
                fontSize: '12px',
                marginBottom: '12px',
                outline: 'none',
                boxSizing: 'border-box'
              }}
              autoFocus
            />

            {isMeasureField(editingFieldLabel) && (
              <div style={{ marginBottom: '12px' }}>
                <div
                  style={{
                    fontSize: '12px',
                    color: '#64748b',
                    marginBottom: '6px',
                    fontWeight: 600
                  }}
                >
                  Metric Format
                </div>

                <select
                  value={editingFieldFormat}
                  onChange={(e) => setEditingFieldFormat(e.target.value)}
                  style={{
                    width: '100%',
                    height: '36px',
                    border: '1px solid #d7dfea',
                    borderRadius: '8px',
                    padding: '0 10px',
                    fontSize: '12px',
                    background: '#fff',
                    boxSizing: 'border-box'
                  }}
                >
                  <option value="number">Number</option>
                  <option value="currency">Currency</option>
                  <option value="percent">Percent</option>
                  <option value="integer">Whole Number</option>
                  <option value="decimal-1">1 Decimal</option>
                  <option value="decimal-2">2 Decimals</option>
                  <option value="decimal-3">3 Decimals</option>
                </select>
              </div>
            )}            

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button
                type="button"
                onClick={cancelFieldLabelEdit}
                style={{
                  height: '32px',
                  padding: '0 12px',
                  border: '1px solid #d8e2ee',
                  borderRadius: '8px',
                  background: '#fff',
                  color: '#475569',
                  fontSize: '12px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
              >
                Cancel
              </button>

              <button
                type="button"
                onClick={() => saveFieldLabel(editingFieldLabel)}
                style={{
                  height: '32px',
                  padding: '0 12px',
                  border: 'none',
                  borderRadius: '8px',
                  background: '#075b67',
                  color: '#fff',
                  fontSize: '12px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}

      {headerFilterMenu && (
        <div
          onClick={() => setHeaderFilterMenu(null)}
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 250
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              position: 'fixed',
              top: (headerFilterMenu.anchorRect?.bottom || 80) + 6,
              left: Math.max(12, (headerFilterMenu.anchorRect?.left || 12) - 180),
              width: '240px',
              maxHeight: '320px',
              overflow: 'hidden',
              border: '1px solid #d8e2ee',
              borderRadius: '12px',
              background: '#ffffff',
              boxShadow: '0 18px 40px rgba(15, 23, 42, 0.16)',
              display: 'flex',
              flexDirection: 'column'
            }}
          >
            <div style={{ padding: '10px', borderBottom: '1px solid #eef2f7' }}>
              <div style={{ fontSize: '12px', fontWeight: 700, color: '#0f172a', marginBottom: '8px' }}>
                Filter {getDisplayLabel(headerFilterMenu.field)}
              </div>

              <input
                type="text"
                value={columnFilterSearch}
                onChange={(e) => setColumnFilterSearch(e.target.value)}
                placeholder="Search values..."
                style={{
                  width: '100%',
                  height: '32px',
                  border: '1px solid #d7dfea',
                  borderRadius: '8px',
                  padding: '0 10px',
                  fontSize: '12px',
                  boxSizing: 'border-box'
                }}
              />
            </div>

            <div style={{ padding: '8px', overflowY: 'auto', maxHeight: '220px' }}>
              {getHeaderFilterOptions(headerFilterMenu.field)
                .filter((value) =>
                  value.toLowerCase().includes(columnFilterSearch.toLowerCase())
                )
                .map((value) => {
                  const checked = isHeaderFilterValueSelected(headerFilterMenu.field, value);

                  return (
                    <label
                      key={value}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        padding: '6px 4px',
                        fontSize: '12px',
                        color: '#334155',
                        cursor: 'pointer'
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleHeaderFilterValue(headerFilterMenu.field, value)}
                      />
                      <span>{value}</span>
                    </label>
                  );
                })}
            </div>

            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                gap: '8px',
                padding: '10px',
                borderTop: '1px solid #eef2f7'
              }}
            >
              <button
                type="button"
                onClick={() => clearHeaderFilter(headerFilterMenu.field)}
                style={{
                  height: '30px',
                  padding: '0 10px',
                  border: '1px solid #d8e2ee',
                  borderRadius: '8px',
                  background: '#fff',
                  fontSize: '12px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
              >
                Clear
              </button>

              <button
                type="button"
                onClick={() => setHeaderFilterMenu(null)}
                style={{
                  height: '30px',
                  padding: '0 10px',
                  border: 'none',
                  borderRadius: '8px',
                  background: '#075b67',
                  color: '#fff',
                  fontSize: '12px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      <div style={{ flex: '0 0 auto', padding: '12px 18px 0' }}>
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
            <div style={{ position: 'relative' }} ref={exportMenuRef}>
              {viewMode === 'straight' ? (
                <button
                  type="button"
                  onClick={handleExportStraightTable}
                  style={{
                    height: '38px',
                    padding: '0 14px',
                    border: 'none',
                    borderRadius: '8px',
                    background: 'linear-gradient(180deg, #24c5c0 0%, #16b3ad 100%)',
                    color: '#fff',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '8px'
                  }}
                >
                  <img
                    src={exportIcon}
                    alt=""
                    aria-hidden="true"
                    style={{
                      width: '14px',
                      height: '14px',
                      display: 'block',
                      objectFit: 'contain'
                    }}
                  />
                  <span>Export Report</span>
                </button>
              ) : (
                <>
                  <button
                    type="button"
                    onClick={() => setExportMenuOpen((prev) => !prev)}
                    style={{
                      height: '38px',
                      padding: '0 14px',
                      border: 'none',
                      borderRadius: '8px',
                      background: 'linear-gradient(180deg, #24c5c0 0%, #16b3ad 100%)',
                      color: '#fff',
                      fontSize: '12px',
                      fontWeight: 700,
                      cursor: 'pointer',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '8px'
                    }}
                  >
                    <img
                      src={exportIcon}
                      alt=""
                      aria-hidden="true"
                      style={{
                        width: '14px',
                        height: '14px',
                        display: 'block',
                        objectFit: 'contain'
                      }}
                    />
                    <span>Export Report</span>
                    <span aria-hidden="true" style={{ fontSize: '10px', lineHeight: 1 }}>
                      {exportMenuOpen ? '▲' : '▼'}
                    </span>
                  </button>

                  {exportMenuOpen && (
                    <div
                      style={{
                        position: 'absolute',
                        top: '44px',
                        right: 0,
                        minWidth: '220px',
                        background: '#fff',
                        border: '1px solid #d6dee7',
                        borderRadius: '12px',
                        boxShadow: '0 16px 36px rgba(15, 23, 42, 0.14)',
                        padding: '8px',
                        zIndex: 120
                      }}
                    >
                      <button
                        type="button"
                        onClick={async () => {
                          setExportMenuOpen(false);
                          await handleExportPivotExpanded();
                        }}
                        style={{
                          width: '100%',
                          height: '40px',
                          border: 'none',
                          borderRadius: '8px',
                          background: '#f7f9fc',
                          color: '#0f172a',
                          fontSize: '12px',
                          fontWeight: 700,
                          cursor: 'pointer',
                          textAlign: 'left',
                          padding: '0 12px'
                        }}
                      >
                        Export Fully Expanded Pivot
                      </button>

                      <button
                        type="button"
                        onClick={async () => {
                          setExportMenuOpen(false);
                          await handleExportPivotRawData();
                        }}
                        style={{
                          width: '100%',
                          height: '40px',
                          border: 'none',
                          borderRadius: '8px',
                          background: '#ffffff',
                          color: '#0f172a',
                          fontSize: '12px',
                          fontWeight: 700,
                          cursor: 'pointer',
                          textAlign: 'left',
                          padding: '0 12px',
                          marginTop: '6px'
                        }}
                      >
                        Export Raw Data
                      </button>
                    </div>
                  )}
                </>
              )}
            </div>

            <button
              type="button"
              onClick={() => setIsInfoOpen(true)}
              style={{
                width: '34px',
                height: '34px',
                border: '1px solid #d6dee7',
                borderRadius: '10px',
                background: '#f7f9fc',
                color: '#0f172a',
                fontSize: '16px',
                fontWeight: 700,
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}
              aria-label="Open information panel"
              title="Information"
            >
              i
            </button>
          </div>
        </div>

        {isInfoOpen && (
          <div
            onClick={() => setIsInfoOpen(false)}
            style={{
              position: 'fixed',
              inset: 0,
              background: 'rgba(15, 23, 42, 0.32)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              zIndex: 200
            }}
          >
            <div
              onClick={(e) => e.stopPropagation()}
              style={{
                width: 'min(720px, 100%)',
                maxHeight: '80vh',
                overflowY: 'auto',
                border: '1px solid #e3e8ef',
                borderRadius: '16px',
                background: '#ffffff',
                boxShadow: '0 24px 60px rgba(15,23,42,0.18)',
                padding: '16px'
              }}
            >
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '12px',
                  marginBottom: '14px'
                }}
              >
                <div style={{ fontSize: '14px', fontWeight: 700, color: '#0f172a' }}>
                  Information Panel
                </div>

                <button
                  type="button"
                  onClick={() => setIsInfoOpen(false)}
                  style={{
                    width: '28px',
                    height: '28px',
                    border: '1px solid #d6dee7',
                    borderRadius: '8px',
                    background: '#f8fafc',
                    color: '#475569',
                    fontSize: '14px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  aria-label="Close information panel"
                >
                  ×
                </button>
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
        <div style={{ flex: '0 0 auto', margin: '0 18px 10px', color: '#b00020', fontSize: '13px' }}>
          {message}
        </div>
      )}

      <div
        style={{
          flex: 1,
          minHeight: 0,
          display: 'grid',
          gridTemplateColumns: isSelectionsPanelOpen ? 'clamp(220px, 24vw, 280px) 1fr' : '32px 1fr',
          gap: '0',
          borderTop: '1px solid #e5e7eb'
        }}
      >
        <aside
          style={{
            borderRight: '1px solid #e5e7eb',
            background: '#ffffff',
            overflow: 'hidden',
            minHeight: 0
          }}
        >
          {isSelectionsPanelOpen ? (
            <div style={{ height: '100%', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  padding: '12px 8px 10px 10px',
                  borderBottom: '1px solid #eef2f7',
                  flex: '0 0 auto'
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

              <div style={{ padding: '10px', overflowY: 'auto', minHeight: 0 }}>
                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Data Source(s):
                  </div>
                  {selectedSources.length > 0 ? (
                    selectedSources.map((source) => (
                      <div key={source} style={{ marginBottom: '8px' }}>
                        <SelectionPill
                          text={source}
                          color="#065f73"
                          options={
                            selectedSources.length > 1
                              ? { onRemove: () => toggleSource(source) }
                              : undefined
                          }
                        />
                      </div>
                    ))
                  ) : (
                    <div style={{ fontSize: '12px', color: '#94a3b8' }}>None selected</div>
                  )}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Date Range:
                  </div>
                  {dateRange.start || dateRange.end ? (
                    <SelectionPill
                      text={`${dateRange.start || '...'} - ${dateRange.end || '...'}`}
                      color="#065f73"
                    />
                  ) : (
                    <div style={{ fontSize: '12px', color: '#94a3b8' }}>Not set</div>
                  )}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Dimensions:
                  </div>

                  {selectedDimensions.length > 0 ? (
                    <DndContext
                      sensors={sensors}
                      collisionDetection={closestCenter}
                      onDragEnd={handleDimensionDragEnd}
                    >
                      <SortableContext
                        items={selectedDimensions}
                        strategy={verticalListSortingStrategy}
                      >
                        <div>
                          {selectedDimensions.map((field, index) => (
                            <SortableSelectionPill
                              key={field}
                              id={field}
                              label={getDisplayLabel(field)}
                              index={index}
                              total={selectedDimensions.length}
                              onMoveUp={() => moveDimension(field, 'up')}
                              onMoveDown={() => moveDimension(field, 'down')}
                              onEdit={() => startEditingFieldLabel(field)}
                              onRemove={() => toggleDimension(field)}
                            />
                          ))}
                        </div>
                      </SortableContext>
                    </DndContext>
                  ) : (
                    <div style={{ fontSize: 12, color: '#94a3b8' }}>No dimensions selected</div>
                  )}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Selected Measures:
                  </div>
                  {selectedMeasures.length > 0 ? (
                    <DndContext
                      sensors={sensors}
                      collisionDetection={closestCenter}
                      onDragEnd={handleMeasureDragEnd}
                    >
                      <SortableContext
                        items={selectedMeasures}
                        strategy={verticalListSortingStrategy}
                      >
                        <div>
                          {selectedMeasures.map((field, index) => (
                            <SortableSelectionPill
                              key={field}
                              id={field}
                              label={getDisplayLabel(field)}
                              index={index}
                              total={selectedMeasures.length}
                              onMoveUp={() => moveMeasure(field, 'up')}
                              onMoveDown={() => moveMeasure(field, 'down')}
                              onEdit={() => startEditingFieldLabel(field)}
                              onRemove={() => toggleMeasure(field)}
                            />
                          ))}
                        </div>
                      </SortableContext>
                    </DndContext>
                  ) : (
                    <div style={{ fontSize: '12px', color: '#94a3b8' }}>None selected</div>
                  )}
                </div>

                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', fontWeight: 700, color: '#6b7280', marginBottom: '8px' }}>
                    Active Filters:
                  </div>
                  {filters.length > 0 ? (
                    filters.map((filter) => (
                      <div key={filter.field} style={{ marginBottom: '8px' }}>
                        <SelectionPill
                          text={`${getDisplayLabel(filter.field)}: ${filter.values.join(', ')}`}
                          color="#065f73"
                          options={{ onRemove: () => removeFilter(filter.field) }}
                        />
                      </div>
                    ))
                  ) : (
                    <div style={{ fontSize: '12px', color: '#94a3b8', marginBottom: '12px' }}>
                      None applied
                    </div>
                  )}
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

        <main
          style={{
            padding: '10px 12px 14px',
            minWidth: 0,
            minHeight: 0,
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden'
          }}
        >
          <div style={{ display: 'flex', gap: '8px', marginBottom: '10px', flex: '0 0 auto' }}>
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

          <div
            style={{
              fontSize: '12px',
              color: '#6b7280',
              marginBottom: '8px',
              padding: '6px 10px',
              border: '1px solid #e5e7eb',
              borderRadius: '8px',
              background: '#f8fafc'
            }}
          >
            {isLoadingData
              ? 'Loading data...'
              : hasActiveFilters
                ? `Showing ${visibleRowCount.toLocaleString()} matching rows`
                : `${loadedRowCount.toLocaleString()} rows loaded`}
          </div>

          <div
            style={{
              flex: 1,
              minHeight: 0,
              minWidth: 0,
              overflow: 'hidden'
            }}
          >
            {viewMode === 'pivot' ? renderPivotTable() : renderStraightTable()}
          </div>
        </main>
      </div>
    </div>
  );
}