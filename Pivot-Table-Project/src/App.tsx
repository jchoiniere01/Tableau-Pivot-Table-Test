import { useEffect, useMemo, useRef, useState } from 'react';

declare const tableau: any;

type PreviewRow = Record<string, any>;

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

type StyleSettings = {
  dimensionHeaderBg: string;
  dimensionHeaderText: string;
  dimensionValueBg: string;
  dimensionValueText: string;
  measureHeaderBg: string;
  measureHeaderText: string;
  measureValueBg: string;
  measureValueText: string;
  totalBg: string;
  totalText: string;
};

type MeasureFormat = {
  formatType: 'number' | 'currency' | 'percent';
  decimals: number;
  prefix: string;
  suffix: string;
  useThousandsSeparator: boolean;
  alignment: 'left' | 'center' | 'right';
};

type MeasureFormatMap = Record<string, MeasureFormat>;

type LayoutSettings = {
  hierarchyWidth: number;
  defaultMeasureWidth: number;
  measureWidths: Record<string, number>;
};

type DisplaySettings = {
  showTableTitle: boolean;
  tableTitle: string;
  hierarchyHeader: string;
};

const DEFAULT_STYLES: StyleSettings = {
  dimensionHeaderBg: '#f3f3f3',
  dimensionHeaderText: '#111111',
  dimensionValueBg: '#ffffff',
  dimensionValueText: '#111111',
  measureHeaderBg: '#f3f3f3',
  measureHeaderText: '#111111',
  measureValueBg: '#ffffff',
  measureValueText: '#111111',
  totalBg: '#f9f9f9',
  totalText: '#111111'
};

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

const DEFAULT_DISPLAY: DisplaySettings = {
  showTableTitle: true,
  tableTitle: 'Pivot Table',
  hierarchyHeader: 'Hierarchy'
};

const MIN_HIERARCHY_WIDTH = 160;
const MIN_MEASURE_WIDTH = 80;

function parseNumber(value: any): number {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;

  const cleaned = String(value).replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

function parseSavedArray(value: any): string[] {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function parseSavedObject<T>(value: any, fallback: T): T {
  if (!value) return fallback;
  try {
    const parsed = JSON.parse(value);
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
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

export default function App() {
  const [message, setMessage] = useState('Starting extension...');
  const [dimensionFields, setDimensionFields] = useState<string[]>([]);
  const [measureFields, setMeasureFields] = useState<string[]>([]);
  const [fieldLabels, setFieldLabels] = useState<Record<string, string>>({});
  const [styleSettings, setStyleSettings] = useState<StyleSettings>(DEFAULT_STYLES);
  const [measureFormats, setMeasureFormats] = useState<MeasureFormatMap>({});
  const [layoutSettings, setLayoutSettings] = useState<LayoutSettings>(DEFAULT_LAYOUT);
  const [displaySettings, setDisplaySettings] = useState<DisplaySettings>(DEFAULT_DISPLAY);
  const [showTotals, setShowTotals] = useState(false);
  const [treeRoots, setTreeRoots] = useState<TreeNode[]>([]);
  const [expandedNodes, setExpandedNodes] = useState<Record<string, boolean>>({});
  const [isResizing, setIsResizing] = useState(false);

  const unregisterHandlersRef = useRef<(() => void)[]>([]);
  const refreshTimerRef = useRef<number | null>(null);

  function clearEventListeners() {
    unregisterHandlersRef.current.forEach((unregister) => {
      try {
        unregister();
      } catch {
      }
    });
    unregisterHandlersRef.current = [];
  }

  function debouncedRefresh() {
    if (refreshTimerRef.current) {
      window.clearTimeout(refreshTimerRef.current);
    }

    refreshTimerRef.current = window.setTimeout(() => {
      loadWorksheetData(true);
    }, 250);
  }

  function getMeasureFormat(field: string): MeasureFormat {
    return {
      ...DEFAULT_MEASURE_FORMAT,
      ...(measureFormats[field] || {})
    };
  }

  function getSavedMeasureWidth(field: string): number {
    return layoutSettings.measureWidths[field] || layoutSettings.defaultMeasureWidth;
  }

  function formatMeasureValue(field: string, value: number): string {
    const fmt = getMeasureFormat(field);
    let numericValue = value;

    if (fmt.formatType === 'percent') {
      numericValue = value * 100;
    }

    const formatted = new Intl.NumberFormat(undefined, {
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: fmt.useThousandsSeparator
    }).format(numericValue);

    const percentSuffix = fmt.formatType === 'percent' ? '%' : '';
    const currencyPrefix = fmt.formatType === 'currency' && !fmt.prefix ? '$' : '';

    return `${fmt.prefix || currencyPrefix}${formatted}${percentSuffix}${fmt.suffix}`;
  }

  function getTotalAlignment(): 'left' | 'center' | 'right' {
    if (measureFields.length === 0) return 'right';
    return getMeasureFormat(measureFields[0]).alignment;
  }

  function getRowTotal(measures: Record<string, number>): number {
    return measureFields.reduce((sum, field) => sum + (measures[field] ?? 0), 0);
  }

  function toggleNode(nodeId: string) {
    setExpandedNodes((prev) => ({
      ...prev,
      [nodeId]: !prev[nodeId]
    }));
  }

  function getDisplayLabel(field: string): string {
    return fieldLabels[field]?.trim() || field;
  }

  function buildTree(allRows: PreviewRow[], dims: string[], measures: string[]): TreeNode[] {
    const roots: TreeNode[] = [];
    const nodeMap = new Map<string, TreeNode>();

    allRows.forEach((row) => {
      let parentChildren = roots;
      let pathValues: string[] = [];

      dims.forEach((dimensionField, level) => {
        const rawValue = row[dimensionField];
        const label = String(rawValue ?? '');
        pathValues = [...pathValues, label];
        const nodeId = pathValues.join(' || ');

        let node = nodeMap.get(nodeId);

        if (!node) {
          node = createNode(
            nodeId,
            label,
            level,
            dimensionField,
            [...pathValues],
            measures
          );

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

        if (expanded[node.id]) {
          visit(node.children);
        }
      });
    }

    visit(nodes);
    return result;
  }

  function getLevelStyles(level: number, hasChildren: boolean) {
    if (level === 0) {
      return {
        hierarchyWeight: 700,
        measureWeight: 700,
        textColor: styleSettings.dimensionValueText
      };
    }

    if (level === 1) {
      return {
        hierarchyWeight: hasChildren ? 600 : 500,
        measureWeight: hasChildren ? 600 : 500,
        textColor: styleSettings.dimensionValueText
      };
    }

    if (level === 2) {
      return {
        hierarchyWeight: hasChildren ? 500 : 400,
        measureWeight: hasChildren ? 500 : 400,
        textColor: styleSettings.dimensionValueText
      };
    }

    return {
      hierarchyWeight: hasChildren ? 500 : 400,
      measureWeight: hasChildren ? 500 : 400,
      textColor: styleSettings.dimensionValueText
    };
  }

  async function saveLayoutSettings(nextLayout: LayoutSettings) {
    try {
      tableau.extensions.settings.set('layoutSettings', JSON.stringify(nextLayout));
      await tableau.extensions.settings.saveAsync();
    } catch {
    }
  }

  async function configure() {
    const popupUrl = `${window.location.origin}/dialog.html`;

    try {
      const result = await tableau.extensions.ui.displayDialogAsync(popupUrl, '', {
        height: 700,
        width: 900
      });

      // Re-read settings directly after dialog closes
      const latestDisplay = parseSavedObject<DisplaySettings>(
        tableau.extensions.settings.get('displaySettings'),
        DEFAULT_DISPLAY
      );

      setDisplaySettings({
        showTableTitle:
          latestDisplay.showTableTitle ?? DEFAULT_DISPLAY.showTableTitle,
        tableTitle: latestDisplay.tableTitle ?? DEFAULT_DISPLAY.tableTitle,
        hierarchyHeader:
          latestDisplay.hierarchyHeader ?? DEFAULT_DISPLAY.hierarchyHeader
      });

      await loadWorksheetData(false);
    } catch (error: any) {
      if (error?.errorCode === tableau.ErrorCodes.DialogClosedByUser) {
        return;
      }

      console.error('Configure dialog error:', error);
      setMessage(`Configuration error: ${error?.message || error}`);
    }
  }

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

    async function onMouseUp(moveEvent: MouseEvent) {
      const deltaX = moveEvent.clientX - startX;
      const proposedWidth = startWidth + deltaX;

      let nextLayout: LayoutSettings;

      if (type === 'hierarchy') {
        nextLayout = {
          ...layoutSettings,
          hierarchyWidth: Math.max(MIN_HIERARCHY_WIDTH, proposedWidth)
        };
      } else {
        nextLayout = {
          ...layoutSettings,
          measureWidths: {
            ...layoutSettings.measureWidths,
            [field || '']: Math.max(MIN_MEASURE_WIDTH, proposedWidth)
          }
        };
      }

      setLayoutSettings(nextLayout);
      await saveLayoutSettings(nextLayout);

      setIsResizing(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';

      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    }

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  async function registerWorksheetListeners(worksheet: any) {
    clearEventListeners();

    const filterHandler = worksheet.addEventListener(
      tableau.TableauEventType.FilterChanged,
      () => {
        debouncedRefresh();
      }
    );

    unregisterHandlersRef.current.push(filterHandler);

    try {
      const summaryHandler = worksheet.addEventListener(
        tableau.TableauEventType.SummaryDataChanged,
        () => {
          debouncedRefresh();
        }
      );

      unregisterHandlersRef.current.push(summaryHandler);
    } catch {
    }
  }

  async function loadWorksheetData(fromAutoRefresh = false) {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const savedWorksheet = tableau.extensions.settings.get('sourceWorksheet');
    const savedDimensions = parseSavedArray(
      tableau.extensions.settings.get('dimensionFields')
    );
    const savedMeasures = parseSavedArray(
      tableau.extensions.settings.get('measureFields')
    );
    const savedLabels = parseSavedObject<Record<string, string>>(
      tableau.extensions.settings.get('fieldLabels'),
      {}
    );
    const savedStyles = parseSavedObject<StyleSettings>(
      tableau.extensions.settings.get('styleSettings'),
      DEFAULT_STYLES
    );
    const savedMeasureFormats = parseSavedObject<MeasureFormatMap>(
      tableau.extensions.settings.get('measureFormats'),
      {}
    );
    const savedLayout = parseSavedObject<LayoutSettings>(
      tableau.extensions.settings.get('layoutSettings'),
      DEFAULT_LAYOUT
    );
    const savedDisplay = parseSavedObject<DisplaySettings>(
      tableau.extensions.settings.get('displaySettings'),
      DEFAULT_DISPLAY
    );
    const savedShowTotals =
      tableau.extensions.settings.get('showTotals') === 'true';

    setDimensionFields(savedDimensions);
    setMeasureFields(savedMeasures);
    setFieldLabels(savedLabels);
    setStyleSettings({ ...DEFAULT_STYLES, ...savedStyles });
    setMeasureFormats(savedMeasureFormats);
    setLayoutSettings({
      hierarchyWidth: savedLayout.hierarchyWidth || DEFAULT_LAYOUT.hierarchyWidth,
      defaultMeasureWidth:
        savedLayout.defaultMeasureWidth || DEFAULT_LAYOUT.defaultMeasureWidth,
      measureWidths: savedLayout.measureWidths || {}
    });
    setDisplaySettings({
      showTableTitle:
        savedDisplay.showTableTitle ?? DEFAULT_DISPLAY.showTableTitle,
      tableTitle: savedDisplay.tableTitle ?? DEFAULT_DISPLAY.tableTitle,
      hierarchyHeader:
        savedDisplay.hierarchyHeader ?? DEFAULT_DISPLAY.hierarchyHeader
    });
    setShowTotals(savedShowTotals);

    if (!savedWorksheet) {
      setMessage('Not configured yet. Use the extension menu and click Configure.');
      setTreeRoots([]);
      clearEventListeners();
      return;
    }

    if (savedDimensions.length === 0 || savedMeasures.length === 0) {
      setMessage('Configuration is incomplete. Use the extension menu and click Configure.');
      setTreeRoots([]);
      clearEventListeners();
      return;
    }

    const worksheet = dashboard.worksheets.find(
      (w: any) => w.name === savedWorksheet
    );

    if (!worksheet) {
      setMessage(`Configured worksheet "${savedWorksheet}" was not found on this dashboard.`);
      setTreeRoots([]);
      clearEventListeners();
      return;
    }

    if (!fromAutoRefresh) {
      setMessage('Loading pivot table...');
    }

    const reader = await worksheet.getSummaryDataReaderAsync();

    try {
      const dataTable = await reader.getAllPagesAsync();

      const allRows = dataTable.data.map((row: any[]) => {
        const obj: PreviewRow = {};
        dataTable.columns.forEach((col: any, index: number) => {
          obj[col.fieldName] = row[index]?.value ?? row[index]?.formattedValue ?? '';
        });
        return obj;
      });

      const roots = buildTree(allRows, savedDimensions, savedMeasures);

      setTreeRoots(roots);

      setExpandedNodes((prev) => {
        const next = { ...prev };
        roots.forEach((node) => {
          if (next[node.id] === undefined) {
            next[node.id] = true;
          }
        });
        return next;
      });

      await registerWorksheetListeners(worksheet);

      if (!fromAutoRefresh) {
        setMessage('');
      }
    } finally {
      await reader.releaseAsync();
    }
  }

  useEffect(() => {
    async function init() {
      try {
        if (typeof tableau === 'undefined') {
          setMessage('Tableau API not found.');
          return;
        }

        await tableau.extensions.initializeAsync({ configure });
        await loadWorksheetData(false);
      } catch (error: any) {
        setMessage(`Error: ${error?.message || error}`);
      }
    }

    init();

    return () => {
      clearEventListeners();
      if (refreshTimerRef.current) {
        window.clearTimeout(refreshTimerRef.current);
      }
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
  }, []);

  const visibleNodes = useMemo(
    () => flattenVisibleNodes(treeRoots, expandedNodes),
    [treeRoots, expandedNodes]
  );

  const hierarchyWidth = Math.max(MIN_HIERARCHY_WIDTH, layoutSettings.hierarchyWidth);

  const measureColumnWidths = useMemo(() => {
    const result: Record<string, number> = {};
    measureFields.forEach((field) => {
      result[field] = Math.max(MIN_MEASURE_WIDTH, getSavedMeasureWidth(field));
    });
    return result;
  }, [measureFields, layoutSettings]);

  const totalColumnWidth =
    showTotals && measureFields[0]
      ? Math.max(MIN_MEASURE_WIDTH, getSavedMeasureWidth(measureFields[0]))
      : 0;

  if (message && visibleNodes.length === 0) {
    return (
      <div
        style={{
          width: '100%',
          padding: '20px',
          boxSizing: 'border-box',
          fontFamily: 'Arial, sans-serif'
        }}
      >
        {displaySettings.showTableTitle && (
          <h1 style={{ marginTop: 0, marginBottom: '12px' }}>
            {displaySettings.tableTitle}
          </h1>
        )}
        <div
          style={{
            border: '1px solid #ddd',
            borderRadius: '8px',
            padding: '16px',
            background: '#fafafa',
            color: '#444'
          }}
        >
          {message}
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        width: '100%',
        padding: '20px',
        boxSizing: 'border-box',
        fontFamily: 'Arial, sans-serif',
        cursor: isResizing ? 'col-resize' : 'default',
        overflow: 'hidden'
      }}
    >
      {displaySettings.showTableTitle && (
        <h1 style={{ marginTop: 0, marginBottom: '16px' }}>
          {displaySettings.tableTitle}
        </h1>
      )}

      <div
        style={{
          width: '100%',
          overflowX: 'auto',
          overflowY: 'hidden'
        }}
      >
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
                  border: '1px solid #ccc',
                  padding: '8px',
                  background: styleSettings.dimensionHeaderBg,
                  color: styleSettings.dimensionHeaderText,
                  textAlign: 'left',
                  width: `${hierarchyWidth}px`,
                  boxSizing: 'border-box'
                }}
              >
                <div
                  style={{
                    paddingRight: '12px',
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis'
                  }}
                >
                  {displaySettings.hierarchyHeader}
                </div>
                <div
                  onMouseDown={(e) => startResize(e, 'hierarchy')}
                  title="Drag to resize"
                  style={{
                    position: 'absolute',
                    top: 0,
                    right: 0,
                    width: '8px',
                    height: '100%',
                    cursor: 'col-resize',
                    zIndex: 2,
                    background: 'transparent'
                  }}
                />
              </th>

              {measureFields.map((field) => {
                const width = measureColumnWidths[field];

                return (
                  <th
                    key={field}
                    style={{
                      position: 'relative',
                      border: '1px solid #ccc',
                      padding: '8px',
                      background: styleSettings.measureHeaderBg,
                      color: styleSettings.measureHeaderText,
                      textAlign: getMeasureFormat(field).alignment,
                      width: `${width}px`,
                      boxSizing: 'border-box'
                    }}
                  >
                    <div
                      style={{
                        paddingRight: '12px',
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}
                    >
                      {getDisplayLabel(field)}
                    </div>
                    <div
                      onMouseDown={(e) => startResize(e, 'measure', field)}
                      title="Drag to resize"
                      style={{
                        position: 'absolute',
                        top: 0,
                        right: 0,
                        width: '8px',
                        height: '100%',
                        cursor: 'col-resize',
                        zIndex: 2,
                        background: 'transparent'
                      }}
                    />
                  </th>
                );
              })}

              {showTotals && (
                <th
                  style={{
                    border: '1px solid #ccc',
                    padding: '8px',
                    background: styleSettings.totalBg,
                    color: styleSettings.totalText,
                    textAlign: getTotalAlignment(),
                    width: `${totalColumnWidth}px`,
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    boxSizing: 'border-box'
                  }}
                >
                  Total
                </th>
              )}
            </tr>
          </thead>
          <tbody>
            {visibleNodes.map((node) => {
              const levelStyles = getLevelStyles(node.level, node.hasChildren);

              return (
                <tr key={node.id}>
                  <td
                    style={{
                      border: '1px solid #ccc',
                      padding: '8px',
                      color: levelStyles.textColor,
                      background: styleSettings.dimensionValueBg,
                      width: `${hierarchyWidth}px`,
                      overflow: 'hidden',
                      boxSizing: 'border-box'
                    }}
                  >
                    <div
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        paddingLeft: `${node.level * 24}px`,
                        overflow: 'hidden',
                        minWidth: 0
                      }}
                    >
                      {node.hasChildren ? (
                        <button
                          onClick={() => toggleNode(node.id)}
                          style={{
                            width: '16px',
                            height: '16px',
                            minWidth: '16px',
                            marginRight: '8px',
                            border: '1px solid #999',
                            background: '#fff',
                            cursor: 'pointer',
                            fontWeight: 700,
                            fontSize: '12px',
                            lineHeight: '12px',
                            padding: 0,
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            color: '#333',
                            borderRadius: '2px'
                          }}
                        >
                          {expandedNodes[node.id] ? '-' : '+'}
                        </button>
                      ) : (
                        <span style={{ display: 'inline-block', width: '32px', minWidth: '32px' }} />
                      )}

                      <div style={{ overflow: 'hidden', minWidth: 0 }}>
                        <div
                          style={{
                            fontWeight: levelStyles.hierarchyWeight,
                            color: styleSettings.dimensionValueText,
                            whiteSpace: 'nowrap',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis'
                          }}
                        >
                          {node.label}
                        </div>
                      </div>
                    </div>
                  </td>

                  {measureFields.map((field) => {
                    const width = measureColumnWidths[field];

                    return (
                      <td
                        key={field}
                        style={{
                          border: '1px solid #ccc',
                          padding: '8px',
                          textAlign: getMeasureFormat(field).alignment,
                          fontWeight: levelStyles.measureWeight,
                          color: styleSettings.measureValueText,
                          background: styleSettings.measureValueBg,
                          width: `${width}px`,
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          boxSizing: 'border-box'
                        }}
                      >
                        {formatMeasureValue(field, node.measures[field] ?? 0)}
                      </td>
                    );
                  })}

                  {showTotals && (
                    <td
                      style={{
                        border: '1px solid #ccc',
                        padding: '8px',
                        textAlign: getTotalAlignment(),
                        fontWeight:
                          levelStyles.measureWeight + 100 > 700
                            ? 700
                            : levelStyles.measureWeight + 100,
                        color: styleSettings.totalText,
                        background: styleSettings.totalBg,
                        width: `${totalColumnWidth}px`,
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        boxSizing: 'border-box'
                      }}
                    >
                      {formatMeasureValue(
                        measureFields[0] || 'total',
                        getRowTotal(node.measures)
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}