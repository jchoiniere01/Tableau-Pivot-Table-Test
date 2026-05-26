import React, { useEffect, useMemo, useRef, useState } from 'react';
import ReactDOM from 'react-dom/client';

declare const tableau: any;

type FieldInfo = {
  fieldName: string;
  dataType: string;
  suggestedRole: 'dimension' | 'measure';
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
type PreviewRow = Record<string, any>;

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
  measureValueText: '#111111'
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

function DialogApp() {
  const [status, setStatus] = useState('Starting dialog...');
  const [worksheets, setWorksheets] = useState<string[]>([]);
  const [fields, setFields] = useState<FieldInfo[]>([]);
  const [worksheetName, setWorksheetName] = useState('');
  const [dimensionFields, setDimensionFields] = useState<string[]>([]);
  const [measureFields, setMeasureFields] = useState<string[]>([]);
  const [dimensionMenuOpen, setDimensionMenuOpen] = useState(false);
  const [measureMenuOpen, setMeasureMenuOpen] = useState(false);
  const [showAllDimensionFields, setShowAllDimensionFields] = useState(false);
  const [showAllMeasureFields, setShowAllMeasureFields] = useState(false);
  const [fieldLabels, setFieldLabels] = useState<Record<string, string>>({});
  const [styleSettings, setStyleSettings] = useState<StyleSettings>(DEFAULT_STYLES);
  const [measureFormats, setMeasureFormats] = useState<MeasureFormatMap>({});
  const [columnsFound, setColumnsFound] = useState<string[]>([]);
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([]);
  const [layoutSettings, setLayoutSettings] = useState<LayoutSettings>(DEFAULT_LAYOUT);
  const [displaySettings, setDisplaySettings] = useState<DisplaySettings>(DEFAULT_DISPLAY);

  const dimensionMenuRef = useRef<HTMLDivElement | null>(null);
  const measureMenuRef = useRef<HTMLDivElement | null>(null);

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

  function classifyField(fieldName: string, dataType: string): 'dimension' | 'measure' {
    const name = fieldName.toLowerCase();

    const dimensionHints = [
      'id',
      'key',
      'code',
      'zip',
      'postal',
      'phone',
      'number',
      'year',
      'month',
      'day',
      'region',
      'state',
      'city',
      'country',
      'category',
      'segment',
      'name'
    ];

    const numericTypes = ['float', 'double', 'integer', 'int', 'real', 'numeric'];
    const textLikeTypes = ['string', 'boolean', 'date', 'datetime'];

    if (dimensionHints.some((hint) => name.includes(hint))) {
      return 'dimension';
    }

    if (numericTypes.includes((dataType || '').toLowerCase())) {
      return 'measure';
    }

    if (textLikeTypes.includes((dataType || '').toLowerCase())) {
      return 'dimension';
    }

    return 'dimension';
  }

  async function loadFieldsForWorksheet(sheetName: string) {
    try {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const worksheet = dashboard.worksheets.find((w: any) => w.name === sheetName);

      if (!worksheet) {
        setFields([]);
        setColumnsFound([]);
        setPreviewRows([]);
        setStatus(`Worksheet "${sheetName}" not found.`);
        return;
      }

      setStatus(`Loading fields from ${sheetName}...`);

      const reader = await worksheet.getSummaryDataReaderAsync();

      try {
        const dataTable = await reader.getAllPagesAsync();

        const fieldInfos: FieldInfo[] = dataTable.columns.map((c: any) => {
          const fieldName = c.fieldName;
          const dataType = String(c.dataType || '').toLowerCase();
          const suggestedRole = classifyField(fieldName, dataType);

          return {
            fieldName,
            dataType,
            suggestedRole
          };
        });

        const colNames = dataTable.columns.map((c: any) => c.fieldName);
        const preview = dataTable.data.slice(0, 5).map((row: any[]) => {
          const obj: PreviewRow = {};
          dataTable.columns.forEach((col: any, index: number) => {
            obj[col.fieldName] = row[index]?.formattedValue ?? row[index]?.value ?? '';
          });
          return obj;
        });

        setFields(fieldInfos);
        setColumnsFound(colNames);
        setPreviewRows(preview);

        const dimensionCount = fieldInfos.filter(
          (f) => f.suggestedRole === 'dimension'
        ).length;
        const measureCount = fieldInfos.filter(
          (f) => f.suggestedRole === 'measure'
        ).length;

        setStatus(
          `Loaded ${fieldInfos.length} fields from ${sheetName}. Suggested dimensions: ${dimensionCount}. Suggested measures: ${measureCount}.`
        );
      } finally {
        await reader.releaseAsync();
      }
    } catch (error: any) {
      setStatus(`Error loading fields: ${error?.message || error}`);
      setFields([]);
      setColumnsFound([]);
      setPreviewRows([]);
    }
  }

  useEffect(() => {
    async function initDialog() {
      try {
        await tableau.extensions.initializeDialogAsync();

        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const sheetNames = dashboard.worksheets.map((w: any) => w.name);
        setWorksheets(sheetNames);

        const savedWorksheet =
          tableau.extensions.settings.get('sourceWorksheet') || '';
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
        const savedShowAllDimensions =
          tableau.extensions.settings.get('showAllDimensionFields') === 'true';
        const savedShowAllMeasures =
          tableau.extensions.settings.get('showAllMeasureFields') === 'true';

        const firstWorksheet = savedWorksheet || sheetNames[0] || '';

        setWorksheetName(firstWorksheet);
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
        setShowAllDimensionFields(savedShowAllDimensions);
        setShowAllMeasureFields(savedShowAllMeasures);

        if (firstWorksheet) {
          await loadFieldsForWorksheet(firstWorksheet);
        } else {
          setStatus('No worksheets found on this dashboard.');
        }
      } catch (error: any) {
        setStatus(`Dialog error: ${error?.message || error}`);
      }
    }

    initDialog();
  }, []);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      const target = event.target as Node;

      if (dimensionMenuRef.current && !dimensionMenuRef.current.contains(target)) {
        setDimensionMenuOpen(false);
      }

      if (measureMenuRef.current && !measureMenuRef.current.contains(target)) {
        setMeasureMenuOpen(false);
      }
    }

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  async function onWorksheetChange(newWorksheet: string) {
    setWorksheetName(newWorksheet);
    setDimensionFields([]);
    setMeasureFields([]);
    setFieldLabels({});
    setMeasureFormats({});
    setDimensionMenuOpen(false);
    setMeasureMenuOpen(false);
    await loadFieldsForWorksheet(newWorksheet);
  }

  function ensureMeasureFormat(field: string) {
    setMeasureFormats((prev) => ({
      ...prev,
      [field]: {
        ...DEFAULT_MEASURE_FORMAT,
        ...(prev[field] || {})
      }
    }));
  }

  function getMeasureWidth(field: string) {
    return layoutSettings.measureWidths[field] || layoutSettings.defaultMeasureWidth;
  }

  function updateHierarchyWidth(value: number) {
    setLayoutSettings((prev) => ({
      ...prev,
      hierarchyWidth: Math.max(120, value || DEFAULT_LAYOUT.hierarchyWidth)
    }));
  }

  function updateDefaultMeasureWidth(value: number) {
    setLayoutSettings((prev) => ({
      ...prev,
      defaultMeasureWidth: Math.max(80, value || DEFAULT_LAYOUT.defaultMeasureWidth)
    }));
  }

  function updateMeasureWidth(field: string, value: number) {
    setLayoutSettings((prev) => ({
      ...prev,
      measureWidths: {
        ...prev.measureWidths,
        [field]: Math.max(80, value || prev.defaultMeasureWidth || DEFAULT_LAYOUT.defaultMeasureWidth)
      }
    }));
  }

  function updateDisplaySetting<K extends keyof DisplaySettings>(
    key: K,
    value: DisplaySettings[K]
  ) {
    setDisplaySettings((prev) => ({
      ...prev,
      [key]: value
    }));
  }

  function resetWidthsToDefault() {
    setLayoutSettings(DEFAULT_LAYOUT);
  }

  function toggleDimension(field: string) {
    setDimensionFields((prev) => {
      if (prev.includes(field)) {
        return prev.filter((f) => f !== field);
      }
      return [...prev, field];
    });

    setMeasureFields((prev) => prev.filter((f) => f !== field));
  }

  function toggleMeasure(field: string) {
    setMeasureFields((prev) => {
      if (prev.includes(field)) {
        return prev.filter((f) => f !== field);
      }
      return [...prev, field];
    });

    ensureMeasureFormat(field);
    setDimensionFields((prev) => prev.filter((f) => f !== field));
  }

  function moveItemUp(list: string[], index: number, setter: (items: string[]) => void) {
    if (index === 0) return;
    const next = [...list];
    [next[index - 1], next[index]] = [next[index], next[index - 1]];
    setter(next);
  }

  function moveItemDown(list: string[], index: number, setter: (items: string[]) => void) {
    if (index === list.length - 1) return;
    const next = [...list];
    [next[index + 1], next[index]] = [next[index], next[index + 1]];
    setter(next);
  }

  function removeDimension(field: string) {
    setDimensionFields((prev) => prev.filter((f) => f !== field));
  }

  function removeMeasure(field: string) {
    setMeasureFields((prev) => prev.filter((f) => f !== field));
  }

  function setFieldLabel(field: string, label: string) {
    setFieldLabels((prev) => ({
      ...prev,
      [field]: label
    }));
  }

  function updateStyleSetting(key: keyof StyleSettings, value: string) {
    setStyleSettings((prev) => ({
      ...prev,
      [key]: value
    }));
  }

  function updateMeasureFormat(
    field: string,
    key: keyof MeasureFormat,
    value: string | number | boolean
  ) {
    setMeasureFormats((prev) => ({
      ...prev,
      [field]: {
        ...DEFAULT_MEASURE_FORMAT,
        ...(prev[field] || {}),
        [key]: value
      }
    }));
  }

  const availableDimensionFields = useMemo(() => {
    return fields
      .filter((field) => {
        if (measureFields.includes(field.fieldName)) return false;
        if (showAllDimensionFields) return true;
        return field.suggestedRole === 'dimension';
      })
      .sort((a, b) => a.fieldName.localeCompare(b.fieldName));
  }, [fields, measureFields, showAllDimensionFields]);

  const availableMeasureFields = useMemo(() => {
    return fields
      .filter((field) => {
        if (dimensionFields.includes(field.fieldName)) return false;
        if (showAllMeasureFields) return true;
        return field.suggestedRole === 'measure';
      })
      .sort((a, b) => a.fieldName.localeCompare(b.fieldName));
  }, [fields, dimensionFields, showAllMeasureFields]);

  async function save() {
    try {
      if (!worksheetName || dimensionFields.length === 0 || measureFields.length === 0) {
        setStatus('Please select a worksheet, at least one dimension, and at least one measure.');
        return;
      }

      tableau.extensions.settings.set('sourceWorksheet', worksheetName);
      tableau.extensions.settings.set('dimensionFields', JSON.stringify(dimensionFields));
      tableau.extensions.settings.set('measureFields', JSON.stringify(measureFields));
      tableau.extensions.settings.set('fieldLabels', JSON.stringify(fieldLabels));
      tableau.extensions.settings.set('styleSettings', JSON.stringify(styleSettings));
      tableau.extensions.settings.set('measureFormats', JSON.stringify(measureFormats));
      tableau.extensions.settings.set('layoutSettings', JSON.stringify(layoutSettings));
      tableau.extensions.settings.set('displaySettings', JSON.stringify(displaySettings));
      tableau.extensions.settings.set('showAllDimensionFields', String(showAllDimensionFields));
      tableau.extensions.settings.set('showAllMeasureFields', String(showAllMeasureFields));

      await tableau.extensions.settings.saveAsync();
      tableau.extensions.ui.closeDialog('saved');
    } catch (error: any) {
      setStatus(`Save error: ${error?.message || error}`);
    }
  }

  const sectionBox: React.CSSProperties = {
    border: '1px solid #cfcfcf',
    borderRadius: '6px',
    padding: '12px',
    background: '#fff'
  };

  const dropdownButton: React.CSSProperties = {
    width: '100%',
    padding: '10px 12px',
    border: '1px solid #bdbdbd',
    background: '#fff',
    textAlign: 'left',
    cursor: 'pointer',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center'
  };

  const menuPanel: React.CSSProperties = {
    position: 'absolute',
    top: '100%',
    left: 0,
    right: 0,
    marginTop: '4px',
    border: '1px solid #bdbdbd',
    background: '#fff',
    zIndex: 20,
    maxHeight: '240px',
    overflowY: 'auto',
    boxShadow: '0 4px 12px rgba(0,0,0,0.12)'
  };

  const rowStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 10px',
    borderBottom: '1px solid #eee'
  };

  const selectedItemStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: '10px',
    padding: '10px',
    border: '1px solid #ddd',
    borderRadius: '6px',
    marginBottom: '8px',
    background: '#fafafa'
  };

  const smallButton: React.CSSProperties = {
    border: '1px solid #bbb',
    background: '#fff',
    padding: '4px 8px',
    cursor: 'pointer',
    borderRadius: '4px'
  };

  const metaStyle: React.CSSProperties = {
    fontSize: '12px',
    color: '#666'
  };

  const colorGrid: React.CSSProperties = {
    display: 'grid',
    gridTemplateColumns: '1fr 120px 120px',
    gap: '10px',
    alignItems: 'center'
  };

  const labelInputStyle: React.CSSProperties = {
    width: '220px',
    padding: '6px 8px',
    border: '1px solid #ccc',
    borderRadius: '4px'
  };

  const measureFormatGrid: React.CSSProperties = {
    display: 'grid',
    gridTemplateColumns: '160px 120px 90px 100px 100px 150px 110px',
    gap: '8px',
    alignItems: 'center',
    fontSize: '13px'
  };

  const widthGrid: React.CSSProperties = {
    display: 'grid',
    gridTemplateColumns: '180px 140px',
    gap: '10px',
    alignItems: 'center'
  };

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h2>Configure Pivot Extension</h2>
      <p>{status}</p>

      <div style={{ marginBottom: '20px' }}>
        <label>Worksheet</label>
        <br />
        <select
          value={worksheetName}
          onChange={(e) => onWorksheetChange(e.target.value)}
          style={{ width: '100%', marginTop: '8px', padding: '10px' }}
        >
          <option value="">-- Select worksheet --</option>
          {worksheets.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Display Text</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          <div style={{ marginBottom: '12px' }}>
            <label>
              <input
                type="checkbox"
                checked={displaySettings.showTableTitle}
                onChange={(e) =>
                  updateDisplaySetting('showTableTitle', e.target.checked)
                }
                style={{ marginRight: '8px' }}
              />
              Show table title
            </label>
          </div>

          <div style={{ marginBottom: '12px' }}>
            <div style={{ marginBottom: '6px' }}>Table title</div>
            <input
              type="text"
              value={displaySettings.tableTitle}
              onChange={(e) =>
                updateDisplaySetting('tableTitle', e.target.value)
              }
              style={{ width: '320px', padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }}
            />
          </div>

          <div>
            <div style={{ marginBottom: '6px' }}>Hierarchy header</div>
            <input
              type="text"
              value={displaySettings.hierarchyHeader}
              onChange={(e) =>
                updateDisplaySetting('hierarchyHeader', e.target.value)
              }
              style={{ width: '320px', padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }}
            />
          </div>
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Columns Found</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          {columnsFound.length === 0 ? (
            <div style={{ color: '#666' }}>No columns loaded.</div>
          ) : (
            <div>{columnsFound.join(', ')}</div>
          )}
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Preview Rows</label>
        <div style={{ ...sectionBox, marginTop: '8px', overflowX: 'auto' }}>
          {previewRows.length === 0 || columnsFound.length === 0 ? (
            <div style={{ color: '#666' }}>No preview rows loaded.</div>
          ) : (
            <table style={{ borderCollapse: 'collapse', width: '100%', minWidth: '700px' }}>
              <thead>
                <tr>
                  {columnsFound.map((col) => (
                    <th
                      key={col}
                      style={{
                        border: '1px solid #ddd',
                        padding: '8px',
                        background: '#f5f5f5',
                        textAlign: 'left'
                      }}
                    >
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewRows.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {columnsFound.map((col) => (
                      <td
                        key={col}
                        style={{
                          border: '1px solid #ddd',
                          padding: '8px'
                        }}
                      >
                        {String(row[col] ?? '')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: '1fr 1fr',
          gap: '20px',
          marginBottom: '24px'
        }}
      >
        <div ref={dimensionMenuRef} style={{ position: 'relative' }}>
          <label>Select Dimensions</label>
          <button
            style={dropdownButton}
            onClick={() => setDimensionMenuOpen((prev) => !prev)}
          >
            <span>
              {dimensionFields.length > 0
                ? `${dimensionFields.length} selected`
                : 'Select Dimensions'}
            </span>
            <span>{dimensionMenuOpen ? '▲' : '▼'}</span>
          </button>

          <div style={{ marginTop: '6px', fontSize: '13px' }}>
            <label>
              <input
                type="checkbox"
                checked={showAllDimensionFields}
                onChange={(e) => setShowAllDimensionFields(e.target.checked)}
                style={{ marginRight: '6px' }}
              />
              Show all fields
            </label>
          </div>

          {dimensionMenuOpen && (
            <div style={menuPanel}>
              {availableDimensionFields.length === 0 && (
                <div style={{ padding: '10px' }}>No dimension fields available.</div>
              )}

              {availableDimensionFields.map((field) => (
                <label key={field.fieldName} style={rowStyle}>
                  <input
                    type="checkbox"
                    checked={dimensionFields.includes(field.fieldName)}
                    onChange={() => toggleDimension(field.fieldName)}
                  />
                  <div>
                    <div>{field.fieldName}</div>
                    <div style={metaStyle}>
                      Type: {field.dataType || 'unknown'} | Suggested: {field.suggestedRole}
                    </div>
                  </div>
                </label>
              ))}
            </div>
          )}
        </div>

        <div ref={measureMenuRef} style={{ position: 'relative' }}>
          <label>Select Measures</label>
          <button
            style={dropdownButton}
            onClick={() => setMeasureMenuOpen((prev) => !prev)}
          >
            <span>
              {measureFields.length > 0
                ? `${measureFields.length} selected`
                : 'Select Measures'}
            </span>
            <span>{measureMenuOpen ? '▲' : '▼'}</span>
          </button>

          <div style={{ marginTop: '6px', fontSize: '13px' }}>
            <label>
              <input
                type="checkbox"
                checked={showAllMeasureFields}
                onChange={(e) => setShowAllMeasureFields(e.target.checked)}
                style={{ marginRight: '6px' }}
              />
              Show all fields
            </label>
          </div>

          {measureMenuOpen && (
            <div style={menuPanel}>
              {availableMeasureFields.length === 0 && (
                <div style={{ padding: '10px' }}>No measure fields available.</div>
              )}

              {availableMeasureFields.map((field) => (
                <label key={field.fieldName} style={rowStyle}>
                  <input
                    type="checkbox"
                    checked={measureFields.includes(field.fieldName)}
                    onChange={() => toggleMeasure(field.fieldName)}
                  />
                  <div>
                    <div>{field.fieldName}</div>
                    <div style={metaStyle}>
                      Type: {field.dataType || 'unknown'} | Suggested: {field.suggestedRole}
                    </div>
                  </div>
                </label>
              ))}
            </div>
          )}
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Selected Dimensions (hierarchy order + display label)</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          {dimensionFields.length === 0 && (
            <div style={{ color: '#666' }}>No dimensions selected.</div>
          )}

          {dimensionFields.map((field, index) => (
            <div key={field} style={selectedItemStyle}>
              <div>
                <div><strong>{field}</strong></div>
                <div style={{ marginTop: '6px' }}>
                  <input
                    type="text"
                    placeholder="Custom header label"
                    value={fieldLabels[field] || ''}
                    onChange={(e) => setFieldLabel(field, e.target.value)}
                    style={labelInputStyle}
                  />
                </div>
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button
                  style={smallButton}
                  onClick={() => moveItemUp(dimensionFields, index, setDimensionFields)}
                >
                  ↑
                </button>
                <button
                  style={smallButton}
                  onClick={() => moveItemDown(dimensionFields, index, setDimensionFields)}
                >
                  ↓
                </button>
                <button style={smallButton} onClick={() => removeDimension(field)}>
                  Remove
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Selected Measures (column order + display label)</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          {measureFields.length === 0 && (
            <div style={{ color: '#666' }}>No measures selected.</div>
          )}

          {measureFields.map((field, index) => (
            <div key={field} style={selectedItemStyle}>
              <div>
                <div><strong>{field}</strong></div>
                <div style={{ marginTop: '6px' }}>
                  <input
                    type="text"
                    placeholder="Custom header label"
                    value={fieldLabels[field] || ''}
                    onChange={(e) => setFieldLabel(field, e.target.value)}
                    style={labelInputStyle}
                  />
                </div>
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button
                  style={smallButton}
                  onClick={() => moveItemUp(measureFields, index, setMeasureFields)}
                >
                  ↑
                </button>
                <button
                  style={smallButton}
                  onClick={() => moveItemDown(measureFields, index, setMeasureFields)}
                >
                  ↓
                </button>
                <button style={smallButton} onClick={() => removeMeasure(field)}>
                  Remove
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Column Widths</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          <div style={{ ...widthGrid, marginBottom: '10px' }}>
            <div>Hierarchy width (px)</div>
            <input
              type="number"
              min={120}
              value={layoutSettings.hierarchyWidth}
              onChange={(e) => updateHierarchyWidth(Number(e.target.value))}
            />
          </div>

          <div style={{ ...widthGrid, marginBottom: '16px' }}>
            <div>Default measure width (px)</div>
            <input
              type="number"
              min={80}
              value={layoutSettings.defaultMeasureWidth}
              onChange={(e) => updateDefaultMeasureWidth(Number(e.target.value))}
            />
          </div>

          {measureFields.length > 0 && (
            <>
              <div style={{ fontWeight: 700, marginBottom: '8px' }}>Per-measure widths</div>

              {measureFields.map((field) => (
                <div key={field} style={{ ...widthGrid, marginBottom: '8px' }}>
                  <div>{fieldLabels[field]?.trim() || field}</div>
                  <input
                    type="number"
                    min={80}
                    value={getMeasureWidth(field)}
                    onChange={(e) => updateMeasureWidth(field, Number(e.target.value))}
                  />
                </div>
              ))}
            </>
          )}

          <div style={{ marginTop: '14px' }}>
            <button style={smallButton} onClick={resetWidthsToDefault}>
              Reset widths to default
            </button>
          </div>
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Formatting</label>
        <div style={{ ...sectionBox, marginTop: '8px' }}>
          <div style={{ ...colorGrid, marginBottom: '12px' }}>
            <strong>Section</strong>
            <strong>Background</strong>
            <strong>Text</strong>
          </div>

          {[
            ['Dimension headers', 'dimensionHeaderBg', 'dimensionHeaderText'],
            ['Dimension values', 'dimensionValueBg', 'dimensionValueText'],
            ['Measure headers', 'measureHeaderBg', 'measureHeaderText'],
            ['Measure values', 'measureValueBg', 'measureValueText']
          ].map(([label, bgKey, textKey]) => (
            <div key={label} style={{ ...colorGrid, marginBottom: '10px' }}>
              <div>{label}</div>
              <input
                type="color"
                value={styleSettings[bgKey as keyof StyleSettings]}
                onChange={(e) =>
                  updateStyleSetting(bgKey as keyof StyleSettings, e.target.value)
                }
              />
              <input
                type="color"
                value={styleSettings[textKey as keyof StyleSettings]}
                onChange={(e) =>
                  updateStyleSetting(textKey as keyof StyleSettings, e.target.value)
                }
              />
            </div>
          ))}
        </div>
      </div>

      <div style={{ marginBottom: '22px' }}>
        <label>Per-Measure Formatting</label>
        <div style={{ ...sectionBox, marginTop: '8px', overflowX: 'auto' }}>
          {measureFields.length === 0 && (
            <div style={{ color: '#666' }}>No measures selected.</div>
          )}

          {measureFields.length > 0 && (
            <>
              <div style={{ ...measureFormatGrid, marginBottom: '10px', fontWeight: 700 }}>
                <div>Measure</div>
                <div>Format</div>
                <div>Decimals</div>
                <div>Prefix</div>
                <div>Suffix</div>
                <div>Thousands</div>
                <div>Align</div>
              </div>

              {measureFields.map((field) => {
                const format = {
                  ...DEFAULT_MEASURE_FORMAT,
                  ...(measureFormats[field] || {})
                };

                return (
                  <div
                    key={field}
                    style={{
                      ...measureFormatGrid,
                      marginBottom: '10px',
                      padding: '8px 0',
                      borderTop: '1px solid #eee'
                    }}
                  >
                    <div>{fieldLabels[field]?.trim() || field}</div>

                    <select
                      value={format.formatType}
                      onChange={(e) =>
                        updateMeasureFormat(field, 'formatType', e.target.value)
                      }
                    >
                      <option value="number">Number</option>
                      <option value="currency">Currency</option>
                      <option value="percent">Percent</option>
                    </select>

                    <input
                      type="number"
                      min={0}
                      max={6}
                      value={format.decimals}
                      onChange={(e) =>
                        updateMeasureFormat(field, 'decimals', Number(e.target.value))
                      }
                      style={{ width: '70px' }}
                    />

                    <input
                      type="text"
                      value={format.prefix}
                      onChange={(e) =>
                        updateMeasureFormat(field, 'prefix', e.target.value)
                      }
                      style={{ width: '90px' }}
                    />

                    <input
                      type="text"
                      value={format.suffix}
                      onChange={(e) =>
                        updateMeasureFormat(field, 'suffix', e.target.value)
                      }
                      style={{ width: '90px' }}
                    />

                    <label>
                      <input
                        type="checkbox"
                        checked={format.useThousandsSeparator}
                        onChange={(e) =>
                          updateMeasureFormat(field, 'useThousandsSeparator', e.target.checked)
                        }
                        style={{ marginRight: '6px' }}
                      />
                      On
                    </label>

                    <select
                      value={format.alignment}
                      onChange={(e) =>
                        updateMeasureFormat(field, 'alignment', e.target.value)
                      }
                    >
                      <option value="left">Left</option>
                      <option value="center">Center</option>
                      <option value="right">Right</option>
                    </select>
                  </div>
                );
              })}
            </>
          )}
        </div>
      </div>

      <button onClick={save}>Save Configuration</button>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')!).render(<DialogApp />);