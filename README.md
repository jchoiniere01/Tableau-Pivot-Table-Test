# Tableau Pivot Table Extension

A React + TypeScript + Vite Tableau Extension that renders Tableau worksheet data in a configurable pivot/flat table interface with drag-and-drop field ordering, editable field labels, measure formatting, filters, export support, and Tableau dashboard integration.

## Overview

This project is a Tableau dashboard extension built with React, TypeScript, and Vite. It connects to a Tableau worksheet through the Tableau Extensions API, reads underlying or summary data, and lets users build an interactive tabular view by selecting dimensions, measures, date ranges, and filters. The extension supports both pivot-style and straight-table outputs and includes UI controls for renaming fields, reordering fields, and formatting measure values.

## Core Features

- Tableau dashboard extension integration.
- React + TypeScript + Vite development workflow.
- Configurable data source and worksheet selection.
- Dimension and measure selection.
- Drag-and-drop reordering for dimensions.
- Drag-and-drop reordering for measures.
- Move up / move down controls for selected fields.
- Editable display labels for dimensions and measures.
- Measure format editing from the field edit dialog.
- Active filter display and removal.
- Date range selection.
- Pivot table and straight table rendering.
- Export-friendly tabular output.
- State-driven UI updates for labels, ordering, and formatting.

## Tableau compatibility

This repository maintains two parallel front-end entry sets so the extension can support both newer and older Tableau environments.

### Tableau 2026 and newer

Use the newer compatibility set for Tableau 2026+:

- `App.latest.tsx`
- `main.latest.tsx`
- `index.latest.css`
- `configure.latest.tsx`

These files contain the implementation intended for Tableau 2026 and newer environments.

### Tableau versions prior to 2026

Use the legacy-compatible set for Tableau versions before 2026:

- `App.tsx`
- `main.tsx`
- `index.css`
- `configure.tsx`

These files are retained for prior-version compatibility.

### Compatibility notes

- The `.latest` files are the preferred implementation path for Tableau 2026 and newer.
- The non-`latest` files are the compatibility path for older Tableau versions.
- Any shared logic should be kept aligned across both tracks unless a Tableau-version-specific difference requires separate behavior.
- When updating features such as drag-and-drop, field-label editing, or measure formatting, verify whether the change must be applied to both file sets.

## Project Structure

```text
Tableau-Pivot-Table-Test/
├── public/
├── src/
│   ├── App.latest.tsx          # Tableau 2026 and newer
│   ├── main.latest.tsx         # Tableau 2026 and newer
│   ├── index.latest.css        # Tableau 2026 and newer
│   ├── configure.latest.tsx    # Tableau 2026 and newer
│   ├── App.tsx                 # Versions prior to Tableau 2026
│   ├── main.tsx                # Versions prior to Tableau 2026
│   ├── index.css               # Versions prior to Tableau 2026
│   ├── configure.tsx           # Versions prior to Tableau 2026
│   └── ...
├── package.json
├── tsconfig.json
├── vite.config.ts
├── eslint.config.js
└── README.md
```

If the project includes a Tableau extension manifest, keep it in the location expected by your packaging or hosting workflow.

## Supported Field Customization

### Labels

Users can edit the display label for a selected dimension or measure. Label overrides are stored in local React state and applied through a display-label helper such as `getDisplayLabel(field)`.

### Measure formats

Measures support editable value formatting from the edit dialog. Recommended supported formats:

- `number`
- `integer`
- `decimal-1`
- `decimal-2`
- `decimal-3`
- `currency`
- `percent`

Format overrides are stored by field name, for example:

```ts
const [fieldFormatOverrides, setFieldFormatOverrides] = useState<Record<string, string>>({});
```

The rendered table must read the saved override when formatting values:

```ts
const format = fieldFormatOverrides[field] || 'number';
```

## Drag-and-Drop Behavior

The extension uses sortable list behavior for selected dimensions and selected measures.

### Dimensions

- Selected dimensions are rendered from `selectedDimensions`.
- Drag-and-drop updates `selectedDimensions` directly.
- Move up and move down buttons use the same underlying reorder helper.

### Measures

- Selected measures are rendered from `selectedMeasures`.
- Drag-and-drop updates `selectedMeasures` directly.
- Measure ordering stays in sync with the rendered table.

### Shared reorder helper

```ts
function moveArrayItem<T>(items: T[], fromIndex: number, toIndex: number): T[] {
  if (
    fromIndex === toIndex ||
    fromIndex < 0 ||
    toIndex < 0 ||
    fromIndex >= items.length ||
    toIndex >= items.length
  ) {
    return items;
  }

  const next = [...items];
  const [moved] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, moved);
  return next;
}
```

## Editing Workflow

The edit dialog supports both label editing and measure-format editing.

### Required state

```ts
const [editingFieldLabel, setEditingFieldLabel] = useState<string | null>(null);
const [editingFieldValue, setEditingFieldValue] = useState('');
const [editingFieldFormat, setEditingFieldFormat] = useState('number');
const [fieldLabelOverrides, setFieldLabelOverrides] = useState<Record<string, string>>({});
const [fieldFormatOverrides, setFieldFormatOverrides] = useState<Record<string, string>>({});
```

### Open edit dialog

```ts
function startEditingFieldLabel(field: string) {
  setEditingFieldLabel(field);
  setEditingFieldValue(getDisplayLabel(field));
  setEditingFieldFormat(fieldFormatOverrides[field] || 'number');
}
```

### Save label and format

```ts
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
```

### Conditional measure-format control

Inside the edit dialog, show the format selector only for measures:

```tsx
{isMeasureField(editingFieldLabel) && (
  <select
    value={editingFieldFormat}
    onChange={(e) => setEditingFieldFormat(e.target.value)}
  >
    <option value="number">Number</option>
    <option value="currency">Currency</option>
    <option value="percent">Percent</option>
    <option value="integer">Whole Number</option>
    <option value="decimal-1">1 Decimal</option>
    <option value="decimal-2">2 Decimals</option>
    <option value="decimal-3">3 Decimals</option>
  </select>
)}
```

## Measure Formatting

Measure formatting must be applied during render, not only saved in state.

Example formatter:

```ts
function formatMeasureValue(field: string, value: number) {
  const format = fieldFormatOverrides[field] || 'number';

  switch (format) {
    case 'currency':
      return value.toLocaleString(undefined, {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    case 'percent':
      return `${(value * 100).toFixed(1)}%`;
    case 'integer':
      return value.toLocaleString(undefined, {
        maximumFractionDigits: 0,
      });
    case 'decimal-1':
      return value.toLocaleString(undefined, {
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      });
    case 'decimal-2':
      return value.toLocaleString(undefined, {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    case 'decimal-3':
      return value.toLocaleString(undefined, {
        minimumFractionDigits: 3,
        maximumFractionDigits: 3,
      });
    case 'number':
    default:
      return value.toLocaleString();
  }
}
```

## Local Development

### Prerequisites

- Node.js 18+ recommended.
- npm or another compatible package manager.
- A Tableau environment that supports dashboard extensions.
- Access to the Tableau Extensions API library or script used by the project.

### Install dependencies

```bash
npm install
```

### Start the Vite dev server

```bash
npm run dev
```

### Build for production

```bash
npm run build
```

### Preview the production build

```bash
npm run preview
```

## Tableau Extension Setup

To run the extension inside Tableau, the following pieces are typically required:

1. A hosted or locally served web application URL for the extension UI.
2. A Tableau extension manifest file (`.trex`) that points to the application URL.
3. Tableau dashboard configuration that loads the extension.
4. Permissions and network settings that allow Tableau to reach the extension host.

A typical workflow is:

1. Start the app locally with Vite or deploy it to a web server.
2. Create or update the `.trex` file so its source URL points to the app.
3. Open the Tableau dashboard.
4. Add the extension object to the dashboard.
5. Select the `.trex` manifest.
6. Authorize the extension if Tableau prompts for permissions.

## Recommended Manifest Notes

Your `.trex` file should reflect the deployed extension URL and any required permissions. A minimal setup usually includes:

- extension name
- description
- author
- version
- source URL
- required Tableau API version

If your current repository already has a working manifest, keep that as the system of record.

## Key Application State

Common state used by the extension includes:

- `selectedSources`
- `selectedDimensions`
- `selectedMeasures`
- `filters`
- `dateRange`
- `fieldLabelOverrides`
- `fieldFormatOverrides`
- `editingFieldLabel`
- `editingFieldValue`
- `editingFieldFormat`

Keeping dimensions and measures as the single source of truth for ordering prevents drag-and-drop UI drift.

## Typical Helper Functions

Examples of helpers commonly used in the project:

- `getDisplayLabel(field)`
- `startEditingFieldLabel(field)`
- `saveFieldLabel(field)`
- `cancelFieldLabelEdit()`
- `formatMeasureValue(field, value)`
- `moveDimension(field, direction)`
- `moveMeasure(field, direction)`
- `moveArrayItem(items, fromIndex, toIndex)`
- `toggleSource(source)`
- `toggleMeasure(field)`
- `removeFilter(field)`
- `isMeasureField(field)`

## UI Behavior Summary

### Selected dimensions

- Reorder by drag-and-drop.
- Reorder by up/down controls.
- Edit display label.
- Remove from selection.

### Selected measures

- Reorder by drag-and-drop.
- Reorder by up/down controls.
- Edit display label.
- Edit metric format.
- Remove from selection.

### Active filters

- Show current filter values.
- Remove individual filters.

### Date range

- Show selected start and end dates.
- Use as part of the dataset query / filtering workflow.

## Troubleshooting

### Format changes do not appear

Check that the formatter reads from the override state:

```ts
const format = fieldFormatOverrides[field] || 'number';
```

If it still references another object such as `fieldFormat[field]`, the UI will not reflect saved format changes.

### TypeScript error: Cannot find name `SelectionPill`

Make sure `SelectionPill` is defined as a real top-level component and that all old `renderSelectionPill(...)` usages have been replaced with `<SelectionPill ... />`.

### TypeScript error: Cannot find name `setFieldFormatOverrides`

Make sure the state is declared with the exact matching setter name:

```ts
const [fieldFormatOverrides, setFieldFormatOverrides] = useState<Record<string, string>>({});
```

### Drag-and-drop reverts or items disappear

Use a single source of truth for each sortable list:

- `selectedDimensions` for dimensions
- `selectedMeasures` for measures

Avoid keeping a second mirrored sortable array unless absolutely necessary.

### Measure format dropdown appears but does nothing

Verify all three of the following:

1. `startEditingFieldLabel()` loads the current format.
2. `saveFieldLabel()` saves to `fieldFormatOverrides`.
3. `formatMeasureValue()` reads from `fieldFormatOverrides`.

## Suggested Scripts

Use the package scripts already present in `package.json`. A standard Vite setup usually includes:

```json
{
  "scripts": {
    "dev": "vite",
    "build": "tsc -b && vite build",
    "preview": "vite preview",
    "lint": "eslint ."
  }
}
```

## Tech Stack

- React
- TypeScript
- Vite
- Tableau Extensions API
- Drag-and-drop library used by the project

## Future Enhancements

- Persist user configuration across Tableau sessions.
- Add per-measure currency codes.
- Add configurable percent precision.
- Add saved views / presets.
- Add multi-column sorting.
- Add totals and subtotals.
- Add conditional formatting.
- Add export to CSV or Excel.

## Notes for Contributors

When making UI changes, keep these rules in mind:

- Use one source of truth per sortable list.
- Prefer reusable components such as `SelectionPill` for selected-field UI.
- Keep formatting logic in the render formatter, not only in modal state.
- Reset modal editing state on save and cancel.
- Keep TypeScript state names and setters consistent.
- Apply compatibility-sensitive changes to the correct Tableau version track.

## License

Add the appropriate license for your organization or project here.