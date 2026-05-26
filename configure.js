let calculations = [];
let contextPayload = { worksheets: [], fields: [] };

function $(id) {
  return document.getElementById(id);
}

function safeJsonParse(raw, fallback) {
  if (!raw) return fallback;
  try {
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

function getFieldLabel(fieldName) {
  const match = contextPayload.fields.find((f) => f.fieldName === fieldName);
  return match ? match.caption : fieldName;
}

function populateWorksheetOptions() {
  const worksheetSelect = $('worksheet');
  worksheetSelect.innerHTML = '';

  contextPayload.worksheets.forEach((worksheetName) => {
    const option = document.createElement('option');
    option.value = worksheetName;
    option.textContent = worksheetName;
    worksheetSelect.appendChild(option);
  });
}

function populateFieldOptions() {
  const fieldSelect = $('field');
  fieldSelect.innerHTML = '';

  contextPayload.fields.forEach((field) => {
    const option = document.createElement('option');
    option.value = field.fieldName;
    option.textContent = field.caption;
    fieldSelect.appendChild(option);
  });
}

function renderCalculations() {
  const container = $('calcList');

  if (!calculations.length) {
    container.className = 'empty';
    container.textContent = 'No calculations added yet.';
    return;
  }

  container.className = '';
  container.innerHTML = '';

  calculations.forEach((calc) => {
    const card = document.createElement('div');
    card.className = 'card';

    const top = document.createElement('div');
    top.className = 'card-top';

    const title = document.createElement('div');
    title.className = 'card-title';
    title.textContent = calc.label;

    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'danger';
    remove.textContent = 'Remove';
    remove.addEventListener('click', () => {
      calculations = calculations.filter((item) => item.id !== calc.id);
      renderCalculations();
    });

    top.appendChild(title);
    top.appendChild(remove);

    const body = document.createElement('div');
    body.style.fontSize = '13px';
    body.style.color = '#475569';
    body.innerHTML = `
      <div><strong>Worksheet:</strong> ${calc.worksheet}</div>
      <div><strong>Calculation:</strong> ${calc.aggregation}(${getFieldLabel(calc.field)})</div>
      <div><strong>Format:</strong> ${calc.format}</div>
    `;

    card.appendChild(top);
    card.appendChild(body);
    container.appendChild(card);
  });
}

function addCalculation() {
  const label = $('label').value.trim();
  const worksheet = $('worksheet').value;
  const field = $('field').value;
  const aggregation = $('aggregation').value;
  const format = $('format').value;

  if (!label || !worksheet || !field || !aggregation || !format) {
    alert('Please complete all fields before adding the calculation.');
    return;
  }

  calculations.push({
    id: `calc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    label,
    worksheet,
    field,
    aggregation,
    format
  });

  $('label').value = '';
  renderCalculations();
}

async function saveAndClose() {
  tableau.extensions.settings.set('infoCalculations', JSON.stringify(calculations));
  await tableau.extensions.settings.saveAsync();
  tableau.extensions.ui.closeDialog('saved');
}

async function init() {
  const openPayload = await tableau.extensions.initializeDialogAsync();
  contextPayload = safeJsonParse(openPayload, { worksheets: [], fields: [] });

  populateWorksheetOptions();
  populateFieldOptions();

  calculations = safeJsonParse(
    tableau.extensions.settings.get('infoCalculations'),
    []
  );

  renderCalculations();

  $('addCalc').addEventListener('click', addCalculation);
  $('cancelBtn').addEventListener('click', () => tableau.extensions.ui.closeDialog('cancel'));
  $('saveBtn').addEventListener('click', saveAndClose);
}

document.addEventListener('DOMContentLoaded', init);