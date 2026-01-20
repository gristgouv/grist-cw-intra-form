const fieldsContainer = document.getElementById('fields');
const addButton = document.getElementById('addBtn');
const configModal = document.getElementById('configModal');
const closeModal = document.getElementById('closeModal');
const elementType = document.getElementById('elementType');
const columnSelect = document.getElementById('columnSelect');
const elementContent = document.getElementById('elementContent');
const addElementBtn = document.getElementById('addElementBtn');
const allElementsContainer = document.getElementById('allElements');
const popupOverlay = document.getElementById('popupOverlay');
const formError = document.getElementById('formError');
const formSuccess = document.getElementById('formSuccess');
const fontSelect = document.getElementById('fontSelect');
const paddingSelect = document.getElementById('paddingSelect');
const formContainer = document.querySelector('.container');

let columns = [];
let columnMetadata = {};
let formElements = [];
let draggedElement = null;
let initInProgress = false;
let globalFont = '';
let globalPadding = '';


grist.ready({
  requiredAccess: 'full',
  onEditOptions: () => configModal.classList.add('show')
});

(async () => {
  console.group('🚀 INIT METADATA');

  // === TESTS DE COMPARAISON (temporaires) ===
  await runComparisonTests();

  columns = await getAllColumnsFromMetadata();
  console.log('✅ Colonnes metadata:', columns);

  columnMetadata = await getColumnMetadata();
  console.log('📘 columnMetadata:', columnMetadata);

  await loadConfiguration();

  console.groupEnd();
})();

// ============================================
// FONCTIONS DE TEST (à supprimer après validation)
// ============================================

async function runComparisonTests() {
  console.group('🧪 TESTS DE COMPARAISON');

  try {
    await testCompareColumns();
  } catch (e) {
    console.error('❌ Test colonnes échoué:', e.message);
  }

  try {
    await testCompareMetadata();
  } catch (e) {
    console.error('❌ Test metadata échoué:', e.message);
  }

  console.groupEnd();
}

// Test 1: Comparer getAllColumnsFromMetadata (ancienne) vs API REST (nouvelle)
async function testCompareColumns() {
  console.group('📋 Test: getAllColumnsFromMetadata');

  // Ancienne méthode (tables système)
  let oldResult = [];
  try {
    const table = await grist.getTable();
    const tableName = await table._platform.getTableId();
    const tables = await grist.docApi.fetchTable('_grist_Tables');
    const columnsTable = await grist.docApi.fetchTable('_grist_Tables_column');
    const tableRef = tables.id[tables.tableId.indexOf(tableName)];

    for (let i = 0; i < columnsTable.parentId.length; i++) {
      if (columnsTable.parentId[i] === tableRef) {
        const colId = columnsTable.colId[i];
        if (colId !== 'id' && colId !== 'manualSort' && !colId.startsWith('gristHelper')) {
          oldResult.push(colId);
        }
      }
    }
    console.log('✅ Ancienne méthode (fetchTable):', oldResult);
  } catch (e) {
    console.warn('⚠️ Ancienne méthode échouée:', e.message);
  }

  // Nouvelle méthode (API REST)
  let newResult = [];
  try {
    const tableId = await grist.selectedTable.getTableId();
    const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
    const response = await fetch(`${tokenInfo.baseUrl}/tables/${tableId}/columns`, {
      headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
    });
    if (response.ok) {
      const data = await response.json();
      newResult = data.columns
        .map(col => col.id)
        .filter(colId => colId !== 'id' && colId !== 'manualSort' && !colId.startsWith('gristHelper'));
    }
    console.log('✅ Nouvelle méthode (API REST):', newResult);
  } catch (e) {
    console.warn('⚠️ Nouvelle méthode échouée:', e.message);
  }

  // Comparaison
  const match = JSON.stringify(oldResult.sort()) === JSON.stringify(newResult.sort());
  console.log(match ? '✅ MATCH: Les deux méthodes retournent le même résultat' : '❌ DIFF: Résultats différents!');

  console.groupEnd();
}

// Test 2: Comparer getColumnMetadata (ancienne) vs API REST (nouvelle)
async function testCompareMetadata() {
  console.group('📋 Test: getColumnMetadata');

  // Ancienne méthode
  let oldResult = {};
  try {
    const table = await grist.getTable();
    const currentTableId = await table._platform.getTableId();
    const docInfo = await grist.docApi.fetchTable('_grist_Tables_column');
    const tablesInfo = await grist.docApi.fetchTable('_grist_Tables');
    const currentTableNumericId = tablesInfo.id[tablesInfo.tableId.indexOf(currentTableId)];

    for (let i = 0; i < docInfo.colId.length; i++) {
      if (docInfo.parentId[i] !== currentTableNumericId) continue;
      const colId = docInfo.colId[i];
      const type = docInfo.type[i];
      let choices = null;

      if (docInfo.widgetOptions?.[i]) {
        try {
          const options = JSON.parse(docInfo.widgetOptions[i]);
          if (options.choices) choices = options.choices;
        } catch (e) { }
      }

      oldResult[colId] = { type, choices };
    }
    console.log('✅ Ancienne méthode (fetchTable):', oldResult);
  } catch (e) {
    console.warn('⚠️ Ancienne méthode échouée:', e.message);
  }

  // Nouvelle méthode (API REST)
  let newResult = {};
  try {
    const tableId = await grist.selectedTable.getTableId();
    const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
    const response = await fetch(`${tokenInfo.baseUrl}/tables/${tableId}/columns`, {
      headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
    });

    if (response.ok) {
      const data = await response.json();
      for (const col of data.columns) {
        const colId = col.id;
        const type = col.fields.type || 'Text';
        let choices = null;

        if (col.fields.widgetOptions) {
          try {
            const options = typeof col.fields.widgetOptions === 'string'
              ? JSON.parse(col.fields.widgetOptions)
              : col.fields.widgetOptions;
            if (options.choices) choices = options.choices;
          } catch (e) { }
        }

        newResult[colId] = { type, choices };
      }
    }
    console.log('✅ Nouvelle méthode (API REST):', newResult);
  } catch (e) {
    console.warn('⚠️ Nouvelle méthode échouée:', e.message);
  }

  // Comparaison (on compare type et choices pour chaque colonne)
  let allMatch = true;
  for (const colId of Object.keys(oldResult)) {
    if (!newResult[colId]) {
      console.warn(`⚠️ Colonne ${colId} manquante dans nouvelle méthode`);
      allMatch = false;
      continue;
    }
    if (oldResult[colId].type !== newResult[colId].type) {
      console.warn(`⚠️ Type différent pour ${colId}: ${oldResult[colId].type} vs ${newResult[colId].type}`);
      allMatch = false;
    }
    if (JSON.stringify(oldResult[colId].choices) !== JSON.stringify(newResult[colId].choices)) {
      console.warn(`⚠️ Choices différents pour ${colId}`);
      allMatch = false;
    }
  }

  console.log(allMatch ? '✅ MATCH: Types et choices identiques' : '❌ DIFF: Des différences ont été trouvées');

  console.groupEnd();
}





async function getAllColumnsFromMetadata() {
  try {
    // Méthode standard via tables système
    const table = await grist.getTable();
    const tableName = await table._platform.getTableId();

    const tables = await grist.docApi.fetchTable('_grist_Tables');
    const columnsTable = await grist.docApi.fetchTable('_grist_Tables_column');

    const tableRef = tables.id[tables.tableId.indexOf(tableName)];
    const cols = [];

    for (let i = 0; i < columnsTable.parentId.length; i++) {
      if (columnsTable.parentId[i] === tableRef) {
        const colId = columnsTable.colId[i];
        if (
          colId !== 'id' &&
          colId !== 'manualSort' &&
          !colId.startsWith('gristHelper')
        ) {
          cols.push(colId);
        }
      }
    }

    return cols;
  } catch (error) {
    // Fallback si permissions bloquent l'accès aux tables système
    console.warn('⚠️ Fallback fetchSelectedTable pour colonnes:', error.message);
    const tableData = await grist.fetchSelectedTable();
    return Object.keys(tableData).filter(colId =>
      colId !== 'id' &&
      colId !== 'manualSort' &&
      !colId.startsWith('gristHelper')
    );
  }
}

async function loadConfiguration() {
  console.group('⚙️ loadConfiguration');

  console.log('columns:', columns);
  if (!columns || columns.length === 0) {
    console.warn('⛔ Abort: columns vides');
    console.groupEnd();
    return;
  }

  let options = await grist.getOptions();
  console.log('📦 options brutes:', options);

  // 🔑 première install = options === null
  if (options === null) {
    console.warn('🆕 Première install détectée');
    options = {};
  }

  const isFirstInstall =
    options.initialized !== true &&
    options.formElements === undefined;

  console.log('isFirstInstall:', isFirstInstall);

  // 🔥 AUTO-INIT UNIQUEMENT SI PREMIÈRE INSTALL
  if (isFirstInstall) {
    console.warn('🔥 Initialisation automatique des colonnes');

    // Filtrer les colonnes formule et pièces jointes
    const editableColumns = columns.filter(col => {
      const meta = columnMetadata[col];
      return !meta?.isFormula && !meta?.isAttachment;
    });

    formElements = editableColumns.map(col => ({
      type: 'field',
      fieldName: col,
      fieldLabel: columnMetadata[col]?.label || col,
      required: false,
      maxLength: null,
      conditional: null
    }));

    await grist.setOptions({
      initialized: true,
      formElements
    });
  } else {
    // ✅ CHARGER LA CONFIG EXISTANTE
    formElements = options.formElements || [];
  }

  // Charger les paramètres globaux
  globalFont = options.globalFont || '';
  globalPadding = options.globalPadding || '';

  // Restaurer les sélecteurs
  if (fontSelect) fontSelect.value = globalFont;
  if (paddingSelect) paddingSelect.value = globalPadding;

  // Appliquer les styles au form
  applyGlobalStyles();

  console.log('formElements FINAL:', formElements);

  // ⚠️ RENDER TOUJOURS
  renderConfigList();
  renderForm();
  updateColumnSelect();

  console.groupEnd();
}

async function saveConfiguration() {
  await grist.setOptions({
    initialized: true,
    formElements,
    globalFont,
    globalPadding
  });
}

function applyGlobalStyles() {
  if (formContainer) {
    formContainer.style.fontFamily = globalFont || '';

    let padding = '';
    switch (globalPadding) {
      case 'small': padding = '12px'; break;
      case 'medium': padding = '24px'; break;
      case 'large': padding = '40px'; break;
    }
    formContainer.style.padding = padding || '24px';
  }
}


function updateColumnSelect() {
  console.group('🔍 updateColumnSelect DEBUG');
  console.log('columns (toutes):', columns);
  console.log('columnMetadata:', columnMetadata);

  const usedColumns = formElements
    .filter(el => el.type === 'field')
    .map(el => el.fieldName);
  console.log('usedColumns (dans form):', usedColumns);

  const availableColumns = columns.filter(col => {
    const meta = columnMetadata[col];
    const isUsed = usedColumns.includes(col);
    const isFormula = meta?.isFormula;
    const isAttachment = meta?.isAttachment;
    console.log(`  - ${col}: used=${isUsed}, formula=${isFormula}, attachment=${isAttachment}, meta=`, meta);
    if (isUsed) return false;
    if (isFormula || isAttachment) return false;
    return true;
  });
  console.log('availableColumns (filtrées):', availableColumns);
  console.groupEnd();

  columnSelect.innerHTML = '';

  if (availableColumns.length === 0) {
    const opt = document.createElement('option');
    opt.textContent = 'Toutes les colonnes sont déjà utilisées';
    opt.disabled = true;
    columnSelect.appendChild(opt);
    return;
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '-- Sélectionner une colonne --';
  columnSelect.appendChild(placeholder);

  availableColumns.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col;
    opt.textContent = col;
    columnSelect.appendChild(opt);
  });
}

function showEditPopup(element, index, propertyName = 'content') {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  const isLabel = propertyName === 'fieldLabel';
  const currentValue = isLabel ? (element.fieldLabel || element.fieldName) : (element.content || '');
  const title = isLabel ? 'Modifier le libellé' : 'Modifier le contenu';

  const popup = document.createElement('div');
  popup.className = 'edit-popup rich-editor';
  popup.innerHTML = `
    <h3>${title}</h3>
    <div class="editor-toolbar">
      <button type="button" class="toolbar-btn" data-cmd="bold" title="Gras (Ctrl+B)"><strong>B</strong></button>
      <button type="button" class="toolbar-btn" data-cmd="italic" title="Italique (Ctrl+I)"><em>I</em></button>
      <button type="button" class="toolbar-btn" data-cmd="underline" title="Souligné (Ctrl+U)"><u>U</u></button>
      <span class="toolbar-sep"></span>
      <button type="button" class="toolbar-btn" data-cmd="insertUnorderedList" title="Liste à puces">•</button>
      <button type="button" class="toolbar-btn" id="linkBtn" title="Lien">🔗</button>
      <button type="button" class="toolbar-btn" id="emojiBtn" title="Emoji">😀</button>
      <span class="toolbar-sep"></span>
      <select class="toolbar-select" id="formatSelect">
        <option value="">Style</option>
        <option value="h1">Titre 1</option>
        <option value="h2">Titre 2</option>
        <option value="h3">Titre 3</option>
        <option value="p">Paragraphe</option>
      </select>
      <div class="color-picker-wrapper">
        <button type="button" class="toolbar-btn" id="colorBtn" title="Couleur du texte"><span class="color-icon">A</span></button>
        <div class="color-picker" id="colorPicker"></div>
      </div>
      <div class="color-picker-wrapper">
        <button type="button" class="toolbar-btn" id="bgColorBtn" title="Couleur de fond"><span class="bg-color-icon">A</span></button>
        <div class="color-picker" id="bgColorPicker"></div>
      </div>
    </div>
    <div class="emoji-picker" id="emojiPickerPopup"></div>
    <div id="editContent" class="rich-editor-content" contenteditable="true">${currentValue}</div>
    <div class="edit-popup-buttons">
      <button class="cancel">Annuler</button>
      <button class="save">Enregistrer</button>
    </div>
  `;

  document.body.appendChild(popup);

  const editor = popup.querySelector('#editContent');
  editor.focus();

  // Initialiser les fonctionnalités de l'éditeur
  initRichEditor(popup, editor);

  popup.querySelector('.cancel').addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });

  popup.querySelector('.save').addEventListener('click', () => {
    const newContent = editor.innerHTML.trim();
    if (newContent && newContent !== '<br>') {
      element[propertyName] = newContent;
      saveConfiguration();
      renderConfigList();
      renderForm();
    }
    popup.remove();
    overlay.classList.remove('show');
  });

  overlay.addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });
}

// Emojis disponibles
const emojis = [
  '😀', '😃', '😄', '😁', '😊', '😍', '🥰', '😘',
  '😂', '🤣', '😉', '😎', '🤔', '😐', '😑', '😶',
  '🤝', '👍', '👎', '👏', '🙏', '💪', '✊', '👊',
  '❤️', '💙', '💚', '💛', '🧡', '💜', '🖤', '💔',
  '📊', '📈', '📉', '💼', '🏢', '⚙️', '🔧', '🛠️',
  '✅', '❌', '⚠️', '⛔', '🚫', '💡', '🔔', '📢',
  '🎯', '🎓', '🏆', '🥇', '⭐', '✨', '🔍', '📝'
];

// Couleurs prédéfinies
const colors = [
  '#000000', '#374151', '#6B7280', '#9CA3AF',
  '#DC2626', '#EF4444', '#EA580C', '#F97316',
  '#F59E0B', '#FBBF24', '#84CC16', '#10B981',
  '#14B8A6', '#06B6D4', '#0EA5E9', '#3B82F6',
  '#6366F1', '#8B5CF6', '#A855F7', '#EC4899'
];

function initRichEditor(popup, editor) {
  // Boutons de formatage
  popup.querySelectorAll('.toolbar-btn[data-cmd]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.execCommand(btn.dataset.cmd, false, null);
      editor.focus();
    });
  });

  // Sélecteur de style
  const formatSelect = popup.querySelector('#formatSelect');
  formatSelect.addEventListener('change', () => {
    if (formatSelect.value) {
      document.execCommand('formatBlock', false, formatSelect.value);
      formatSelect.value = '';
      editor.focus();
    }
  });

  // Bouton lien
  popup.querySelector('#linkBtn').addEventListener('click', () => {
    const url = prompt('URL du lien:');
    if (url) {
      const selection = window.getSelection();
      const text = selection.toString() || url;
      const normalizedUrl = url.match(/^https?:\/\//) ? url : 'https://' + url;
      document.execCommand('insertHTML', false, `<a href="${normalizedUrl}" target="_blank">${text}</a>`);
    }
    editor.focus();
  });

  // Emoji picker
  const emojiBtn = popup.querySelector('#emojiBtn');
  const emojiPicker = popup.querySelector('#emojiPickerPopup');
  emojiPicker.innerHTML = emojis.map(e => `<button type="button" class="emoji-btn">${e}</button>`).join('');

  emojiBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    emojiPicker.classList.toggle('show');
  });

  emojiPicker.querySelectorAll('.emoji-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.execCommand('insertText', false, btn.textContent);
      emojiPicker.classList.remove('show');
      editor.focus();
    });
  });

  // Color picker
  const colorBtn = popup.querySelector('#colorBtn');
  const colorPicker = popup.querySelector('#colorPicker');
  colorPicker.innerHTML = colors.map(c => `<button type="button" class="color-btn" style="background:${c}" data-color="${c}"></button>`).join('');

  colorBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    colorPicker.classList.toggle('show');
    popup.querySelector('#bgColorPicker').classList.remove('show');
  });

  colorPicker.querySelectorAll('.color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.execCommand('foreColor', false, btn.dataset.color);
      colorPicker.classList.remove('show');
      editor.focus();
    });
  });

  // Background color picker
  const bgColorBtn = popup.querySelector('#bgColorBtn');
  const bgColorPicker = popup.querySelector('#bgColorPicker');
  bgColorPicker.innerHTML = colors.map(c => {
    const r = parseInt(c.slice(1, 3), 16);
    const g = parseInt(c.slice(3, 5), 16);
    const b = parseInt(c.slice(5, 7), 16);
    return `<button type="button" class="color-btn" style="background:${c}" data-color="rgba(${r},${g},${b},0.2)"></button>`;
  }).join('');

  bgColorBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    bgColorPicker.classList.toggle('show');
    colorPicker.classList.remove('show');
  });

  bgColorPicker.querySelectorAll('.color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.execCommand('backColor', false, btn.dataset.color);
      bgColorPicker.classList.remove('show');
      editor.focus();
    });
  });

  // Fermer les pickers en cliquant ailleurs
  popup.addEventListener('click', (e) => {
    if (!e.target.closest('.color-picker-wrapper') && !e.target.closest('#emojiBtn')) {
      colorPicker.classList.remove('show');
      bgColorPicker.classList.remove('show');
      emojiPicker.classList.remove('show');
    }
  });

  // Raccourcis clavier
  editor.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'b':
          e.preventDefault();
          document.execCommand('bold');
          break;
        case 'i':
          e.preventDefault();
          document.execCommand('italic');
          break;
        case 'u':
          e.preventDefault();
          document.execCommand('underline');
          break;
      }
    }
  });
}

function showValidationPopup(element, index) {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  const popup = document.createElement('div');
  popup.className = 'validation-popup';

  const hasValidation = element.maxLength !== null && element.maxLength !== undefined;

  popup.innerHTML = `
    <h3>Validation du champ</h3>
    <label>Nombre max de caractères :</label>
    <input type="number" id="maxLengthInput" value="${element.maxLength || ''}">
    <div class="validation-popup-buttons">
      <div>
        ${hasValidation ? '<button class="delete">Supprimer</button>' : ''}
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="cancel">Annuler</button>
        <button class="save">Enregistrer</button>
      </div>
    </div>
  `;

  document.body.appendChild(popup);

  popup.querySelector('.cancel').addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });

  if (hasValidation) {
    popup.querySelector('.delete').addEventListener('click', () => {
      element.maxLength = null;
      saveConfiguration();
      renderConfigList();
      popup.remove();
      overlay.classList.remove('show');
    });
  }

  popup.querySelector('.save').addEventListener('click', () => {
    const maxLength = document.getElementById('maxLengthInput').value;
    element.maxLength = maxLength ? parseInt(maxLength) : null;
    saveConfiguration();
    renderConfigList();
    popup.remove();
    overlay.classList.remove('show');
  });

  overlay.addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });
}

function showFilterPopup(element, index) {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  // Récupérer les champs qui sont des listes simples (Choice ou Ref)
  const conditionalFields = formElements.filter(el => {
    if (el.type !== 'field') return false;
    const meta = columnMetadata[el.fieldName];
    if (!meta) return false;
    return (meta.choices && meta.choices.length > 0 && !meta.isMultiple) ||
      (meta.isRef && !meta.isMultiple && meta.refChoices.length > 0);
  });

  const hasConditional = element.conditional !== null && element.conditional !== undefined;

  const popup = document.createElement('div');
  popup.className = 'filter-popup';

  let fieldsOptions = '<option value="">-- Sélectionner un champ --</option>';
  conditionalFields.forEach(field => {
    fieldsOptions += `<option value="${field.fieldName}">${field.fieldName}</option>`;
  });

  popup.innerHTML = `
    <h3>Affichage conditionnel</h3>
    <div class="filter-row">
      <label>Ce champ est affiché si :</label>
    </div>
    <div class="filter-row">
      <select id="conditionalField">${fieldsOptions}</select>
    </div>
    <div class="filter-row">
      <label>est égal à</label>
    </div>
    <div class="filter-row">
      <select id="conditionalValue">
        <option value="">-- Sélectionner une valeur --</option>
      </select>
    </div>
    <div class="filter-popup-buttons">
      <div>
        ${hasConditional ? '<button class="delete">Supprimer</button>' : ''}
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="cancel">Annuler</button>
        <button class="save">Enregistrer</button>
      </div>
    </div>
  `;

  document.body.appendChild(popup);

  const conditionalFieldSelect = popup.querySelector('#conditionalField');
  const conditionalValueSelect = popup.querySelector('#conditionalValue');

  // Pré-remplir si déjà configuré
  if (element.conditional) {
    conditionalFieldSelect.value = element.conditional.field;
    updateConditionalValues(element.conditional.field, conditionalValueSelect);
    setTimeout(() => {
      conditionalValueSelect.value = element.conditional.value;
    }, 0);
  }

  conditionalFieldSelect.addEventListener('change', (e) => {
    updateConditionalValues(e.target.value, conditionalValueSelect);
  });

  popup.querySelector('.cancel').addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });

  if (hasConditional) {
    popup.querySelector('.delete').addEventListener('click', () => {
      element.conditional = null;
      saveConfiguration();
      renderConfigList();
      renderForm();
      popup.remove();
      overlay.classList.remove('show');
    });
  }

  popup.querySelector('.save').addEventListener('click', () => {
    const field = conditionalFieldSelect.value;
    const value = conditionalValueSelect.value;

    if (field && value) {
      element.conditional = { field, value };
    } else {
      element.conditional = null;
    }

    saveConfiguration();
    renderConfigList();
    renderForm();
    popup.remove();
    overlay.classList.remove('show');
  });

  overlay.addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });
}

function updateConditionalValues(fieldName, selectElement) {
  if (!fieldName) {
    selectElement.innerHTML = '<option value="">-- Sélectionner une valeur --</option>';
    return;
  }

  const meta = columnMetadata[fieldName];
  if (!meta) return;

  selectElement.innerHTML = '<option value="">-- Sélectionner une valeur --</option>';

  if (meta.refChoices && meta.refChoices.length > 0) {
    meta.refChoices.forEach(choice => {
      const opt = document.createElement('option');
      opt.value = choice.id;
      opt.textContent = choice.label;
      selectElement.appendChild(opt);
    });
  } else if (meta.choices && meta.choices.length > 0) {
    meta.choices.forEach(choice => {
      const opt = document.createElement('option');
      opt.value = choice;
      opt.textContent = choice;
      selectElement.appendChild(opt);
    });
  }
}

function createEditButton(tooltip, onClick) {
  const btn = document.createElement('div');
  btn.className = 'icon-btn';
  btn.innerHTML = `✎<span class="tooltip">${tooltip}</span>`;
  btn.onclick = onClick;
  return btn;
}

function createControls() {
  const controls = document.createElement('div');
  controls.className = 'element-controls';
  return controls;
}

function createDeleteButton(index) {
  const btn = document.createElement('button');
  btn.textContent = '🗑️';
  btn.onclick = () => {
    formElements.splice(index, 1);
    saveConfiguration();
    renderConfigList();
    renderForm();
    updateColumnSelect();
  };
  return btn;
}

function renderConfigList() {
  allElementsContainer.innerHTML = '';

  formElements.forEach((element, index) => {
    const div = document.createElement('div');
    div.className = 'element-item';
    div.draggable = true;
    div.dataset.index = index;

    const dragHandle = document.createElement('div');
    dragHandle.className = 'drag-handle';
    dragHandle.innerHTML = '⋮⋮';
    div.appendChild(dragHandle);

    const contentWrapper = document.createElement('div');
    contentWrapper.className = 'element-content-wrapper';

    const preview = document.createElement('div');
    preview.className = 'element-preview';

    if (element.type === 'field') {
      const meta = columnMetadata[element.fieldName] || {};
      // Champ texte ou numérique (pour validation max caractères)
      const isTextOrNumericField = !meta.isBool && !meta.isDate && !meta.isMultiple &&
        (!meta.choices || meta.choices.length === 0) &&
        (!meta.isRef || meta.refChoices.length === 0);
      // Champ texte pur (pour multiligne)
      const isPureTextField = isTextOrNumericField && !meta.isNumeric && !meta.isInt;

      // Label
      const labelWrapper = document.createElement('div');
      labelWrapper.className = 'field-label-wrapper';

      const labelSpan = document.createElement('span');
      labelSpan.className = 'field-label-text';
      labelSpan.innerHTML = element.fieldLabel || element.fieldName;

      labelWrapper.appendChild(labelSpan);

      preview.className = 'element-preview field';
      preview.textContent = `Id colonne : ${element.fieldName}`;

      contentWrapper.appendChild(labelWrapper);
      contentWrapper.appendChild(preview);

      const controls = createControls();

      // Icône édition du libellé
      controls.appendChild(createEditButton('Modifier le libellé', () => showEditPopup(element, index, 'fieldLabel')));

      // Icône requis (*)
      const requiredBtn = document.createElement('div');
      requiredBtn.className = 'icon-btn' + (element.required ? ' active' : '');
      requiredBtn.innerHTML = `
        <span class="icon-star">✦</span>
        <span class="tooltip">Rendre le champ obligatoire</span>
      `;
      requiredBtn.onclick = () => {
        element.required = !element.required;
        saveConfiguration();
        renderConfigList();
        renderForm();
      };
      controls.appendChild(requiredBtn);

      // Icône validation (texte et numérique)
      if (isTextOrNumericField) {
        const validationBtn = document.createElement('div');
        validationBtn.className = 'icon-btn' + (element.maxLength ? ' active' : '');
        validationBtn.innerHTML = `
          ✓
          <span class="tooltip">Critère de validation</span>
        `;
        validationBtn.onclick = () => {
          showValidationPopup(element, index);
        };
        controls.appendChild(validationBtn);
      }

      // Icône multiligne (texte pur uniquement)
      if (isPureTextField) {
        const multilineBtn = document.createElement('div');
        multilineBtn.className = 'icon-btn' + (element.multiline ? ' active' : '');
        multilineBtn.innerHTML = `
          ≡
          <span class="tooltip">Multiligne</span>
        `;
        multilineBtn.onclick = () => {
          element.multiline = !element.multiline;
          saveConfiguration();
          renderConfigList();
          renderForm();
        };
        controls.appendChild(multilineBtn);
      }

      // Icône filtre
      const filterBtn = document.createElement('div');
      filterBtn.className = 'icon-btn' + (element.conditional ? ' active' : '');
      filterBtn.innerHTML = `
        ⚡
        <span class="tooltip">Affichage conditionnel</span>
      `;
      filterBtn.onclick = () => {
        showFilterPopup(element, index);
      };
      controls.appendChild(filterBtn);

      controls.appendChild(createDeleteButton(index));

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'separator') {
      preview.className = 'element-preview separator';
      contentWrapper.appendChild(preview);

      const controls = createControls();
      controls.appendChild(createDeleteButton(index));

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'title') {
      preview.className = 'element-preview title';
      preview.innerHTML = element.content;
      contentWrapper.appendChild(preview);

      const controls = createControls();
      controls.appendChild(createEditButton('Modifier', () => showEditPopup(element, index)));
      controls.appendChild(createDeleteButton(index));

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'text') {
      preview.className = 'element-preview text';
      preview.innerHTML = element.content;
      contentWrapper.appendChild(preview);

      const controls = createControls();
      controls.appendChild(createEditButton('Modifier', () => showEditPopup(element, index)));
      controls.appendChild(createDeleteButton(index));

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    }

    div.addEventListener('dragstart', function (e) {
      draggedElement = this;
      this.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });

    div.addEventListener('dragend', function () {
      this.classList.remove('dragging');
      draggedElement = null;

      // Nettoyer tous les indicateurs visuels
      document.querySelectorAll('.element-item').forEach(el => {
        el.classList.remove('drag-over-top', 'drag-over-bottom');
      });
    });

    div.addEventListener('dragover', function (e) {
      if (e.preventDefault) e.preventDefault();
      e.dataTransfer.dropEffect = 'move';

      if (draggedElement && draggedElement !== this) {
        // Nettoyer les classes précédentes
        document.querySelectorAll('.element-item').forEach(el => {
          el.classList.remove('drag-over-top', 'drag-over-bottom');
        });

        // Calculer la position relative de la souris dans l'élément
        const rect = this.getBoundingClientRect();
        const midpoint = rect.top + rect.height / 2;

        if (e.clientY < midpoint) {
          this.classList.add('drag-over-top');
        } else {
          this.classList.add('drag-over-bottom');
        }
      }

      return false;
    });

    div.addEventListener('dragleave', function (e) {
      // Vérifier si on quitte vraiment l'élément (et pas juste un enfant)
      if (e.target === this) {
        this.classList.remove('drag-over-top', 'drag-over-bottom');
      }
    });

    div.addEventListener('drop', function (e) {
      if (e.stopPropagation) e.stopPropagation();

      // Nettoyer tous les indicateurs visuels
      document.querySelectorAll('.element-item').forEach(el => {
        el.classList.remove('drag-over-top', 'drag-over-bottom');
      });

      if (draggedElement && draggedElement !== this) {
        const draggedIndex = parseInt(draggedElement.dataset.index);
        const targetIndex = parseInt(this.dataset.index);

        // Calculer la position relative de la souris dans l'élément
        const rect = this.getBoundingClientRect();
        const midpoint = rect.top + rect.height / 2;

        let insertPosition;
        if (e.clientY < midpoint) {
          // Insérer avant l'élément cible
          insertPosition = targetIndex;
        } else {
          // Insérer après l'élément cible
          insertPosition = targetIndex + 1;
        }

        // Ajuster la position si on déplace vers le bas
        if (draggedIndex < insertPosition) {
          insertPosition--;
        }

        const temp = formElements[draggedIndex];
        formElements.splice(draggedIndex, 1);
        formElements.splice(insertPosition, 0, temp);

        saveConfiguration();
        renderConfigList();
        renderForm();
      }

      return false;
    });

    allElementsContainer.appendChild(div);
  });
}

elementType.addEventListener('change', () => {
  columnSelect.style.display = 'none';
  elementContent.style.display = 'none';

  if (elementType.value === 'column') {
    updateColumnSelect();
    columnSelect.style.display = 'block';
  } else if (elementType.value === 'separator') {
    // Rien
  } else if (elementType.value) {
    elementContent.style.display = 'block';
    elementContent.placeholder = 'Texte';
  }
});

addElementBtn.addEventListener('click', () => {
  const type = elementType.value;
  if (!type) return;

  if (type === 'column') {
    const col = columnSelect.value;
    if (!col) {
      alert('Veuillez sélectionner une colonne');
      return;
    }
    formElements.push({
      type: 'field',
      fieldName: col,
      fieldLabel: columnMetadata[col]?.label || col,
      required: false,
      maxLength: null,
      conditional: null
    });
  } else if (type === 'separator') {
    formElements.push({ type: 'separator', content: '' });
  } else {
    const content = elementContent.value.trim();
    if (!content) {
      alert('Veuillez saisir un contenu');
      return;
    }
    formElements.push({ type, content });
  }

  saveConfiguration();
  renderConfigList();
  renderForm();

  elementType.value = '';
  columnSelect.value = '';
  elementContent.value = '';
  columnSelect.style.display = 'none';
  elementContent.style.display = 'none';
});

closeModal.addEventListener('click', () => configModal.classList.remove('show'));
configModal.addEventListener('click', (e) => {
  if (e.target === configModal) configModal.classList.remove('show');
});

// Event listeners pour les paramètres globaux
fontSelect.addEventListener('change', () => {
  globalFont = fontSelect.value;
  applyGlobalStyles();
  saveConfiguration();
});

paddingSelect.addEventListener('change', () => {
  globalPadding = paddingSelect.value;
  applyGlobalStyles();
  saveConfiguration();
});

async function getColumnMetadata() {
  try {
    const table = await grist.getTable();
    const currentTableId = await table._platform.getTableId();
    const docInfo = await grist.docApi.fetchTable('_grist_Tables_column');
    const tablesInfo = await grist.docApi.fetchTable('_grist_Tables');

    console.log('🔬 DEBUG docInfo keys:', Object.keys(docInfo));
    console.log('🔬 DEBUG docInfo.isFormula:', docInfo.isFormula);
    console.log('🔬 DEBUG docInfo.formula:', docInfo.formula);

    const metadata = {};

    const currentTableNumericId = tablesInfo.id[tablesInfo.tableId.indexOf(currentTableId)];

    for (let i = 0; i < docInfo.colId.length; i++) {
      if (docInfo.parentId[i] !== currentTableNumericId) continue;

      const colId = docInfo.colId[i];
      const type = docInfo.type[i];

      // Debug pour voir les valeurs brutes
      if (['hihi', 'NUMERUS', 'NO_FORMULA', 'INIT_FORMULA', 'FORMULA'].includes(colId)) {
        console.log(`🔬 RAW ${colId}: isFormula=${docInfo.isFormula?.[i]}, formula="${docInfo.formula?.[i]}", recalcWhen=${docInfo.recalcWhen?.[i]}, recalcDeps=${docInfo.recalcDeps?.[i]}`);
      }

      let choices = null;
      let refTable = null;
      let refChoices = [];

      if (docInfo.widgetOptions?.[i]) {
        try {
          const options = JSON.parse(docInfo.widgetOptions[i]);
          if (options.choices) choices = options.choices;
        } catch (e) { }
      }

      if (type.startsWith('Ref:')) {
        refTable = type.substring(4);
      } else if (type.startsWith('RefList:')) {
        refTable = type.substring(8);
      }

      if (refTable) {
        try {
          const refData = await grist.docApi.fetchTable(refTable);
          refChoices = refData.id.map((id, idx) => ({
            id: id,
            label: refData[Object.keys(refData).find(k => k !== 'id' && k !== 'manualSort')]?.[idx] || id
          }));
        } catch (e) { }
      }

      metadata[colId] = {
        type,
        choices,
        label: docInfo.label?.[i] || colId,
        isMultiple: type === 'ChoiceList' || type.startsWith('RefList:'),
        isRef: type.startsWith('Ref:') || type.startsWith('RefList:'),
        refTable,
        refChoices,
        isBool: type === 'Bool',
        isDate: type === 'Date' || type === 'DateTime',
        isNumeric: type === 'Numeric',
        isInt: type === 'Int',
        isFormula: docInfo.isFormula?.[i] === true,
        isAttachment: type === 'Attachments'
      };
    }

    return metadata;
  } catch (error) {
    // Fallback via API REST si permissions bloquent tables système
    console.warn('⚠️ Fallback API REST pour metadata:', error.message);
    return getColumnMetadataViaREST();
  }
}

// Fallback API REST pour les métadonnées
async function getColumnMetadataViaREST() {
  try {
    console.log('🔄 getColumnMetadataViaREST: début');

    const tableId = await grist.selectedTable.getTableId();
    console.log('📍 tableId:', tableId);

    const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
    console.log('🔑 baseUrl:', tokenInfo.baseUrl);

    const url = `${tokenInfo.baseUrl}/tables/${tableId}/columns`;
    console.log('🌐 Fetching:', url);

    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
    });

    console.log('📡 Response status:', response.status);
    if (!response.ok) throw new Error(`API error: ${response.status}`);

    const data = await response.json();
    console.log('📦 API data:', data);

    const metadata = {};

    for (const col of data.columns) {
      const colId = col.id;
      if (colId === 'id' || colId === 'manualSort' || colId.startsWith('gristHelper')) continue;

      const type = col.fields.type || 'Text';
      let choices = null;
      let refTable = null;
      let refChoices = [];

      if (col.fields.widgetOptions) {
        try {
          const options = typeof col.fields.widgetOptions === 'string'
            ? JSON.parse(col.fields.widgetOptions)
            : col.fields.widgetOptions;
          if (options.choices) choices = options.choices;
        } catch (e) { }
      }

      if (type.startsWith('Ref:')) {
        refTable = type.substring(4);
      } else if (type.startsWith('RefList:')) {
        refTable = type.substring(8);
      }

      if (refTable) {
        try {
          const refResponse = await fetch(`${tokenInfo.baseUrl}/tables/${refTable}/records`, {
            headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
          });
          if (refResponse.ok) {
            const refData = await refResponse.json();
            refChoices = refData.records.map(record => {
              const labelKey = Object.keys(record.fields).find(k => k !== 'id' && k !== 'manualSort');
              return { id: record.id, label: labelKey ? record.fields[labelKey] : record.id };
            });
          }
        } catch (e) { }
      }

      metadata[colId] = {
        type,
        choices,
        label: col.fields.label || colId,
        isMultiple: type === 'ChoiceList' || type.startsWith('RefList:'),
        isRef: type.startsWith('Ref:') || type.startsWith('RefList:'),
        refTable,
        refChoices,
        isBool: type === 'Bool',
        isDate: type === 'Date' || type === 'DateTime',
        isNumeric: type === 'Numeric',
        isInt: type === 'Int',
        isFormula: col.fields.isFormula === true,
        isAttachment: type === 'Attachments'
      };
    }

    return metadata;
  } catch (error) {
    console.error("❌ Erreur API REST metadata:", error);
    console.error("❌ Error name:", error.name);
    console.error("❌ Error message:", error.message);
    console.error("❌ Error stack:", error.stack);
    return {};
  }
}

function createInputForColumn(col, meta, element = {}) {
  let inp;

  if (meta.isBool) {
    inp = document.createElement('input');
    inp.type = 'checkbox';
    inp.id = `input_${col}`;
    return inp;
  }

  if (meta.isDate) {
    inp = document.createElement('input');
    inp.type = 'date';
    inp.id = `input_${col}`;
    return inp;
  }

  if (meta.isNumeric || meta.isInt) {
    inp = document.createElement('input');
    inp.type = 'text';
    inp.id = `input_${col}`;
    return inp;
  }

  if (meta.isMultiple) {
    const sel = document.createElement('select');
    sel.multiple = true;
    sel.id = `input_${col}`;
    const opts = meta.refChoices.length > 0 ? meta.refChoices : (meta.choices || []).map(c => ({ id: c, label: c }));
    opts.forEach(opt => {
      const o = document.createElement('option');
      o.value = opt.id;
      o.textContent = opt.label;
      sel.appendChild(o);
    });

    // Permettre la sélection/déselection au simple clic
    sel.addEventListener('mousedown', function (e) {
      e.preventDefault();
      const option = e.target;
      if (option.tagName === 'OPTION') {
        option.selected = !option.selected;
        sel.focus();
        // Déclencher l'événement change pour la mise à jour conditionnelle
        sel.dispatchEvent(new Event('change'));
      }
    });

    return sel;
  }

  if ((meta.choices && meta.choices.length > 0) || (meta.isRef && meta.refChoices.length > 0)) {
    inp = document.createElement('select');
    inp.id = `input_${col}`;
    const empty = document.createElement('option');
    empty.value = '';
    empty.textContent = '-- Sélectionner --';
    inp.appendChild(empty);
    const opts = meta.refChoices.length > 0 ? meta.refChoices : meta.choices.map(c => ({ id: c, label: c }));
    opts.forEach(opt => {
      const o = document.createElement('option');
      o.value = opt.id;
      o.textContent = opt.label;
      inp.appendChild(o);
    });
    return inp;
  }

  // Champ texte multiligne ou simple
  if (element.multiline) {
    inp = document.createElement('textarea');
    inp.id = `input_${col}`;
    inp.rows = 4;
    return inp;
  }

  inp = document.createElement('input');
  inp.type = 'text';
  inp.id = `input_${col}`;
  return inp;
}

function getInputValue(col, meta) {
  const inp = document.getElementById(`input_${col}`);
  if (!inp) return null;

  if (meta.isBool) return inp.checked;

  if (meta.isMultiple) {
    const selected = Array.from(inp.selectedOptions).map(opt => opt.value);
    const values = meta.isRef ? selected.map(v => parseInt(v)) : selected;
    return ["L", ...values];
  }

  if (meta.isRef) return inp.value ? parseInt(inp.value) : null;
  if (meta.isNumeric || meta.isInt) return inp.value ? parseFloat(inp.value) : null;

  return inp.value;
}

function validateField(col, meta, element) {
  const inp = document.getElementById(`input_${col}`);
  const err = document.getElementById(`error_${col}`);

  inp.classList.remove('error');
  if (err) err.classList.remove('show');

  // Vérifier si le champ est requis
  if (element.required) {
    if (meta.isBool) {
      // Les checkboxes sont toujours valides (checked ou non)
    } else if (meta.isMultiple) {
      if (inp.selectedOptions.length === 0) {
        inp.classList.add('error');
        if (err) {
          err.textContent = 'Ce champ est requis';
          err.classList.add('show');
        }
        return false;
      }
    } else {
      const val = inp.value.trim();
      if (val === '') {
        inp.classList.add('error');
        if (err) {
          err.textContent = 'Ce champ est requis';
          err.classList.add('show');
        }
        return false;
      }
    }
  }

  // Validation numérique
  if (meta.isNumeric || meta.isInt) {
    const val = inp.value.trim();
    if (val === '') return true;

    const num = parseFloat(val);
    if (isNaN(num)) {
      inp.classList.add('error');
      if (err) {
        err.textContent = 'Valeur numérique requise';
        err.classList.add('show');
      }
      return false;
    }

    if (meta.isInt && !Number.isInteger(num)) {
      inp.classList.add('error');
      if (err) {
        err.textContent = 'Valeur entière requise';
        err.classList.add('show');
      }
      return false;
    }
  }

  // Validation de la longueur maximale
  if (element.maxLength && !meta.isBool && !meta.isDate && !meta.isMultiple) {
    const val = inp.value;
    if (val.length > element.maxLength) {
      inp.classList.add('error');
      if (err) {
        err.textContent = `Ce champ doit contenir ${element.maxLength} caractères maximum`;
        err.classList.add('show');
      }
      return false;
    }
  }

  return true;
}

function shouldShowField(element) {
  if (!element.conditional) return true;

  const conditionalField = element.conditional.field;
  const conditionalValue = element.conditional.value;

  const inp = document.getElementById(`input_${conditionalField}`);
  if (!inp) return true;

  const meta = columnMetadata[conditionalField];
  if (!meta) return true;

  // Récupérer la valeur actuelle du champ conditionnel
  let currentValue;
  if (meta.isRef) {
    currentValue = inp.value ? parseInt(inp.value) : null;
  } else {
    currentValue = inp.value;
  }

  // Comparer avec la valeur conditionnelle
  const expectedValue = meta.isRef ? parseInt(conditionalValue) : conditionalValue;

  return currentValue == expectedValue;
}

function renderForm() {
  fieldsContainer.innerHTML = '';

  formElements.forEach(element => {
    if (element.type === 'separator') {
      const sep = document.createElement('div');
      sep.className = 'separator';
      fieldsContainer.appendChild(sep);
    } else if (element.type === 'title') {
      const title = document.createElement('div');
      title.className = 'custom-title';
      title.innerHTML = element.content;
      fieldsContainer.appendChild(title);
    } else if (element.type === 'text') {
      const text = document.createElement('div');
      text.className = 'custom-text';
      text.innerHTML = element.content;
      fieldsContainer.appendChild(text);
    } else if (element.type === 'field') {
      const col = element.fieldName;
      const meta = columnMetadata[col] || {};

      const fieldDiv = document.createElement('div');
      fieldDiv.className = meta.isBool ? 'field checkbox-field' : 'field';
      fieldDiv.id = `field_${col}`;

      // Vérifier si le champ doit être affiché
      if (!shouldShowField(element)) {
        fieldDiv.classList.add('hidden');
      }

      const label = document.createElement('label');
      label.innerHTML = element.fieldLabel || col;
      if (element.required) {
        label.innerHTML += ' <span class="required-star">*</span>';
      }

      const inp = createInputForColumn(col, meta, element);

      // Ajouter un event listener pour mettre à jour l'affichage conditionnel
      inp.addEventListener('change', () => {
        updateConditionalFields();
      });

      if (meta.isBool) {
        fieldDiv.appendChild(label);
        fieldDiv.appendChild(inp);
      } else {
        fieldDiv.appendChild(label);
        fieldDiv.appendChild(inp);
        const err = document.createElement('div');
        err.className = 'error-message';
        err.id = `error_${col}`;
        fieldDiv.appendChild(err);
      }

      fieldsContainer.appendChild(fieldDiv);
    }
  });
}

function updateConditionalFields() {
  formElements.forEach(element => {
    if (element.type === 'field') {
      const fieldDiv = document.getElementById(`field_${element.fieldName}`);
      if (fieldDiv) {
        if (shouldShowField(element)) {
          fieldDiv.classList.remove('hidden');
        } else {
          fieldDiv.classList.add('hidden');
        }
      }
    }
  });
}

console.log('🧪 Champs rendus:',
  Array.from(document.querySelectorAll('.field')).map(f => f.id)
);

grist.onRecords(() => {
  console.log('🔄 Records updated');
  renderForm();
});


addButton.addEventListener('click', async () => {
  let valid = true;
  let errorMessages = [];

  formError.classList.remove('show');
  formSuccess.classList.remove('show');

  formElements.forEach(element => {
    if (element.type === 'field') {
      const col = element.fieldName;
      const meta = columnMetadata[col] || {};

      // Ne valider que si le champ est visible
      if (shouldShowField(element)) {
        if (!validateField(col, meta, element)) {
          valid = false;
          errorMessages.push(`${element.fieldLabel || col}`);
        }
      }
    }
  });

  if (!valid) {
    formError.textContent = 'Il y a une ou plusieurs erreurs dans le formulaire, veuillez vérifier les champs';
    formError.classList.add('show');
    return;
  }

  const fields = {};
  formElements.forEach(element => {
    if (element.type === 'field') {
      const col = element.fieldName;
      const meta = columnMetadata[col] || {};

      // Ne collecter que les champs visibles
      if (shouldShowField(element)) {
        fields[col] = getInputValue(col, meta);
      }
    }
  });

  try {
    await grist.selectedTable.create({ fields });

    // Afficher le message de succès
    formSuccess.classList.add('show');
    setTimeout(() => {
      formSuccess.classList.remove('show');
    }, 3000);

    // Réinitialiser le formulaire
    formElements.forEach(element => {
      if (element.type === 'field') {
        const col = element.fieldName;
        const inp = document.getElementById(`input_${col}`);
        const meta = columnMetadata[col] || {};

        if (meta.isBool) {
          inp.checked = false;
        } else if (meta.isMultiple) {
          Array.from(inp.options).forEach(opt => opt.selected = false);
        } else {
          inp.value = '';
        }
      }
    });

    // Mettre à jour l'affichage conditionnel après réinitialisation
    updateConditionalFields();
  } catch (error) {
    console.error("Erreur:", error);
    alert("Erreur: " + error.message);
  }
});