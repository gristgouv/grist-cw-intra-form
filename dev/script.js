// =============================================================================
// GRIST CUSTOM FORM WIDGET
// A configurable form widget for Grist that supports drag-drop ordering,
// conditional fields, attachments, rich text editing, and field validation.
// =============================================================================

// -----------------------------------------------------------------------------
// DOM REFERENCES
// -----------------------------------------------------------------------------
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

// -----------------------------------------------------------------------------
// STATE
// -----------------------------------------------------------------------------
let columns = [];              // List of column IDs from current table
let columnMetadata = {};       // Metadata for each column (type, choices, etc.)
let formElements = [];         // Form configuration (fields, separators, text)
let draggedElement = null;     // Currently dragged element in config modal
let globalFont = '';           // Selected font family
let globalPadding = '';        // Selected padding size
let pendingAttachments = {};   // Temporary storage for file uploads per column

// -----------------------------------------------------------------------------
// ATTACHMENT HANDLING
// -----------------------------------------------------------------------------

// Render the list of pending attachments for a column
function renderAttachmentList(col, container) {
  container.innerHTML = '';

  pendingAttachments[col].forEach((file, index) => {
    const fileItem = document.createElement('div');
    fileItem.className = 'attachment-item';

    const fileName = document.createElement('span');
    fileName.className = 'attachment-name';
    fileName.textContent = file.name;

    const fileSize = document.createElement('span');
    fileSize.className = 'attachment-size';
    fileSize.textContent = formatFileSize(file.size);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'attachment-remove';
    removeBtn.textContent = '×';
    removeBtn.addEventListener('click', () => {
      pendingAttachments[col].splice(index, 1);
      renderAttachmentList(col, container);
    });

    fileItem.appendChild(fileName);
    fileItem.appendChild(fileSize);
    fileItem.appendChild(removeBtn);
    container.appendChild(fileItem);
  });
}

// Format file size in human-readable format (bytes, Ko, Mo)
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' o';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' Ko';
  return (bytes / (1024 * 1024)).toFixed(1) + ' Mo';
}

// Upload pending attachments to Grist and return attachment IDs in list format
async function uploadAttachments(col) {
  const files = pendingAttachments[col] || [];
  if (files.length === 0) return null;

  const attachmentIds = [];
  const tokenInfo = await grist.docApi.getAccessToken({ readOnly: false });

  for (const file of files) {
    const formData = new FormData();
    formData.append('upload', file, file.name);

    const response = await fetch(`${tokenInfo.baseUrl}/attachments?auth=${tokenInfo.token}`, {
      method: 'POST',
      body: formData,
      headers: { 'X-Requested-With': 'XMLHttpRequest' },
    });

    if (!response.ok) {
      throw new Error(`Le téléchargement a échoué : ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    attachmentIds.push(result[0]);
  }

  // Return in Grist list format: ['L', id1, id2, ...]
  return ['L', ...attachmentIds];
}

// -----------------------------------------------------------------------------
// GRIST INITIALIZATION
// -----------------------------------------------------------------------------


grist.ready({
  requiredAccess: 'full',
  // Callback when user clicks gear icon > "Open configuration" in widget menu
  onEditOptions: () => configModal.classList.add('show')
});

// Fetch table structure then load saved config
(async () => {
  columnMetadata = await getColumnMetadata();
  columns = Object.keys(columnMetadata);
  await loadConfiguration();
})();

// -----------------------------------------------------------------------------
// COLUMN & METADATA FETCHING
// -----------------------------------------------------------------------------

// Fetch detailed metadata for all columns (type, choices, refs, etc.)
// Returns: { colId: { type, choices, isRef, refChoices, isBool, ... }, ... }
async function getColumnMetadata() {
  try {
    // Get current table name (eg "Table1")
    const table = await grist.getTable();
    const currentTableId = await table._platform.getTableId();

    // _grist_Tables_column: list of all columns across all tables
    // eg   {
    //     id: [1, 2, 3, 4, 5, 6, 7],
    //     colId: ['Nom', 'Email', 'Date', 'Montant', 'Titre', 'Prix', 'Stock'],
    //     parentId: [1, 1, 2, 2, 3, 3, 3]
    //   }
    const colsInfo = await grist.docApi.fetchTable('_grist_Tables_column');

    // _grist_Tables: list of all tables in the document
    // eg   {
    //     id: [1, 2, 3],  // numeric ref
    //     tableId: ['Clients', 'Commandes', 'Produits']
    //   }
    const tablesInfo = await grist.docApi.fetchTable('_grist_Tables');

    const metadata = {};

    // Convert tableId to numeric ref (eg "Clients" → 1)
    const currentTableNumericId = tablesInfo.id[tablesInfo.tableId.indexOf(currentTableId)];

    // Used for visibleCol: numeric column ID -> column info
    // Example: visibleCol=5 means "display column with id=5" for Ref fields
    const colById = {};
    for (let i = 0; i < colsInfo.id.length; i++) {
      colById[colsInfo.id[i]] = {
        colId: colsInfo.colId[i],
        parentId: colsInfo.parentId[i]
      };
    }

    // -----------------------------------------------------------------------------
    // LOOP THROUGH COLUMNS BELONGING TO CURRENT TABLE
    // -----------------------------------------------------------------------------
    for (let i = 0; i < colsInfo.colId.length; i++) {
      if (colsInfo.parentId[i] !== currentTableNumericId) continue;

      const colId = colsInfo.colId[i];

      // Exclude system columns (id, manualSort, gristHelper_*)
      if (colId === 'id' || colId === 'manualSort' || colId.startsWith('gristHelper')) continue;

      const type = colsInfo.type[i];  // eg: "Text", "Int", "Ref:Clients", "ChoiceList"
      let choices = null;
      let refTable = null;
      let refChoices = [];

      // For Choice/ChoiceList columns: extract choices from widgetOptions JSON
      // Example: {"choices": ["Option A", "Option B", "Option C"]}
      if (colsInfo.widgetOptions?.[i]) {
        try {
          const options = JSON.parse(colsInfo.widgetOptions[i]);
          if (options.choices) choices = options.choices;
        } catch (e) { }
      }

      // For Ref/RefList columns: extract target table name from type
      // "Ref:Clients" -> refTable = "Clients"
      // "RefList:Products" -> refTable = "Products"
      if (type.startsWith('Ref:')) {
        refTable = type.substring(4);
      } else if (type.startsWith('RefList:')) {
        refTable = type.substring(8);
      }

      // For Ref/RefList: fetch target table and build dropdown choices
      if (refTable) {
        try {
          const refData = await grist.docApi.fetchTable(refTable);

          // visibleCol: for Reference columns: numeric ID of the display column
          let displayColId = null;
          const visibleColRef = colsInfo.visibleCol?.[i];

          // Resolve numeric ID -> column name using our index
          if (visibleColRef && visibleColRef !== 0 && colById[visibleColRef]) {
            displayColId = colById[visibleColRef].colId;
          }

          // Fallback: use first non-system column if visibleCol not set
          if (!displayColId || !refData[displayColId]) {
            displayColId = Object.keys(refData).find(k => k !== 'id' && k !== 'manualSort');
          }

          // Build choices array: [{id: 1, label: "Client A"}, {id: 2, label: "Client B"}]
          refChoices = refData.id.map((id, idx) => ({
            id: id,
            label: displayColId && refData[displayColId] ? refData[displayColId][idx] : id
          }));
        } catch (e) { }
      }

      // Store all metadata for this column
      metadata[colId] = {
        type,
        choices,                    // For Choice/ChoiceList: ["A", "B", "C"]
        label: colsInfo.label?.[i] || colId,
        isMultiple: type === 'ChoiceList' || type.startsWith('RefList:'),
        isRef: type.startsWith('Ref:') || type.startsWith('RefList:'),
        refTable,                   // Target table name for Ref/RefList
        refChoices,                 // [{id, label}] for Ref/RefList dropdowns
        isBool: type === 'Bool',
        isDate: type === 'Date' || type === 'DateTime',
        isNumeric: type === 'Numeric',
        isInt: type === 'Int',
        isFormula: colsInfo.isFormula?.[i] === true && colsInfo.formula?.[i]?.length > 0,
        isAttachment: type === 'Attachments'
      };
    }

    return metadata;
  } catch (error) {
    return {};
  }
}

// -----------------------------------------------------------------------------
// CONFIGURATION PERSISTENCE
// -----------------------------------------------------------------------------

// Load form configuration from Grist widget options
async function loadConfiguration() {
  const options = await grist.getOptions() || {};
  const isFirstInstall = !options.initialized && !options.formElements;

  if (isFirstInstall) {
    // Auto-initialize with all editable (non-formula) columns
    const editableColumns = columns.filter(col => {
      const meta = columnMetadata[col];
      return !meta?.isFormula;
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
    // Load existing configuration
    formElements = options.formElements || [];
  }

  // Load global style settings
  globalFont = options.globalFont || '';
  globalPadding = options.globalPadding || '';

  // Restore UI state
  if (fontSelect) fontSelect.value = globalFont;
  if (paddingSelect) paddingSelect.value = globalPadding;

  applyGlobalStyles();
  renderConfigList();
  renderForm();
  updateColumnSelect();
}

// Save form configuration to Grist widget options
async function saveConfiguration() {
  await grist.setOptions({
    initialized: true,
    formElements,
    globalFont,
    globalPadding
  });
}

// Apply global font and padding styles to form container
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

// -----------------------------------------------------------------------------
// COLUMN SELECT (for adding new fields)
// -----------------------------------------------------------------------------

// Update column dropdown to show only unused, non-formula columns
// Called when opening the config modal or after adding/removing a field
function updateColumnSelect() {
  // Get list of columns already used in the form
  const usedColumns = formElements
    .filter(el => el.type === 'field')
    .map(el => el.fieldName);

  // Filter out used columns and formula columns (which can't be edited)
  const availableColumns = columns.filter(col => {
    if (usedColumns.includes(col)) return false;
    const meta = columnMetadata[col];
    if (meta?.isFormula) return false;
    return true;
  });

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

// -----------------------------------------------------------------------------
// RICH TEXT EDITOR POPUP
// -----------------------------------------------------------------------------

// Available emojis for rich text editor
const emojis = [
  '😀', '😃', '😄', '😁', '😊', '😍', '🥰', '😘',
  '😂', '🤣', '😉', '😎', '🤔', '😐', '😑', '😶',
  '🤝', '👍', '👎', '👏', '🙏', '💪', '✊', '👊',
  '❤️', '💙', '💚', '💛', '🧡', '💜', '🖤', '💔',
  '📊', '📈', '📉', '💼', '🏢', '⚙️', '🔧', '🛠️',
  '✅', '❌', '⚠️', '⛔', '🚫', '💡', '🔔', '📢',
  '🎯', '🎓', '🏆', '🥇', '⭐', '✨', '🔍', '📝'
];

// Available colors for text and background
const colors = [
  '#000000', '#374151', '#6B7280', '#9CA3AF',
  '#DC2626', '#EF4444', '#EA580C', '#F97316',
  '#F59E0B', '#FBBF24', '#84CC16', '#10B981',
  '#14B8A6', '#06B6D4', '#0EA5E9', '#3B82F6',
  '#6366F1', '#8B5CF6', '#A855F7', '#EC4899'
];

// Show rich text editor popup for editing field labels or text content
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

// Initialize rich text editor toolbar functionality
function initRichEditor(popup, editor) {
  // Format buttons (bold, italic, underline, list)
  popup.querySelectorAll('.toolbar-btn[data-cmd]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.execCommand(btn.dataset.cmd, false, null);
      editor.focus();
    });
  });

  // Block format selector (h1, h2, h3, p)
  const formatSelect = popup.querySelector('#formatSelect');
  formatSelect.addEventListener('change', () => {
    if (formatSelect.value) {
      document.execCommand('formatBlock', false, formatSelect.value);
      formatSelect.value = '';
      editor.focus();
    }
  });

  // Link insertion
  popup.querySelector('#linkBtn').addEventListener('click', () => {
    const url = prompt('URL du lien:');
    if (url) {
      // If text was selected before clicking the link icon
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

  // Text color picker
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

  // Background color picker (20% opacity)
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

  // Close pickers when clicking elsewhere
  popup.addEventListener('click', (e) => {
    if (!e.target.closest('.color-picker-wrapper') && !e.target.closest('#emojiBtn')) {
      colorPicker.classList.remove('show');
      bgColorPicker.classList.remove('show');
      emojiPicker.classList.remove('show');
    }
  });

  // Keyboard shortcuts
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

// -----------------------------------------------------------------------------
// VALIDATION POPUP
// -----------------------------------------------------------------------------

// Show popup to configure max length validation for a field
// Allows user to set a character limit that will be enforced on form submission
function showValidationPopup(element, index) {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  const popup = document.createElement('div');
  popup.className = 'validation-popup';

  // Check if this field already has a maxLength validation configured
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

  // Delete button only exists if a validation was already configured
  if (hasValidation) {
    popup.querySelector('.delete').addEventListener('click', () => {
      // Remove the maxLength validation from this field
      element.maxLength = null;
      saveConfiguration();
      renderConfigList();
      popup.remove();
      overlay.classList.remove('show');
    });
  }

  // Save the validation configuration
  popup.querySelector('.save').addEventListener('click', () => {
    const maxLength = document.getElementById('maxLengthInput').value;
    // Convert to integer or null if empty
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

// -----------------------------------------------------------------------------
// CONDITIONAL DISPLAY POPUP
// -----------------------------------------------------------------------------

// Show popup to configure conditional display rules for a field
function showFilterPopup(element, index) {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  // Get fields that can be used as conditions (single Choice or Ref)
  const conditionalFields = formElements.filter(el => {
    if (el.type !== 'field') return false;
    const meta = columnMetadata[el.fieldName];
    if (!meta) return false;
    return (meta.choices && meta.choices.length > 0 && !meta.isMultiple) ||
      (meta.isRef && !meta.isMultiple && meta.refChoices.length > 0);
  });

  // Check if this field already has a conditional rule configured
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
      <select id="conditionalOperator">
        <option value="equals">est égal à</option>
        <option value="notEquals">n'est pas égal à</option>
      </select>
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
  const conditionalOperatorSelect = popup.querySelector('#conditionalOperator');
  const conditionalValueSelect = popup.querySelector('#conditionalValue');

  // Pre-fill if already configured
  if (element.conditional) {
    conditionalFieldSelect.value = element.conditional.field;
    conditionalOperatorSelect.value = element.conditional.operator || 'equals';
    updateConditionalValues(element.conditional.field, conditionalValueSelect);
    setTimeout(() => {
      conditionalValueSelect.value = element.conditional.value;
    }, 0);
  }

  // When user selects a different field, update the available values
  conditionalFieldSelect.addEventListener('change', (e) => {
    updateConditionalValues(e.target.value, conditionalValueSelect);
  });

  popup.querySelector('.cancel').addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });

  // Delete button only exists if a conditional rule was already configured
  if (hasConditional) {
    popup.querySelector('.delete').addEventListener('click', () => {
      // Remove the conditional rule from this field
      element.conditional = null;
      saveConfiguration();
      renderConfigList();
      renderForm();
      popup.remove();
      overlay.classList.remove('show');
    });
  }

  // Save the conditional rule configuration
  popup.querySelector('.save').addEventListener('click', () => {
    const field = conditionalFieldSelect.value;
    const operator = conditionalOperatorSelect.value;
    const value = conditionalValueSelect.value;

    // Only save if both field and value are selected, otherwise clear the rule
    if (field && value) {
      element.conditional = { field, operator, value };
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

// Populate value dropdown based on selected conditional field
// Shows either reference choices (for Ref columns) or choice options (for Choice columns)
function updateConditionalValues(fieldName, selectElement) {
  // Reset to empty placeholder if no field selected
  if (!fieldName) {
    selectElement.innerHTML = '<option value="">-- Sélectionner une valeur --</option>';
    return;
  }

  const meta = columnMetadata[fieldName];
  if (!meta) return;

  // Reset dropdown before populating
  selectElement.innerHTML = '<option value="">-- Sélectionner une valeur --</option>';

  // For Ref columns: use refChoices (id + label from referenced table)
  if (meta.refChoices && meta.refChoices.length > 0) {
    meta.refChoices.forEach(choice => {
      const opt = document.createElement('option');
      opt.value = choice.id;
      opt.textContent = choice.label;
      selectElement.appendChild(opt);
    });
  // For Choice columns: use choices array directly
  } else if (meta.choices && meta.choices.length > 0) {
    meta.choices.forEach(choice => {
      const opt = document.createElement('option');
      opt.value = choice;
      opt.textContent = choice;
      selectElement.appendChild(opt);
    });
  }
}

// -----------------------------------------------------------------------------
// FORM CONFIGURATION : LIST UI HELPERS
// -----------------------------------------------------------------------------

// Create an edit button with tooltip
function createEditButton(tooltip, onClick) {
  const btn = document.createElement('div');
  btn.className = 'icon-btn';
  btn.innerHTML = `✎<span class="tooltip">${tooltip}</span>`;
  btn.onclick = onClick;
  return btn;
}

// Create controls container
function createControls() {
  const controls = document.createElement('div');
  controls.className = 'element-controls';
  return controls;
}

// Create delete button for form element
// Removes the element from form config and refreshes all views
function createDeleteButton(index) {
  const btn = document.createElement('button');
  btn.textContent = '🗑️';
  btn.onclick = () => {
    formElements.splice(index, 1);
    saveConfiguration();
    renderConfigList();
    renderForm();
    updateColumnSelect();  // Column becomes available again
  };
  return btn;
}

// -----------------------------------------------------------------------------
// FORM CONFIGURATION : LIST & BLOCKS RENDERING
// -----------------------------------------------------------------------------

// Render the configuration list with drag-drop support
// Displays all form elements in the config modal with their controls
// (edit label, required toggle, validation, multiline, conditional display, delete)
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

      // Determine field capabilities for showing appropriate controls
      const isTextOrNumericField = !meta.isBool && !meta.isDate && !meta.isMultiple && !meta.isAttachment &&
        (!meta.choices || meta.choices.length === 0) &&
        (!meta.isRef || meta.refChoices.length === 0);
      const isPureTextField = isTextOrNumericField && !meta.isNumeric && !meta.isInt;

      // Label display
      const labelWrapper = document.createElement('div');
      labelWrapper.className = 'field-label-wrapper';

      const labelSpan = document.createElement('span');
      labelSpan.className = 'field-label-text';
      labelSpan.innerHTML = element.fieldLabel || element.fieldName;

      labelWrapper.appendChild(labelSpan);

      // Display column id
      preview.className = 'element-preview field';
      preview.textContent = `Id colonne : ${element.fieldName}`;

      contentWrapper.appendChild(labelWrapper);
      contentWrapper.appendChild(preview);

      const controls = createControls();

      // Edit label button
      controls.appendChild(createEditButton('Modifier le libellé', () => showEditPopup(element, index, 'fieldLabel')));

      // Required toggle button
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

      // Max length validation button (text and numeric fields only)
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

      // Multiline toggle button (pure text fields only)
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

      // Conditional display button
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
    } else if (element.type === 'title' || element.type === 'text') {
      // Both title and text are rendered the same way
      preview.className = 'element-preview text';
      preview.innerHTML = element.content;
      contentWrapper.appendChild(preview);

      const controls = createControls();
      controls.appendChild(createEditButton('Modifier', () => showEditPopup(element, index)));
      controls.appendChild(createDeleteButton(index));

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    }

    // Drag and drop event handlers
    div.addEventListener('dragstart', function (e) {
      draggedElement = this;
      this.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });

    div.addEventListener('dragend', function () {
      this.classList.remove('dragging');
      draggedElement = null;

      document.querySelectorAll('.element-item').forEach(el => {
        el.classList.remove('drag-over-top', 'drag-over-bottom');
      });
    });

    div.addEventListener('dragover', function (e) {
      if (e.preventDefault) e.preventDefault();
      e.dataTransfer.dropEffect = 'move';

      if (draggedElement && draggedElement !== this) {
        document.querySelectorAll('.element-item').forEach(el => {
          el.classList.remove('drag-over-top', 'drag-over-bottom');
        });

        // Show drop indicator based on mouse position
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
      if (e.target === this) {
        this.classList.remove('drag-over-top', 'drag-over-bottom');
      }
    });

    div.addEventListener('drop', function (e) {
      if (e.stopPropagation) e.stopPropagation();

      document.querySelectorAll('.element-item').forEach(el => {
        el.classList.remove('drag-over-top', 'drag-over-bottom');
      });

      if (draggedElement && draggedElement !== this) {
        const draggedIndex = parseInt(draggedElement.dataset.index);
        const targetIndex = parseInt(this.dataset.index);

        const rect = this.getBoundingClientRect();
        const midpoint = rect.top + rect.height / 2;

        let insertPosition;
        if (e.clientY < midpoint) {
          insertPosition = targetIndex;
        } else {
          insertPosition = targetIndex + 1;
        }

        // Adjust for removal of dragged element
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

// -----------------------------------------------------------------------------
// FORM CONFIGURATION : ADD NEW ELEMENT
// -----------------------------------------------------------------------------

// When user selects an element type in the "Add element" dropdown:
// - "column" → show column picker dropdown
// - "separator"
// - "text" → show text input field

elementType.addEventListener('change', () => {
  // Hide all secondary inputs first
  columnSelect.style.display = 'none';
  elementContent.style.display = 'none';

  if (elementType.value === 'column') {
    updateColumnSelect();
    columnSelect.style.display = 'block';
  } else if (elementType.value === 'separator') {
    // Separator needs no extra input
  } else if (elementType.value) {
    elementContent.style.display = 'block';
    elementContent.placeholder = 'Texte';
  }
});

// When user clicks "Add" button: create the element and add it to form config
addElementBtn.addEventListener('click', () => {
  const type = elementType.value;
  if (!type) return;

  if (type === 'column') {
    const col = columnSelect.value;
    if (!col) {
      alert('Veuillez sélectionner une colonne');
      return;
    }
    // Add field element linked to this column
    formElements.push({
      type: 'field',
      fieldName: col,
      fieldLabel: columnMetadata[col]?.label || col,
      required: false,
      maxLength: null,
      conditional: null
    });
  } else if (type === 'separator') {
    // Add horizontal separator
    formElements.push({ type: 'separator', content: '' });
  } else {
    // Add text block
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

  // Reset the "Add element" panel
  elementType.value = '';
  columnSelect.value = '';
  elementContent.value = '';
  columnSelect.style.display = 'none';
  elementContent.style.display = 'none';
});

// -----------------------------------------------------------------------------
// FORM CONFIGURATION : MODAL CLOSE & STYLE SETTINGS
// -----------------------------------------------------------------------------

// Close modal when clicking X button or outside the modal
closeModal.addEventListener('click', () => configModal.classList.remove('show'));
configModal.addEventListener('click', (e) => {
  if (e.target === configModal) configModal.classList.remove('show');
});

// Font selection: apply to form and save
fontSelect.addEventListener('change', () => {
  globalFont = fontSelect.value;
  applyGlobalStyles();
  saveConfiguration();
});

// Padding selection: apply to form and save
paddingSelect.addEventListener('change', () => {
  globalPadding = paddingSelect.value;
  applyGlobalStyles();
  saveConfiguration();
});

// -----------------------------------------------------------------------------
// FORM INPUT CREATION
// -----------------------------------------------------------------------------

// Create appropriate input element based on column type
// Returns different HTML elements depending on the column's Grist type:
// - Attachments: file upload with drag-drop
// - Bool: checkbox
// - Date: date picker
// - Numeric/Int: text input (allows comma as decimal separator)
// - ChoiceList/RefList: multi-select dropdown
// - Choice/Ref: single-select dropdown
// - Text: input or textarea (if multiline enabled)
function createInputForColumn(col, meta, element = {}) {
  let inp;

  // Attachment columns: custom file upload UI
  if (meta.isAttachment) {
    pendingAttachments[col] = [];

    const container = document.createElement('div');
    container.className = 'attachment-input';
    container.id = `input_${col}`;

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.multiple = true;
    fileInput.style.display = 'none';
    fileInput.id = `file_${col}`;

    const uploadBtn = document.createElement('button');
    uploadBtn.type = 'button';
    uploadBtn.className = 'attachment-btn';
    uploadBtn.textContent = 'Charger une PJ';

    const fileList = document.createElement('div');
    fileList.className = 'attachment-list';
    fileList.id = `files_${col}`;

    uploadBtn.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', () => {
      const files = Array.from(fileInput.files);
      const maxSize = 10 * 1024 * 1024; // 10 MB
      const maxFiles = 3;

      files.forEach(file => {
        if (pendingAttachments[col].length >= maxFiles) {
          const errorDiv = document.getElementById(`error_${col}`);
          if (errorDiv) {
            errorDiv.textContent = `Maximum ${maxFiles} pièces jointes par champ`;
            errorDiv.classList.add('show');
            setTimeout(() => errorDiv.classList.remove('show'), 3000);
          }
          return;
        }

        if (file.size > maxSize) {
          const errorDiv = document.getElementById(`error_${col}`);
          if (errorDiv) {
            errorDiv.textContent = `Le fichier "${file.name}" dépasse 10 Mo`;
            errorDiv.classList.add('show');
            setTimeout(() => errorDiv.classList.remove('show'), 3000);
          }
          return;
        }

        // Prevent duplicates
        if (pendingAttachments[col].some(f => f.name === file.name && f.size === file.size)) {
          return;
        }

        pendingAttachments[col].push(file);
        renderAttachmentList(col, fileList);
      });

      // Reset to allow re-selecting same file
      fileInput.value = '';
    });

    container.appendChild(fileInput);
    container.appendChild(uploadBtn);
    container.appendChild(fileList);

    return container;
  }

  // Boolean columns: simple checkbox
  if (meta.isBool) {
    inp = document.createElement('input');
    inp.type = 'checkbox';
    inp.id = `input_${col}`;
    return inp;
  }

  // Date columns: native date picker
  if (meta.isDate) {
    inp = document.createElement('input');
    inp.type = 'date';
    inp.id = `input_${col}`;
    return inp;
  }

  // Numeric columns: text input to allow comma as decimal separator (French format)
  if (meta.isNumeric || meta.isInt) {
    inp = document.createElement('input');
    inp.type = 'text';
    inp.id = `input_${col}`;
    return inp;
  }

  // ChoiceList or RefList columns: multi-select dropdown
  if (meta.isMultiple) {
    const sel = document.createElement('select');
    sel.multiple = true;
    sel.id = `input_${col}`;
    // Use refChoices for RefList, or map choices for ChoiceList
    const opts = meta.refChoices.length > 0 ? meta.refChoices : (meta.choices || []).map(c => ({ id: c, label: c }));
    opts.forEach(opt => {
      const o = document.createElement('option');
      o.value = opt.id;
      o.textContent = opt.label;
      sel.appendChild(o);
    });

    // Allow toggle selection with single click
    sel.addEventListener('mousedown', function (e) {
      e.preventDefault();
      const option = e.target;
      if (option.tagName === 'OPTION') {
        option.selected = !option.selected;
        sel.focus();
        sel.dispatchEvent(new Event('change'));
      }
    });

    return sel;
  }

  // Choice or Ref columns: single-select dropdown
  if ((meta.choices && meta.choices.length > 0) || (meta.isRef && meta.refChoices.length > 0)) {
    inp = document.createElement('select');
    inp.id = `input_${col}`;
    // Add empty placeholder option
    const empty = document.createElement('option');
    empty.value = '';
    empty.textContent = '-- Sélectionner --';
    inp.appendChild(empty);
    // Use refChoices for Ref, or map choices for Choice
    const opts = meta.refChoices.length > 0 ? meta.refChoices : meta.choices.map(c => ({ id: c, label: c }));
    opts.forEach(opt => {
      const o = document.createElement('option');
      o.value = opt.id;
      o.textContent = opt.label;
      inp.appendChild(o);
    });
    return inp;
  }

  // Text field: textarea if multiline enabled, otherwise simple input
  if (element.multiline) {
    inp = document.createElement('textarea');
    inp.id = `input_${col}`;
    inp.rows = 4;
    return inp;
  }

  // Default: single-line text input
  inp = document.createElement('input');
  inp.type = 'text';
  inp.id = `input_${col}`;
  return inp;
}

// -----------------------------------------------------------------------------
// FORM VALUE HANDLING
// -----------------------------------------------------------------------------

// Sanitize text input: trim, remove control chars, limit length
function sanitizeText(value) {
  if (typeof value !== 'string') return value;
  return value
    .trim()
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
    .substring(0, 50000);
}

// Get value from form input, converted to appropriate Grist type
// Handles type conversion for each column type before sending to Grist API
function getInputValue(col, meta) {
  const inp = document.getElementById(`input_${col}`);
  if (!inp) return null;

  // Boolean: return checkbox state
  if (meta.isBool) return inp.checked;

  // Attachments are handled separately via uploadAttachments
  if (meta.isAttachment) return null;

  // ChoiceList/RefList: return Grist list format ['L', val1, val2, ...]
  if (meta.isMultiple) {
    const selected = Array.from(inp.selectedOptions).map(opt => opt.value);
    const values = meta.isRef ? selected.map(v => parseInt(v)) : selected;
    return ["L", ...values];
  }

  // Ref: return integer ID or null
  if (meta.isRef) return inp.value ? parseInt(inp.value) : null;

  // Numeric: parse float, accepting comma as decimal separator
  if (meta.isNumeric || meta.isInt) {
    const val = sanitizeText(inp.value);
    return val ? parseFloat(val.replace(',', '.')) : null;
  }

  // Default (text): return sanitized string
  return sanitizeText(inp.value);
}

// -----------------------------------------------------------------------------
// FORM VALIDATION
// -----------------------------------------------------------------------------

// Validate a single field, show error message if invalid
// Checks: required fields, numeric format, integer format, max length
// Returns true if valid, false if invalid
function validateField(col, meta, element) {
  const inp = document.getElementById(`input_${col}`);
  const err = document.getElementById(`error_${col}`);

  // Clear previous error state
  inp.classList.remove('error');
  if (err) err.classList.remove('show');

  // Required field validation
  if (element.required) {
    if (meta.isBool) {
      if (!inp.checked) {
        inp.classList.add('error');
        if (err) {
          err.textContent = 'Ce champ doit être coché';
          err.classList.add('show');
        }
        return false;
      }
    } else if (meta.isAttachment) {
      if (!pendingAttachments[col] || pendingAttachments[col].length === 0) {
        inp.classList.add('error');
        if (err) {
          err.textContent = 'Au moins une pièce jointe est requise';
          err.classList.add('show');
        }
        return false;
      }
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

  // Numeric validation
  if (meta.isNumeric || meta.isInt) {
    const val = inp.value.trim();
    if (val === '') return true;

    const normalizedVal = val.replace(',', '.');
    const num = parseFloat(normalizedVal);
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

  // Max length validation
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

// -----------------------------------------------------------------------------
// CONDITIONAL FIELD DISPLAY (only available for column types Choice and Ref)
// -----------------------------------------------------------------------------

// Determines if a field should be visible based on its conditional rule
// A conditional rule is: "show this field if [otherField] [equals/notEquals] [value]"
// Returns true if: no condition set, or condition is satisfied
function shouldShowField(element) {
  if (!element.conditional) return true;

  const conditionalField = element.conditional.field;    // Field we depend on (eg "Status")
  const conditionalValue = element.conditional.value;    // Expected value (eg "Active" or ref ID)
  const conditionalOperator = element.conditional.operator; // "equal to" or "not equal to"

  // Get the current value of the field we depend on
  const inp = document.getElementById(`input_${conditionalField}`);
  if (!inp) return true;

  const meta = columnMetadata[conditionalField];
  if (!meta) return true;

  // For Ref columns: compare numeric IDs (not display labels)
  // HTML select values are strings, so we convert to int for proper comparison
  let currentValue;
  if (meta.isRef) {
    currentValue = inp.value ? parseInt(inp.value) : null;
  } else {
    currentValue = inp.value;
  }

  const expectedValue = meta.isRef ? parseInt(conditionalValue) : conditionalValue;

  // Evaluate the condition
  if (conditionalOperator === 'notEquals') {
    return currentValue != expectedValue;
  }
  return currentValue == expectedValue;
}

// Re-evaluate all conditional rules and update field visibility
// Called whenever a field value changes
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

// -----------------------------------------------------------------------------
// FORM RENDERING
// -----------------------------------------------------------------------------

// Render the actual form based on configuration
// Creates form elements (fields, separators, text blocks) in the DOM
function renderForm() {
  fieldsContainer.innerHTML = '';

  formElements.forEach(element => {
    if (element.type === 'separator') {
      const sep = document.createElement('div');
      sep.className = 'separator';
      fieldsContainer.appendChild(sep);
    } else if (element.type === 'title') {
      // Legacy support for 'title' type
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

      if (!shouldShowField(element)) {
        fieldDiv.classList.add('hidden');
      }

      const label = document.createElement('label');
      label.innerHTML = element.fieldLabel || col;
      if (element.required) {
        label.innerHTML += ' <span class="required-star">*</span>';
      }

      const inp = createInputForColumn(col, meta, element);

      // Update conditional fields when value changes
      inp.addEventListener('change', () => {
        updateConditionalFields();
      });

      if (meta.isBool) {
        // Checkbox layout: input first, then label
        fieldDiv.appendChild(inp);
        fieldDiv.appendChild(label);
        const err = document.createElement('div');
        err.className = 'error-message';
        err.id = `error_${col}`;
        fieldDiv.appendChild(err);
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

// Re-render form when records change (to update ref choices)
grist.onRecords(() => {
  renderForm();
});

// -----------------------------------------------------------------------------
// FORM SUBMISSION
// -----------------------------------------------------------------------------

// Handle form submission: validate, upload attachments, create record, reset form
addButton.addEventListener('click', async () => {
  let valid = true;

  // Hide any previous error/success messages
  formError.classList.remove('show');
  formSuccess.classList.remove('show');

  // Validate all visible fields (hidden conditional fields are skipped)
  formElements.forEach(element => {
    if (element.type === 'field') {
      const col = element.fieldName;
      const meta = columnMetadata[col] || {};

      if (shouldShowField(element)) {
        if (!validateField(col, meta, element)) {
          valid = false;
        }
      }
    }
  });

  if (!valid) {
    formError.textContent = 'Il y a une ou plusieurs erreurs dans le formulaire, veuillez vérifier les champs';
    formError.classList.add('show');
    return;
  }

  // Collect field values (except attachments)
  const fields = {};
  formElements.forEach(element => {
    if (element.type === 'field') {
      const col = element.fieldName;
      const meta = columnMetadata[col] || {};

      if (shouldShowField(element) && !meta.isAttachment) {
        fields[col] = getInputValue(col, meta);
      }
    }
  });

  try {
    // Upload attachments first
    for (const element of formElements) {
      if (element.type === 'field') {
        const col = element.fieldName;
        const meta = columnMetadata[col] || {};

        if (meta.isAttachment && shouldShowField(element)) {
          const attachmentValue = await uploadAttachments(col);
          if (attachmentValue) {
            fields[col] = attachmentValue;
          }
        }
      }
    }

    // Create new record
    await grist.selectedTable.create({ fields });

    // Show success message
    formSuccess.classList.add('show');
    setTimeout(() => {
      formSuccess.classList.remove('show');
    }, 3000);

    // Reset form: clear all inputs based on their type
    formElements.forEach(element => {
      if (element.type === 'field') {
        const col = element.fieldName;
        const inp = document.getElementById(`input_${col}`);
        const meta = columnMetadata[col] || {};

        if (meta.isBool) {
          inp.checked = false;
        } else if (meta.isAttachment) {
          // Clear pending files and file list display
          pendingAttachments[col] = [];
          const fileList = document.getElementById(`files_${col}`);
          if (fileList) fileList.innerHTML = '';
        } else if (meta.isMultiple) {
          // Deselect all options in multi-select
          Array.from(inp.options).forEach(opt => opt.selected = false);
        } else {
          inp.value = '';
        }
      }
    });

    // Re-evaluate conditional fields after reset
    updateConditionalFields();

  } catch (error) {
    formError.textContent = "Erreur: " + error.message;
    formError.classList.add('show');
  }
});
