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

let columns = [];
let columnMetadata = {};
let formElements = [];
let draggedElement = null;
let initInProgress = false;


grist.ready({
  requiredAccess: 'full',
  onEditOptions: () => configModal.classList.add('show')
});

(async () => {
  console.group('🚀 INIT METADATA');

  columns = await getAllColumnsFromMetadata();
  console.log('✅ Colonnes metadata:', columns);

  columnMetadata = await getColumnMetadata();
  console.log('📘 columnMetadata:', columnMetadata);

  await loadConfiguration();

  console.groupEnd();
})();





async function getAllColumnsFromMetadata() {
  try {
    // Utiliser fetchSelectedTable qui respecte les permissions de table
    const records = await grist.fetchSelectedTable();
    console.log('📊 fetchSelectedTable OK, nb records:', records.length);

    if (!records || records.length === 0) {
      console.warn('⚠️ Aucun record, impossible d\'inférer les colonnes');
      return [];
    }

    // Extraire les colonnes depuis le premier record
    const cols = Object.keys(records[0]).filter(colId =>
      colId !== 'id' &&
      colId !== 'manualSort' &&
      !colId.startsWith('gristHelper')
    );

    console.log('✅ Colonnes détectées depuis données:', cols);
    return cols;
  } catch (error) {
    console.error('❌ Erreur fetchSelectedTable:', error);
    return [];
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

    formElements = columns.map(col => ({
      type: 'field',
      fieldName: col,
      fieldLabel: col,
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
    formElements
  });
}


function updateColumnSelect() {
  const usedColumns = formElements
    .filter(el => el.type === 'field')
    .map(el => el.fieldName);

  const availableColumns = columns.filter(
    col => !usedColumns.includes(col)
  );

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

function showEditPopup(element, index) {
  const overlay = document.getElementById('popupOverlay');
  overlay.classList.add('show');

  const popup = document.createElement('div');
  popup.className = 'edit-popup';
  popup.innerHTML = `
    <h3>Modifier le contenu</h3>
    <textarea id="editContent">${element.content || ''}</textarea>
    <div class="edit-popup-buttons">
      <button class="cancel">Annuler</button>
      <button class="save">Enregistrer</button>
    </div>
  `;

  document.body.appendChild(popup);

  const textarea = popup.querySelector('#editContent');
  textarea.focus();
  textarea.setSelectionRange(textarea.value.length, textarea.value.length);

  popup.querySelector('.cancel').addEventListener('click', () => {
    popup.remove();
    overlay.classList.remove('show');
  });

  popup.querySelector('.save').addEventListener('click', () => {
    const newContent = document.getElementById('editContent').value.trim();
    if (newContent) {
      element.content = newContent;
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
      const isTextField = !meta.isBool && !meta.isDate && !meta.isMultiple &&
        (!meta.choices || meta.choices.length === 0) &&
        (!meta.isRef || meta.refChoices.length === 0);

      const labelInput = document.createElement('input');
      labelInput.type = 'text';
      labelInput.className = 'field-label-input';
      labelInput.value = element.fieldLabel || element.fieldName;
      labelInput.onchange = (e) => {
        element.fieldLabel = e.target.value;
        saveConfiguration();
        renderForm();
      };

      labelInput.addEventListener('mousedown', e => {
        e.stopPropagation();
      });

      labelInput.addEventListener('click', e => {
        e.stopPropagation();
      });

      labelInput.addEventListener('focus', () => {
        div.draggable = false;
      });

      labelInput.addEventListener('blur', () => {
        div.draggable = true;
      });


      preview.className = 'element-preview field';
      preview.textContent = `Id colonne : ${element.fieldName}`;

      contentWrapper.appendChild(labelInput);
      contentWrapper.appendChild(preview);

      const controls = document.createElement('div');
      controls.className = 'element-controls';

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

      // Icône validation (texte uniquement)
      if (isTextField) {
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

      // Bouton supprimer
      const deleteBtn = document.createElement('button');
      deleteBtn.textContent = '🗑️';
      deleteBtn.onclick = () => {
        formElements.splice(index, 1);
        saveConfiguration();
        renderConfigList();
        renderForm();
        updateColumnSelect();
      };
      controls.appendChild(deleteBtn);

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'separator') {
      preview.className = 'element-preview separator';
      contentWrapper.appendChild(preview);

      const controls = document.createElement('div');
      controls.className = 'element-controls';

      const deleteBtn = document.createElement('button');
      deleteBtn.textContent = '🗑️';
      deleteBtn.onclick = () => {
        formElements.splice(index, 1);
        saveConfiguration();
        renderConfigList();
        renderForm();
        updateColumnSelect();
      };
      controls.appendChild(deleteBtn);

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'title') {
      preview.className = 'element-preview title';
      preview.innerHTML = `
        ${element.content}
        <span class="edit-icon" onclick="event.stopPropagation();">✎</span>
      `;
      preview.querySelector('.edit-icon').onclick = (e) => {
        e.stopPropagation();
        showEditPopup(element, index);
      };
      contentWrapper.appendChild(preview);

      const controls = document.createElement('div');
      controls.className = 'element-controls';

      const deleteBtn = document.createElement('button');
      deleteBtn.textContent = '🗑️';
      deleteBtn.onclick = () => {
        formElements.splice(index, 1);
        saveConfiguration();
        renderConfigList();
        renderForm();
        updateColumnSelect();
      };
      controls.appendChild(deleteBtn);

      div.appendChild(contentWrapper);
      div.appendChild(controls);
    } else if (element.type === 'text') {
      preview.className = 'element-preview text';
      preview.innerHTML = `
        ${element.content}
        <span class="edit-icon" onclick="event.stopPropagation();">✎</span>
      `;
      preview.querySelector('.edit-icon').onclick = (e) => {
        e.stopPropagation();
        showEditPopup(element, index);
      };
      contentWrapper.appendChild(preview);

      const controls = document.createElement('div');
      controls.className = 'element-controls';

      const deleteBtn = document.createElement('button');
      deleteBtn.textContent = '🗑️';
      deleteBtn.onclick = () => {
        formElements.splice(index, 1);
        saveConfiguration();
        renderConfigList();
        renderForm();
        updateColumnSelect();
      };
      controls.appendChild(deleteBtn);

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
    elementContent.placeholder = elementType.value === 'title' ? 'Titre' : 'Texte';
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
      fieldLabel: col,
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

async function getColumnMetadata() {
  try {
    // Utiliser l'API REST qui respecte les permissions de table
    const tableId = await grist.selectedTable.getTableId();
    console.log('📍 TableId pour metadata:', tableId);

    const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
    console.log('🔑 Token obtenu pour metadata');

    const response = await fetch(`${tokenInfo.baseUrl}/tables/${tableId}/columns`, {
      headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.status}`);
    }

    const data = await response.json();
    console.log('📡 API columns metadata response:', data);

    const metadata = {};

    for (const col of data.columns) {
      const colId = col.id;
      if (colId === 'id' || colId === 'manualSort' || colId.startsWith('gristHelper')) {
        continue;
      }

      const type = col.fields.type || 'Text';
      let choices = null;
      let refTable = null;
      let refChoices = [];

      // Parser widgetOptions si présent
      if (col.fields.widgetOptions) {
        try {
          const options = typeof col.fields.widgetOptions === 'string'
            ? JSON.parse(col.fields.widgetOptions)
            : col.fields.widgetOptions;
          if (options.choices) choices = options.choices;
        } catch (e) { }
      }

      // Déterminer la table référencée
      if (type.startsWith('Ref:')) {
        refTable = type.substring(4);
      } else if (type.startsWith('RefList:')) {
        refTable = type.substring(8);
      }

      // Récupérer les choix de la table référencée via API REST
      if (refTable) {
        try {
          const refResponse = await fetch(`${tokenInfo.baseUrl}/tables/${refTable}/records`, {
            headers: { 'Authorization': `Bearer ${tokenInfo.token}` }
          });
          if (refResponse.ok) {
            const refData = await refResponse.json();
            refChoices = refData.records.map(record => {
              // Trouver la première colonne non-id pour le label
              const labelKey = Object.keys(record.fields).find(k => k !== 'id' && k !== 'manualSort');
              return {
                id: record.id,
                label: labelKey ? record.fields[labelKey] : record.id
              };
            });
          }
        } catch (e) {
          console.warn(`⚠️ Impossible de charger les refs pour ${refTable}:`, e);
        }
      }

      metadata[colId] = {
        type,
        choices,
        isMultiple: type === 'ChoiceList' || type.startsWith('RefList:'),
        isRef: type.startsWith('Ref:') || type.startsWith('RefList:'),
        refTable,
        refChoices,
        isBool: type === 'Bool',
        isDate: type === 'Date' || type === 'DateTime',
        isNumeric: type === 'Numeric',
        isInt: type === 'Int'
      };
    }

    return metadata;
  } catch (error) {
    console.error("❌ Erreur API REST metadata:", error);
    // Fallback: inférer depuis les données
    return getColumnMetadataFromData();
  }
}

// Fallback: inférer les métadonnées depuis les données
async function getColumnMetadataFromData() {
  console.warn('⚠️ Fallback: inférence metadata depuis données');
  try {
    const records = await grist.fetchSelectedTable();
    if (!records || records.length === 0) {
      return {};
    }

    const metadata = {};
    const sample = records[0];

    for (const colId of Object.keys(sample)) {
      if (colId === 'id' || colId === 'manualSort' || colId.startsWith('gristHelper')) {
        continue;
      }

      const value = sample[colId];
      let type = 'Text';
      let isBool = false;
      let isDate = false;
      let isNumeric = false;
      let isInt = false;
      let isMultiple = false;
      let isRef = false;

      // Inférer le type depuis la valeur
      if (typeof value === 'boolean') {
        type = 'Bool';
        isBool = true;
      } else if (Array.isArray(value) && value[0] === 'L') {
        // RefList ou ChoiceList
        type = 'ChoiceList';
        isMultiple = true;
      } else if (typeof value === 'number') {
        type = 'Numeric';
        isNumeric = true;
        if (Number.isInteger(value)) {
          type = 'Int';
          isInt = true;
        }
      }

      metadata[colId] = {
        type,
        choices: null,
        isMultiple,
        isRef,
        refTable: null,
        refChoices: [],
        isBool,
        isDate,
        isNumeric,
        isInt
      };
    }

    console.log('✅ Metadata inférée:', metadata);
    return metadata;
  } catch (error) {
    console.error("❌ Erreur inférence metadata:", error);
    return {};
  }
}

function createInputForColumn(col, meta) {
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
      title.textContent = element.content;
      fieldsContainer.appendChild(title);
    } else if (element.type === 'text') {
      const text = document.createElement('div');
      text.className = 'custom-text';
      text.textContent = element.content;
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
      label.textContent = element.fieldLabel || col;
      if (element.required) {
        label.textContent += ' *';
      }

      const inp = createInputForColumn(col, meta);

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