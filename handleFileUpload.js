// Global memory state for immediate rendering
let loadedDirectoryRecords = [];
let loadedSpecialRecords = [];
let customTemplates = [];
let activeClientIndex = -1;

const predefinedTemplates = {
  'edge-velo': [],
  'fortigate': [],
  'general': [],
  'meraki': [],
  'custom': []
};

// Default Initial Record for Blue Beacon Special Customers
const defaultBlueBeaconCustomer = [
  {
    locNum: "BB-001",
    locName: "Blue Beacon Corporate HQ",
    address: "500 E 8th St, Salina, KS 67401",
    contactGm: "Kevin GM",
    contactOm: "Joyce OM",
    phone: "(800) 733-8439",
    mobile: "(314) 748-5622",
    connType: "Coax / Fiber",
    carrier: "Cox / DSR / AT&T",
    speed: "100 Mbps / 1 Gbps"
  }
];

// Status Cycling Configuration
const statusCycle = ['dsr', 'green', 'cox', 'att'];
const statusClassMap = {
  'dsr': 'status-dsr',
  'green': 'status-green',
  'cox': 'status-cox',
  'att': 'status-att'
};

document.addEventListener("DOMContentLoaded", function() {
  loadDirectoryFromLocalStorage();
  loadSpecialFromLocalStorage();
  loadCustomTemplates();
  bindLocCellEvents();
});

// Helper Function: HTML Sanitization
function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// Navigation Tabs Switcher
function switchTab(tabId) {
  ['directory', 'special', 'templates', 'runbooks'].forEach(id => {
    document.getElementById(`tab-${id}`)?.classList.add('hidden');
    const btn = document.getElementById(`btn-${id}`);
    if (btn) {
      btn.classList.remove('bg-slate-800', 'text-blue-400', 'text-amber-400', 'text-white');
      btn.classList.add('text-slate-400');
    }
  });
  
  const targetTab = document.getElementById(`tab-${tabId}`);
  const activeBtn = document.getElementById(`btn-${tabId}`);
  if (targetTab) targetTab.classList.remove('hidden');
  
  if (activeBtn) {
    activeBtn.classList.remove('text-slate-400');
    activeBtn.classList.add('bg-slate-800');
    activeBtn.classList.add(tabId === 'special' ? 'text-amber-400' : 'text-blue-400');
  }
}

// Search Filter
function filterDirectory(inputId, tableId) {
  const input = document.getElementById(inputId).value.toLowerCase();
  const rows = document.querySelectorAll(`#${tableId} tbody tr`);
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(input) ? "" : "none";
  });
}

// Clipboard Helpers
function copyToClipboard(elementId) {
  const elem = document.getElementById(elementId);
  if (!elem) return;
  const text = elem.innerText;
  navigator.clipboard.writeText(text).then(() => {
    alert("Text copied to clipboard!");
  });
}

function copyTextDirectly(text) {
  navigator.clipboard.writeText(text).then(() => {
    alert("Link copied to clipboard!");
  });
}

// Case & Space Insensitive Field Mapper
function getFieldValue(row, possibleKeys) {
  if (!row) return "-";

  const normalizedRow = {};
  for (let k in row) {
    if (Object.prototype.hasOwnProperty.call(row, k)) {
      normalizedRow[k.trim().toLowerCase()] = row[k];
    }
  }

  for (let key of possibleKeys) {
    const targetKey = key.trim().toLowerCase();
    const val = normalizedRow[targetKey];
    if (val !== undefined && val !== null && String(val).trim() !== "" && String(val).trim() !== "undefined") {
      return String(val).trim();
    }
  }

  return "-";
}

/* -------------------------------------------------------------
   PRIMARY CONTACT DIRECTORY MANAGEMENT
   ------------------------------------------------------------- */

function handleFileUploadDirectory(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (typeof XLSX === "undefined") {
    alert("Error: SheetJS (XLSX) library is not loaded in your HTML head!");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });

      if (!jsonData || jsonData.length === 0) {
        alert("The uploaded file contains no data.");
        return;
      }

      loadedDirectoryRecords = jsonData.map(row => ({
        id: getFieldValue(row, ["ACCOUNT ID", "Account ID", "Account_ID", "Account", "ID"]),
        name: getFieldValue(row, ["ACCOUNT NAME", "Account Name", "Account_Name", "Company", "Name"]),
        contactName: getFieldValue(row, ["PRIMARY CONTACT", "Primary Contact", "Contact Name", "Contact", "Name", "GM"]),
        contactPhone: getFieldValue(row, ["PHONE NUMBER", "Phone Number", "Phone", "Mobile", "Cell", "Contact Phone"]),
        contactEmail: getFieldValue(row, ["EMAIL ADDRESS", "Email Address", "Email", "Contact Email", "Primary Email"]),
        notes: getFieldValue(row, ["SPECIAL HANDLING NOTES", "Special Handling Notes", "SpecialHandlingNotes", "Notes", "Handling Notes"])
      }));

      try {
        localStorage.setItem('ops_directory_data', JSON.stringify(loadedDirectoryRecords));
      } catch (err) {
        console.warn("LocalStorage quota exceeded: Retained in active memory.");
      }

      renderDirectoryRows(loadedDirectoryRecords);
      alert(`Successfully imported ${loadedDirectoryRecords.length} client records!`);
    } catch (err) {
      console.error("Error processing file:", err);
      alert("Error reading file. Check browser console (F12).");
    } finally {
      event.target.value = '';
    }
  };

  reader.readAsArrayBuffer(file);
}

function renderDirectoryRows(records) {
  const tbody = document.querySelector("#directoryTable tbody");
  if (!tbody) return;
  
  tbody.innerHTML = "";
  const dataToRender = records && records.length > 0 ? records : loadedDirectoryRecords;

  if (!dataToRender || dataToRender.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="p-4 text-center text-slate-400">No records found. Click "+ Add Client" or import an Excel file.</td></tr>`;
    return;
  }

  const fragment = document.createDocumentFragment();

  dataToRender.forEach((rec, index) => {
    const tr = document.createElement("tr");
    tr.className = "hover:bg-slate-50 transition border-b border-slate-100";
    tr.innerHTML = `
      <td class="p-4 font-mono font-semibold">
        <button onclick="openTemplateModal(${index})" class="text-blue-600 hover:text-blue-800 hover:underline focus:outline-none flex items-center gap-1 cursor-pointer">
          <span>${escapeHtml(rec.id)}</span>
          <span class="text-[10px] bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded font-sans">Send Email</span>
        </button>
      </td>
      <td class="p-4 font-semibold text-slate-800">${escapeHtml(rec.name)}</td>
      <td class="p-4 font-medium text-slate-700">${escapeHtml(rec.contactName || '-')}</td>
      <td class="p-4 font-mono text-xs text-slate-700">${escapeHtml(rec.contactPhone || '-')}</td>
      <td class="p-4 font-mono text-xs text-blue-600">${escapeHtml(rec.contactEmail || '-')}</td>
      <td class="p-4 text-slate-600 text-xs">${escapeHtml(rec.notes || '-')}</td>
      <td class="p-4 text-center space-x-2">
        <button onclick="openClientModal(${index})" class="text-xs text-blue-600 hover:text-blue-800 font-medium cursor-pointer">Edit</button>
        <button onclick="deleteClient(${index})" class="text-xs text-rose-600 hover:text-rose-800 font-medium cursor-pointer">Delete</button>
      </td>
    `;
    fragment.appendChild(tr);
  });

  tbody.appendChild(fragment);
}

function loadDirectoryFromLocalStorage() {
  const savedData = localStorage.getItem('ops_directory_data');
  if (savedData) {
    try {
      loadedDirectoryRecords = JSON.parse(savedData);
      renderDirectoryRows(loadedDirectoryRecords);
    } catch (e) {
      renderDirectoryRows([]);
    }
  } else {
    renderDirectoryRows([]);
  }
}

function openClientModal(index = -1) {
  const modal = document.getElementById("clientModal");
  const modalTitle = document.getElementById("modalTitle");
  const editIndex = document.getElementById("editIndex");

  if (!modal) return;
  editIndex.value = index;

  if (index >= 0) {
    const rec = loadedDirectoryRecords[index];
    if (rec) {
      if (modalTitle) modalTitle.innerText = "Edit Client & Primary Contact";
      document.getElementById("inputAccountId").value = rec.id || "";
      document.getElementById("inputAccountName").value = rec.name || "";
      document.getElementById("inputContactName").value = rec.contactName || "";
      document.getElementById("inputContactPhone").value = rec.contactPhone || "";
      document.getElementById("inputContactEmail").value = rec.contactEmail || "";
      document.getElementById("inputNotes").value = rec.notes || "";
    }
  } else {
    if (modalTitle) modalTitle.innerText = "Add Client & Primary Contact";
    const form = document.getElementById("clientForm");
    if (form) form.reset();
  }

  modal.classList.remove("hidden");
}

function closeClientModal() {
  const modal = document.getElementById("clientModal");
  if (modal) modal.classList.add("hidden");
}

function saveClient(event) {
  event.preventDefault();

  const editIndex = parseInt(document.getElementById("editIndex").value);
  const newRecord = {
    id: document.getElementById("inputAccountId").value.trim(),
    name: document.getElementById("inputAccountName").value.trim(),
    contactName: document.getElementById("inputContactName").value.trim(),
    contactPhone: document.getElementById("inputContactPhone").value.trim(),
    contactEmail: document.getElementById("inputContactEmail").value.trim(),
    notes: document.getElementById("inputNotes").value.trim()
  };

  if (editIndex >= 0) {
    loadedDirectoryRecords[editIndex] = newRecord;
  } else {
    loadedDirectoryRecords.push(newRecord);
  }

  try {
    localStorage.setItem('ops_directory_data', JSON.stringify(loadedDirectoryRecords));
  } catch (err) {
    console.warn("LocalStorage full, saved to session memory.");
  }

  renderDirectoryRows(loadedDirectoryRecords);
  closeClientModal();
}

function deleteClient(index) {
  if (confirm("Are you sure you want to delete this client record?")) {
    loadedDirectoryRecords.splice(index, 1);
    try {
      localStorage.setItem('ops_directory_data', JSON.stringify(loadedDirectoryRecords));
    } catch (err) {}
    renderDirectoryRows(loadedDirectoryRecords);
  }
}

/* -------------------------------------------------------------
   SPECIAL CUSTOMERS IMPORTER & RENDERER (BLUE BEACON)
   ------------------------------------------------------------- */

function handleFileUploadSpecial(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (typeof XLSX === "undefined") {
    alert("Error: SheetJS (XLSX) library is not loaded in your HTML head!");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });

      if (!jsonData || jsonData.length === 0) {
        alert("The uploaded file contains no data.");
        return;
      }

      loadedSpecialRecords = jsonData.map(row => ({
        locNum: getFieldValue(row, ["Loc #", "Loc#", "Loc Number", "Location #", "ID"]),
        locName: getFieldValue(row, ["Location Name", "Location", "Account Name", "Name"]),
        address: getFieldValue(row, ["Address", "Street Address", "Location Address"]),
        contactGm: getFieldValue(row, ["Contact Name (GM)", "GM Contact", "GM Name", "GM"]),
        contactOm: getFieldValue(row, ["Contact Name (OM)", "OM Contact", "OM Name", "OM"]),
        phone: getFieldValue(row, ["Phone", "Phone Number", "Telephone"]),
        mobile: getFieldValue(row, ["Mobile Number", "Mobile", "Cell", "Cell Phone"]),
        connType: getFieldValue(row, ["Connection Type(Coax, Fiber)", "Connection Type", "Conn Type", "Type"]),
        carrier: getFieldValue(row, ["Carrier Name", "Carrier", "Provider", "Circuit 1 Provider"]),
        speed: getFieldValue(row, ["Connection Speed", "Speed", "Circuit Speed", "Bandwidth"])
      }));

      try {
        localStorage.setItem('ops_special_data', JSON.stringify(loadedSpecialRecords));
      } catch (err) {
        console.warn("LocalStorage full, saved to session memory.");
      }

      renderSpecialRows(loadedSpecialRecords);
      alert(`Successfully imported ${loadedSpecialRecords.length} location records!`);
    } catch (err) {
      console.error("Error processing VIP file:", err);
      alert("Error reading file. Check browser console (F12).");
    } finally {
      event.target.value = '';
    }
  };

  reader.readAsArrayBuffer(file);
}

function renderSpecialRows(records) {
  const tbody = document.querySelector("#specialTable tbody");
  if (!tbody) return;

  tbody.innerHTML = "";
  const dataToRender = records && records.length > 0 ? records : loadedSpecialRecords;

  if (!dataToRender || dataToRender.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" class="p-4 text-center text-slate-400">No special customer records found.</td></tr>`;
    return;
  }

  const fragment = document.createDocumentFragment();

  dataToRender.forEach(rec => {
    const tr = document.createElement("tr");
    tr.className = "hover:bg-amber-50/50 transition border-b border-slate-100";
    tr.innerHTML = `
      <td class="p-3 font-mono text-amber-600 font-semibold loc-cell status-att" data-status="att">${escapeHtml(rec.locNum)}</td>
      <td class="p-3 font-semibold text-slate-800">${escapeHtml(rec.locName)}</td>
      <td class="p-3 text-slate-600 text-xs">${escapeHtml(rec.address)}</td>
      <td class="p-3 text-slate-700">${escapeHtml(rec.contactGm)}</td>
      <td class="p-3 text-slate-700">${escapeHtml(rec.contactOm)}</td>
      <td class="p-3 font-mono text-xs text-slate-700">${escapeHtml(rec.phone)}</td>
      <td class="p-3 font-mono text-xs text-slate-700">${escapeHtml(rec.mobile)}</td>
      <td class="p-3 text-xs text-slate-600">${escapeHtml(rec.connType)}</td>
      <td class="p-3 text-xs font-medium text-slate-800">${escapeHtml(rec.carrier)}</td>
      <td class="p-3 font-mono text-xs text-slate-600">${escapeHtml(rec.speed)}</td>
    `;
    fragment.appendChild(tr);
  });

  tbody.appendChild(fragment);
}

function loadSpecialFromLocalStorage() {
  const savedData = localStorage.getItem('ops_special_data');
  if (savedData) {
    try {
      loadedSpecialRecords = JSON.parse(savedData);
      renderSpecialRows(loadedSpecialRecords);
    } catch (e) {
      loadedSpecialRecords = defaultBlueBeaconCustomer;
      renderSpecialRows(loadedSpecialRecords);
    }
  } else {
    loadedSpecialRecords = defaultBlueBeaconCustomer;
    localStorage.setItem('ops_special_data', JSON.stringify(defaultBlueBeaconCustomer));
    renderSpecialRows(defaultBlueBeaconCustomer);
  }
}

function clearImportedData(storageKey, tableId) {
  if (confirm("Are you sure you want to clear all imported data for this hub?")) {
    localStorage.removeItem(storageKey);
    if (storageKey === 'ops_directory_data') loadedDirectoryRecords = [];
    if (storageKey === 'ops_special_data') loadedSpecialRecords = [];
    
    const tbody = document.querySelector(`#${tableId} tbody`);
    if (tbody) tbody.innerHTML = "";
    alert("Directory data cleared!");
  }
}

/* -------------------------------------------------------------
   TEMPLATE CREATION & CUSTOM TEMPLATE MANAGEMENT
   ------------------------------------------------------------- */

function openNewTemplateModal(editId = null) {
  const modal = document.getElementById('newTemplateModal');
  const title = document.getElementById('templateModalTitle');
  const editInput = document.getElementById('editTemplateId');

  if (editId) {
    const tpl = customTemplates.find(t => t.id === editId);
    if (tpl) {
      title.innerText = "Edit Email Template";
      editInput.value = tpl.id;
      document.getElementById('tplTitleInput').value = tpl.title;
      document.getElementById('tplCategoryInput').value = tpl.category;
      document.getElementById('tplDescInput').value = tpl.desc;
      document.getElementById('tplBodyInput').value = tpl.body;
    }
  } else {
    title.innerText = "Create New Email Template";
    editInput.value = "";
    document.getElementById('newTemplateForm')?.reset();
  }
  modal?.classList.remove('hidden');
}

function closeNewTemplateModal() {
  document.getElementById('newTemplateForm')?.reset();
  document.getElementById('editTemplateId').value = "";
  document.getElementById('newTemplateModal')?.classList.add('hidden');
}

function saveNewTemplate(e) {
  e.preventDefault();
  const editId = document.getElementById('editTemplateId').value;
  const title = document.getElementById('tplTitleInput').value.trim();
  const category = document.getElementById('tplCategoryInput').value;
  const desc = document.getElementById('tplDescInput').value.trim();
  const body = document.getElementById('tplBodyInput').value.trim();

  if (editId) {
    const index = customTemplates.findIndex(t => t.id === editId);
    if (index !== -1) {
      customTemplates[index] = { id: editId, title, category, desc, body };
      localStorage.setItem('ops_custom_templates', JSON.stringify(customTemplates));
      reRenderAllCustomTemplates();
      alert("Template updated successfully!");
    }
  } else {
    const templateId = `tpl-custom-${Date.now()}`;
    const newTpl = { id: templateId, title, category, desc, body };
    customTemplates.push(newTpl);
    localStorage.setItem('ops_custom_templates', JSON.stringify(customTemplates));
    renderCustomTemplateCard(newTpl);
    alert("Template created successfully!");
  }

  closeNewTemplateModal();
}

function loadCustomTemplates() {
  const saved = localStorage.getItem('ops_custom_templates');
  if (saved) {
    try {
      customTemplates = JSON.parse(saved);
      customTemplates.forEach(tpl => renderCustomTemplateCard(tpl));
    } catch(e) {
      console.error("Error loading custom templates", e);
    }
  }
}

function reRenderAllCustomTemplates() {
  const grid = document.getElementById('templatesGrid');
  if (grid) grid.innerHTML = "";
  customTemplates.forEach(tpl => renderCustomTemplateCard(tpl));
}

function renderCustomTemplateCard(tpl) {
  const grid = document.getElementById('templatesGrid');
  if (!grid) return;

  const card = document.createElement('div');
  card.className = "tpl-card bg-white p-6 rounded-lg shadow border border-slate-200 flex flex-col justify-between";
  card.setAttribute('data-category', tpl.category);
  card.id = `card-${tpl.id}`;

  card.innerHTML = `
    <div>
      <div class="flex items-center justify-between mb-2">
        <h3 class="font-bold text-lg text-slate-800">${escapeHtml(tpl.title)}</h3>
        <div class="flex items-center gap-1">
          <span class="bg-blue-100 text-blue-800 text-[10px] font-semibold px-2 py-0.5 rounded capitalize">${escapeHtml(tpl.category)}</span>
          <button onclick="openNewTemplateModal('${tpl.id}')" title="Edit Template" class="text-blue-600 hover:text-blue-800 font-semibold px-1 text-xs cursor-pointer">Edit</button>
          <button onclick="deleteCustomTemplate('${tpl.id}')" title="Delete Template" class="text-rose-500 hover:text-rose-700 font-bold px-1 text-sm cursor-pointer">&times;</button>
        </div>
      </div>
      <p class="text-xs text-slate-500 mb-4">${escapeHtml(tpl.desc)}</p>
      <div class="bg-slate-50 p-3 rounded border border-slate-200 font-mono text-xs text-slate-700 mb-4 whitespace-pre-wrap max-h-72 overflow-y-auto" id="${tpl.id}">${escapeHtml(tpl.body)}</div>
    </div>
    <button type="button" onclick="copyToClipboard('${tpl.id}')" class="bg-blue-600 hover:bg-blue-700 text-white text-xs px-4 py-2 rounded transition font-medium w-full cursor-pointer">Copy Template</button>
  `;
  grid.appendChild(card);
}

function deleteCustomTemplate(id) {
  if (confirm("Delete this template?")) {
    customTemplates = customTemplates.filter(t => t.id !== id);
    localStorage.setItem('ops_custom_templates', JSON.stringify(customTemplates));
    document.getElementById(`card-${id}`)?.remove();
  }
}

function filterTemplates(category) {
  const buttons = document.querySelectorAll('.tpl-filter-btn');
  buttons.forEach(btn => {
    btn.className = 'tpl-filter-btn px-3 py-1.5 rounded-lg text-xs font-semibold transition bg-slate-200 text-slate-700 hover:bg-slate-300 cursor-pointer';
  });

  const activeBtn = document.getElementById(`tpl-btn-${category}`);
  if (activeBtn) {
    activeBtn.className = 'tpl-filter-btn px-3 py-1.5 rounded-lg text-xs font-semibold transition bg-blue-600 text-white cursor-pointer';
  }

  const cards = document.querySelectorAll('.tpl-card');
  cards.forEach(card => {
    if (category === 'all' || card.getAttribute('data-category') === category) {
      card.classList.remove('hidden');
    } else {
      card.classList.add('hidden');
    }
  });
}

/* -------------------------------------------------------------
   DYNAMIC DUAL-MODE SITE & CASE ROW MANAGEMENT
   ------------------------------------------------------------- */

function addSiteLocationRow(defaultCase = '', defaultLoc = '') {
  const container = document.getElementById('siteLocationRowsList');
  if (!container) return;

  const currentCount = container.querySelectorAll('.site-loc-row').length + 1;
  const rowId = Date.now() + Math.floor(Math.random() * 1000);

  const rowDiv = document.createElement('div');
  rowDiv.className = 'site-loc-row flex items-center gap-2 mb-2';
  rowDiv.id = `site-loc-row-${rowId}`;

  rowDiv.innerHTML = `
    <input type="text" id="siteCaseInput_${rowId}" value="${escapeHtml(defaultCase)}" onkeyup="updatePopoutTemplateText()" 
           placeholder="${currentCount} Case #:" class="border rounded px-2 py-1 text-xs w-1/3">
    <input type="text" id="siteLocInput_${rowId}" value="${escapeHtml(defaultLoc)}" onkeyup="updatePopoutTemplateText()" 
           placeholder="site location ${currentCount}:" class="border rounded px-2 py-1 text-xs w-2/3">
    <button type="button" onclick="removeSiteLocationRow('${rowId}')" class="text-rose-500 hover:text-rose-700 font-bold px-1.5 text-sm cursor-pointer">&times;</button>
  `;

  container.appendChild(rowDiv);
  updatePopoutTemplateText();
}

function removeSiteLocationRow(rowId) {
  const row = document.getElementById(`site-loc-row-${rowId}`);
  if (row) {
    row.remove();
    updatePopoutTemplateText();
  }
}

function clearSiteLocationRows() {
  const container = document.getElementById('siteLocationRowsList');
  if (container) container.innerHTML = '';
}

function handleTemplateSelectChange() {
  const templateId = document.getElementById('templateSelect')?.value;
  const singleContainer = document.getElementById('singleLocationContainer');
  const multiContainer = document.getElementById('multiLocationContainer');

  const isMultiLink = templateId === 'tpl-multi-link-down' || (templateId && templateId.toLowerCase().includes('multiple link down'));

  if (isMultiLink) {
    if (singleContainer) singleContainer.classList.add('hidden');
    if (multiContainer) multiContainer.classList.remove('hidden');
    
    const list = document.getElementById('siteLocationRowsList');
    if (list && list.children.length === 0) {
      addSiteLocationRow();
    }
  } else {
    if (singleContainer) singleContainer.classList.remove('hidden');
    if (multiContainer) multiContainer.classList.add('hidden');
  }

  updatePopoutTemplateText();
}

/* -------------------------------------------------------------
   INTERACTIVE UI HANDLERS & MODALS
   ------------------------------------------------------------- */

function bindLocCellEvents() {
  document.addEventListener('click', function(e) {
    if (e.target && e.target.classList.contains('loc-cell')) {
      const cell = e.target;
      let currentStatus = cell.getAttribute('data-status') || 'att';
      let currentIndex = statusCycle.indexOf(currentStatus);
      let nextIndex = (currentIndex + 1) % statusCycle.length;
      let nextStatus = statusCycle[nextIndex];

      Object.values(statusClassMap).forEach(cls => cell.classList.remove(cls));
      cell.classList.add(statusClassMap[nextStatus]);
      cell.setAttribute('data-status', nextStatus);
    }
  });
}

function populateContactSelect(selectedIdx = -1) {
  const contactSelect = document.getElementById('contactSelect');
  if (!contactSelect) return;
  contactSelect.innerHTML = "";

  if (loadedDirectoryRecords.length === 0) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "-- No Contacts Found in Directory --";
    contactSelect.appendChild(opt);
    return;
  }

  loadedDirectoryRecords.forEach((rec, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    const displayName = rec.name && rec.name !== '-' ? rec.name : (rec.contactName || rec.id);
    opt.textContent = `${displayName} (${rec.id})`;
    if (idx === selectedIdx) {
      opt.selected = true;
    }
    contactSelect.appendChild(opt);
  });
}

function openTemplateModal(clientIndex = -1) {
  activeClientIndex = clientIndex;
  const categorySelect = document.getElementById('categorySelect');
  const templateSelect = document.getElementById('templateSelect');
  
  populateContactSelect(clientIndex >= 0 ? clientIndex : 0);

  if (categorySelect) categorySelect.value = "";
  if (templateSelect) {
    templateSelect.innerHTML = "";
    templateSelect.classList.add("hidden");
  }
  
  const siteInput = document.getElementById('siteLocationInput');
  if (siteInput) siteInput.value = "";

  const etrMulti = document.getElementById('etrMultiInput');
  if (etrMulti) etrMulti.value = "";

  clearSiteLocationRows();
  
  document.getElementById('popoutTemplatePreview').innerText = "";
  document.getElementById('templateModal')?.classList.remove('hidden');
}

function closeTemplateModal() {
  document.getElementById('templateModal')?.classList.add('hidden');
}

function handleCategorySelectChange() {
  const category = document.getElementById('categorySelect').value;
  const templateSelect = document.getElementById('templateSelect');
  templateSelect.innerHTML = "";

  let items = [...(predefinedTemplates[category] || [])];
  
  const customMatches = customTemplates.filter(t => t.category === category);
  customMatches.forEach(t => items.push({ id: t.id, title: t.title }));

  if (items.length === 0) {
    templateSelect.classList.add("hidden");
    document.getElementById('popoutTemplatePreview').innerText = "No templates found for this category.";
    return;
  }

  templateSelect.classList.remove("hidden");
  
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.disabled = true;
  defaultOption.selected = true;
  defaultOption.textContent = "-- Select Email Template --";
  templateSelect.appendChild(defaultOption);

  items.forEach(item => {
    const opt = document.createElement("option");
    opt.value = item.id;
    opt.textContent = item.title;
    templateSelect.appendChild(opt);
  });

  document.getElementById('popoutTemplatePreview').innerText = "";
}

function handleCarrierChange(index) {
  const select = document.getElementById(`circuitSelect${index}`);
  const customInput = document.getElementById(`customCarrierInput${index}`);
  if (select && customInput) {
    if (select.value === '__custom__') {
      customInput.classList.remove('hidden');
      customInput.focus();
    } else {
      customInput.classList.add('hidden');
    }
  }
  updatePopoutTemplateText();
}

function clearAllCarriersAndPorts() {
  for (let i = 1; i <= 4; i++) {
    const select = document.getElementById(`circuitSelect${i}`);
    const customInput = document.getElementById(`customCarrierInput${i}`);
    const portSelect = document.getElementById(`portSelect${i}`);
    
    if (select) select.value = "";
    if (customInput) {
      customInput.value = "";
      customInput.classList.add('hidden');
    }
    if (portSelect) portSelect.value = "";
  }

  const singleCarrierInput = document.querySelector('input[placeholder*="Carrier"], #carrierInput, .carrier-input') || 
                             Array.from(document.querySelectorAll('input')).find(i => i.previousElementSibling?.textContent?.includes('Carrier'));
  if (singleCarrierInput) {
    singleCarrierInput.value = "";
  }

  updatePopoutTemplateText();
}

function getSelectedCarrier(index) {
  const select = document.getElementById(`circuitSelect${index}`);
  if (select && select.value) {
    if (select.value === '__custom__') {
      return document.getElementById(`customCarrierInput${index}`)?.value.trim() || "";
    }
    return select.value;
  }
  return "";
}

function getCombinedCarrierAndPorts() {
  const pairs = [];
  for (let i = 1; i <= 4; i++) {
    const carrier = getSelectedCarrier(i);
    const port = document.getElementById(`portSelect${i}`)?.value;

    if (carrier || port) {
      const cStr = carrier || "[Carrier]";
      const pStr = port || "[Port]";
      pairs.push(`${cStr} (Port ${pStr})`);
    }
  }
  return pairs.length > 0 ? pairs.join(", ") : "[Carrier/Port Details]";
}

/* -------------------------------------------------------------
   DYNAMIC TEMPLATE PREVIEW & PLACEHOLDER REPLACER
   ------------------------------------------------------------- */

function updatePopoutTemplateText() {
  const templateId = document.getElementById('templateSelect')?.value;
  if (!templateId) return;
  
  const rawElem = document.getElementById(templateId);
  if (!rawElem) return;

  let carrierValue = getSelectedCarrier(1);
  if (!carrierValue) {
    const singleCarrierInput = document.querySelector('input[placeholder*="Carrier"], #carrierInput, .carrier-input') || 
                               Array.from(document.querySelectorAll('input')).find(i => i.previousElementSibling?.textContent?.includes('Carrier'));
    if (singleCarrierInput && singleCarrierInput.value.trim() !== "") {
      carrierValue = singleCarrierInput.value.trim();
    }
  }
  if (!carrierValue) carrierValue = "[Carrier]";

  const c1 = carrierValue;
  const c2 = getSelectedCarrier(2) || c1;
  const c3 = getSelectedCarrier(3) || c1;
  const c4 = getSelectedCarrier(4) || c1;

  const p1 = document.getElementById('portSelect1')?.value || "[Port 1]";
  const p2 = document.getElementById('portSelect2')?.value || "[Port 2]";
  const p3 = document.getElementById('portSelect3')?.value || "[Port 3]";
  const p4 = document.getElementById('portSelect4')?.value || "[Port 4]";

  const combinedCarrierPorts = getCombinedCarrierAndPorts();
  const siteLocation = document.getElementById('siteLocationInput')?.value.trim() || "[Site Location]";
  const etrValue = document.getElementById('etrMultiInput')?.value.trim() || document.querySelector('#etrInput')?.value?.trim() || "[ETR]";

  const companyInput = document.getElementById('companyNameInput') || 
                       document.querySelector('input[placeholder*="Company"], .company-input') || 
                       Array.from(document.querySelectorAll('input')).find(i => i.previousElementSibling?.textContent?.includes('Company Name'));
  
  const selectedContactIdx = parseInt(document.getElementById('contactSelect')?.value);
  let clientName = "[Client Name]";

  if (!isNaN(selectedContactIdx) && loadedDirectoryRecords[selectedContactIdx]) {
    clientName = loadedDirectoryRecords[selectedContactIdx].name;
  } else if (companyInput && companyInput.value.trim() !== "") {
    clientName = companyInput.value.trim();
  }

  let text = rawElem.innerText;

  const siteRows = document.querySelectorAll('#siteLocationRowsList .site-loc-row, .site-loc-row');
  const siteLocList = [];

  if (siteRows.length > 0) {
    siteRows.forEach((row, index) => {
      const idx = index + 1;
      const inputs = row.querySelectorAll('input[type="text"]');
      const caseNum = inputs[0]?.value.trim() || `[Case ${idx}]`;
      const locName = inputs[1]?.value.trim() || `[Location ${idx}]`;

      const caseRegex = new RegExp(`{{Case_${idx}}}`, 'g');
      const locRegex = new RegExp(`{{Site_Location_${idx}}}`, 'g');

      text = text.replace(caseRegex, caseNum);
      text = text.replace(locRegex, locName);

      if (idx === 1) {
        text = text.replace(/{{Case_#}}/g, caseNum);
        text = text.replace(/{{Site_Location}}/g, locName);
      }

      siteLocList.push(`- Case #: ${caseNum}, Location ${idx}: ${locName}`);
    });
  } else {
    text = text
      .replace(/{{Case_1}}/g, "[Case 1]")
      .replace(/{{Site_Location_1}}/g, "[Location 1]")
      .replace(/{{Case_2}}/g, "[Case 2]")
      .replace(/{{Site_Location_2}}/g, "[Location 2]")
      .replace(/{{Case_3}}/g, "[Case 3]")
      .replace(/{{Site_Location_3}}/g, "[Location 3]");
  }

  text = text.replace(/{{Site_Locations_List}}/g, siteLocList.join('\n'));

  text = text
    .replace(/{{Carrier_Type_1}}/g, c1)
    .replace(/{{Carrier_Type_2}}/g, c2)
    .replace(/{{Carrier_Type_3}}/g, c3)
    .replace(/{{Carrier_Type_4}}/g, c4)
    .replace(/{{Carrier_Type}}/g, c1)
    .replace(/{{Circuit_Type}}/g, c1)
    .replace(/{{Connection_1}}/g, `${c1} (Port ${p1})`)
    .replace(/{{Connection_2}}/g, `${c2} (Port ${p2})`)
    .replace(/{{Connection_3}}/g, `${c3} (Port ${p3})`)
    .replace(/{{Connection_4}}/g, `${c4} (Port ${p4})`)
    .replace(/{{Carrier_Port_Summary}}/g, combinedCarrierPorts)
    .replace(/{{Port1}}/g, p1)
    .replace(/{{Port2}}/g, p2)
    .replace(/{{Port3}}/g, p3)
    .replace(/{{Port4}}/g, p4)
    .replace(/{{Port}}/g, p1)
    .replace(/{{Full_Address}}/g, siteLocation)
    .replace(/{{location}}/g, siteLocation)
    .replace(/{{ETR}}/g, etrValue);

  text = text
    .replace(/Dear {{Client_Name}},/g, `Dear ${clientName},`)
    .replace(/Dear {{Account_Name}},/g, `Dear ${clientName},`)
    .replace(/{{Account_Name}}/g, clientName)
    .replace(/{{Client_Name}}/g, clientName);

  if (!isNaN(selectedContactIdx) && loadedDirectoryRecords[selectedContactIdx]) {
    const selectedClient = loadedDirectoryRecords[selectedContactIdx];
    text = text
      .replace(/{{Account_ID}}/g, selectedClient.id)
      .replace(/{{Customer_Phone}}/g, selectedClient.contactPhone || '-')
      .replace(/{{Customer_Email}}/g, selectedClient.contactEmail || '-');
  }

  document.getElementById('popoutTemplatePreview').innerText = text;
}

function copyPopoutTemplate() {
  const text = document.getElementById('popoutTemplatePreview').innerText;
  if (!text) {
    alert("Please select a template first.");
    return;
  }
  navigator.clipboard.writeText(text).then(() => {
    alert("Populated template copied to clipboard!");
    closeTemplateModal();
  });
}
