const DB_NAME = "bie-db-v1";
const DB_VERSION = 1;
const STORE_NAME = "state";
const CHECKLIST_TEMPLATE_URL = "./checklist-bie.xlsx";

const defectOptions = [
  "Manguera rota",
  "Hay un obstáculo",
  "Cristal roto",
  "Manómetro mal",
  "Válvula mal",
  "Lanza mal",
  "Devanadera mal",
  "Sin presión",
  "Presión baja",
  "Sin señal",
  "Señal caducada",
  "Armario en mal estado",
];

const checklistOptions = [
  ["A", "Etiqueta de revisión visible y legible"],
  ["B", "Está señalizada"],
  ["C", "Es accesible"],
  ["D", "Elementos metálicos sin corrosión visible"],
  ["E", "Armario en buen estado"],
  ["F", "Cristal/puerta en perfecto estado"],
  ["G", "Manómetro con presión adecuada"],
  ["H", "Conjunto limpio para manejo adecuado"],
  ["I", "Bisagras y cierre engrasados"],
  ["J", "Lugar visible"],
  ["K", "Correctamente fijada"],
  ["L", "Manguera se desenrolla con facilidad"],
  ["M", "Manguera en buen estado"],
  ["N", "Lanza estanca y funcional"],
  ["O", "Válvula opera correctamente"],
  ["P", "Acoples estancos y juntas correctas"],
  ["Q", "Manómetro contrasta correctamente"],
  ["S", "Retimbrado de manguera menor a 5 años"],
  ["T", "Caudal constante y suficiente"],
  ["U", "Manguera certificada UNE-EN 694/14540"],
  ["V", "Racores certificados UNE 23400"],
  ["W", "Manguera sin fugas a 15 Kg/cm2"],
  ["R1", "Red de tuberías sin daños ni corrosión"],
  ["R2", "Soportación de tubería y equipos correcta"],
  ["R3", "Válvulas de sectorización abiertas y operativas"],
  ["R4", "Existe contador para incendios"],
  ["R5", "Alimentación exclusiva para incendios"],
  ["R6", "Detector de flujo funciona correctamente"],
  ["R7", "Válvula motorizada por detección funciona correctamente"],
];

const fields = [
  "cliente",
  "edificio",
  "cantidad",
  "ubicacion",
  "modelo",
  "fabricanteMarca",
  "numeroSerie",
  "fechaFabricacion",
  "fechaProximoRetimbrado",
  "operacionesRealizadas",
  "cadu20A",
  "tipoRevision",
  "estadoEquipo",
  "observaciones",
  "senal",
];

let records = [];
let currentPhotos = ["", ""];

const $ = (id) => document.getElementById(id);

function safeText(value) {
  return value === undefined || value === null ? "" : String(value);
}

function createId() {
  return `rec-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function fillYearLists() {
  const fabricacion = $("fechaFabricacion");
  const retimbrado = $("fechaProximoRetimbrado");
  fabricacion.innerHTML = `<option value="">Sin indicar</option>`;
  retimbrado.innerHTML = `<option value="">Sin indicar</option>`;
  for (let year = 2005; year <= 2026; year += 1) {
    fabricacion.insertAdjacentHTML("beforeend", `<option value="${year}">${year}</option>`);
  }
  for (let year = 2010; year <= 2026; year += 1) {
    retimbrado.insertAdjacentHTML("beforeend", `<option value="${year}">${year}</option>`);
  }
}

function normalizeDefects(defects) {
  if (!Array.isArray(defects)) return [];
  return defects
    .map((defect) => {
      const text = safeText(defect).trim();
      const withoutPeriod = text.endsWith(".") ? text.slice(0, -1) : text;
      if (withoutPeriod === "Armario roto y presión baja") return "Armario en mal estado";
      return withoutPeriod;
    })
    .filter((defect) => defectOptions.includes(defect));
}

function normalizeChecklist(values) {
  const valid = new Set(checklistOptions.map(([code]) => code));
  return (Array.isArray(values) ? values : []).map((value) => safeText(value)).filter((value) => valid.has(value));
}

function normalizeYearValue(value) {
  const text = safeText(value);
  const match = text.match(/\b(19\d{2}|20\d{2})\b/);
  return match ? match[1] : text;
}

function cleanRecord(record = {}) {
  return {
    id: record.id || createId(),
    cliente: safeText(record.cliente),
    edificio: safeText(record.edificio ?? record.edificioCodigo),
    cantidad: safeText(record.cantidad),
    ubicacion: safeText(record.ubicacion),
    modelo: safeText(record.modelo),
    fabricanteMarca: safeText(record.fabricanteMarca),
    numeroSerie: safeText(record.numeroSerie),
    fechaFabricacion: normalizeYearValue(record.fechaFabricacion),
    fechaProximoRetimbrado: normalizeYearValue(record.fechaProximoRetimbrado),
    operacionesRealizadas: safeText(record.operacionesRealizadas),
    cadu20A: safeText(record.cadu20A),
    tipoRevision: safeText(record.tipoRevision),
    estadoEquipo: safeText(record.estadoEquipo),
    observaciones: safeText(record.observaciones),
    senal: safeText(record.senal),
    defectos: normalizeDefects(record.defectos),
    checklist: normalizeChecklist(record.checklist),
    photos: Array.isArray(record.photos) ? [safeText(record.photos[0]), safeText(record.photos[1])] : ["", ""],
    visto: Boolean(record.visto),
    origen: record.origen || "excel",
  };
}

function excelCellToText(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-ES");
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return excelCellToText(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
  }
  return String(value);
}

function resolveChecklistModel(values) {
  if (safeText(values[2]).trim()) return "25";
  if (safeText(values[3]).trim()) return "45";
  if (safeText(values[4]).trim()) return "25T45";
  return safeText(values[6]);
}

function rowToImportedRecord(rowValues, index, metadata = {}) {
  const values = [];
  for (let col = 1; col <= 10; col += 1) values[col] = excelCellToText(rowValues[col]);
  return cleanRecord({
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    cliente: metadata.cliente,
    edificio: metadata.domicilio,
    cantidad: values[1],
    modelo: resolveChecklistModel(values),
    fabricanteMarca: values[5],
    fechaFabricacion: values[6],
    fechaProximoRetimbrado: values[7],
    operacionesRealizadas: values[8],
    cadu20A: values[9],
    origen: "importado",
  });
}

function normalizeHeader(value) {
  return safeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function truthyExcelValue(value) {
  const text = normalizeHeader(excelCellToText(value));
  return Boolean(text && !["no", "false", "0", "sin indicar"].includes(text));
}

function buildHeaderMap(rowValues) {
  const map = new Map();
  rowValues.forEach((value, index) => {
    const key = normalizeHeader(excelCellToText(value));
    if (key && !map.has(key)) map.set(key, index);
  });
  return map;
}

function headerMatches(header, candidate) {
  const key = normalizeHeader(header);
  const wanted = normalizeHeader(candidate);
  if (!key || !wanted) return false;
  if (key === wanted || key.includes(wanted) || wanted.includes(key)) return true;
  if (wanted.includes("syco")) return key.includes("syco");
  if (wanted.includes("serie")) return key.includes("serie");
  if (wanted.includes("ubicacion")) return key.includes("ubicacion");
  if (wanted.includes("fabricante") || wanted.includes("marca")) return key.includes("fabricante") || key.includes("marca");
  if (wanted.includes("fabricacion")) return key.includes("fabricacion") || key.includes("fabri");
  if (wanted.includes("retimbrado")) return key.includes("retimbre") || key.includes("retimbrado");
  if (wanted.includes("senal")) return key.includes("senal");
  return false;
}

function importedValue(rowValues, headerMap, candidates) {
  for (const candidate of candidates) {
    const col = headerMap.get(normalizeHeader(candidate));
    if (col !== undefined) return excelCellToText(rowValues[col]);
  }
  for (const candidate of candidates) {
    for (const [header, col] of headerMap.entries()) {
      if (headerMatches(header, candidate)) return excelCellToText(rowValues[col]);
    }
  }
  return "";
}

function selectedFromText(text, options) {
  const normalized = normalizeHeader(text);
  return options.filter((option) => normalized.includes(normalizeHeader(option)));
}

function findAppExportHeader(sheet) {
  let best = null;
  const wanted = ["cliente", "edificio", "syco", "ubicacion", "modelo", "serie", "foto"];
  const maxRow = Math.min(sheet.rowCount, 30);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const headerMap = buildHeaderMap(row.values);
    const keys = Array.from(headerMap.keys());
    const score = wanted.reduce((total, word) => total + (keys.some((key) => key.includes(word)) ? 1 : 0), 0);
    if (!best || score > best.score) best = { rowNumber, headerMap, score };
  }
  return best && best.score >= 2 ? best : null;
}

function importAppExportRows(sheet) {
  const headerInfo = findAppExportHeader(sheet);
  if (!headerInfo) return [];
  const { rowNumber: headerRowNumber, headerMap } = headerInfo;

  const imported = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowNumber) return;
    const rowValues = row.values;
    const selectedDefects = new Set(selectedFromText(importedValue(rowValues, headerMap, ["Defectos encontrados"]), defectOptions));
    defectOptions.forEach((defect) => {
      const col = headerMap.get(normalizeHeader(defect));
      if (col !== undefined && truthyExcelValue(rowValues[col])) selectedDefects.add(defect);
    });

    const checklistText = importedValue(rowValues, headerMap, ["Comprobaciones checklist"]);
    const selectedChecklist = new Set(
      checklistText
        .split(/[\/,;]+/)
        .map((item) => safeText(item).trim().split(/\s+/)[0])
        .filter(Boolean)
    );
    checklistOptions.forEach(([code, text]) => {
      const col = headerMap.get(normalizeHeader(`${code} - ${text}`));
      if (col !== undefined && truthyExcelValue(rowValues[col])) selectedChecklist.add(code);
    });

    const record = cleanRecord({
      id: `import-${Date.now()}-${rowNumber}-${Math.random().toString(16).slice(2)}`,
      cliente: importedValue(rowValues, headerMap, ["Cliente"]),
      edificio: importedValue(rowValues, headerMap, ["Edificio"]),
      cantidad: importedValue(rowValues, headerMap, ["Número SYCo", "Nº SYCo", "Numero SYCo"]),
      ubicacion: importedValue(rowValues, headerMap, ["Ubicación", "Ubicacion"]),
      modelo: importedValue(rowValues, headerMap, ["Modelo"]),
      fabricanteMarca: importedValue(rowValues, headerMap, ["Fabricante/Marca", "Fabricante", "Marca"]),
      numeroSerie: importedValue(rowValues, headerMap, ["Nº serie", "N° serie", "Numero serie"]),
      fechaFabricacion: importedValue(rowValues, headerMap, ["Fecha / año fabricación", "Fabricación", "Fabricacion"]),
      fechaProximoRetimbrado: importedValue(rowValues, headerMap, ["Fecha retimbrado", "Retimbrado"]),
      operacionesRealizadas: importedValue(rowValues, headerMap, ["Op. realizadas", "Operaciones realizadas"]),
      cadu20A: importedValue(rowValues, headerMap, ["Cadu 20 A"]),
      tipoRevision: importedValue(rowValues, headerMap, ["Tipo revisión", "Tipo revision"]),
      estadoEquipo: importedValue(rowValues, headerMap, ["Estado equipo"]),
      observaciones: importedValue(rowValues, headerMap, ["Observaciones"]),
      senal: importedValue(rowValues, headerMap, ["Señal", "Senal"]),
      defectos: Array.from(selectedDefects),
      checklist: normalizeChecklist(Array.from(selectedChecklist)),
      visto: truthyExcelValue(importedValue(rowValues, headerMap, ["Visto"])),
      origen: "importado",
    });
    const hasData = [record.cliente, record.edificio, record.cantidad, record.ubicacion, record.modelo, record.numeroSerie].some((value) => safeText(value).trim());
    if (hasData) imported.push(record);
  });
  return imported;
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readState() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get("records");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function writeState(value) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(value, "records");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadRecords() {
  try {
    const saved = await readState();
    if (Array.isArray(saved)) {
      records = saved.map(cleanRecord);
      return;
    }
  } catch {}
  records = (window.INITIAL_BIE_LISTADOS || []).map(cleanRecord);
  await saveRecords();
}

async function saveRecords() {
  records = records.map(cleanRecord);
  updateStats();
  await writeState(records);
}

function updateStats() {
  const total = records.length;
  const seen = records.filter((record) => record.visto).length;
  $("totalCount").textContent = total;
  $("seenCount").textContent = seen;
  $("pendingCount").textContent = total - seen;
  updateLastNumberUsed();
}

function getLastNumberUsed() {
  const numbers = records.map((record) => safeText(record.cantidad).trim()).filter(Boolean);
  if (!numbers.length) return "-";
  numbers.sort((a, b) => compareText(a, b));
  return numbers[numbers.length - 1];
}

function updateLastNumberUsed() {
  const target = $("lastNumberUsed");
  if (target) target.textContent = getLastNumberUsed();
}

function clearFilters() {
  ["filterCliente", "filterEdificio", "filterNumero", "filterSerie"].forEach((id) => {
    const input = $(id);
    if (input) input.value = "";
  });
  if ($("sortOrder")) $("sortOrder").value = "none";
  if ($("seenFilter")) $("seenFilter").value = "all";
}

function showView(name) {
  $("homeView").classList.toggle("hidden", name !== "home");
  $("listView").classList.toggle("hidden", name !== "list");
  $("formView").classList.toggle("hidden", name !== "form");
  if (name === "list") renderTable();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function compareText(a, b) {
  return safeText(a).localeCompare(safeText(b), "es", { numeric: true, sensitivity: "base" });
}

function filteredRecords() {
  const filterCliente = $("filterCliente").value.trim().toLowerCase();
  const filterEdificio = $("filterEdificio").value.trim().toLowerCase();
  const filterNumero = $("filterNumero").value.trim().toLowerCase();
  const filterSerie = $("filterSerie").value.trim().toLowerCase();
  const seenFilter = $("seenFilter").value;
  const sortOrder = $("sortOrder").value;

  const rows = records.filter((record) => {
    if (seenFilter === "seen" && !record.visto) return false;
    if (seenFilter === "pending" && record.visto) return false;
    if (filterCliente && !safeText(record.cliente).toLowerCase().includes(filterCliente)) return false;
    if (filterEdificio && ![record.edificio, record.ubicacion].join(" ").toLowerCase().includes(filterEdificio)) return false;
    if (filterNumero && !safeText(record.cantidad).toLowerCase().includes(filterNumero)) return false;
    if (filterSerie && !safeText(record.numeroSerie).toLowerCase().includes(filterSerie)) return false;
    return true;
  });

  if (sortOrder === "cliente") rows.sort((a, b) => compareText(a.cliente, b.cliente) || compareText(a.edificio, b.edificio));
  if (sortOrder === "edificio") rows.sort((a, b) => compareText(a.edificio, b.edificio) || compareText(a.cantidad, b.cantidad));
  if (sortOrder === "numero") rows.sort((a, b) => compareText(a.cantidad, b.cantidad) || compareText(a.edificio, b.edificio));
  return rows;
}

function renderTable() {
  const body = $("recordsBody");
  const rows = filteredRecords();
  body.innerHTML = "";
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="21">No hay registros con ese filtro.</td></tr>`;
    return;
  }
  for (const record of rows) {
    const defects = record.defectos.length ? record.defectos.join(" / ") : "-";
    const checklist = record.checklist.length ? record.checklist.join(" / ") : "-";
    const photo1 = record.photos[0] ? `<img class="tablePhoto" src="${record.photos[0]}" alt="Foto 1">` : `<span class="noPhoto">—</span>`;
    const photo2 = record.photos[1] ? `<img class="tablePhoto" src="${record.photos[1]}" alt="Foto 2">` : `<span class="noPhoto">—</span>`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${safeText(record.cliente) || "-"}</td>
      <td>${safeText(record.edificio) || "-"}</td>
      <td><strong>${safeText(record.cantidad) || "-"}</strong></td>
      <td>${safeText(record.ubicacion) || "-"}</td>
      <td>${safeText(record.modelo) || "-"}</td>
      <td>${safeText(record.fabricanteMarca) || "-"}</td>
      <td>${safeText(record.numeroSerie) || "-"}</td>
      <td>${safeText(record.fechaFabricacion) || "-"}</td>
      <td>${safeText(record.fechaProximoRetimbrado) || "-"}</td>
      <td>${safeText(record.operacionesRealizadas) || "-"}</td>
      <td>${safeText(record.cadu20A) || "-"}</td>
      <td>${safeText(record.tipoRevision) || "-"}</td>
      <td>${safeText(record.estadoEquipo) || "-"}</td>
      <td>${safeText(record.observaciones) || "-"}</td>
      <td>${safeText(record.senal) || "-"}</td>
      <td>${defects}</td>
      <td>${checklist}</td>
      <td>${photo1}</td>
      <td>${photo2}</td>
      <td><span class="${record.visto ? "ok" : "pending"}">${record.visto ? "Sí" : "No"}</span></td>
      <td><button class="editBtn" data-edit="${record.id}">Ver / corregir</button></td>
    `;
    body.appendChild(tr);
  }
  body.querySelectorAll("[data-edit]").forEach((button) => {
    button.addEventListener("click", () => openForm(button.dataset.edit));
  });
}

function renderDefects(selected = []) {
  const box = $("defectsList");
  box.innerHTML = "";
  for (const option of defectOptions) {
    const label = document.createElement("label");
    label.className = "checkItem";
    label.innerHTML = `<input type="checkbox" value="${option}"><span>${option}</span>`;
    label.querySelector("input").checked = selected.includes(option);
    box.appendChild(label);
  }
}

function renderChecklist(selected = []) {
  const box = $("checklistList");
  box.innerHTML = "";
  for (const [code, text] of checklistOptions) {
    const label = document.createElement("label");
    label.className = "checkItem checklistItem";
    label.innerHTML = `<input type="checkbox" value="${code}"><span><strong>${code}</strong> ${text}</span>`;
    label.querySelector("input").checked = selected.includes(code);
    box.appendChild(label);
  }
}

function setPhotoPreview(index, dataUrl) {
  const photo = safeText(dataUrl);
  const img = $(`photoPreview${index + 1}`);
  const text = $(`photoBox${index + 1}`).querySelector("span");
  currentPhotos[index] = photo;
  img.src = photo;
  img.classList.toggle("hidden", !photo);
  text.classList.toggle("hidden", Boolean(photo));
  $(`deletePhoto${index + 1}`).disabled = !photo;
}

function resizePhoto(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const maxSide = 1200;
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.72));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function openForm(id = null) {
  const record = id ? records.find((item) => item.id === id) : null;
  $("recordId").value = record?.id || "";
  $("formTitle").textContent = record ? "Ver y corregir BIE" : "Meter dato nuevo";
  $("formKicker").textContent = record ? "REGISTRO EXISTENTE" : "NUEVO REGISTRO";
  $("deleteBtn").classList.toggle("hidden", !record);
  for (const key of fields) $(key).value = safeText(record?.[key]);
  $("visto").checked = Boolean(record?.visto);
  renderDefects(record?.defectos || []);
  renderChecklist(record?.checklist || []);
  const photos = Array.isArray(record?.photos) ? record.photos : ["", ""];
  setPhotoPreview(0, photos[0]);
  setPhotoPreview(1, photos[1]);
  updateLastNumberUsed();
  showView("form");
}

function collectForm() {
  const currentId = $("recordId").value;
  const editingExisting = Boolean(currentId && records.some((item) => item.id === currentId));
  const record = { id: editingExisting ? currentId : createId(), origen: editingExisting ? "editado" : "manual" };
  for (const key of fields) record[key] = $(key).value.trim();
  record.defectos = Array.from($("defectsList").querySelectorAll("input:checked")).map((input) => input.value);
  record.checklist = Array.from($("checklistList").querySelectorAll("input:checked")).map((input) => input.value);
  record.photos = [currentPhotos[0] || "", currentPhotos[1] || ""];
  record.visto = $("visto").checked;
  return cleanRecord(record);
}

async function saveForm(event) {
  event.preventDefault();
  const wasEditing = Boolean($("recordId").value && records.some((item) => item.id === $("recordId").value));
  const record = collectForm();
  const index = records.findIndex((item) => item.id === record.id);
  if (index >= 0) records[index] = record;
  else records.unshift(record);
  await saveRecords();
  clearFilters();
  if (wasEditing) showView("list");
  else openForm();
}

async function deleteCurrent() {
  const id = $("recordId").value;
  if (!id) return;
  if (!confirm("¿Seguro que quieres eliminar este registro?")) return;
  records = records.filter((record) => record.id !== id);
  await saveRecords();
  showView("list");
}

async function deleteAllRecords() {
  if (!records.length) {
    alert("No hay registros para eliminar.");
    return;
  }
  if (!confirm(`¿Seguro que quieres eliminar los ${records.length} registros existentes? Esta acción no se puede deshacer.`)) return;
  records = [];
  await saveRecords();
  renderTable();
  showView("list");
  $("importStatus").textContent = "Registros existentes eliminados. Ya puedes importar otro cliente.";
}

async function importExcelFile(file) {
  if (!window.ExcelJS) return alert("No se ha cargado el lector de Excel.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) return alert("No encuentro ninguna hoja en ese Excel.");
  const metadata = {
    cliente: excelCellToText(sheet.getCell("E2").value),
    domicilio: excelCellToText(sheet.getCell("E3").value),
  };
  let imported = importAppExportRows(sheet);
  if (!imported.length) {
    imported = [];
    sheet.eachRow((row, rowNumber) => {
      const firstCell = excelCellToText(row.values[1]).trim();
      if (!/^syco\s*\d+/i.test(firstCell)) return;
      const record = rowToImportedRecord(row.values, rowNumber, metadata);
      const hasData = [record.cantidad, record.modelo, record.fabricanteMarca, record.fechaFabricacion, record.fechaProximoRetimbrado].some((value) => safeText(value).trim());
      if (!hasData) return;
      imported.push(record);
    });
  }
  if (!imported.length) {
    $("importStatus").textContent = "No se encontraron registros para importar.";
    return alert("No se encontraron registros para importar.");
  }
  records = [...imported, ...records];
  await saveRecords();
  clearFilters();
  $("importStatus").textContent = `Importados ${imported.length}. Los repetidos se han mantenido.`;
  alert(`Importación correcta.\nImportados: ${imported.length}\nLos repetidos se han mantenido.`);
}

function defectFlag(selected, defect) {
  return selected.includes(defect) ? "Sí" : "";
}

async function downloadExcel() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BIE";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("BIE");
  const columns = [
    ["cliente", "Cliente", 22],
    ["edificio", "Edificio", 14],
    ["cantidad", "Número SYCo", 18],
    ["ubicacion", "Ubicación", 42],
    ["modelo", "Modelo", 20],
    ["fabricanteMarca", "Fabricante/Marca", 22],
    ["numeroSerie", "Nº serie", 18],
    ["fechaFabricacion", "Fecha / año fabricación", 22],
    ["fechaProximoRetimbrado", "Fecha retimbrado", 20],
    ["operacionesRealizadas", "Op. realizadas", 24],
    ["cadu20A", "Cadu 20 A", 14],
    ["tipoRevision", "Tipo revisión", 18],
    ["estadoEquipo", "Estado equipo", 18],
    ["observaciones", "Observaciones", 34],
    ["senal", "Señal", 14],
    ["defectos", "Defectos encontrados", 42],
    ["defectoMangueraRota", "Manguera rota", 20],
    ["defectoObstaculo", "Hay un obstáculo", 20],
    ["defectoCristalRoto", "Cristal roto", 18],
    ["defectoManometroMal", "Manómetro mal", 18],
    ["defectoValvulaMal", "Válvula mal", 18],
    ["defectoLanzaMal", "Lanza mal", 18],
    ["defectoDevanaderaMal", "Devanadera mal", 20],
    ["defectoSinPresion", "Sin presión", 18],
    ["defectoPresionBaja", "Presión baja", 18],
    ["defectoSinSenal", "Sin señal", 16],
    ["defectoSenalCaducada", "Señal caducada", 20],
    ["defectoArmarioMalEstado", "Armario en mal estado", 26],
    ["checklist", "Comprobaciones checklist", 34],
    ...checklistOptions.map(([code, text]) => [`check_${code}`, `${code} - ${text}`, 28]),
    ["foto1", "Foto 1", 22],
    ["foto2", "Foto 2", 22],
    ["visto", "Visto", 10],
  ];
  sheet.columns = columns.map(([key, header, width]) => ({ key, header, width }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FF0F3A5F" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF93C5FD" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(1).height = 30;

  for (const record of filteredRecords()) {
    const selected = record.defectos || [];
    const selectedChecklist = record.checklist || [];
    const checklistFlags = Object.fromEntries(checklistOptions.map(([code]) => [`check_${code}`, selectedChecklist.includes(code) ? "Sí" : ""]));
    const row = sheet.addRow({
      ...record,
      defectos: selected.join(" / "),
      checklist: selectedChecklist.join(" / "),
      defectoMangueraRota: defectFlag(selected, "Manguera rota"),
      defectoObstaculo: defectFlag(selected, "Hay un obstáculo"),
      defectoCristalRoto: defectFlag(selected, "Cristal roto"),
      defectoManometroMal: defectFlag(selected, "Manómetro mal"),
      defectoValvulaMal: defectFlag(selected, "Válvula mal"),
      defectoLanzaMal: defectFlag(selected, "Lanza mal"),
      defectoDevanaderaMal: defectFlag(selected, "Devanadera mal"),
      defectoSinPresion: defectFlag(selected, "Sin presión"),
      defectoPresionBaja: defectFlag(selected, "Presión baja"),
      defectoSinSenal: defectFlag(selected, "Sin señal"),
      defectoSenalCaducada: defectFlag(selected, "Señal caducada"),
      defectoArmarioMalEstado: defectFlag(selected, "Armario en mal estado"),
      ...checklistFlags,
      foto1: record.photos[0] ? "Foto 1" : "",
      foto2: record.photos[1] ? "Foto 2" : "",
      visto: record.visto ? "Sí" : "No",
    });
    if (record.photos[0] || record.photos[1]) row.height = 92;
    [0, 1].forEach((photoIndex) => {
      const photo = record.photos[photoIndex];
      if (!photo) return;
      const imageId = workbook.addImage({ base64: photo, extension: "jpeg" });
      const col = columns.findIndex(([key]) => key === (photoIndex === 0 ? "foto1" : "foto2"));
      sheet.addImage(imageId, { tl: { col, row: row.number - 1 }, ext: { width: 120, height: 85 }, editAs: "oneCell" });
    });
  }
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE6E0DA" } },
        left: { style: "thin", color: { argb: "FFE6E0DA" } },
        bottom: { style: "thin", color: { argb: "FFE6E0DA" } },
        right: { style: "thin", color: { argb: "FFE6E0DA" } },
      };
      cell.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && rowNumber % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFAF8F5" } };
    });
  });
  const blob = new Blob([await workbook.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `BIE_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function downloadWorkbook(workbook, fileName) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });
}

function allUsableRecords() {
  return records
    .map(cleanRecord)
    .filter((record) => [record.cliente, record.edificio, record.cantidad, record.ubicacion, record.numeroSerie].some((value) => safeText(value).trim()))
    .sort((a, b) =>
      safeText(a.cliente).localeCompare(safeText(b.cliente), "es", { numeric: true, sensitivity: "base" }) ||
      safeText(a.edificio).localeCompare(safeText(b.edificio), "es", { numeric: true, sensitivity: "base" }) ||
      safeText(a.cantidad).localeCompare(safeText(b.cantidad), "es", { numeric: true, sensitivity: "base" })
    );
}

function isCorrectiveRecord(record) {
  const estado = safeText(record.estadoEquipo).trim().toUpperCase();
  return Boolean(
    (record.defectos || []).length ||
    (record.checklist || []).length ||
    (record.observaciones || "").trim() ||
    (estado && estado !== "OK" && estado !== "CORRECTO" && estado !== "BUENO")
  );
}

async function downloadCorrectivas() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const correctivas = allUsableRecords().filter(isCorrectiveRecord);
  if (!correctivas.length) return alert("No hay registros con correctivas para descargar.");

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BIE";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Correctivas");
  const columns = [
    ["cliente", "Cliente", 22],
    ["edificio", "Edificio", 18],
    ["cantidad", "Número SYCo", 18],
    ["ubicacion", "Ubicación", 42],
    ["modelo", "Modelo", 14],
    ["fabricanteMarca", "Fabricante/Marca", 22],
    ["numeroSerie", "Nº serie", 18],
    ["defectos", "Defectos encontrados", 42],
    ["checklist", "Checklist marcado", 42],
    ["estadoEquipo", "Estado equipo", 18],
    ["observaciones", "Observaciones", 42],
    ["senal", "Señal", 16],
    ["foto1", "Foto 1", 22],
    ["foto2", "Foto 2", 22],
    ["visto", "Visto", 10],
  ];
  sheet.columns = columns.map(([key, header, width]) => ({ key, header, width }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FF0F3A5F" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFBFDBFE" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(1).height = 30;

  for (const record of correctivas) {
    const selectedChecklist = record.checklist || [];
    const checklistText = selectedChecklist
      .map((code) => {
        const option = checklistOptions.find(([optionCode]) => optionCode === code);
        return option ? `${code} - ${option[1]}` : code;
      })
      .join(" / ");
    const row = sheet.addRow({
      ...record,
      defectos: (record.defectos || []).join(" / "),
      checklist: checklistText,
      foto1: record.photos[0] ? "Foto 1" : "",
      foto2: record.photos[1] ? "Foto 2" : "",
      visto: record.visto ? "Sí" : "No",
    });
    if (record.photos[0] || record.photos[1]) row.height = 92;
    [0, 1].forEach((photoIndex) => {
      const photo = record.photos[photoIndex];
      if (!photo) return;
      const imageId = workbook.addImage({ base64: photo, extension: "jpeg" });
      const col = columns.findIndex(([key]) => key === (photoIndex === 0 ? "foto1" : "foto2"));
      sheet.addImage(imageId, { tl: { col, row: row.number - 1 }, ext: { width: 120, height: 85 }, editAs: "oneCell" });
    });
  }

  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E2F0" } },
        left: { style: "thin", color: { argb: "FFD8E2F0" } },
        bottom: { style: "thin", color: { argb: "FFD8E2F0" } },
        right: { style: "thin", color: { argb: "FFD8E2F0" } },
      };
      cell.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && rowNumber % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FBFF" } };
    });
  });

  return downloadWorkbook(workbook, `Correctivas_BIE_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function copyCell(sourceCell, targetCell) {
  targetCell.value = sourceCell.value;
  targetCell.style = JSON.parse(JSON.stringify(sourceCell.style || {}));
  if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
}

function copyTemplateSheet(workbook, sourceSheet, name) {
  const sheet = workbook.addWorksheet(name);
  sheet.pageSetup = JSON.parse(JSON.stringify(sourceSheet.pageSetup || {}));
  sheet.pageMargins = JSON.parse(JSON.stringify(sourceSheet.pageMargins || {}));
  sheet.headerFooter = JSON.parse(JSON.stringify(sourceSheet.headerFooter || {}));
  sheet.views = JSON.parse(JSON.stringify(sourceSheet.views || []));
  sheet.properties = JSON.parse(JSON.stringify(sourceSheet.properties || {}));
  for (let col = 1; col <= sourceSheet.columnCount; col += 1) {
    const sourceCol = sourceSheet.getColumn(col);
    const targetCol = sheet.getColumn(col);
    targetCol.width = sourceCol.width;
    targetCol.hidden = sourceCol.hidden;
    targetCol.outlineLevel = sourceCol.outlineLevel;
  }
  for (let rowNumber = 1; rowNumber <= sourceSheet.rowCount; rowNumber += 1) {
    const sourceRow = sourceSheet.getRow(rowNumber);
    const targetRow = sheet.getRow(rowNumber);
    targetRow.height = sourceRow.height;
    for (let col = 1; col <= sourceSheet.columnCount; col += 1) {
      copyCell(sourceRow.getCell(col), targetRow.getCell(col));
    }
  }
  (sourceSheet.model?.merges || []).forEach((range) => sheet.mergeCells(range));
  return sheet;
}

function checklistPages() {
  const rows = allUsableRecords();
  const pages = [];
  for (let start = 0; start < rows.length; start += 20) {
    pages.push(rows.slice(start, start + 20));
  }
  return pages;
}

function pageTitle(rows, index) {
  const cliente = safeText(rows.find((record) => safeText(record.cliente).trim())?.cliente).trim();
  const edificios = Array.from(new Set(rows.map((record) => safeText(record.edificio).trim()).filter(Boolean)));
  const edificio = edificios.length === 1 ? edificios[0] : edificios.length > 1 ? "Varios edificios" : "";
  return [cliente, edificio].filter(Boolean).join(" - ") || `Checklist ${index}`;
}

function safeSheetName(value, index) {
  const name = safeText(value).replace(/[\\/*?:[\]]/g, " ").replace(/\s+/g, " ").trim() || `Checklist ${index}`;
  return name.slice(0, 25) + (index > 1 ? ` ${index}` : "");
}

function clearChecklistRows(sheet) {
  for (let row = 8; row <= 27; row += 1) {
    for (let col = 1; col <= 40; col += 1) {
      sheet.getRow(row).getCell(col).value = "";
    }
  }
}

function checklistObservation(record) {
  return [
    (record.defectos || []).join(" / "),
    safeText(record.observaciones).trim(),
  ].filter(Boolean).join(" / ");
}

function fillChecklistSheet(sheet, title, rows) {
  sheet.getCell("E2").value = safeText(rows.find((record) => safeText(record.cliente).trim())?.cliente);
  sheet.getCell("E3").value = safeText(rows.find((record) => safeText(record.edificio).trim())?.edificio);
  clearChecklistRows(sheet);
  rows.forEach((record, index) => {
    const rowNumber = 8 + index;
    const row = sheet.getRow(rowNumber);
    const model = safeText(record.modelo).replace(/\s+/g, "").toUpperCase();
    const selectedChecklist = new Set(record.checklist || []);
    row.getCell(1).value = safeText(record.cantidad) || `SYCo ${index + 1}`;
    row.getCell(2).value = model === "25" ? "X" : "";
    row.getCell(3).value = model === "45" ? "X" : "";
    row.getCell(4).value = model === "25T45" || model === "25+T45" ? "X" : "";
    row.getCell(5).value = safeText(record.fabricanteMarca);
    row.getCell(6).value = safeText(record.fechaFabricacion);
    row.getCell(7).value = safeText(record.fechaProximoRetimbrado);
    row.getCell(8).value = safeText(record.operacionesRealizadas);
    row.getCell(9).value = safeText(record.cadu20A);
    const codeColumns = {
      A: 10, B: 11, C: 12, D: 13, E: 14, F: 15, G: 16, H: 17, I: 18, J: 19,
      K: 20, L: 21, M: 22, N: 23, O: 24, P: 25, Q: 26, S: 27, T: 28, U: 29,
      V: 30, R1: 31, R2: 32, R3: 33, R4: 34, R5: 35, R6: 36, R7: 37, W: 38,
    };
    Object.entries(codeColumns).forEach(([code, col]) => {
      row.getCell(col).value = selectedChecklist.has(code) ? "X" : "";
    });
    row.getCell(39).value = safeText(record.ubicacion);
    row.getCell(40).value = checklistObservation(record);
    row.getCell(40).alignment = { ...(row.getCell(40).alignment || {}), wrapText: true, vertical: "top" };
    if (row.getCell(40).value) row.height = Math.max(row.height || 18, 30);
  });
}

async function downloadChecklist() {
  if (!window.ExcelJS) return alert("No se ha cargado el generador de Excel.");
  const pages = checklistPages();
  if (!pages.length) return alert("No hay registros para generar el checklist.");
  const response = await fetch(CHECKLIST_TEMPLATE_URL);
  if (!response.ok) return alert("No se ha podido cargar la plantilla del checklist.");

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const templateSheet = workbook.worksheets[0];
  pages.forEach((rows, index) => {
    const title = pageTitle(rows, index + 1);
    const sheet = index === 0 ? templateSheet : copyTemplateSheet(workbook, templateSheet, safeSheetName(title, index + 1));
    sheet.name = safeSheetName(title, index + 1);
    fillChecklistSheet(sheet, title, rows);
  });

  return downloadWorkbook(workbook, `Checklist_BIE_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function bindEvents() {
  $("openListBtn").addEventListener("click", () => showView("list"));
  $("newRecordBtn").addEventListener("click", () => openForm());
  $("newRecordFromListBtn").addEventListener("click", () => openForm());
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
  $("downloadCorrectivasBtn").addEventListener("click", downloadCorrectivas);
  $("downloadChecklistBtn").addEventListener("click", downloadChecklist);
  $("deleteAllBtn").addEventListener("click", deleteAllRecords);
  $("deleteAllFromTableBtn").addEventListener("click", deleteAllRecords);
  $("viewTableFromFormBtn").addEventListener("click", () => showView("list"));
  $("importExcelBtn").addEventListener("click", () => $("importExcelInput").click());
  $("importExcelInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      $("importStatus").textContent = "Importando Excel...";
      await importExcelFile(file);
      renderTable();
    } catch (error) {
      console.error(error);
      $("importStatus").textContent = "No se ha podido importar el Excel.";
      alert("No se ha podido importar el Excel. Revisa que tenga el mismo formato.");
    } finally {
      event.target.value = "";
    }
  });
  ["filterCliente", "filterEdificio", "filterNumero", "filterSerie", "sortOrder", "seenFilter"].forEach((id) => {
    $(id).addEventListener("input", renderTable);
    $(id).addEventListener("change", renderTable);
  });
  $("recordForm").addEventListener("submit", saveForm);
  $("deleteBtn").addEventListener("click", deleteCurrent);
  [0, 1].forEach((index) => {
    $(`photoInput${index + 1}`).addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        setPhotoPreview(index, await resizePhoto(file));
      } catch {
        alert("No he podido cargar esa foto. Prueba con otra imagen.");
      } finally {
        event.target.value = "";
      }
    });
    $(`deletePhoto${index + 1}`).addEventListener("click", () => setPhotoPreview(index, ""));
  });
  document.querySelectorAll("[data-back]").forEach((button) => button.addEventListener("click", () => showView(button.dataset.back)));
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./sw.js").catch(() => {}));
}

async function init() {
  fillYearLists();
  await loadRecords();
  bindEvents();
  updateStats();
}

init();
