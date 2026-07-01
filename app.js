const DB_NAME = "bie-db-v1";
const DB_VERSION = 1;
const STORE_NAME = "state";

const defectOptions = [
  "Manguera rota",
  "Hay un obstáculo",
  "Cristal roto",
  "Manómetro mal",
  "Válvula mal",
  "Lanza mal",
  "Devanadera mal",
  "Sin presión",
  "Sin señal",
  "Señal caducada",
  "Armario roto y presión baja",
];

const fields = [
  "cliente",
  "edificio",
  "cantidad",
  "ubicacion",
  "modelo",
  "numeroSerie",
  "fechaFabricacion",
  "fechaProximoRetimbrado",
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
      if (text.endsWith(".")) return text.slice(0, -1);
      return text;
    })
    .filter((defect) => defectOptions.includes(defect));
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
    numeroSerie: safeText(record.numeroSerie),
    fechaFabricacion: normalizeYearValue(record.fechaFabricacion),
    fechaProximoRetimbrado: normalizeYearValue(record.fechaProximoRetimbrado),
    observaciones: safeText(record.observaciones),
    senal: safeText(record.senal),
    defectos: normalizeDefects(record.defectos),
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

function rowToImportedRecord(rowValues, index) {
  const values = [];
  for (let col = 1; col <= 10; col += 1) values[col] = excelCellToText(rowValues[col]);
  return cleanRecord({
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    cantidad: values[2],
    edificio: values[3],
    ubicacion: values[4],
    modelo: values[6],
    numeroSerie: values[7],
    fechaFabricacion: values[9],
    fechaProximoRetimbrado: values[10],
    origen: "importado",
  });
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
    body.innerHTML = `<tr><td colspan="15">No hay registros con ese filtro.</td></tr>`;
    return;
  }
  for (const record of rows) {
    const defects = record.defectos.length ? record.defectos.join(" / ") : "-";
    const photo1 = record.photos[0] ? `<img class="tablePhoto" src="${record.photos[0]}" alt="Foto 1">` : `<span class="noPhoto">—</span>`;
    const photo2 = record.photos[1] ? `<img class="tablePhoto" src="${record.photos[1]}" alt="Foto 2">` : `<span class="noPhoto">—</span>`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${safeText(record.cliente) || "-"}</td>
      <td>${safeText(record.edificio) || "-"}</td>
      <td><strong>${safeText(record.cantidad) || "-"}</strong></td>
      <td>${safeText(record.ubicacion) || "-"}</td>
      <td>${safeText(record.modelo) || "-"}</td>
      <td>${safeText(record.numeroSerie) || "-"}</td>
      <td>${safeText(record.fechaFabricacion) || "-"}</td>
      <td>${safeText(record.fechaProximoRetimbrado) || "-"}</td>
      <td>${safeText(record.observaciones) || "-"}</td>
      <td>${safeText(record.senal) || "-"}</td>
      <td>${defects}</td>
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
  const photos = Array.isArray(record?.photos) ? record.photos : ["", ""];
  setPhotoPreview(0, photos[0]);
  setPhotoPreview(1, photos[1]);
  showView("form");
}

function collectForm() {
  const record = { id: $("recordId").value || createId(), origen: $("recordId").value ? "editado" : "manual" };
  for (const key of fields) record[key] = $(key).value.trim();
  record.defectos = Array.from($("defectsList").querySelectorAll("input:checked")).map((input) => input.value);
  record.photos = [currentPhotos[0] || "", currentPhotos[1] || ""];
  record.visto = $("visto").checked;
  return cleanRecord(record);
}

async function saveForm(event) {
  event.preventDefault();
  const record = collectForm();
  const index = records.findIndex((item) => item.id === record.id);
  if (index >= 0) records[index] = record;
  else records.unshift(record);
  await saveRecords();
  showView("list");
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
  const imported = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = rowToImportedRecord(row.values, rowNumber);
    const hasData = [record.edificio, record.cantidad, record.ubicacion, record.modelo, record.numeroSerie].some((value) => safeText(value).trim());
    if (!hasData) return;
    imported.push(record);
  });
  if (!imported.length) {
    $("importStatus").textContent = "No se encontraron registros para importar.";
    return alert("No se encontraron registros para importar.");
  }
  records = [...imported, ...records];
  await saveRecords();
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
    ["numeroSerie", "Nº serie", 18],
    ["fechaFabricacion", "Fecha / año fabricación", 22],
    ["fechaProximoRetimbrado", "Fecha retimbrado", 20],
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
    ["defectoSinSenal", "Sin señal", 16],
    ["defectoSenalCaducada", "Señal caducada", 20],
    ["defectoArmarioRotoPresionBaja", "Armario roto y presión baja", 30],
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
    const row = sheet.addRow({
      ...record,
      defectos: selected.join(" / "),
      defectoMangueraRota: defectFlag(selected, "Manguera rota"),
      defectoObstaculo: defectFlag(selected, "Hay un obstáculo"),
      defectoCristalRoto: defectFlag(selected, "Cristal roto"),
      defectoManometroMal: defectFlag(selected, "Manómetro mal"),
      defectoValvulaMal: defectFlag(selected, "Válvula mal"),
      defectoLanzaMal: defectFlag(selected, "Lanza mal"),
      defectoDevanaderaMal: defectFlag(selected, "Devanadera mal"),
      defectoSinPresion: defectFlag(selected, "Sin presión"),
      defectoSinSenal: defectFlag(selected, "Sin señal"),
      defectoSenalCaducada: defectFlag(selected, "Señal caducada"),
      defectoArmarioRotoPresionBaja: defectFlag(selected, "Armario roto y presión baja"),
      foto1: record.photos[0] ? "Foto 1" : "",
      foto2: record.photos[1] ? "Foto 2" : "",
      visto: record.visto ? "Sí" : "No",
    });
    if (record.photos[0] || record.photos[1]) row.height = 92;
    [0, 1].forEach((photoIndex) => {
      const photo = record.photos[photoIndex];
      if (!photo) return;
      const imageId = workbook.addImage({ base64: photo, extension: "jpeg" });
      const col = photoIndex === 0 ? 22 : 23;
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

function bindEvents() {
  $("openListBtn").addEventListener("click", () => showView("list"));
  $("newRecordBtn").addEventListener("click", () => openForm());
  $("newRecordFromListBtn").addEventListener("click", () => openForm());
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
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
