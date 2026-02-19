const ARES_BASE = "https://ares.gov.cz/ekonomicke-subjekty-v-be/rest";

const form = document.getElementById("powerForm");
const layout = document.querySelector(".layout");
const previewShell = document.querySelector(".preview-shell");
const message = document.getElementById("formMessage");
const actionsError = document.getElementById("actionsError");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const togglePreviewBtn = document.getElementById("togglePreviewBtn");
const ARIAL_REGULAR_URL = "./assets/fonts/arial.ttf";
const ARIAL_BOLD_URL = "./assets/fonts/arialbd.ttf";
const ARIAL_ITALIC_URL = "./assets/fonts/ariali.ttf";
const ARIAL_BOLDITALIC_URL = "./assets/fonts/arialbi.ttf";
let arialFontPromise = null;

const fieldIds = [
  "grantorIco",
  "grantorId",
  "grantorName",
  "grantorAddress",
  "attorneyIco",
  "attorneyId",
  "attorneyName",
  "attorneyAddress",
  "vehicleBrand",
  "vehicleModel",
  "vehicleSpz",
  "vehicleVin",
  "signPlace",
  "signDate"
];

const fields = Object.fromEntries(fieldIds.map((id) => [id, document.getElementById(id)]));

initialize();

function initialize() {
  fields.signDate.value = formatDate(new Date());

  setupSubjectLookup("grantor");
  setupSubjectLookup("attorney");

  fields.vehicleVin.addEventListener("input", () => {
    fields.vehicleVin.value = sanitizeVin(fields.vehicleVin.value);
  });
  fields.vehicleSpz.addEventListener("input", () => {
    fields.vehicleSpz.value = sanitizeSpz(fields.vehicleSpz.value);
  });
  fields.grantorIco.addEventListener("input", () => {
    fields.grantorIco.value = sanitizeIco(fields.grantorIco.value);
  });
  fields.attorneyIco.addEventListener("input", () => {
    fields.attorneyIco.value = sanitizeIco(fields.attorneyIco.value);
  });
  fields.signDate.addEventListener("input", () => {
    fields.signDate.value = sanitizeDate(fields.signDate.value);
  });

  form.addEventListener("input", () => {
    clearMessages();
    updatePreview();
  });
  form.addEventListener("change", updatePreview);

  document.getElementById("detectPlaceBtn").addEventListener("click", () => detectCurrentPlace(false));
  downloadPdfBtn.addEventListener("click", handlePdfDownload);
  setupPreviewToggle();
  setDefaultActionsChecked();

  updatePreview();
}

function setupPreviewToggle() {
  if (!togglePreviewBtn || !previewShell || !layout) {
    return;
  }

  togglePreviewBtn.addEventListener("click", () => {
    const isHidden = !previewShell.classList.contains("hidden");
    previewShell.classList.toggle("hidden", isHidden);
    layout.classList.toggle("layout-preview-hidden", isHidden);
    togglePreviewBtn.textContent = isHidden ? "Zobrazit náhled" : "Skrýt náhled";
    togglePreviewBtn.setAttribute("aria-expanded", String(!isHidden));
  });
}

function setDefaultActionsChecked() {
  document.querySelectorAll('input[name="actions"]').forEach((checkbox) => {
    checkbox.checked = true;
  });
}

function setupSubjectLookup(prefix) {
  const lookup = document.getElementById(`${prefix}Lookup`);
  const suggestions = document.getElementById(`${prefix}Suggestions`);
  const ico = document.getElementById(`${prefix}Ico`);
  const id = document.getElementById(`${prefix}Id`);
  const name = document.getElementById(`${prefix}Name`);
  const address = document.getElementById(`${prefix}Address`);

  let searchTimer = null;
  let requestId = 0;
  let activeItems = [];

  lookup.addEventListener("input", () => {
    const query = lookup.value.trim();
    clearTimeout(searchTimer);
    hideSuggestions(suggestions);

    if (query.length < 2) {
      return;
    }

    searchTimer = setTimeout(async () => {
      requestId += 1;
      const currentRequest = requestId;

      const found = await searchAres(query);
      if (currentRequest !== requestId) {
        return;
      }

      activeItems = found;
      renderSuggestions({
        suggestions,
        items: activeItems,
        onSelect: (item) => {
          applySubject(item, { lookup, ico, id, name, address });
          hideSuggestions(suggestions);
          updatePreview();
        }
      });
    }, 260);
  });

  ico.addEventListener("blur", async () => {
    const cleanIco = sanitizeIco(ico.value);
    ico.value = cleanIco;
    if (!/^\d{8}$/.test(cleanIco)) {
      return;
    }

    const detail = await fetchByIco(cleanIco);
    if (!detail) {
      return;
    }

    applySubject(detail, { lookup, ico, id, name, address });
    updatePreview();
  });

  document.addEventListener("click", (event) => {
    if (!event.target.closest(`#${lookup.id}`) && !event.target.closest(`#${suggestions.id}`)) {
      hideSuggestions(suggestions);
    }
  });
}

async function searchAres(query) {
  try {
    const payload = { start: 0, pocet: 8 };
    if (/^\d{8}$/.test(query)) {
      payload.ico = [query];
    } else {
      payload.obchodniJmeno = query;
    }

    const response = await fetch(`${ARES_BASE}/ekonomicke-subjekty/vyhledat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    return data.ekonomickeSubjekty || [];
  } catch (_) {
    return [];
  }
}

async function fetchByIco(ico) {
  try {
    const response = await fetch(`${ARES_BASE}/ekonomicke-subjekty/${ico}`);
    if (!response.ok) {
      return null;
    }
    return await response.json();
  } catch (_) {
    return null;
  }
}

function applySubject(subject, refs) {
  refs.ico.value = subject.ico || refs.ico.value || "";
  refs.name.value = subject.obchodniJmeno || refs.name.value || "";
  refs.address.value = formatAddress(subject.sidlo) || refs.address.value || "";

  if (!refs.id.value.trim()) {
    refs.id.value = subject.ico || "";
  }

  const lookupName = subject.obchodniJmeno || "";
  const lookupIco = subject.ico ? ` (${subject.ico})` : "";
  refs.lookup.value = `${lookupName}${lookupIco}`.trim();
}

function renderSuggestions({ suggestions, items, onSelect }) {
  suggestions.innerHTML = "";

  if (!items.length) {
    hideSuggestions(suggestions);
    return;
  }

  items.forEach((item) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "suggestion-item";
    btn.innerHTML = `
      <span class="suggestion-title">${escapeHtml(item.obchodniJmeno || "Bez názvu")}</span>
      <span class="suggestion-subtitle">IČO ${escapeHtml(item.ico || "-")} • ${escapeHtml(formatAddress(item.sidlo) || "adresa nedostupná")}</span>
    `;
    btn.addEventListener("click", () => onSelect(item));
    suggestions.appendChild(btn);
  });

  suggestions.classList.add("visible");
}

function hideSuggestions(suggestions) {
  suggestions.classList.remove("visible");
  suggestions.innerHTML = "";
}

function formatAddress(sidlo) {
  if (!sidlo) {
    return "";
  }

  if (sidlo.textovaAdresa) {
    return formatPostCode(sidlo.textovaAdresa);
  }

  const parts = [];
  const street = sidlo.nazevUlice ? sidlo.nazevUlice : "";
  const house = sidlo.cisloDomovni ? String(sidlo.cisloDomovni) : "";
  const orient = sidlo.cisloOrientacni ? `/${sidlo.cisloOrientacni}` : "";
  const city = sidlo.nazevObce ? sidlo.nazevObce : "";
  const psc = sidlo.psc ? formatPscNumber(sidlo.psc) : "";

  if (street || house) {
    parts.push(`${street} ${house}${orient}`.trim());
  }
  if (psc || city) {
    parts.push(`${psc} ${city}`.trim());
  }

  return parts.join(", ").trim();
}

function formatPscNumber(psc) {
  const clean = String(psc).replace(/\D/g, "");
  if (clean.length !== 5) {
    return clean;
  }
  return `${clean.slice(0, 3)} ${clean.slice(3)}`;
}

function formatPostCode(text) {
  return text.replace(/\b(\d{3})(\d{2})\b/g, "$1 $2");
}

function updatePreview() {
  setText("previewGrantorName", valueOrBlank(fields.grantorName.value));
  setText("previewGrantorId", valueOrBlank(fields.grantorId.value || fields.grantorIco.value));
  setText("previewGrantorAddress", valueOrBlank(fields.grantorAddress.value));

  setText("previewAttorneyName", valueOrBlank(fields.attorneyName.value));
  setText("previewAttorneyId", valueOrBlank(fields.attorneyId.value || fields.attorneyIco.value));
  setText("previewAttorneyAddress", valueOrBlank(fields.attorneyAddress.value));

  const typeModel = [fields.vehicleBrand.value.trim(), fields.vehicleModel.value.trim()].filter(Boolean).join(" ");
  setText("previewVehicleTypeModel", valueOrBlank(typeModel));

  const spz = fields.vehicleSpz.value.trim();
  setText("previewVehicleSpz", valueOrBlank(spz));
  const previewSpzRow = document.getElementById("previewSpzRow");
  if (previewSpzRow) {
    previewSpzRow.style.display = spz ? "grid" : "none";
  }

  setText("previewVehicleVin", valueOrBlank(fields.vehicleVin.value));

  setText("previewSignPlace", valueOrDots(fields.signPlace.value, "Praha"));
  setText("previewSignDate", valueOrDots(fields.signDate.value));

  const selectedActions = getSelectedActions();
  const actionsList = document.getElementById("previewActions");
  actionsList.innerHTML = "";

  if (!selectedActions.length) {
    const li = document.createElement("li");
    li.textContent = "....................................";
    actionsList.appendChild(li);
    return;
  }

  selectedActions.forEach((action) => {
    const li = document.createElement("li");
    li.textContent = action;
    actionsList.appendChild(li);
  });
}

function getSelectedActions() {
  return Array.from(document.querySelectorAll('input[name="actions"]:checked')).map((item) => item.value);
}

async function handlePdfDownload() {
  clearMessages();
  updatePreview();

  if (!form.checkValidity()) {
    form.reportValidity();
    return;
  }

  if (!isValidDate(fields.signDate.value)) {
    setMessage("Datum musí být ve formátu dd.mm.rrrr.", true);
    fields.signDate.focus();
    return;
  }

  const selectedActions = getSelectedActions();
  if (!selectedActions.length) {
    actionsError.textContent = "Vyberte alespoň jeden úkon.";
    return;
  }
  actionsError.textContent = "";

  if (!window.jspdf || !window.jspdf.jsPDF) {
    setMessage("Nepodařilo se načíst knihovnu pro PDF. Obnovte stránku.", true);
    return;
  }

  const fileDate = fields.signDate.value.replace(/\./g, "-");
  const filename = `plna-moc-prevod-vozidla-${fileDate || "dokument"}.pdf`;

  downloadPdfBtn.disabled = true;
  setMessage("Generuji PDF...");

  try {
    const pdf = await buildTextPdf();
    pdf.save(filename);

    setMessage("PDF bylo staženo.");
  } catch (_) {
    setMessage("Vytvoření PDF se nezdařilo. Zkuste to znovu.", true);
  } finally {
    downloadPdfBtn.disabled = false;
  }
}

async function buildTextPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const arial = await ensureArialFontData();
  registerArialFont(doc, arial);

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  const marginLeft = 14;
  const marginRight = 14;
  const marginTop = 18;
  const marginBottom = 14;
  const labelX = marginLeft;
  const valueX = 62;
  const valueEndX = pageWidth - marginRight;
  const contentWidth = pageWidth - marginLeft - marginRight;
  const rowSpacing = 8;
  const lineHeight = 5;

  let y = marginTop;

  doc.setFont("Arial", "bold");
  doc.setFontSize(19);
  doc.text("Plná moc", pageWidth / 2, y, { align: "center" });
  y += 12;

  const ensureSpace = (requiredHeight) => {
    if (y + requiredHeight <= pageHeight - marginBottom) {
      return;
    }
    doc.addPage();
    y = marginTop;
  };

  const drawDottedLine = (startX, endX, atY) => {
    doc.setDrawColor(120);
    doc.setLineWidth(0.2);
    doc.setLineDashPattern([0.8, 0.8], 0);
    doc.line(startX, atY, endX, atY);
    doc.setLineDashPattern([], 0);
  };

  const drawLabeledValueRow = (label, rawValue) => {
    const value = valueOrDots(rawValue);
    const wrapped = doc.splitTextToSize(value, valueEndX - valueX);
    ensureSpace(rowSpacing + (wrapped.length - 1) * lineHeight);

    doc.setFont("Arial", "normal");
    doc.setFontSize(12);
    doc.text(label, labelX, y);
    doc.text(wrapped, valueX, y);

    for (let i = 0; i < wrapped.length; i += 1) {
      drawDottedLine(valueX, valueEndX, y + 1.5 + i * lineHeight);
    }

    y += rowSpacing + (wrapped.length - 1) * lineHeight;
  };

  const drawWrappedParagraph = (text, bottomSpace = 3) => {
    const wrapped = doc.splitTextToSize(text, contentWidth);
    ensureSpace(wrapped.length * lineHeight + bottomSpace);
    doc.setFont("Arial", "normal");
    doc.setFontSize(12);
    doc.text(wrapped, labelX, y);
    y += wrapped.length * lineHeight + bottomSpace;
  };

  const drawBullet = (text) => {
    const wrapped = doc.splitTextToSize(`•  ${text}`, contentWidth - 8);
    ensureSpace(wrapped.length * lineHeight + 1);
    doc.setFont("Arial", "normal");
    doc.setFontSize(12);
    doc.text(wrapped, labelX + 6, y);
    y += wrapped.length * lineHeight + 1;
  };

  drawLabeledValueRow("Já, níže podepsaný", fields.grantorName.value.trim());
  drawLabeledValueRow("se sídlem (bytem)", fields.grantorAddress.value.trim());
  drawLabeledValueRow("IČ/RČ", (fields.grantorId.value || fields.grantorIco.value).trim());

  ensureSpace(10);
  doc.setFont("Arial", "bold");
  doc.setFontSize(12);
  doc.text("(dále jako „zmocnitel“)", labelX, y);
  y += 10;

  ensureSpace(9);
  doc.setFont("Arial", "italic");
  doc.setFontSize(12);
  doc.text("zmocňuji", labelX, y);
  drawDottedLine(labelX, labelX + doc.getTextWidth("zmocňuji"), y + 0.8);
  y += 10;

  drawLabeledValueRow("Společnost/osobu", fields.attorneyName.value.trim());
  drawLabeledValueRow("se sídlem/bytem", fields.attorneyAddress.value.trim());
  drawLabeledValueRow("IČ/RČ", (fields.attorneyId.value || fields.attorneyIco.value).trim());

  ensureSpace(10);
  doc.setFont("Arial", "bold");
  doc.setFontSize(12);
  doc.text("(dále jako „zmocněnec“)", labelX, y);
  y += 10;

  drawWrappedParagraph("Aby mě v plném rozsahu zastupoval(a) a činil(a) veškeré úkony související s:", 2);
  getSelectedActions().forEach((action) => drawBullet(action));
  y += 8;

  const typeModel = [fields.vehicleBrand.value.trim(), fields.vehicleModel.value.trim()].filter(Boolean).join(" ");
  drawLabeledValueRow("Typ a model vozu", typeModel);

  const spz = fields.vehicleSpz.value.trim();
  if (spz) {
    drawLabeledValueRow("SPZ / RZ", spz);
  }

  drawLabeledValueRow("Číslo karoserie - VIN", fields.vehicleVin.value.trim());
  y += 5;

  drawWrappedParagraph(
    "tato plná moc není ve výše uvedeném rozsahu ničím omezena. Zplnomocnění v plném rozsahu zmocněnec přijímá.",
    20
  );

  ensureSpace(12);
  doc.setFont("Arial", "normal");
  doc.setFontSize(12);

  let leftX = labelX;
  doc.text("V ", leftX, y);
  leftX += doc.getTextWidth("V ");
  doc.setFont("Arial", "bold");
  doc.text(valueOrDots(fields.signPlace.value, "Praha"), leftX, y);
  leftX += doc.getTextWidth(valueOrDots(fields.signPlace.value, "Praha"));

  doc.setFont("Arial", "normal");
  doc.text(" dne ", leftX, y);
  leftX += doc.getTextWidth(" dne ");

  doc.setFont("Arial", "bold");
  doc.text(valueOrDots(fields.signDate.value), leftX, y);

  const signatureLineStartX = pageWidth - marginRight - 66;
  drawDottedLine(signatureLineStartX, valueEndX, y + 1.2);

  doc.setFont("Arial", "italic");
  doc.setFontSize(10.5);
  doc.text("(zmocnitel)", (signatureLineStartX + valueEndX) / 2, y + 6, { align: "center" });

  return doc;
}

async function ensureArialFontData() {
  if (window.__ARIAL_FONT_DATA) {
    return window.__ARIAL_FONT_DATA;
  }

  if (!arialFontPromise) {
    arialFontPromise = Promise.all([
      fetchFontAsBase64(ARIAL_REGULAR_URL),
      fetchFontAsBase64(ARIAL_BOLD_URL),
      fetchFontAsBase64(ARIAL_ITALIC_URL),
      fetchFontAsBase64(ARIAL_BOLDITALIC_URL)
    ]).then(([regular, bold, italic, bolditalic]) => ({
      regular,
      bold,
      italic,
      bolditalic
    }));
  }

  return arialFontPromise;
}

async function fetchFontAsBase64(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Font download failed: ${response.status}`);
  }

  const buffer = await response.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";

  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }

  return btoa(binary);
}

function registerArialFont(doc, fontData) {
  doc.addFileToVFS("Arial-Regular.ttf", fontData.regular);
  doc.addFileToVFS("Arial-Bold.ttf", fontData.bold);
  doc.addFileToVFS("Arial-Italic.ttf", fontData.italic);
  doc.addFileToVFS("Arial-BoldItalic.ttf", fontData.bolditalic || fontData.italic);

  doc.addFont("Arial-Regular.ttf", "Arial", "normal");
  doc.addFont("Arial-Bold.ttf", "Arial", "bold");
  doc.addFont("Arial-Italic.ttf", "Arial", "italic");
  doc.addFont("Arial-BoldItalic.ttf", "Arial", "bolditalic");
  doc.setFont("Arial", "normal");
  doc.setCharSpace(0);
}

function detectCurrentPlace(silent) {
  if (!navigator.geolocation) {
    if (!silent) {
      setMessage("Tento prohlížeč nepodporuje geolokaci.");
    }
    return;
  }

  navigator.geolocation.getCurrentPosition(
    async (position) => {
      const city = await getCityByCoords(position.coords.latitude, position.coords.longitude);
      if (city) {
        fields.signPlace.value = city;
        updatePreview();
        if (!silent) {
          setMessage(`Místo podpisu předvyplněno: ${city}`);
        }
      } else if (!silent) {
        setMessage("Místo nebylo možné určit, ponecháno ruční vyplnění.");
      }
    },
    () => {
      if (!silent) {
        setMessage("Přístup k lokaci byl zamítnut, ponecháno ruční vyplnění.");
      }
    },
    { enableHighAccuracy: false, timeout: 8000, maximumAge: 600000 }
  );
}

async function getCityByCoords(lat, lon) {
  try {
    const url = new URL("https://nominatim.openstreetmap.org/reverse");
    url.searchParams.set("format", "jsonv2");
    url.searchParams.set("lat", String(lat));
    url.searchParams.set("lon", String(lon));
    url.searchParams.set("accept-language", "cs");

    const response = await fetch(url.toString());
    if (!response.ok) {
      return "";
    }

    const data = await response.json();
    const address = data.address || {};
    return address.city || address.town || address.village || address.municipality || "";
  } catch (_) {
    return "";
  }
}

function clearMessages() {
  message.textContent = "";
  message.style.color = "";
  actionsError.textContent = "";
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.style.color = isError ? "#ad2e24" : "#3b4a5a";
}

function valueOrDots(value, fallback = "....................................") {
  const text = value.trim();
  return text || fallback;
}

function valueOrBlank(value) {
  return value.trim();
}

function setText(id, value) {
  const node = document.getElementById(id);
  if (node) {
    node.textContent = value;
  }
}

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function sanitizeIco(value) {
  return value.replace(/\D/g, "").slice(0, 8);
}

function sanitizeVin(value) {
  return value.toUpperCase().replace(/[^A-HJ-NPR-Z0-9]/g, "").slice(0, 17);
}

function sanitizeSpz(value) {
  return value.toUpperCase().replace(/[^0-9A-Z]/g, "").slice(0, 10);
}

function sanitizeDate(value) {
  const clean = value.replace(/[^\d.]/g, "").slice(0, 10);
  const digits = clean.replace(/\./g, "");
  if (digits.length <= 2) {
    return digits;
  }
  if (digits.length <= 4) {
    return `${digits.slice(0, 2)}.${digits.slice(2)}`;
  }
  return `${digits.slice(0, 2)}.${digits.slice(2, 4)}.${digits.slice(4, 8)}`;
}

function isValidDate(value) {
  const match = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(value);
  if (!match) {
    return false;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(year, month - 1, day);

  return (
    date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
  );
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
