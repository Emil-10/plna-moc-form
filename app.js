const ARES_BASE = "https://ares.gov.cz/ekonomicke-subjekty-v-be/rest";
const MORPHODITA_BASE = "https://lindat.mff.cuni.cz/services/morphodita/api";
const MORPHODITA_MODEL = "czech-morfflex2.1-pdtc2.0-250909";
const DK_VIN_PROXY = "/api/vin";
const DK_TIMEOUT_MS = 8000;
const VIN_LENGTH = 17;
const VIN_AUTO_LOOKUP_DELAY_MS = 450;
const DEFAULT_SIGN_PLACE = "Praha";
const DEFAULT_SIGN_PLACE_LOCATIVE = "Praze";
const MORPHODITA_TIMEOUT_MS = 3200;
const LOCATIVE_CASE = "6";

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
const placeCaseCache = new Map();
let placeCaseState = {
  mode: "declined",
  basePlace: DEFAULT_SIGN_PLACE,
  locativePlace: DEFAULT_SIGN_PLACE_LOCATIVE
};
let placeCaseResolveTimer = null;
let placeCaseResolveRequestId = 0;
let vinAutoLookupTimer = null;
let vinLookupInFlightVin = "";
let vinLookupLastSuccessfulVin = "";
placeCaseCache.set(DEFAULT_SIGN_PLACE.toLocaleLowerCase("cs-CZ"), {
  mode: "declined",
  basePlace: DEFAULT_SIGN_PLACE,
  locativePlace: DEFAULT_SIGN_PLACE_LOCATIVE
});

const fieldIds = [
  "grantorId",
  "grantorName",
  "grantorAddress",
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
    const sanitized = sanitizeVin(fields.vehicleVin.value);
    fields.vehicleVin.value = sanitized;
    setVinCustomValidity(sanitized);
    scheduleAutoVinLookup();
  });
  setVinCustomValidity(fields.vehicleVin.value.trim());
  fields.vehicleSpz.addEventListener("input", () => {
    fields.vehicleSpz.value = sanitizeSpz(fields.vehicleSpz.value);
  });
  fields.signDate.addEventListener("input", () => {
    fields.signDate.value = sanitizeDate(fields.signDate.value);
  });
  fields.signPlace.addEventListener("input", () => {
    schedulePlaceCaseResolution();
  });
  fields.signPlace.addEventListener("blur", () => {
    schedulePlaceCaseResolution(true);
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
  schedulePlaceCaseResolution(true);

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

function schedulePlaceCaseResolution(immediate = false) {
  clearTimeout(placeCaseResolveTimer);

  if (immediate) {
    resolvePlaceCaseForCurrentInput().catch(() => {
      // Ignore and keep fallback rendering.
    });
    return;
  }

  placeCaseResolveTimer = setTimeout(() => {
    resolvePlaceCaseForCurrentInput().catch(() => {
      // Ignore and keep fallback rendering.
    });
  }, 280);
}

async function resolvePlaceCaseForCurrentInput() {
  const requestId = ++placeCaseResolveRequestId;
  const basePlace = normalizePlaceInput(fields.signPlace.value, DEFAULT_SIGN_PLACE);
  const resolved = await resolvePlaceCase(basePlace);

  if (requestId !== placeCaseResolveRequestId) {
    return resolved;
  }

  placeCaseState = resolved;
  updatePreview();
  return resolved;
}

function normalizePlaceInput(rawValue, fallback) {
  const text = String(rawValue || "").replace(/\s+/g, " ").trim();
  return text || fallback;
}

async function resolvePlaceCase(basePlace) {
  const normalizedPlace = normalizePlaceInput(basePlace, DEFAULT_SIGN_PLACE);
  const cacheKey = normalizedPlace.toLocaleLowerCase("cs-CZ");
  const cached = placeCaseCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  try {
    const locativePlace = await declinePlaceToLocative(normalizedPlace);
    if (locativePlace) {
      const resolved = {
        mode: "declined",
        basePlace: normalizedPlace,
        locativePlace
      };
      placeCaseCache.set(cacheKey, resolved);
      return resolved;
    }
  } catch (_) {
    // Use manual fallback signature when the external service is unavailable.
  }

  return {
    mode: "fallback",
    basePlace: normalizedPlace,
    locativePlace: DEFAULT_SIGN_PLACE_LOCATIVE
  };
}

async function declinePlaceToLocative(placeName) {
  const tagged = await callMorphodita("tag", placeName);
  const sentences = Array.isArray(tagged?.result) ? tagged.result : [];
  const tokens = Array.isArray(sentences[0]) ? sentences[0] : [];

  if (!tokens.length) {
    return "";
  }

  const replacements = new Map();
  const generationLines = [];
  const generationMeta = [];
  let activeCase = LOCATIVE_CASE;

  tokens.forEach((token, index) => {
    const rawToken = String(token?.token || "");
    const rawTag = String(token?.tag || "");
    if (!rawToken || !rawTag) {
      return;
    }

    if (rawTag.startsWith("R")) {
      const prepositionCase = getRequiredCaseFromPrepositionTag(rawTag);
      activeCase = prepositionCase || activeCase;
      return;
    }

    if (rawTag.startsWith("Z")) {
      activeCase = LOCATIVE_CASE;
      return;
    }

    if (!shouldInflectPlaceToken(rawTag)) {
      return;
    }

    const currentCase = getCaseFromTag(rawTag);
    if (!currentCase || currentCase === activeCase) {
      return;
    }

    const targetTag = replaceCaseInTag(rawTag, activeCase);
    if (!targetTag) {
      return;
    }

    const lemma = String(token?.lemma || rawToken).trim() || rawToken;
    generationLines.push(`${lemma}\t${targetTag}`);
    generationMeta.push({
      index,
      sourceToken: rawToken,
      targetTag
    });
  });

  if (generationLines.length) {
    const generated = await callMorphodita("generate", generationLines.join("\n"));
    const generatedRows = Array.isArray(generated?.result) ? generated.result : [];

    generationMeta.forEach((task, rowIndex) => {
      const forms = Array.isArray(generatedRows[rowIndex]) ? generatedRows[rowIndex] : [];
      const exact = forms.find((item) => item?.tag === task.targetTag && item?.form);
      const first = forms.find((item) => item?.form);
      const inflected = exact?.form || first?.form;

      if (inflected) {
        replacements.set(task.index, applySourceCasing(task.sourceToken, String(inflected)));
      }
    });
  }

  let output = "";
  tokens.forEach((token, index) => {
    const sourceToken = String(token?.token || "");
    const tokenText = replacements.has(index) ? replacements.get(index) : sourceToken;
    output += tokenText;
    output += typeof token?.space === "string" ? token.space : "";
  });

  return output.trim();
}

async function callMorphodita(method, data) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), MORPHODITA_TIMEOUT_MS);

  try {
    const body = new URLSearchParams();
    body.set("model", MORPHODITA_MODEL);
    body.set("data", data);
    body.set("output", "json");

    const response = await fetch(`${MORPHODITA_BASE}/${method}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body,
      signal: controller.signal
    });

    if (!response.ok) {
      throw new Error(`MorphoDiTa ${method} failed with status ${response.status}`);
    }

    return await response.json();
  } finally {
    clearTimeout(timeout);
  }
}

function shouldInflectPlaceToken(tag) {
  if (typeof tag !== "string" || tag.length < 5) {
    return false;
  }

  const pos = tag[0];
  return ["A", "C", "N", "P"].includes(pos) && /[1-7]/.test(tag[4]);
}

function getCaseFromTag(tag) {
  if (typeof tag !== "string" || tag.length < 5) {
    return "";
  }
  return /[1-7]/.test(tag[4]) ? tag[4] : "";
}

function replaceCaseInTag(tag, targetCase) {
  if (typeof tag !== "string" || tag.length < 5 || !/[1-7]/.test(targetCase)) {
    return "";
  }

  const chars = tag.split("");
  chars[4] = targetCase;
  return chars.join("");
}

function getRequiredCaseFromPrepositionTag(tag) {
  if (typeof tag !== "string" || tag.length < 5 || tag[0] !== "R") {
    return "";
  }
  return /[1-7]/.test(tag[4]) ? tag[4] : "";
}

function applySourceCasing(sourceToken, inflectedToken) {
  if (!sourceToken) {
    return inflectedToken;
  }

  const sourceUpper = sourceToken.toLocaleUpperCase("cs-CZ");
  const sourceLower = sourceToken.toLocaleLowerCase("cs-CZ");

  if (sourceToken === sourceUpper && sourceToken !== sourceLower) {
    return inflectedToken.toLocaleUpperCase("cs-CZ");
  }

  if (
    sourceToken[0] &&
    sourceToken[0] === sourceToken[0].toLocaleUpperCase("cs-CZ") &&
    sourceToken.slice(1) === sourceToken.slice(1).toLocaleLowerCase("cs-CZ")
  ) {
    return inflectedToken[0].toLocaleUpperCase("cs-CZ") + inflectedToken.slice(1);
  }

  return inflectedToken;
}

function setDefaultActionsChecked() {
  document.querySelectorAll('input[name="actions"]').forEach((checkbox) => {
    checkbox.checked = true;
  });
}

function setupSubjectLookup(prefix) {
  const suggestions = document.getElementById(`${prefix}Suggestions`);
  const id = document.getElementById(`${prefix}Id`);
  const name = document.getElementById(`${prefix}Name`);
  const address = document.getElementById(`${prefix}Address`);

  let searchTimer = null;
  let requestId = 0;

  name.addEventListener("input", () => {
    const query = name.value.trim();
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

      renderSuggestions({
        suggestions,
        items: found,
        onSelect: (item) => {
          applySubject(item, { id, name, address });
          hideSuggestions(suggestions);
          updatePreview();
        }
      });
    }, 260);
  });

  id.addEventListener("blur", async () => {
    const ico = extractIcoCandidate(id.value);
    if (!ico) {
      return;
    }

    const detail = await fetchByIco(ico);
    if (!detail) {
      return;
    }

    applySubject(detail, { id, name, address });
    updatePreview();
  });

  document.addEventListener("click", (event) => {
    if (
      !event.target.closest(`#${name.id}`) &&
      !event.target.closest(`#${id.id}`) &&
      !event.target.closest(`#${suggestions.id}`)
    ) {
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
  refs.name.value = subject.obchodniJmeno || refs.name.value || "";
  refs.address.value = formatAddress(subject.sidlo) || refs.address.value || "";

  if (subject.ico) {
    refs.id.value = subject.ico;
  }
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

  const parts = [];
  const street = formatStreetPart(sidlo);
  const city = resolveCityPart(sidlo);
  const district = resolveDistrictPart(sidlo, city);
  const psc = sidlo.psc ? formatPscNumber(sidlo.psc) : "";

  if (street) {
    parts.push(street);
  }
  if (psc || city) {
    parts.push(`${psc} ${city}`.trim());
  }
  if (district) {
    if (parts.length) {
      const last = parts.pop();
      parts.push(`${last} - ${district}`);
    } else {
      parts.push(district);
    }
  }

  if (parts.length) {
    return parts.join(", ").trim();
  }

  return sidlo.textovaAdresa ? formatPostCode(sidlo.textovaAdresa) : "";
}

function formatStreetPart(sidlo) {
  const street = normalizeAddressToken(sidlo.nazevUlice);
  const house = sidlo.cisloDomovni ? String(sidlo.cisloDomovni) : "";
  const orient = sidlo.cisloOrientacni ? `/${sidlo.cisloOrientacni}` : "";
  return `${street} ${house}${orient}`.trim();
}

function resolveCityPart(sidlo) {
  const city = normalizeAddressToken(sidlo.nazevObce);
  const candidates = [
    normalizeAddressToken(sidlo.nazevSpravnihoObvodu),
    normalizeAddressToken(sidlo.nazevMestskehoObvodu),
    city
  ].filter(Boolean);

  if (!candidates.length) {
    return "";
  }

  if (city) {
    const extended = candidates.find(
      (candidate) =>
        candidate !== city &&
        (candidate.startsWith(`${city} `) || candidate.startsWith(`${city}-`))
    );
    if (extended) {
      return extended;
    }
  }

  return city || candidates[0];
}

function resolveDistrictPart(sidlo, cityPart) {
  const city = normalizeAddressToken(cityPart);
  const candidates = [
    normalizeAddressToken(sidlo.nazevCastiObce),
    normalizeAddressToken(sidlo.nazevMestskeCastiObvodu)
  ];

  for (let candidate of candidates) {
    if (!candidate) {
      continue;
    }

    if (city && candidate.startsWith(`${city}-`)) {
      candidate = candidate.slice(city.length + 1).trim();
    }

    if (!candidate) {
      continue;
    }

    if (city && (candidate === city || city.includes(candidate) || candidate.includes(city))) {
      continue;
    }

    return candidate;
  }

  return "";
}

function normalizeAddressToken(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
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
  setText("previewGrantorId", valueOrBlank(fields.grantorId.value));
  setText("previewGrantorAddress", valueOrBlank(fields.grantorAddress.value));

  setText("previewAttorneyName", valueOrBlank(fields.attorneyName.value));
  setText("previewAttorneyId", valueOrBlank(fields.attorneyId.value));
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

  const basePlace = normalizePlaceInput(fields.signPlace.value, DEFAULT_SIGN_PLACE);
  const signDateText = valueOrDots(fields.signDate.value, formatDate(new Date()));
  const useResolvedPlace =
    placeCaseState.mode === "declined" &&
    placeCaseState.basePlace === basePlace &&
    !!placeCaseState.locativePlace;
  const useFallbackSignature =
    placeCaseState.mode === "fallback" &&
    placeCaseState.basePlace === basePlace;

  setText("previewSignPlace", useResolvedPlace ? placeCaseState.locativePlace : basePlace);
  setText("previewSignDate", signDateText);
  setText("previewPlaceDateFallbackText", `${DEFAULT_SIGN_PLACE}, ${signDateText}`);

  const previewPlaceDatePrimary = document.getElementById("previewPlaceDatePrimary");
  const previewPlaceDateFallback = document.getElementById("previewPlaceDateFallback");
  if (previewPlaceDatePrimary && previewPlaceDateFallback) {
    previewPlaceDatePrimary.hidden = useFallbackSignature;
    previewPlaceDateFallback.hidden = !useFallbackSignature;
  }

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

  const vinPart = sanitizeVin(fields.vehicleVin.value) || "VINnumber";
  const datePart = isValidDate(fields.signDate.value) ? fields.signDate.value : formatDate(new Date());
  const filename = `PM_prevod_vozidla_${vinPart}_${datePart}.pdf`;

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
  drawLabeledValueRow("IČ/RČ", fields.grantorId.value.trim());
  drawLabeledValueRow("se sídlem (bytem)", fields.grantorAddress.value.trim());

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
  drawLabeledValueRow("IČ/RČ", fields.attorneyId.value.trim());
  drawLabeledValueRow("se sídlem/bytem", fields.attorneyAddress.value.trim());

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

  const basePlace = normalizePlaceInput(fields.signPlace.value, DEFAULT_SIGN_PLACE);
  let resolvedPlaceCase = placeCaseState;
  if (resolvedPlaceCase.basePlace !== basePlace) {
    resolvedPlaceCase = await resolvePlaceCase(basePlace);
    placeCaseState = resolvedPlaceCase;
  }

  const signDateText = valueOrDots(fields.signDate.value, formatDate(new Date()));
  const useFallbackSignature = resolvedPlaceCase.mode === "fallback";

  ensureSpace(useFallbackSignature ? 18 : 12);
  doc.setFont("Arial", "normal");
  doc.setFontSize(12);

  if (useFallbackSignature) {
    doc.text(`${DEFAULT_SIGN_PLACE}, ${signDateText}`, labelX, y);
    const placeDateLineEndX = labelX + 66;
    drawDottedLine(labelX, placeDateLineEndX, y + 4.6);

    doc.setFont("Arial", "italic");
    doc.setFontSize(10.5);
    doc.text("(Místo, datum)", (labelX + placeDateLineEndX) / 2, y + 9.3, { align: "center" });

    doc.setFont("Arial", "normal");
    doc.setFontSize(12);
  } else {
    let leftX = labelX;
    doc.text("V ", leftX, y);
    leftX += doc.getTextWidth("V ");
    doc.setFont("Arial", "bold");
    doc.text(resolvedPlaceCase.locativePlace, leftX, y);
    leftX += doc.getTextWidth(resolvedPlaceCase.locativePlace);

    doc.setFont("Arial", "normal");
    doc.text(" dne ", leftX, y);
    leftX += doc.getTextWidth(" dne ");

    doc.setFont("Arial", "bold");
    doc.text(signDateText, leftX, y);
  }

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
        schedulePlaceCaseResolution(true);
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

function extractIcoCandidate(value) {
  const compact = String(value || "").replace(/\s+/g, "");
  return /^\d{8}$/.test(compact) ? compact : "";
}

function sanitizeVin(value) {
  return value.toUpperCase().replace(/[^A-HJ-NPR-Z0-9]/g, "").slice(0, VIN_LENGTH);
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

/* ── Datová kostka – VIN lookup ── */

function setVinCustomValidity(vin) {
  if (vin && vin.length !== VIN_LENGTH) {
    fields.vehicleVin.setCustomValidity(`VIN musí mít přesně ${VIN_LENGTH} znaků.`);
    return;
  }
  fields.vehicleVin.setCustomValidity("");
}

function scheduleAutoVinLookup() {
  clearTimeout(vinAutoLookupTimer);

  const vin = fields.vehicleVin.value.trim();
  if (!vin) {
    vinLookupLastSuccessfulVin = "";
    setVinLookupMessage("");
    return;
  }

  if (vin.length !== VIN_LENGTH) {
    setVinLookupMessage("");
    return;
  }

  if (vin === vinLookupLastSuccessfulVin || vin === vinLookupInFlightVin) {
    return;
  }

  vinAutoLookupTimer = setTimeout(() => {
    handleVinLookup({ auto: true, focusOnInvalid: false });
  }, VIN_AUTO_LOOKUP_DELAY_MS);
}

async function handleVinLookup({ auto = false, focusOnInvalid = true } = {}) {
  const vin = fields.vehicleVin.value.trim();
  if (vin.length !== VIN_LENGTH) {
    if (!auto) {
      setVinLookupMessage(`VIN musí mít přesně ${VIN_LENGTH} znaků.`, true);
    }
    if (focusOnInvalid) {
      fields.vehicleVin.focus();
    }
    return;
  }
  if (window.location.protocol === "file:") {
    setVinLookupMessage("Pro načítání z registru otevřete aplikaci přes Pages URL (např. http://127.0.0.1:8788), ne jako file://.", true);
    return;
  }

  vinLookupInFlightVin = vin;
  setVinLookupMessage(auto ? "Ověřuji VIN…" : "Načítám data o vozidle…");

  try {
    const data = await fetchVehicleByVin(vin);
    if (fields.vehicleVin.value.trim() !== vin) {
      return;
    }
    if (!data) {
      setVinLookupMessage("Vozidlo nebylo nalezeno nebo nastala chyba.", true);
      return;
    }

    if (data.TovarniZnacka) {
      fields.vehicleBrand.value = titleCaseCs(data.TovarniZnacka);
    }
    if (data.ObchodniOznaceni) {
      fields.vehicleModel.value = titleCaseCs(data.ObchodniOznaceni);
    }
    vinLookupLastSuccessfulVin = vin;

    updatePreview();
    setVinLookupMessage("Data o vozidle načtena z registru.");
  } catch (error) {
    const detail = String(error?.message || "").trim();
    setVinLookupMessage(detail ? `Načtení dat se nezdařilo: ${detail}` : "Načtení dat se nezdařilo.", true);
  } finally {
    if (vinLookupInFlightVin === vin) {
      vinLookupInFlightVin = "";
    }
  }
}

async function fetchVehicleByVin(vin) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), DK_TIMEOUT_MS);

  try {
    const response = await fetch(`${DK_VIN_PROXY}?vin=${encodeURIComponent(vin)}`, {
      signal: controller.signal
    });

    if (!response.ok) {
      let detail = "";
      try {
        const errorPayload = await response.json();
        detail = String(errorPayload?.error || errorPayload?.message || "");
      } catch (_) {
        // ignore parse issues
      }
      throw new Error(detail || `HTTP ${response.status}`);
    }

    const result = await response.json();
    if (result.Status !== 1 || !result.Data) {
      return null;
    }

    return result.Data;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error("Vypršel časový limit dotazu.");
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function titleCaseCs(text) {
  if (!text) {
    return "";
  }
  return text
    .toLocaleLowerCase("cs-CZ")
    .replace(/(^|[\s-])\S/g, (match) => match.toLocaleUpperCase("cs-CZ"));
}

function setVinLookupMessage(text, isError = false) {
  const el = document.getElementById("vinLookupMessage");
  if (!el) {
    return;
  }
  el.textContent = text;
  el.style.color = isError ? "var(--danger)" : "";
}
