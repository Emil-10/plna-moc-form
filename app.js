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
const signPowerBtn = document.getElementById("signPowerBtn");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const togglePreviewBtn = document.getElementById("togglePreviewBtn");
const signatureModal = document.getElementById("signatureModal");
const closeSignatureModalBtn = document.getElementById("closeSignatureModalBtn");
const signatureCanvas = document.getElementById("signatureCanvas");
const clearSignatureBtn = document.getElementById("clearSignatureBtn");
const saveSignatureBtn = document.getElementById("saveSignatureBtn");
const generateStampBtn = document.getElementById("generateStampBtn");
const clearStampBtn = document.getElementById("clearStampBtn");
const signatureModalMessage = document.getElementById("signatureModalMessage");
const signatureModalBody = document.getElementById("signatureModalBody");
const previewSignStage = document.getElementById("previewSignStage");
const previewSignatureImage = document.getElementById("previewSignatureImage");
const previewStamp = document.getElementById("previewStamp");
const signatureComposeStage = document.getElementById("signatureComposeStage");
const signatureComposeImage = document.getElementById("signatureComposeImage");
const signatureComposeStamp = document.getElementById("signatureComposeStamp");
const ARIAL_REGULAR_URL = "./assets/fonts/arial.ttf";
const ARIAL_BOLD_URL = "./assets/fonts/arialbd.ttf";
const ARIAL_ITALIC_URL = "./assets/fonts/ariali.ttf";
const ARIAL_BOLDITALIC_URL = "./assets/fonts/arialbi.ttf";
const SIGNATURE_CANVAS_HEIGHT = 240;
const SIGNATURE_PAD_LINE_WIDTH = 2.2;
const PDF_PT_TO_MM = 0.352778;
const SIGNATURE_STAGE_WIDTH_MM = 66;
const SIGNATURE_STAGE_HEIGHT_MM = 24;
const SIGNATURE_STAGE_ASPECT_RATIO = SIGNATURE_STAGE_WIDTH_MM / SIGNATURE_STAGE_HEIGHT_MM;
const SIGNATURE_IMAGE_WIDTH_RATIO = 0.85;
const SIGNATURE_STAGE_PADDING_RATIO_X = 0.01;
const SIGNATURE_STAGE_PADDING_RATIO_Y = 0.01;
const SIGNATURE_EXPORT_PADDING_PX = 8;
const STAMP_SCALE = 1.2;
const DEFAULT_SIGNATURE_LAYOUT = Object.freeze({
  signature: { x: 0.54, y: 0.44 },
  stamp: { x: 0.54, y: 0.66 }
});
const DEFAULT_SIGNATURE_SIZE = Object.freeze({
  widthRatio: 0.32,
  heightRatio: 0.18
});
let arialFontPromise = null;
let signaturePadContext = null;
let signaturePadDrawing = false;
let signaturePadHasInk = false;
let signatureResizeTimer = null;
let signatureDragState = null;
let signatureInkBounds = null;
const signatureState = {
  imageDataUrl: "",
  stampLines: [],
  layout: cloneSignatureLayout(),
  size: cloneSignatureSize()
};
const signatureDraftState = {
  imageDataUrl: "",
  stampLines: [],
  layout: cloneSignatureLayout(),
  size: cloneSignatureSize()
};
const subjectLookupMeta = {
  grantor: { dic: "" },
  attorney: { dic: "" }
};
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
  setupSignatureModal();
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

function setupSignatureModal() {
  if (!signPowerBtn || !signatureModal || !signatureCanvas) {
    return;
  }

  signPowerBtn.addEventListener("click", openSignatureModal);
  closeSignatureModalBtn?.addEventListener("click", closeSignatureModal);
  clearSignatureBtn?.addEventListener("click", () => {
    clearSignatureWorkspace();
    setSignatureModalMessage("Podpis i razítko byly smazány.");
  });
  saveSignatureBtn?.addEventListener("click", saveSignatureFromModal);
  generateStampBtn?.addEventListener("click", generateStampInModal);

  signatureCanvas.addEventListener("pointerdown", handleSignaturePointerDown);
  signatureCanvas.addEventListener("pointermove", handleSignaturePointerMove);
  signatureCanvas.addEventListener("pointerup", handleSignaturePointerUp);
  signatureCanvas.addEventListener("pointercancel", handleSignaturePointerUp);
  signatureCanvas.addEventListener("pointerleave", handleSignaturePointerUp);

  signatureModal.addEventListener("click", (event) => {
    if (event.target === signatureModal) {
      closeSignatureModal();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !signatureModal.hidden) {
      closeSignatureModal();
    }
  });

  window.addEventListener("resize", () => {
    if (signatureModal.hidden) {
      return;
    }
    clearTimeout(signatureResizeTimer);
    signatureResizeTimer = setTimeout(() => {
      resizeSignatureCanvas(true);
      renderComposeStage();
    }, 120);
  });

  setupSignatureComposeDragging();
}

function openSignatureModal() {
  if (!signatureModal || !signatureCanvas) {
    return;
  }

  signatureDraftState.imageDataUrl = signatureState.imageDataUrl;
  signatureDraftState.stampLines = [...signatureState.stampLines];
  signatureDraftState.layout = cloneSignatureLayout(signatureState.layout);
  signatureDraftState.size = cloneSignatureSize(signatureState.size);

  signatureModal.hidden = false;
  document.body.classList.add("modal-open");
  if (signatureModalBody) {
    signatureModalBody.scrollTop = 0;
  }
  setSignatureModalMessage("");
  renderComposeStage();

  requestAnimationFrame(() => {
    resizeSignatureCanvas(false);
    clearSignatureCanvasSurface();
    signaturePadHasInk = false;
    renderComposeStage();
  });
}

function closeSignatureModal() {
  if (!signatureModal || signatureModal.hidden) {
    return;
  }

  signatureModal.hidden = true;
  signaturePadDrawing = false;
  signatureDragState = null;
  document.body.classList.remove("modal-open");
}

function saveSignatureFromModal() {
  syncDraftSignatureFromCanvas();
  signatureState.imageDataUrl = signatureDraftState.imageDataUrl;
  signatureState.stampLines = normalizeStampLines(signatureDraftState.stampLines);
  signatureState.layout = cloneSignatureLayout(signatureDraftState.layout);
  signatureState.size = cloneSignatureSize(signatureDraftState.size);

  closeSignatureModal();
  updatePreview();

  const hasSignature = Boolean(signatureState.imageDataUrl);
  const hasStamp = signatureState.stampLines.length > 0;
  if (hasSignature && hasStamp) {
    setMessage("Podpis a razítko byly uloženy.");
    return;
  }
  if (hasSignature) {
    setMessage("Podpis byl uložen.");
    return;
  }
  if (hasStamp) {
    setMessage("Razítko bylo uloženo.");
    return;
  }
  setMessage("Podpis i razítko byly smazány.");
}

function generateStampInModal() {
  const lines = buildStampLinesFromGrantor();
  if (!lines.length) {
    setSignatureModalMessage("Pro vygenerování razítka doplňte údaje zmocnitele.", true);
    return;
  }

  signatureDraftState.stampLines = lines;
  renderComposeStage();
  setSignatureModalMessage("Razítko je připravené. Tažením můžete upravit jeho pozici.");
}

function clearStampInModal() {
  if (!signatureDraftState.stampLines.length) {
    setSignatureModalMessage("Razítko už je smazané.");
    return;
  }

  signatureDraftState.stampLines = [];
  renderComposeStage();
  setSignatureModalMessage("Razítko bylo smazáno.");
}

function syncDraftSignatureFromCanvas() {
  if (!signatureCanvas) {
    return;
  }

  if (!signaturePadHasInk) {
    renderComposeStage();
    return;
  }

  const exportedSignature = exportTrimmedSignatureImageData();
  signatureDraftState.imageDataUrl = exportedSignature?.dataUrl || "";
  signatureDraftState.size = cloneSignatureSize(exportedSignature?.size);
  clearSignatureCanvasSurface();
  signaturePadHasInk = false;
  signatureInkBounds = null;
  renderComposeStage();
}

function resizeSignatureCanvas(preserveDrawing) {
  if (!signatureCanvas) {
    return;
  }

  let snapshot = null;
  if (preserveDrawing && signaturePadHasInk && signatureCanvas.width && signatureCanvas.height) {
    snapshot = document.createElement("canvas");
    snapshot.width = signatureCanvas.width;
    snapshot.height = signatureCanvas.height;
    const snapshotContext = snapshot.getContext("2d");
    snapshotContext?.drawImage(signatureCanvas, 0, 0);
  }

  const shellWidth = signatureCanvas.parentElement?.clientWidth || 640;
  const cssWidth = Math.max(280, Math.floor(shellWidth));
  const computedHeight = Number.parseFloat(window.getComputedStyle(signatureCanvas).height);
  const cssHeight =
    Number.isFinite(computedHeight) && computedHeight > 0
      ? Math.round(computedHeight)
      : SIGNATURE_CANVAS_HEIGHT;
  const scale = Math.max(window.devicePixelRatio || 1, 1);

  signatureCanvas.width = Math.floor(cssWidth * scale);
  signatureCanvas.height = Math.floor(cssHeight * scale);
  signatureCanvas.style.width = `${cssWidth}px`;
  signatureCanvas.style.height = `${cssHeight}px`;

  signaturePadContext = signatureCanvas.getContext("2d");
  if (!signaturePadContext) {
    return;
  }

  signaturePadContext.setTransform(scale, 0, 0, scale, 0, 0);
  signaturePadContext.lineCap = "round";
  signaturePadContext.lineJoin = "round";
  signaturePadContext.lineWidth = SIGNATURE_PAD_LINE_WIDTH;
  signaturePadContext.strokeStyle = "#1858bf";
  signaturePadContext.clearRect(0, 0, cssWidth, cssHeight);

  if (snapshot) {
    signaturePadContext.drawImage(snapshot, 0, 0, cssWidth, cssHeight);
    signaturePadHasInk = true;
  } else if (!preserveDrawing) {
    signaturePadHasInk = false;
    signatureInkBounds = null;
  }
}

function getSignatureCanvasContext() {
  if (!signaturePadContext) {
    resizeSignatureCanvas(false);
  }
  return signaturePadContext;
}

function clearSignatureWorkspace() {
  signatureDraftState.imageDataUrl = "";
  signatureDraftState.stampLines = [];
  signatureDraftState.layout = cloneSignatureLayout();
  signatureDraftState.size = cloneSignatureSize();
  clearSignatureCanvasSurface();
  signaturePadHasInk = false;
  signatureInkBounds = null;
  renderComposeStage();
}

function clearSignatureCanvasSurface() {
  const context = getSignatureCanvasContext();
  if (!context || !signatureCanvas) {
    return;
  }

  const bounds = signatureCanvas.getBoundingClientRect();
  context.clearRect(0, 0, bounds.width, bounds.height);
}

function drawSignatureImageToCanvas(dataUrl) {
  if (!dataUrl) {
    clearSignatureCanvasSurface();
    return;
  }

  const context = getSignatureCanvasContext();
  if (!context || !signatureCanvas) {
    return;
  }

  const image = new Image();
  image.onload = () => {
    const bounds = signatureCanvas.getBoundingClientRect();
    context.clearRect(0, 0, bounds.width, bounds.height);
    context.drawImage(image, 0, 0, bounds.width, bounds.height);
    signaturePadHasInk = true;
  };
  image.onerror = () => {
    clearSignatureCanvasSurface();
  };
  image.src = dataUrl;
}

function handleSignaturePointerDown(event) {
  if (!signatureCanvas || signatureModal?.hidden || signatureDraftState.imageDataUrl) {
    return;
  }
  if (event.pointerType === "mouse" && event.button !== 0) {
    return;
  }

  const context = getSignatureCanvasContext();
  if (!context) {
    return;
  }

  event.preventDefault();
  signatureCanvas.setPointerCapture?.(event.pointerId);
  const point = getSignatureCanvasPoint(event);
  updateSignatureInkBounds(point.x, point.y);
  context.beginPath();
  context.moveTo(point.x, point.y);
  context.lineTo(point.x + 0.01, point.y + 0.01);
  context.stroke();

  signaturePadDrawing = true;
  signaturePadHasInk = true;
}

function handleSignaturePointerMove(event) {
  if (!signaturePadDrawing || !signatureCanvas) {
    return;
  }

  const context = getSignatureCanvasContext();
  if (!context) {
    return;
  }

  event.preventDefault();
  const point = getSignatureCanvasPoint(event);
  updateSignatureInkBounds(point.x, point.y);
  context.lineTo(point.x, point.y);
  context.stroke();
}

function handleSignaturePointerUp(event) {
  if (!signaturePadDrawing || !signatureCanvas) {
    return;
  }

  signaturePadDrawing = false;
  signatureCanvas.releasePointerCapture?.(event.pointerId);
  syncDraftSignatureFromCanvas();
}

function getSignatureCanvasPoint(event) {
  const bounds = signatureCanvas.getBoundingClientRect();
  return {
    x: event.clientX - bounds.left,
    y: event.clientY - bounds.top
  };
}

function updateSignatureInkBounds(x, y) {
  const halfLineWidth = SIGNATURE_PAD_LINE_WIDTH * 0.5 + 1.2;
  const minX = x - halfLineWidth;
  const maxX = x + halfLineWidth;
  const minY = y - halfLineWidth;
  const maxY = y + halfLineWidth;

  if (!signatureInkBounds) {
    signatureInkBounds = { minX, maxX, minY, maxY };
    return;
  }

  signatureInkBounds.minX = Math.min(signatureInkBounds.minX, minX);
  signatureInkBounds.maxX = Math.max(signatureInkBounds.maxX, maxX);
  signatureInkBounds.minY = Math.min(signatureInkBounds.minY, minY);
  signatureInkBounds.maxY = Math.max(signatureInkBounds.maxY, maxY);
}

function exportTrimmedSignatureImageData() {
  if (!signatureCanvas || !signatureInkBounds) {
    return null;
  }

  const displayWidth = signatureCanvas.clientWidth || signatureCanvas.getBoundingClientRect().width;
  const displayHeight = signatureCanvas.clientHeight || signatureCanvas.getBoundingClientRect().height;
  if (!displayWidth || !displayHeight || !signatureCanvas.width || !signatureCanvas.height) {
    return null;
  }

  const scaleX = signatureCanvas.width / displayWidth;
  const scaleY = signatureCanvas.height / displayHeight;
  const sourceX = Math.max(0, Math.floor((signatureInkBounds.minX - SIGNATURE_EXPORT_PADDING_PX) * scaleX));
  const sourceY = Math.max(0, Math.floor((signatureInkBounds.minY - SIGNATURE_EXPORT_PADDING_PX) * scaleY));
  const sourceMaxX = Math.min(
    signatureCanvas.width,
    Math.ceil((signatureInkBounds.maxX + SIGNATURE_EXPORT_PADDING_PX) * scaleX)
  );
  const sourceMaxY = Math.min(
    signatureCanvas.height,
    Math.ceil((signatureInkBounds.maxY + SIGNATURE_EXPORT_PADDING_PX) * scaleY)
  );
  const cropWidth = sourceMaxX - sourceX;
  const cropHeight = sourceMaxY - sourceY;

  if (cropWidth <= 0 || cropHeight <= 0) {
    return null;
  }

  const exportCanvas = document.createElement("canvas");
  exportCanvas.width = cropWidth;
  exportCanvas.height = cropHeight;
  const exportContext = exportCanvas.getContext("2d");
  if (!exportContext) {
    return null;
  }

  exportContext.drawImage(
    signatureCanvas,
    sourceX,
    sourceY,
    cropWidth,
    cropHeight,
    0,
    0,
    cropWidth,
    cropHeight
  );
  return {
    dataUrl: exportCanvas.toDataURL("image/png"),
    size: {
      widthRatio: clamp(cropWidth / signatureCanvas.width, 0.08, 0.96),
      heightRatio: clamp(cropHeight / signatureCanvas.height, 0.06, 0.9)
    }
  };
}

function setupSignatureComposeDragging() {
  if (!signatureComposeStage) {
    return;
  }

  signatureComposeStage.addEventListener("pointerdown", startSignatureObjectDrag);
  signatureComposeStage.addEventListener("pointermove", updateSignatureObjectDrag);
  signatureComposeStage.addEventListener("pointerup", stopSignatureObjectDrag);
  signatureComposeStage.addEventListener("pointercancel", stopSignatureObjectDrag);
  signatureComposeStage.addEventListener("pointerleave", stopSignatureObjectDrag);
}

function startSignatureObjectDrag(event) {
  if (!signatureComposeStage || signatureModal?.hidden) {
    return;
  }
  if (event.pointerType === "mouse" && event.button !== 0) {
    return;
  }

  const dragElement = event.target.closest("[data-drag-target]");
  if (!dragElement || !signatureComposeStage.contains(dragElement)) {
    return;
  }

  const target = dragElement.getAttribute("data-drag-target");
  if (target !== "signature" && target !== "stamp") {
    return;
  }

  event.preventDefault();
  signatureComposeStage.setPointerCapture?.(event.pointerId);
  signatureDragState = {
    pointerId: event.pointerId,
    target,
    element: dragElement,
    startX: event.clientX,
    startY: event.clientY,
    origin: {
      x: signatureDraftState.layout[target].x,
      y: signatureDraftState.layout[target].y
    }
  };
  dragElement.classList.add("is-dragging");
}

function updateSignatureObjectDrag(event) {
  if (!signatureDragState || event.pointerId !== signatureDragState.pointerId || !signatureComposeStage) {
    return;
  }

  const rect = signatureComposeStage.getBoundingClientRect();
  if (!rect.width || !rect.height) {
    return;
  }

  event.preventDefault();
  const dx = (event.clientX - signatureDragState.startX) / rect.width;
  const dy = (event.clientY - signatureDragState.startY) / rect.height;
  const target = signatureDragState.target;
  const element = target === "signature" ? signatureComposeImage : signatureComposeStamp;
  const nextPosition = clampStageLayoutPosition(
    signatureComposeStage,
    element,
    {
      x: signatureDragState.origin.x + dx,
      y: signatureDragState.origin.y + dy
    }
  );

  signatureDraftState.layout[target] = nextPosition;
  setStageObjectPosition(element, nextPosition);
}

function stopSignatureObjectDrag(event) {
  if (!signatureDragState) {
    return;
  }
  if (event.pointerId !== undefined && event.pointerId !== signatureDragState.pointerId) {
    return;
  }

  signatureDragState.element?.classList.remove("is-dragging");
  signatureComposeStage?.releasePointerCapture?.(signatureDragState.pointerId);
  signatureDragState = null;
}

function renderComposeStage() {
  renderStampLines(signatureComposeStamp, signatureDraftState.stampLines);

  if (signatureComposeImage) {
    if (signatureDraftState.imageDataUrl) {
      signatureComposeImage.src = signatureDraftState.imageDataUrl;
      signatureComposeImage.hidden = false;
    } else {
      signatureComposeImage.hidden = true;
      signatureComposeImage.removeAttribute("src");
    }
    setSignatureImageSize(signatureComposeImage, signatureDraftState.size);
    signatureDraftState.layout.signature = setStageObjectPosition(
      signatureComposeImage,
      signatureDraftState.layout.signature
    );
  }

  signatureDraftState.layout.stamp = setStageObjectPosition(
    signatureComposeStamp,
    signatureDraftState.layout.stamp
  );

  updateSignatureCanvasMode();
}

function updateSignatureCanvasMode() {
  if (!signatureCanvas) {
    return;
  }

  signatureCanvas.classList.toggle("is-locked", Boolean(signatureDraftState.imageDataUrl));
}

function renderSavedSignatureAndStamp() {
  renderStampLines(previewStamp, signatureState.stampLines);

  if (previewSignatureImage) {
    if (signatureState.imageDataUrl) {
      previewSignatureImage.src = signatureState.imageDataUrl;
      previewSignatureImage.hidden = false;
    } else {
      previewSignatureImage.hidden = true;
      previewSignatureImage.removeAttribute("src");
    }
    setSignatureImageSize(previewSignatureImage, signatureState.size);
    setStageObjectPosition(previewSignatureImage, signatureState.layout.signature);
  }

  setStageObjectPosition(previewStamp, signatureState.layout.stamp);

  if (previewSignStage) {
    previewSignStage.hidden = !(signatureState.imageDataUrl || signatureState.stampLines.length);
  }
}

function renderStampLines(container, lines) {
  if (!container) {
    return;
  }

  const normalized = normalizeStampLines(lines);
  container.innerHTML = "";
  container.hidden = !normalized.length;

  normalized.forEach((line, index) => {
    const row = document.createElement("span");
    row.className = `stamp-line${index === 0 ? " stamp-line-main" : ""}`;
    row.textContent = line;
    container.appendChild(row);
  });
}

function setStageObjectPosition(element, position) {
  if (!element) {
    return { x: 0.5, y: 0.5 };
  }

  const stage = element.closest(".pm-sign-stage");
  const normalized = clampStageLayoutPosition(stage, element, position);
  const x = normalized.x;
  const y = normalized.y;
  element.style.left = `${(x * 100).toFixed(2)}%`;
  element.style.top = `${(y * 100).toFixed(2)}%`;
  return normalized;
}

function clampStageLayoutPosition(stage, element, position) {
  const bounds = getStageLayoutBounds(stage, element);
  return {
    x: clamp(position?.x ?? 0.5, bounds.minX, bounds.maxX),
    y: clamp(position?.y ?? 0.5, bounds.minY, bounds.maxY)
  };
}

function getStageLayoutBounds(stage, element) {
  const fallback = {
    minX: SIGNATURE_STAGE_PADDING_RATIO_X,
    maxX: 1 - SIGNATURE_STAGE_PADDING_RATIO_X,
    minY: SIGNATURE_STAGE_PADDING_RATIO_Y,
    maxY: 1 - SIGNATURE_STAGE_PADDING_RATIO_Y
  };

  if (!stage || !element || element.hidden) {
    return fallback;
  }

  const stageRect = stage.getBoundingClientRect();
  const elementRect = element.getBoundingClientRect();
  if (!stageRect.width || !stageRect.height || !elementRect.width || !elementRect.height) {
    return fallback;
  }

  const halfWidthRatio = clamp(elementRect.width / stageRect.width / 2, 0, 0.5);
  const halfHeightRatio = clamp(elementRect.height / stageRect.height / 2, 0, 0.5);

  let minX = halfWidthRatio + SIGNATURE_STAGE_PADDING_RATIO_X;
  let maxX = 1 - halfWidthRatio - SIGNATURE_STAGE_PADDING_RATIO_X;
  let minY = halfHeightRatio + SIGNATURE_STAGE_PADDING_RATIO_Y;
  let maxY = 1 - halfHeightRatio - SIGNATURE_STAGE_PADDING_RATIO_Y;

  minX = clamp(minX, 0, 1);
  maxX = clamp(maxX, minX, 1);
  minY = clamp(minY, 0, 1);
  maxY = clamp(maxY, minY, 1);

  return { minX, maxX, minY, maxY };
}

function setSignatureModalMessage(text, isError = false) {
  if (!signatureModalMessage) {
    return;
  }

  signatureModalMessage.textContent = text;
  signatureModalMessage.classList.toggle("is-error", Boolean(text) && isError);
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
  const clearRegistryMeta = () => {
    if (subjectLookupMeta[prefix]) {
      subjectLookupMeta[prefix].dic = "";
    }
  };

  name.addEventListener("input", () => {
    clearRegistryMeta();
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
          applySubject(item, { id, name, address }, prefix);
          hideSuggestions(suggestions);
          updatePreview();
        }
      });
    }, 260);
  });

  id.addEventListener("input", clearRegistryMeta);
  address.addEventListener("input", clearRegistryMeta);

  id.addEventListener("blur", async () => {
    const ico = extractIcoCandidate(id.value);
    if (!ico) {
      return;
    }

    const detail = await fetchByIco(ico);
    if (!detail) {
      return;
    }

    applySubject(detail, { id, name, address }, prefix);
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

function applySubject(subject, refs, prefix = "") {
  refs.name.value = subject.obchodniJmeno || refs.name.value || "";
  refs.address.value = formatAddress(subject.sidlo) || refs.address.value || "";

  if (subject.ico) {
    refs.id.value = subject.ico;
  }

  if (prefix && subjectLookupMeta[prefix]) {
    subjectLookupMeta[prefix].dic = normalizeDic(subject?.dic);
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

function buildStampLinesFromGrantor() {
  const lines = [];
  const grantorName = fields.grantorName.value.trim();
  if (grantorName) {
    lines.push(grantorName);
  }

  splitAddressForStamp(fields.grantorAddress.value).forEach((line) => {
    lines.push(line);
  });

  const identityParts = [];
  const ico = extractIcoCandidate(fields.grantorId.value);
  if (ico) {
    identityParts.push(`IČ: ${ico}`);
  }

  const dic = getGrantorDicFromRegistry();
  if (dic) {
    identityParts.push(`DIČ: ${dic}`);
  }

  if (identityParts.length) {
    lines.push(identityParts.join(" "));
  }

  return normalizeStampLines(lines);
}

function splitAddressForStamp(rawAddress) {
  const segments = String(rawAddress || "")
    .split(",")
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  if (segments.length <= 1) {
    return segments;
  }

  return [segments[0], segments.slice(1).join(", ")];
}

function getGrantorDicFromRegistry() {
  const registryDic = normalizeDic(subjectLookupMeta.grantor?.dic);
  if (!registryDic) {
    return "";
  }

  const currentIco = extractIcoCandidate(fields.grantorId.value);
  if (!currentIco) {
    return registryDic;
  }

  const registryIcoPart = registryDic.replace(/^CZ/i, "");
  return registryIcoPart === currentIco ? registryDic : "";
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
  } else {
    selectedActions.forEach((action) => {
      const li = document.createElement("li");
      li.textContent = action;
      actionsList.appendChild(li);
    });
  }

  renderSavedSignatureAndStamp();
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
  if (signPowerBtn) {
    signPowerBtn.disabled = true;
  }
  setMessage("Generuji PDF...");

  try {
    const pdf = await buildTextPdf();
    pdf.save(filename);

    setMessage("PDF bylo staženo.");
  } catch (_) {
    setMessage("Vytvoření PDF se nezdařilo. Zkuste to znovu.", true);
  } finally {
    downloadPdfBtn.disabled = false;
    if (signPowerBtn) {
      signPowerBtn.disabled = false;
    }
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

  const pdfStampLines = normalizeStampLines(signatureState.stampLines);
  const hasPdfSignature = Boolean(signatureState.imageDataUrl);
  const hasPdfStamp = pdfStampLines.length > 0;
  const hasPdfSignatureLayer = hasPdfSignature || hasPdfStamp;
  const signatureStageHeightMm = hasPdfSignatureLayer ? SIGNATURE_STAGE_HEIGHT_MM : 0;
  const signatureBottomSpacingMm = hasPdfSignatureLayer ? signatureStageHeightMm + 12 : 20;

  drawWrappedParagraph(
    "tato plná moc není ve výše uvedeném rozsahu ničím omezena. Zplnomocnění v plném rozsahu zmocněnec přijímá.",
    signatureBottomSpacingMm
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

  const signatureLineStartX = pageWidth - marginRight - SIGNATURE_STAGE_WIDTH_MM;
  const signatureAreaWidth = valueEndX - signatureLineStartX;
  const signatureLineY = y + 1.2;
  drawDottedLine(signatureLineStartX, valueEndX, signatureLineY);

  if (hasPdfSignatureLayer) {
    const stageLeftX = signatureLineStartX;
    const stageTopY = signatureLineY - signatureStageHeightMm - 0.8;
    const stageWidthMm = signatureAreaWidth;
    const stageHeightMm = signatureStageHeightMm;

    if (hasPdfStamp) {
      const stampLineStepMm = 3.9 * STAMP_SCALE;
      const stampMaxWidthMm = stageWidthMm - 3.4;
      const preparedStampLines = [];

      pdfStampLines.forEach((line, index) => {
        const isMainLine = index === 0;
        const fontStyle = isMainLine ? "bold" : "normal";
        const fontSize = (isMainLine ? 10.5 : 9.4) * STAMP_SCALE;
        const fitted = fitTextForPdf(doc, line, stampMaxWidthMm, fontStyle, fontSize);
        if (!fitted) {
          return;
        }

        doc.setFont("Arial", fontStyle);
        doc.setFontSize(fontSize);
        preparedStampLines.push({
          text: fitted,
          fontStyle,
          fontSize,
          widthMm: doc.getTextWidth(fitted)
        });
      });

      if (preparedStampLines.length) {
        const stampLayout = signatureState.layout?.stamp || DEFAULT_SIGNATURE_LAYOUT.stamp;
        const maxStampLineWidthMm = preparedStampLines.reduce((max, line) => Math.max(max, line.widthMm), 0);
        const stampBlockHeightMm =
          (preparedStampLines.length - 1) * stampLineStepMm +
          preparedStampLines[preparedStampLines.length - 1].fontSize * PDF_PT_TO_MM;
        const stampHalfWidthRatio = clamp((maxStampLineWidthMm / 2 + 0.7) / stageWidthMm, 0, 0.48);
        const stampHalfHeightRatio = clamp((stampBlockHeightMm / 2 + 0.6) / stageHeightMm, 0, 0.48);
        const stampCenterXRatio = clamp(
          stampLayout.x,
          stampHalfWidthRatio + SIGNATURE_STAGE_PADDING_RATIO_X,
          1 - stampHalfWidthRatio - SIGNATURE_STAGE_PADDING_RATIO_X
        );
        const stampCenterYRatio = clamp(
          stampLayout.y,
          stampHalfHeightRatio + SIGNATURE_STAGE_PADDING_RATIO_Y,
          1 - stampHalfHeightRatio - SIGNATURE_STAGE_PADDING_RATIO_Y
        );
        const stampCenterX = stageLeftX + stampCenterXRatio * stageWidthMm;
        const stampCenterY = stageTopY + stampCenterYRatio * stageHeightMm;
        const stampBaselineOffsetMm = ((preparedStampLines.length - 1) * stampLineStepMm) / 2;
        let stampY = stampCenterY - stampBaselineOffsetMm;

        preparedStampLines.forEach((line) => {
          doc.setFont("Arial", line.fontStyle);
          doc.setFontSize(line.fontSize);
          doc.text(line.text, stampCenterX, stampY, { align: "center" });
          stampY += stampLineStepMm;
        });
      }

      doc.setFont("Arial", "normal");
      doc.setFontSize(12);
    }

    if (hasPdfSignature) {
      const signatureLayout = signatureState.layout?.signature || DEFAULT_SIGNATURE_LAYOUT.signature;
      const signatureSize = cloneSignatureSize(signatureState.size);
      const signatureAspectRatio = await getImageAspectRatio(signatureState.imageDataUrl, SIGNATURE_STAGE_ASPECT_RATIO);
      const signatureBoxWidthMm = stageWidthMm * signatureSize.widthRatio;
      const signatureBoxHeightMm = stageHeightMm * signatureSize.heightRatio;
      let signatureWidthMm = signatureBoxWidthMm;
      let signatureHeightMm = signatureWidthMm / signatureAspectRatio;

      if (signatureHeightMm > signatureBoxHeightMm) {
        signatureHeightMm = signatureBoxHeightMm;
        signatureWidthMm = signatureHeightMm * signatureAspectRatio;
      }

      const signatureHalfWidthRatio = clamp(signatureWidthMm / stageWidthMm / 2, 0, 0.49);
      const signatureHalfHeightRatio = clamp(signatureHeightMm / stageHeightMm / 2, 0, 0.49);
      const signatureCenterXRatio = clamp(
        signatureLayout.x,
        signatureHalfWidthRatio + SIGNATURE_STAGE_PADDING_RATIO_X,
        1 - signatureHalfWidthRatio - SIGNATURE_STAGE_PADDING_RATIO_X
      );
      const signatureCenterYRatio = clamp(
        signatureLayout.y,
        signatureHalfHeightRatio + SIGNATURE_STAGE_PADDING_RATIO_Y,
        1 - signatureHalfHeightRatio - SIGNATURE_STAGE_PADDING_RATIO_Y
      );

      const signatureX = stageLeftX + signatureCenterXRatio * stageWidthMm - signatureWidthMm / 2;
      const signatureY = stageTopY + signatureCenterYRatio * stageHeightMm - signatureHeightMm / 2;

      try {
        doc.addImage(
          signatureState.imageDataUrl,
          "PNG",
          signatureX,
          signatureY,
          signatureWidthMm,
          signatureHeightMm,
          undefined,
          "FAST"
        );
      } catch (_) {
        // Ignore invalid image data and keep generating the PDF.
      }
    }
  }

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

function normalizeDic(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .toUpperCase();
}

function normalizeStampLines(lines) {
  if (!Array.isArray(lines)) {
    return [];
  }

  return lines
    .map((line) => String(line || "").replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .slice(0, 4);
}

function getImageAspectRatio(dataUrl, fallback = SIGNATURE_STAGE_ASPECT_RATIO) {
  if (!dataUrl) {
    return Promise.resolve(fallback);
  }

  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      const width = Number(image.naturalWidth);
      const height = Number(image.naturalHeight);
      if (!width || !height) {
        resolve(fallback);
        return;
      }
      resolve(width / height);
    };
    image.onerror = () => resolve(fallback);
    image.src = dataUrl;
  });
}

function fitTextForPdf(doc, rawText, maxWidth, fontStyle = "normal", fontSize = 9.5) {
  const text = String(rawText || "").trim();
  if (!text) {
    return "";
  }

  doc.setFont("Arial", fontStyle);
  doc.setFontSize(fontSize);

  if (doc.getTextWidth(text) <= maxWidth) {
    return text;
  }

  let shortened = text;
  while (shortened.length > 2 && doc.getTextWidth(`${shortened}…`) > maxWidth) {
    shortened = shortened.slice(0, -1).trimEnd();
  }
  return shortened ? `${shortened}…` : "";
}

function cloneSignatureLayout(layout = DEFAULT_SIGNATURE_LAYOUT) {
  return {
    signature: {
      x: clamp(layout?.signature?.x ?? DEFAULT_SIGNATURE_LAYOUT.signature.x, 0, 1),
      y: clamp(layout?.signature?.y ?? DEFAULT_SIGNATURE_LAYOUT.signature.y, 0, 1)
    },
    stamp: {
      x: clamp(layout?.stamp?.x ?? DEFAULT_SIGNATURE_LAYOUT.stamp.x, 0, 1),
      y: clamp(layout?.stamp?.y ?? DEFAULT_SIGNATURE_LAYOUT.stamp.y, 0, 1)
    }
  };
}

function cloneSignatureSize(size = DEFAULT_SIGNATURE_SIZE) {
  return {
    widthRatio: clamp(size?.widthRatio ?? DEFAULT_SIGNATURE_SIZE.widthRatio, 0.08, 0.96),
    heightRatio: clamp(size?.heightRatio ?? DEFAULT_SIGNATURE_SIZE.heightRatio, 0.06, 0.9)
  };
}

function clamp(value, min, max) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return min;
  }
  return Math.min(max, Math.max(min, numeric));
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
      setVinLookupMessage(`VIN musí mít přesně ${VIN_LENGTH} znaků.`, "error");
    }
    if (focusOnInvalid) {
      fields.vehicleVin.focus();
    }
    return;
  }
  if (window.location.protocol === "file:") {
    setVinLookupMessage(
      "VIN registr nefunguje při otevření přes file://. Spusťte aplikaci přes Pages/Wrangler, například na http://127.0.0.1:8788/.",
      "error"
    );
    return;
  }

  vinLookupInFlightVin = vin;
  setVinLookupMessage(auto ? "Ověřuji VIN v registru vozidel…" : "Načítám data o vozidle z registru…", "loading");

  try {
    const data = await fetchVehicleByVin(vin);
    if (fields.vehicleVin.value.trim() !== vin) {
      return;
    }
    if (!data) {
      setVinLookupMessage("Vozidlo nebylo v registru nalezeno nebo registr nevrátil použitelná data.", "error");
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
    setVinLookupMessage("Data o vozidle byla úspěšně načtena z registru.", "success");
  } catch (error) {
    const detail = String(error?.message || "").trim();
    setVinLookupMessage(detail ? `Načtení dat se nezdařilo: ${detail}` : "Načtení dat se nezdařilo.", "error");
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
      const contentType = String(response.headers.get("content-type") || "").toLowerCase();
      let detail = "";
      if (contentType.includes("application/json")) {
        const errorPayload = await response.json();
        detail = String(errorPayload?.error || errorPayload?.message || "");
      }

      if (response.status === 404 && !contentType.includes("application/json")) {
        throw new Error(
          "Lokální VIN služba na tomto serveru neběží. Při spuštění přes jednoduchý statický server se /api/vin nenačte. Použijte Pages/Wrangler na http://127.0.0.1:8788/."
        );
      }

      if (response.status === 404 && !detail) {
        throw new Error("Vozidlo nebylo v registru nalezeno.");
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

function setVinLookupMessage(text, tone = "info") {
  const el = document.getElementById("vinLookupMessage");
  if (!el) {
    return;
  }

  el.textContent = text;
  el.classList.remove("is-visible", "is-error", "is-success", "is-loading", "is-info");

  if (!text) {
    return;
  }

  const normalizedTone = tone === true ? "error" : tone === false ? "info" : String(tone || "info");
  el.classList.add("is-visible");
  el.classList.add(
    normalizedTone === "error" || normalizedTone === "success" || normalizedTone === "loading"
      ? `is-${normalizedTone}`
      : "is-info"
  );
}

function setSignatureImageSize(element, size = DEFAULT_SIGNATURE_SIZE) {
  if (!element) {
    return;
  }

  const normalized = cloneSignatureSize(size);
  element.style.width = `${(normalized.widthRatio * 100).toFixed(2)}%`;
  element.style.height = `${(normalized.heightRatio * 100).toFixed(2)}%`;
}
