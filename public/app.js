// app.js — archivo completo para reemplazar (correcciones: descarga programática, logs, progreso, persistencia)

// Elementos del DOM
const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const downloadLink = document.getElementById("downloadLink");
const statusDiv = document.getElementById("status");
const styleSelect = document.getElementById("styleSelect");

const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");

// Tipos soportados
const supportedTypes = {
  "text/plain": "Texto",
  "image/jpeg": "Imagen",
  "image/png": "Imagen",
  "image/gif": "Imagen",
};

// Detectar tipo de archivo por MIME o extensión
function getFileType(file) {
  if (supportedTypes[file.type]) return supportedTypes[file.type];
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (extension === 'txt') return 'Texto';
  if (['jpg', 'jpeg', 'png', 'gif'].includes(extension)) return 'Imagen';
  return null;
}

// Leer archivo como DataURL (para imágenes)
function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo de imagen."));
    reader.readAsDataURL(file);
  });
}

// Obtener clase de PptxGenJS cargada en el navegador
function getPptxGenClass() {
  const lib = window.PptxGenJS || window.pptxgen || window.pptxgenjs || window.PptxGenJS?.default;
  if (typeof lib === "function") return lib;
  if (lib && typeof lib.default === "function") return lib.default;
  if (lib && typeof lib.PptxGenJS === "function") return lib.PptxGenJS;
  throw new Error("La librería PPTX no está disponible en este entorno.");
}

// --- Lectura de estilos desde la página para mapear a masters ---
function getCssValue(varName, fallback) {
  try {
    const root = document.documentElement;
    const val = getComputedStyle(root).getPropertyValue(varName).trim();
    if (val) return val;
  } catch (e) {
    // ignore
  }
  return fallback;
}

function getStyleConfigFromPage() {
  const primary = getCssValue('--primary-color', '#003366');
  const accent = getCssValue('--accent-color', '#FF6A00');
  const bg = getCssValue('--bg-color', '#FFFFFF');
  const text = getCssValue('--text-color', '#333333');
  const headerFont = getCssValue('--header-font', 'Arial');
  const bodyFont = getCssValue('--body-font', 'Calibri');

  return {
    primary: primary.replace(/\s/g, '') || '#003366',
    accent: accent.replace(/\s/g, '') || '#FF6A00',
    bg: bg.replace(/\s/g, '') || '#FFFFFF',
    text: text.replace(/\s/g, '') || '#333333',
    headerFont: headerFont || 'Arial',
    bodyFont: bodyFont || 'Calibri'
  };
}

// Define masters dinámicos basados en CSS
function defineMastersFromCss(pres) {
  const cfg = getStyleConfigFromPage();

  pres.defineSlideMaster({
    title: "TEMPLATE_SIMPLE",
    background: { color: cfg.bg },
    objects: [
      { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: cfg.primary } } },
      { text: { text: "", options: { x: 0.3, y: 0.1, w: 9, h: 0.5, color: "#FFFFFF", fontFace: cfg.headerFont, fontSize: 20, bold: true } } },
      { text: { text: "", options: { x: 0.3, y: 6.8, w: 9, h: 0.3, color: cfg.text, fontFace: cfg.bodyFont, fontSize: 10 } } }
    ]
  });

  pres.defineSlideMaster({
    title: "TEMPLATE_MODERN",
    background: { color: cfg.bg },
    objects: [
      { rect: { x: 0, y: 6.6, w: "100%", h: 0.9, fill: { color: cfg.primary } } },
      { text: { text: "", options: { x: 0.3, y: 0.2, w: 9, h: 0.5, color: cfg.primary, fontFace: cfg.headerFont, fontSize: 22 } } }
    ]
  });

  pres.defineSlideMaster({
    title: "TEMPLATE_DARK",
    background: { color: "#1F1F1F" },
    objects: [
      { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: cfg.primary } } },
      { text: { text: "", options: { x: 0.3, y: 0.1, w: 9, h: 0.5, color: "#FFFFFF", fontFace: cfg.headerFont, fontSize: 20 } } }
    ]
  });
}

// Fallback: masters estáticos
function defineMastersStatic(pres) {
  pres.defineSlideMaster({
    title: "TEMPLATE_SIMPLE",
    background: { color: "FFFFFF" },
    objects: [
      { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: "003366" } } },
      { text: { text: "", options: { x: 0.3, y: 0.1, w: 9, h: 0.5, color: "FFFFFF", fontSize: 20, bold: true } } },
      { text: { text: "", options: { x: 0.3, y: 6.8, w: 9, h: 0.3, color: "666666", fontSize: 10 } } }
    ]
  });

  pres.defineSlideMaster({
    title: "TEMPLATE_MODERN",
    background: { color: "F7F7F7" },
    objects: [
      { rect: { x: 0, y: 6.6, w: "100%", h: 0.9, fill: { color: "222222" } } },
      { text: { text: "", options: { x: 0.3, y: 0.2, w: 9, h: 0.5, color: "222222", fontSize: 22 } } }
    ]
  });

  pres.defineSlideMaster({
    title: "TEMPLATE_DARK",
    background: { color: "1F1F1F" },
    objects: [
      { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: "111111" } } },
      { text: { text: "", options: { x: 0.3, y: 0.1, w: 9, h: 0.5, color: "FFFFFF", fontSize: 20 } } }
    ]
  });
}

// Crear diapositiva de imagen
async function createImageSlide(pptx, file, masterName) {
  const dataUrl = await fileToDataUrl(file);
  const slide = pptx.addSlide({ masterName });
  slide.addImage({ data: dataUrl, x: 0.5, y: 0.8, w: 9, h: 4.2 });
  slide.addText(file.name, {
    x: 0.5,
    y: 5.2,
    w: 9,
    h: 0.4,
    fontSize: 12,
    color: "666666",
    align: "left",
  });
}

// Crear diapositiva de texto
async function createTextSlide(pptx, file, masterName) {
  const text = await file.text();
  const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
  const slide = pptx.addSlide({ masterName });

  if (lines.length === 0) {
    slide.addText("(archivo de texto vacío)", {
      x: 0.5,
      y: 0.8,
      w: 9,
      h: 1,
      fontSize: 18,
      color: "363636",
    });
    return;
  }

  const baseY = 0.8;
  const lineHeight = 0.6;
  lines.forEach((line, index) => {
    slide.addText(line, {
      x: 0.5,
      y: baseY + index * lineHeight,
      w: 9,
      h: 0.6,
      fontSize: 16,
      color: "363636",
      align: "left",
      wrap: true,
    });
  });
}

// Crear diapositiva para archivos no soportados
function createUnsupportedSlide(pptx, file, masterName) {
  const slide = pptx.addSlide({ masterName });
  slide.addText(`Archivo: ${file.name}`, {
    x: 0.5,
    y: 0.8,
    w: 9,
    h: 0.6,
    fontSize: 18,
    color: "ff0000",
    bold: true,
  });
  slide.addText("Tipo no soportado.", {
    x: 0.5,
    y: 1.6,
    w: 9,
    h: 0.6,
    fontSize: 16,
    color: "ff0000",
  });
}

// --- Progreso y persistencia de selección ---
try {
  const saved = localStorage.getItem("pptx_selected_style");
  if (saved && styleSelect) styleSelect.value = saved;
} catch (e) {
  // localStorage puede no estar disponible; ignorar
}

if (styleSelect) {
  styleSelect.addEventListener("change", () => {
    try {
      localStorage.setItem("pptx_selected_style", styleSelect.value);
    } catch (e) {
      // ignorar errores de almacenamiento
    }
  });
}

function showProgress(totalFiles) {
  if (!progressContainer || !progressBar || !progressText) return;
  progressContainer.classList.remove("hidden");
  progressContainer.setAttribute("aria-hidden", "false");
  progressBar.value = 0;
  progressBar.max = totalFiles;
  progressText.textContent = `0 / ${totalFiles}`;
}

function updateProgress(processed, total) {
  if (!progressBar || !progressText) return;
  progressBar.value = processed;
  progressText.textContent = `${processed} / ${total}`;
}

function hideProgress() {
  if (!progressContainer) return;
  progressContainer.classList.add("hidden");
  progressContainer.setAttribute("aria-hidden", "true");
}

// --- Evento principal: conversión ---
convertBtn.addEventListener("click", async () => {
  const files = Array.from(fileInput.files || []);
  if (files.length === 0) {
    statusDiv.textContent = "Por favor, selecciona al menos un archivo.";
    return;
  }

  convertBtn.disabled = true;
  statusDiv.textContent = "Iniciando conversión...";
  console.log("Conversión iniciada. Archivos seleccionados:", files.length);

  if (downloadLink) {
    downloadLink.classList.remove("visible");
    downloadLink.href = "";
    downloadLink.classList.add("hidden");
    downloadLink.setAttribute("aria-hidden", "true");
  }

  let pptx;
  try {
    const PptxGenClass = getPptxGenClass();
    pptx = new PptxGenClass();

    // Intentar definir masters dinámicos desde CSS; si falla, usar estáticos
    try {
      defineMastersFromCss(pptx);
    } catch (e) {
      console.warn("defineMastersFromCss falló, usando masters estáticos:", e);
      defineMastersStatic(pptx);
    }
  } catch (err) {
    statusDiv.textContent = "No se pudo inicializar la librería PPTX: " + (err?.message || err);
    console.error("Error inicializando PptxGenJS:", err);
    convertBtn.disabled = false;
    return;
  }

  const selectedMaster = (styleSelect && styleSelect.value) ? styleSelect.value : "TEMPLATE_SIMPLE";
  try {
    try { localStorage.setItem("pptx_selected_style", selectedMaster); } catch (e) {}

    showProgress(files.length);

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      statusDiv.textContent = `Procesando ${i + 1}/${files.length}: ${file.name}`;
      console.log(`Procesando archivo ${i + 1}/${files.length}:`, file.name);

      const fileKind = getFileType(file);
      if (fileKind === "Imagen") {
        await createImageSlide(pptx, file, selectedMaster);
      } else if (fileKind === "Texto") {
        await createTextSlide(pptx, file, selectedMaster);
      } else {
        createUnsupportedSlide(pptx, file, selectedMaster);
      }

      updateProgress(i + 1, files.length);
    }

    statusDiv.textContent = "Generando PPTX...";
    console.log("Generando blob PPTX...");
    const blob = await pptx.write("blob");
    console.log("Blob generado, tamaño (bytes):", blob.size);

    // Crear URL y forzar descarga programáticamente
    const url = URL.createObjectURL(blob);
    try {
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;
      a.download = "converted.pptx";
      document.body.appendChild(a);

      console.log("Disparando descarga automática...");
      a.click();
      a.remove();

      // Mantener enlace visible como respaldo
      if (downloadLink) {
        downloadLink.href = url;
        downloadLink.download = "converted.pptx";
        downloadLink.classList.add("visible");
        downloadLink.classList.remove("hidden");
        downloadLink.setAttribute("aria-hidden", "false");
      }

      statusDiv.textContent = "Conversión completada. Si la descarga no inició, usa el enlace de descarga.";
    } catch (err) {
      console.error("Error al forzar descarga automática:", err);
      // Fallback: mostrar enlace para que el usuario haga click manualmente
      if (downloadLink) {
        downloadLink.href = url;
        downloadLink.download = "converted.pptx";
        downloadLink.classList.add("visible");
        downloadLink.classList.remove("hidden");
        downloadLink.setAttribute("aria-hidden", "false");
      }
      statusDiv.textContent = "Conversión completada. Haz clic en descargar.";
    } finally {
      // Revocar la URL tras un pequeño retardo para dar tiempo al navegador a iniciar la descarga
      setTimeout(() => {
        try { URL.revokeObjectURL(url); console.log("URL revocada"); } catch (e) { /* ignore */ }
      }, 3000);
    }
  } catch (err) {
    statusDiv.textContent = "Error en la conversión: " + (err?.message || err);
    console.error("Error durante la conversión:", err);
  } finally {
    convertBtn.disabled = false;
    setTimeout(hideProgress, 800);
  }
});
