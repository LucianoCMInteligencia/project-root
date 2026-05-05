const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const downloadLink = document.getElementById("downloadLink");
const statusDiv = document.getElementById("status");

const supportedTypes = {
  "text/plain": "Texto",
  "image/jpeg": "Imagen",
  "image/png": "Imagen",
  "image/gif": "Imagen",
};

function getFileType(file) {
  if (supportedTypes[file.type]) return supportedTypes[file.type];
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (extension === 'txt') return 'Texto';
  if (['jpg', 'jpeg', 'png', 'gif'].includes(extension)) return 'Imagen';
  return null;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("No se pudo leer el archivo de imagen."));
    reader.readAsDataURL(file);
  });
}

function getPptxGenClass() {
  const lib = window.PptxGenJS || window.pptxgen || window.pptxgenjs || window.PptxGenJS?.default;
  if (typeof lib === "function") return lib;
  if (lib && typeof lib.default === "function") return lib.default;
  if (lib && typeof lib.PptxGenJS === "function") return lib.PptxGenJS;
  throw new Error("La librería PPTX no está disponible en este entorno.");
}

async function createImageSlide(pptx, file) {
  const dataUrl = await fileToDataUrl(file);
  const slide = pptx.addSlide();
  slide.addImage({ data: dataUrl, x: 0.5, y: 0.5, w: 9, h: 4.8 });
  slide.addText(file.name, {
    x: 0.5,
    y: 5.4,
    w: 9,
    h: 0.4,
    fontSize: 12,
    color: "666666",
    align: "left",
  });
}

async function createTextSlide(pptx, file) {
  const text = await file.text();
  const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
  const slide = pptx.addSlide();

  if (lines.length === 0) {
    slide.addText("(archivo de texto vacío)", {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontSize: 18,
      color: "363636",
    });
    return;
  }

  const baseY = 0.5;
  const lineHeight = 0.7;
  lines.forEach((line, index) => {
    slide.addText(line, {
      x: 0.5,
      y: baseY + index * lineHeight,
      w: 9,
      h: 0.6,
      fontSize: 18,
      color: "363636",
      align: "left",
      wrap: true,
    });
  });
}

function createUnsupportedSlide(pptx, file) {
  const slide = pptx.addSlide();
  slide.addText(`Archivo: ${file.name}`, {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.6,
    fontSize: 18,
    color: "ff0000",
    bold: true,
  });
  slide.addText("Tipo no soportado.", {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 0.6,
    fontSize: 16,
    color: "ff0000",
  });
}

convertBtn.addEventListener("click", async () => {
  const files = Array.from(fileInput.files || []);
  if (files.length === 0) {
    statusDiv.textContent = "Por favor, selecciona al menos un archivo.";
    return;
  }

  convertBtn.disabled = true;
  statusDiv.textContent = "Iniciando conversión...";
  downloadLink.classList.remove("visible");
  downloadLink.href = "";

  let pptx;
  try {
    const PptxGenClass = getPptxGenClass();
    pptx = new PptxGenClass();
  } catch (err) {
    statusDiv.textContent = "No se pudo inicializar la librería PPTX: " + (err?.message || err);
    convertBtn.disabled = false;
    return;
  }

  try {
    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      statusDiv.textContent = `Procesando ${i + 1}/${files.length}: ${file.name}`;

      const fileKind = getFileType(file);
      if (fileKind === "Imagen") {
        await createImageSlide(pptx, file);
      } else if (fileKind === "Texto") {
        await createTextSlide(pptx, file);
      } else {
        createUnsupportedSlide(pptx, file);
      }
    }

    statusDiv.textContent = "Generando PPTX...";
    const blob = await pptx.write("blob");
    const url = URL.createObjectURL(blob);
    downloadLink.href = url;
    downloadLink.download = "converted.pptx";
    downloadLink.classList.add("visible");
    statusDiv.textContent = "Conversión completada. Haz clic en descargar.";
  } catch (err) {
    statusDiv.textContent = "Error en la conversión: " + (err?.message || err);
  } finally {
    convertBtn.disabled = false;
  }
});
