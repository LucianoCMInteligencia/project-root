const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const downloadLink = document.getElementById("downloadLink");
const statusDiv = document.getElementById("status");

convertBtn.addEventListener("click", async () => {
  const files = fileInput.files;
  if (files.length === 0) {
    statusDiv.textContent = "Por favor, selecciona al menos un archivo.";
    return;
  }

  statusDiv.textContent = "Convirtiendo...";

  const pptx = new PptxGenJS();

  for (const file of files) {
    const slide = pptx.addSlide();

    if (file.type.startsWith("image/")) {
      // Para imágenes, añadir como imagen en la diapositiva
      const url = URL.createObjectURL(file);
      slide.addImage({ path: url, x: 0, y: 0, w: 10, h: 5.625 }); // Tamaño A4 aproximado
    } else if (file.type === "text/plain") {
      // Para texto, leer y añadir como texto
      const text = await file.text();
      slide.addText(text, { x: 0.5, y: 0.5, w: 9, h: 4.625, fontSize: 18 });
    } else {
      // Para otros tipos, mostrar mensaje
      slide.addText(`Archivo: ${file.name}\nTipo no soportado directamente.`, { x: 0.5, y: 0.5, w: 9, h: 4.625 });
    }
  }

  // Generar el archivo PPTX
  pptx.writeFile({ fileName: "converted.pptx" }).then(() => {
    statusDiv.textContent = "Conversión completada. Descarga el archivo.";
    downloadLink.style.display = "block";
  }).catch((err) => {
    statusDiv.textContent = "Error en la conversión: " + err.message;
  });
});
