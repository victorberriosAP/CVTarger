
const fileInput = document.querySelector("#cv-file");
const uploadZone = document.querySelector("#upload-zone");
const uploadTitle = document.querySelector("#upload-title");
const fileStatus = document.querySelector("#file-status");
const jobOffer = document.querySelector("#job-offer");
const charCount = document.querySelector("#char-count");
const adaptButton = document.querySelector("#adapt-button");
const formMessage = document.querySelector("#form-message");

const resultTitle = document.querySelector("#result-title");
const resultDescription = document.querySelector("#result-description");
const resultContent = document.querySelector("#result-content");
const adaptedCV = document.querySelector("#adapted-cv");
const resultSkeleton = document.querySelector("#result-skeleton");
const copyButton = document.querySelector("#copy-result");
const downloadWordButton = document.querySelector("#download-word");
let wordBase64 = "";
let textoPlanoCV = ""; // Versión sin símbolos (**, ##) para el botón "Copiar".

const MAX_FILE_SIZE = 10 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ["pdf", "docx"];

// ======================================================
// FORMATO DEL CV EN LA PÁGINA
// Estas funciones hacen que la vista previa se vea igual
// que el archivo Word: nombre centrado, títulos azules en
// mayúsculas, negritas, cursivas y viñetas reales.
// Siguen las mismas reglas que usa el servidor para
// crear el Word (archivo adaptar-cv.js).
// ======================================================

// Títulos habituales de un currículum (igual que en el servidor).
const SECCIONES_CV = [
  "perfil profesional",
  "resumen profesional",
  "objetivo profesional",
  "experiencia laboral",
  "experiencia relevante",
  "experiencia profesional",
  "experiencia",
  "educación",
  "formación académica",
  "formación",
  "habilidades",
  "competencias",
  "competencias técnicas",
  "certificaciones",
  "idiomas",
  "proyectos",
  "logros",
  "información adicional",
  "referencias",
];

// Protege la página: convierte símbolos como < y > en texto
// inofensivo para que nunca se ejecute código extraño.
function escaparHtml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Quita las marcas de título, viñeta y escapes de una línea.
function limpiarLinea(texto) {
  return texto
    .trim()
    .replace(/^#{1,6}\s+/, "")
    .replace(/^[-*•]\s+/, "")
    .replace(/\\([*_#-])/g, "$1")
    .trim();
}

// Quita todas las marcas de negrita y cursiva de un texto.
function quitarMarcas(texto) {
  return texto
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/\*(.*?)\*/g, "$1")
    .replace(/_(.*?)_/g, "$1");
}

// Indica si una línea es el título de una sección (Experiencia, Educación...).
function esEncabezado(texto) {
  const titulo = texto
    .replace(/:$/, "")
    .replace(/\*\*/g, "")
    .replace(/_/g, "")
    .trim()
    .toLocaleLowerCase("es");

  return SECCIONES_CV.includes(titulo);
}

// Convierte **negrita**, *cursiva* y _cursiva_ en HTML real.
function formatoEnLinea(texto) {
  const partes = escaparHtml(texto)
    .split(/(\*\*.*?\*\*|\*.*?\*|_.*?_)/g)
    .filter(Boolean);

  return partes
    .map((parte) => {
      if (/^\*\*.*\*\*$/.test(parte)) {
        return `<strong>${parte.slice(2, -2)}</strong>`;
      }

      if (/^\*.*\*$/.test(parte) || /^_.*_$/.test(parte)) {
        return `<em>${parte.slice(1, -1)}</em>`;
      }

      return parte;
    })
    .join("");
}

// Lee el texto de la IA línea por línea y decide qué es cada una:
// nombre, contacto, título de sección, viñeta o texto normal.
function interpretarCV(texto) {
  const lineas = texto.replace(/\r\n/g, "\n").split("\n");
  const bloques = [];

  let lineasContacto = 0;
  let comenzoContenido = false;

  for (const linea of lineas) {
    const original = linea.trim();

    // Ignorar líneas vacías y separadores (---).
    if (!original || /^\\?---+$/.test(original)) continue;

    const esTituloMarkdown = /^#{1,6}\s+/.test(original);
    const esVineta = /^[-*•]\s+/.test(original);
    const contenido = limpiarLinea(original);

    if (!contenido) continue;

    // La primera línea es el nombre del candidato.
    if (bloques.length === 0) {
      bloques.push({ tipo: "nombre", texto: contenido });
      continue;
    }

    // Títulos de sección.
    if (esTituloMarkdown || esEncabezado(contenido)) {
      comenzoContenido = true;
      bloques.push({ tipo: "titulo", texto: contenido });
      continue;
    }

    // Especialidades y datos de contacto, antes de la primera sección.
    if (!comenzoContenido && lineasContacto < 3) {
      lineasContacto++;
      bloques.push({ tipo: "contacto", texto: contenido, primera: lineasContacto === 1 });
      continue;
    }

    // Todo lo demás: viñetas o texto normal.
    bloques.push({ tipo: esVineta ? "vineta" : "normal", texto: contenido });
  }

  return bloques;
}

// Crea el HTML de la hoja de papel con el CV.
function construirHtmlCV(texto) {
  const bloques = interpretarCV(texto);
  let html = "";
  let listaAbierta = false;

  for (const bloque of bloques) {
    // Si termina una serie de viñetas, se cierra la lista.
    if (bloque.tipo !== "vineta" && listaAbierta) {
      html += "</ul>";
      listaAbierta = false;
    }

    if (bloque.tipo === "nombre") {
      html += `<h3 class="cv-nombre">${escaparHtml(quitarMarcas(bloque.texto))}</h3>`;
    } else if (bloque.tipo === "titulo") {
      html += `<h4 class="cv-titulo">${escaparHtml(quitarMarcas(bloque.texto)).toLocaleUpperCase("es")}</h4>`;
    } else if (bloque.tipo === "contacto") {
      html += `<p class="cv-contacto${bloque.primera ? " cv-primera" : ""}">${formatoEnLinea(bloque.texto)}</p>`;
    } else if (bloque.tipo === "vineta") {
      if (!listaAbierta) {
        html += "<ul>";
        listaAbierta = true;
      }
      html += `<li>${formatoEnLinea(bloque.texto)}</li>`;
    } else {
      html += `<p class="cv-texto">${formatoEnLinea(bloque.texto)}</p>`;
    }
  }

  if (listaAbierta) html += "</ul>";

  return html;
}

// Crea una versión en texto simple (sin ** ni ##) para copiar y pegar.
function construirTextoPlanoCV(texto) {
  const lineas = [];

  for (const bloque of interpretarCV(texto)) {
    const limpio = quitarMarcas(bloque.texto);

    if (bloque.tipo === "titulo") {
      lineas.push("", limpio.toLocaleUpperCase("es"));
    } else if (bloque.tipo === "vineta") {
      lineas.push(`• ${limpio}`);
    } else {
      lineas.push(limpio);
    }
  }

  return lineas.join("\n").trim();
}

let selectedFile = null;

function showMessage(message, success = false) {
  formMessage.textContent = message;
  formMessage.classList.remove("d-none", "alert-warning", "alert-success");
  formMessage.classList.add(success ? "alert-success" : "alert-warning");
}

function clearMessage() {
  formMessage.textContent = "";
  formMessage.classList.add("d-none");
}

function setFile(file) {
  clearMessage();

  if (!file) return;

  const extension = file.name.split(".").pop().toLowerCase();

  if (!ALLOWED_EXTENSIONS.includes(extension)) {
    selectedFile = null;
    fileInput.value = "";
    uploadTitle.textContent = "Arrastra tu CV aquí";
    fileStatus.textContent = "Selecciona un archivo PDF o DOCX.";
    showMessage("El archivo debe estar en formato PDF o DOCX.");
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    selectedFile = null;
    fileInput.value = "";
    uploadTitle.textContent = "Arrastra tu CV aquí";
    fileStatus.textContent = "El archivo supera el límite de 10 MB.";
    showMessage("El tamaño máximo permitido es 10 MB.");
    return;
  }

  selectedFile = file;
  uploadTitle.textContent = file.name;
  fileStatus.textContent =
    `Archivo seleccionado · ${(file.size / (1024 * 1024)).toFixed(2)} MB`;
}

fileInput.addEventListener("change", (event) => {
  setFile(event.target.files[0]);
});

["dragenter", "dragover"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.remove("dragging");
  });
});

uploadZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;

  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    fileInput.files = transfer.files;
  } catch (_) {
    // El archivo seleccionado igualmente se valida.
  }

  setFile(file);
});

jobOffer.addEventListener("input", () => {
  const count = jobOffer.value.length;
  charCount.textContent =
    `${count.toLocaleString("es-CL")} ${count === 1 ? "carácter" : "caracteres"}`;
  clearMessage();
});

adaptButton.addEventListener("click", async () => {
  clearMessage();

  if (!selectedFile) {
    showMessage("Primero selecciona tu currículum en PDF o DOCX.");
    fileInput.focus();
    return;
  }

  if (!jobOffer.value.trim()) {
    showMessage("Pega la oferta laboral antes de continuar.");
    jobOffer.focus();
    return;
  }

  const textoOriginal = adaptButton.textContent;

  adaptButton.disabled = true;
  adaptButton.textContent = "Adaptando tu CV...";
  resultSkeleton.classList.remove("d-none");
  resultContent.classList.add("d-none");
  resultTitle.textContent = "Estamos adaptando tu currículum";
  resultDescription.textContent =
    "Esto puede tardar unos segundos. No cierres esta página.";

  // Baja hasta la zona de resultados para que se vea el progreso.
  resultTitle.scrollIntoView({ behavior: "smooth", block: "center" });

  try {
    const datos = new FormData();
    datos.append("cv", selectedFile);
    datos.append("oferta", jobOffer.value.trim());

    const respuesta = await fetch("/api/adaptar-cv", {
      method: "POST",
      body: datos,
    });

    // Si el servidor responde algo que no es JSON (por ejemplo una
    // página de error de Vercel), no se rompe: se muestra el código.
    let resultado = {};
    try {
      resultado = await respuesta.json();
    } catch (_) {
      // Respuesta sin formato JSON.
    }

    if (!respuesta.ok) {
      throw new Error(
        resultado.error ||
          `El servidor respondió con el código ${respuesta.status}. Revisa los registros en Vercel.`
      );
    }

    if (!resultado.resultado) {
      throw new Error("La IA no devolvió un currículum. Inténtalo de nuevo.");
    }

    // Dibuja el CV como una hoja con el mismo formato del Word.
    adaptedCV.innerHTML = construirHtmlCV(resultado.resultado);

    // Guarda la versión sin símbolos para el botón "Copiar".
    textoPlanoCV = construirTextoPlanoCV(resultado.resultado);

    wordBase64 = resultado.docxBase64 || "";
    downloadWordButton.classList.toggle("d-none", !wordBase64);
    resultTitle.textContent = "¡Tu CV está adaptado!";
    resultDescription.textContent =
      "Revisa el contenido antes de utilizarlo en tu postulación.";
    resultContent.classList.remove("d-none");
    resultSkeleton.classList.add("d-none");

    showMessage("Tu currículum se adaptó correctamente.", true);
  } catch (error) {
    // Deja el detalle del error en la consola del navegador (F12).
    console.error("Error al adaptar el CV:", error);

    resultTitle.textContent = "No pudimos adaptar tu CV";
    resultDescription.textContent =
      "Comprueba tu conexión y vuelve a intentarlo.";
    resultSkeleton.classList.add("d-none");

    showMessage(error.message || "Ocurrió un error inesperado.");
  } finally {
    adaptButton.disabled = false;
    adaptButton.textContent = textoOriginal;
  }
});

copyButton.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(textoPlanoCV);
    showMessage("CV adaptado copiado al portapapeles.", true);
  } catch (_) {
    showMessage("No se pudo copiar automáticamente. Selecciona el texto y cópialo manualmente.");
  }
});

//Descargar Word Modificado
downloadWordButton.addEventListener("click", () => {
  if (!wordBase64) {
    showMessage("Primero adapta tu currículum.");
    return;
  }

  const bytes = Uint8Array.from(
    atob(wordBase64),
    (caracter) => caracter.charCodeAt(0)
  );

  const archivo = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  const url = URL.createObjectURL(archivo);
  const enlace = document.createElement("a");

  enlace.href = url;
  enlace.download = "CV_adaptado.docx";
  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});
