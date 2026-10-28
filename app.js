const fileInput = document.querySelector("#cv-file");
const uploadZone = document.querySelector("#upload-zone");
const uploadTitle = document.querySelector("#upload-title");
const fileStatus = document.querySelector("#file-status");
const jobOffer = document.querySelector("#job-offer");
const charCount = document.querySelector("#char-count");
const adaptButton = document.querySelector("#adapt-button");

// ======================================================
// ENTRADA DE LA OFERTA: DESCRIPCIÓN O URL
// Se crean los controles desde JavaScript para conservar
// el HTML existente.
// IMPORTANTE: el backend también debe aceptar "urlOferta".
// ======================================================
const offerLabel = document.querySelector('label[for="job-offer"]');
const offerCountRow = charCount ? charCount.closest(".d-flex") : null;

const offerUrlInput = document.createElement("input");
offerUrlInput.type = "url";
offerUrlInput.id = "job-offer-url";
offerUrlInput.name = "urlOferta";
offerUrlInput.className = "form-control";
offerUrlInput.placeholder = "https://sitio.com/oferta-laboral";
offerUrlInput.autocomplete = "url";
offerUrlInput.inputMode = "url";
offerUrlInput.setAttribute("aria-label", "URL de la oferta laboral");

const offerUrlHelp = document.createElement("p");
offerUrlHelp.className = "small text-secondary mt-2 mb-3";
offerUrlHelp.textContent =
  "Usa el enlace público del anuncio. Si el sitio requiere iniciar sesión o bloquea el acceso, pega la descripción manualmente.";

// Botones para cambiar entre descripción y URL.
const offerModeControls = document.createElement("div");
offerModeControls.className = "d-flex flex-wrap gap-2 mb-3";
offerModeControls.setAttribute("role", "group");
offerModeControls.setAttribute(
  "aria-label",
  "Forma de ingresar la oferta laboral"
);

offerModeControls.innerHTML = `
  <button
    type="button"
    class="btn btn-sm btn-primary"
    id="offer-mode-text"
    aria-pressed="true"
  >
    Pegar descripción
  </button>

  <button
    type="button"
    class="btn btn-sm btn-outline-primary"
    id="offer-mode-url"
    aria-pressed="false"
  >
    Usar URL
  </button>
`;

const offerTextModeButton =
  offerModeControls.querySelector("#offer-mode-text");

const offerUrlModeButton =
  offerModeControls.querySelector("#offer-mode-url");

// Por defecto se utiliza la descripción manual.
let offerMode = "text";

// Inserta los controles junto al campo de oferta existente.
if (offerLabel && jobOffer) {
  offerLabel.parentNode.insertBefore(
    offerModeControls,
    offerLabel
  );

  offerLabel.parentNode.insertBefore(
    offerUrlInput,
    offerLabel
  );

  offerLabel.parentNode.insertBefore(
    offerUrlHelp,
    offerLabel
  );
}

// Oculta inicialmente los campos de URL.
offerUrlInput.classList.add("d-none");
offerUrlHelp.classList.add("d-none");

// Cambia entre el campo de texto y el campo de URL.
function setOfferMode(mode) {
  offerMode = mode;

  const usarUrl = mode === "url";

  offerLabel?.classList.toggle("d-none", usarUrl);
  jobOffer?.classList.toggle("d-none", usarUrl);
  offerCountRow?.classList.toggle("d-none", usarUrl);

  offerUrlInput.classList.toggle("d-none", !usarUrl);
  offerUrlHelp.classList.toggle("d-none", !usarUrl);

  // Actualiza los estilos de los botones.
  offerTextModeButton.className =
    `btn btn-sm ${
      usarUrl ? "btn-outline-primary" : "btn-primary"
    }`;

  offerUrlModeButton.className =
    `btn btn-sm ${
      usarUrl ? "btn-primary" : "btn-outline-primary"
    }`;

  // Actualiza los estados accesibles.
  offerTextModeButton.setAttribute(
    "aria-pressed",
    String(!usarUrl)
  );

  offerUrlModeButton.setAttribute(
    "aria-pressed",
    String(usarUrl)
  );

  clearMessage();
}

// Eventos para cambiar de modalidad.
offerTextModeButton.addEventListener("click", () => {
  setOfferMode("text");
});

offerUrlModeButton.addEventListener("click", () => {
  setOfferMode("url");
});

offerUrlInput.addEventListener("input", clearMessage);

// ======================================================
// ELEMENTOS DE LA INTERFAZ
// ======================================================
const formMessage = document.querySelector("#form-message");

const resultTitle = document.querySelector("#result-title");
const resultDescription =
  document.querySelector("#result-description");
const resultContent = document.querySelector("#result-content");
const adaptedCV = document.querySelector("#adapted-cv");
const resultSkeleton = document.querySelector("#result-skeleton");

const copyButton = document.querySelector("#copy-result");
const downloadWordButton =
  document.querySelector("#download-word");

let wordBase64 = "";
let textoPlanoCV = "";

// Límite de tamaño y formatos admitidos.
const MAX_FILE_SIZE = 10 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ["pdf", "docx"];

// ======================================================
// FORMATO DEL CV EN LA PÁGINA
// ======================================================

// Títulos habituales de las secciones de un currículum.
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

// Convierte caracteres especiales en texto seguro para HTML.
function escaparHtml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Elimina marcas Markdown de encabezados y viñetas.
function limpiarLinea(texto) {
  return texto
    .trim()
    .replace(/^#{1,6}\s+/, "")
    .replace(/^[-*•]\s+/, "")
    .replace(/\\([*_#-])/g, "$1")
    .trim();
}

// Elimina las marcas de negrita y cursiva.
function quitarMarcas(texto) {
  return texto
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/\*(.*?)\*/g, "$1")
    .replace(/_(.*?)_/g, "$1");
}

// Comprueba si una línea corresponde a una sección del CV.
function esEncabezado(texto) {
  const titulo = texto
    .replace(/:$/, "")
    .replace(/\*\*/g, "")
    .replace(/_/g, "")
    .trim()
    .toLocaleLowerCase("es");

  return SECCIONES_CV.includes(titulo);
}

// Convierte marcas Markdown en etiquetas HTML.
function formatoEnLinea(texto) {
  const partes = escaparHtml(texto)
    .split(/(\*\*.*?\*\*|\*.*?\*|_.*?_)/g)
    .filter(Boolean);

  return partes
    .map((parte) => {
      if (/^\*\*.*\*\*$/.test(parte)) {
        return `<strong>${parte.slice(2, -2)}</strong>`;
      }

      if (
        /^\*.*\*$/.test(parte) ||
        /^_.*_$/.test(parte)
      ) {
        return `<em>${parte.slice(1, -1)}</em>`;
      }

      return parte;
    })
    .join("");
}

// Interpreta el contenido del CV línea por línea.
function interpretarCV(texto) {
  const lineas = texto.replace(/\r\n/g, "\n").split("\n");
  const bloques = [];

  let lineasContacto = 0;
  let comenzoContenido = false;

  for (const linea of lineas) {
    const original = linea.trim();

    // Ignora líneas vacías y separadores.
    if (!original || /^\\?---+$/.test(original)) {
      continue;
    }

    const esTituloMarkdown = /^#{1,6}\s+/.test(original);
    const esVineta = /^[-*•]\s+/.test(original);
    const contenido = limpiarLinea(original);

    if (!contenido) continue;

    // La primera línea se considera el nombre.
    if (bloques.length === 0) {
      bloques.push({
        tipo: "nombre",
        texto: contenido,
      });
      continue;
    }

    // Identifica los títulos de sección.
    if (esTituloMarkdown || esEncabezado(contenido)) {
      comenzoContenido = true;

      bloques.push({
        tipo: "titulo",
        texto: contenido,
      });

      continue;
    }

    // Las primeras líneas pueden contener datos de contacto.
    if (!comenzoContenido && lineasContacto < 3) {
      lineasContacto++;

      bloques.push({
        tipo: "contacto",
        texto: contenido,
        primera: lineasContacto === 1,
      });

      continue;
    }

    // El resto es una viñeta o un párrafo.
    bloques.push({
      tipo: esVineta ? "vineta" : "normal",
      texto: contenido,
    });
  }

  return bloques;
}

// Construye el HTML de la vista previa del currículum.
function construirHtmlCV(texto) {
  const bloques = interpretarCV(texto);

  let html = "";
  let listaAbierta = false;

  for (const bloque of bloques) {
    // Cierra una lista cuando termina la serie de viñetas.
    if (bloque.tipo !== "vineta" && listaAbierta) {
      html += "</ul>";
      listaAbierta = false;
    }

    if (bloque.tipo === "nombre") {
      html += `
        <h3 class="cv-nombre">
          ${escaparHtml(quitarMarcas(bloque.texto))}
        </h3>
      `;
    } else if (bloque.tipo === "titulo") {
      html += `
        <h4 class="cv-titulo">
          ${escaparHtml(
            quitarMarcas(bloque.texto).toLocaleUpperCase("es")
          )}
        </h4>
      `;
    } else if (bloque.tipo === "contacto") {
      html += `
        <p class="cv-contacto${
          bloque.primera ? " cv-primera" : ""
        }">
          ${formatoEnLinea(bloque.texto)}
        </p>
      `;
    } else if (bloque.tipo === "vineta") {
      if (!listaAbierta) {
        html += "<ul>";
        listaAbierta = true;
      }

      html += `<li>${formatoEnLinea(bloque.texto)}</li>`;
    } else {
      html += `
        <p class="cv-texto">
          ${formatoEnLinea(bloque.texto)}
        </p>
      `;
    }
  }

  if (listaAbierta) {
    html += "</ul>";
  }

  return html;
}

// Construye una versión de texto plano para copiar.
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

// ======================================================
// VALIDACIÓN Y CARGA DEL CURRÍCULUM
// ======================================================
let selectedFile = null;

// Muestra un mensaje en la interfaz.
function showMessage(message, success = false) {
  formMessage.textContent = message;

  formMessage.classList.remove(
    "d-none",
    "alert-warning",
    "alert-success"
  );

  formMessage.classList.add(
    success ? "alert-success" : "alert-warning"
  );
}

// Limpia el mensaje anterior.
function clearMessage() {
  formMessage.textContent = "";
  formMessage.classList.add("d-none");
}

// Valida el archivo seleccionado.
function setFile(file) {
  clearMessage();

  if (!file) return;

  const extension = file.name.split(".").pop().toLowerCase();

  if (!ALLOWED_EXTENSIONS.includes(extension)) {
    selectedFile = null;
    fileInput.value = "";

    uploadTitle.textContent = "Arrastra tu CV aquí";
    fileStatus.textContent =
      "Selecciona un archivo PDF o DOCX.";

    showMessage(
      "El archivo debe estar en formato PDF o DOCX."
    );

    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    selectedFile = null;
    fileInput.value = "";

    uploadTitle.textContent = "Arrastra tu CV aquí";
    fileStatus.textContent =
      "El archivo supera el límite de 10 MB.";

    showMessage("El tamaño máximo permitido es 10 MB.");

    return;
  }

  selectedFile = file;

  uploadTitle.textContent = file.name;

  fileStatus.textContent =
    `Archivo seleccionado · ${
      (file.size / (1024 * 1024)).toFixed(2)
    } MB`;
}

// Evento al seleccionar un archivo.
fileInput.addEventListener("change", (event) => {
  setFile(event.target.files[0]);
});

// Efectos visuales para arrastrar archivos.
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

// Permite cargar el CV arrastrándolo a la zona de carga.
uploadZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];

  if (!file) return;

  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    fileInput.files = transfer.files;
  } catch (_) {
    // El archivo se valida aunque DataTransfer no esté disponible.
  }

  setFile(file);
});

// Actualiza el contador de caracteres de la descripción.
jobOffer.addEventListener("input", () => {
  const count = jobOffer.value.length;

  charCount.textContent =
    `${count.toLocaleString("es-CL")} ${
      count === 1 ? "carácter" : "caracteres"
    }`;

  clearMessage();
});

// ======================================================
// ADAPTAR EL CV
// Permite enviar una descripción o una URL al backend.
// ======================================================
adaptButton.addEventListener("click", async () => {
  clearMessage();

  // El CV es obligatorio.
  if (!selectedFile) {
    showMessage(
      "Primero selecciona tu currículum en PDF o DOCX."
    );

    fileInput.focus();
    return;
  }

  let ofertaTexto = "";
  let ofertaUrl = "";

  // Valida el campo correspondiente a la modalidad elegida.
  if (offerMode === "text") {
    ofertaTexto = jobOffer.value.trim();

    if (!ofertaTexto) {
      showMessage(
        "Pega la descripción de la oferta laboral antes de continuar."
      );

      jobOffer.focus();
      return;
    }
  } else {
    ofertaUrl = offerUrlInput.value.trim();

    if (!ofertaUrl) {
      showMessage(
        "Introduce la URL de la oferta laboral antes de continuar."
      );

      offerUrlInput.focus();
      return;
    }

    // Validación básica de la URL en el navegador.
    // El backend debe repetir la validación por seguridad.
    try {
      const parsedUrl = new URL(ofertaUrl);

      if (!["http:", "https:"].includes(parsedUrl.protocol)) {
        throw new Error("Protocolo no permitido");
      }

      if (
        !parsedUrl.hostname ||
        !parsedUrl.hostname.includes(".")
      ) {
        throw new Error("Dominio no válido");
      }
    } catch (_) {
      showMessage(
        "Introduce una URL válida que comience con https:// o http://."
      );

      offerUrlInput.focus();
      return;
    }
  }

  // Guarda el texto original del botón.
  const textoOriginal = adaptButton.textContent;

  // Muestra el estado de carga.
  adaptButton.disabled = true;
  adaptButton.textContent = "Adaptando tu CV...";

  resultSkeleton.classList.remove("d-none");
  resultContent.classList.add("d-none");

  resultTitle.textContent =
    "Estamos adaptando tu currículum";

  resultDescription.textContent =
    "Esto puede tardar unos segundos. No cierres esta página.";

  resultTitle.scrollIntoView({
    behavior: "smooth",
    block: "center",
  });

  try {
    // Prepara los datos que se enviarán al servidor.
    const datos = new FormData();

    datos.append("cv", selectedFile);

    // Envía únicamente el tipo de entrada elegido.
    if (offerMode === "text") {
      datos.append("oferta", ofertaTexto);
    } else {
      datos.append("urlOferta", ofertaUrl);
    }

    // Llama a la API que adapta el currículum.
    const respuesta = await fetch("/api/adaptar-cv", {
      method: "POST",
      body: datos,
    });

    // Intenta leer la respuesta como JSON.
    let resultado = {};

    try {
      resultado = await respuesta.json();
    } catch (_) {
      // La respuesta del servidor no era JSON válido.
    }

    // Maneja los errores devueltos por la API.
    if (!respuesta.ok) {
      throw new Error(
        resultado.error ||
          `El servidor respondió con el código ${respuesta.status}. Revisa los registros en Vercel.`
      );
    }

    if (!resultado.resultado) {
      throw new Error(
        "La IA no devolvió un currículum. Inténtalo de nuevo."
      );
    }

    // Muestra el currículum adaptado en la página.
    adaptedCV.innerHTML = construirHtmlCV(
      resultado.resultado
    );

    // Prepara el texto plano para el botón de copiar.
    textoPlanoCV = construirTextoPlanoCV(
      resultado.resultado
    );

    // Guarda el archivo Word que devuelve el servidor.
    wordBase64 = resultado.docxBase64 || "";

    downloadWordButton.classList.toggle(
      "d-none",
      !wordBase64
    );

    resultTitle.textContent = "¡Tu CV está adaptado!";

    resultDescription.textContent =
      "Revisa el contenido antes de utilizarlo en tu postulación.";

    resultContent.classList.remove("d-none");
    resultSkeleton.classList.add("d-none");

    showMessage(
      "Tu currículum se adaptó correctamente.",
      true
    );
  } catch (error) {
    // Registra el error para poder revisarlo en la consola.
    console.error("Error al adaptar el CV:", error);

    resultTitle.textContent =
      "No pudimos adaptar tu CV";

    resultDescription.textContent =
      "Comprueba tu conexión y vuelve a intentarlo.";

    resultSkeleton.classList.add("d-none");

    // Si no se pudo leer la URL, pasa al modo "Pegar descripción".
    if (
      offerMode === "url" &&
      (error.message || "").includes("pegando la descripción")
    ) {
      setOfferMode("text");
      jobOffer.focus();

      resultDescription.textContent =
        "Pega la descripción de la oferta y vuelve a intentarlo.";

      showMessage(
        "No pudimos leer esa URL (muchos portales bloquean el acceso automático). Pega aquí la descripción de la oferta."
      );
    } else {
      showMessage(
        error.message || "Ocurrió un error inesperado."
      );
    }
  } finally {
    // Restablece el botón incluso si ocurre un error.
    adaptButton.disabled = false;
    adaptButton.textContent = textoOriginal;
  }
});

// ======================================================
// COPIAR EL CV ADAPTADO
// ======================================================
copyButton.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(textoPlanoCV);

    showMessage(
      "CV adaptado copiado al portapapeles.",
      true
    );
  } catch (_) {
    showMessage(
      "No se pudo copiar automáticamente. Selecciona el texto y cópialo manualmente."
    );
  }
});

// ======================================================
// DESCARGAR EL CV EN WORD
// ======================================================
downloadWordButton.addEventListener("click", () => {
  if (!wordBase64) {
    showMessage("Primero adapta tu currículum.");
    return;
  }

  // Convierte el contenido Base64 en bytes.
  const bytes = Uint8Array.from(
    atob(wordBase64),
    (caracter) => caracter.charCodeAt(0)
  );

  // Crea el archivo Word.
  const archivo = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  // Crea un enlace temporal para iniciar la descarga.
  const url = URL.createObjectURL(archivo);
  const enlace = document.createElement("a");

  enlace.href = url;
  enlace.download = "CV_adaptado.docx";

  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();

  // Libera la URL temporal.
  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);
});