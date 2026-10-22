import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  LevelFormat,
} from "docx";

import Groq from "groq-sdk";
import mammoth from "mammoth";

// IMPORTANTE (arregla el error "DOMMatrix is not defined" en Vercel):
// la línea de "pdf-parse/worker" debe ir SIEMPRE antes que la de
// "pdf-parse". Así Vercel incluye el componente @napi-rs/canvas
// que necesita la librería para leer PDFs.
import { CanvasFactory } from "pdf-parse/worker";
import { PDFParse } from "pdf-parse";

import { formidable } from "formidable";
import { readFile } from "node:fs/promises";

// ======================================================
// 1. CONFIGURACIÓN DE GROQ
// Lee la clave API desde las variables de entorno.
// Nunca escribas la clave directamente en este archivo.
// ======================================================

const apiKey = process.env.GROQ_API_KEY;

if (!apiKey || apiKey.includes("[SENSITIVE]")) {
  throw new Error("Falta configurar GROQ_API_KEY.");
}

const groq = new Groq({ apiKey });

// ======================================================
// 2. CONFIGURACIÓN DE LA API
// Permite que formidable procese archivos multipart.
// ======================================================

export const config = {
  api: {
    bodyParser: false,
  },
};

// ======================================================
// 3. LEER LA SOLICITUD
// Obtiene la oferta laboral y el archivo enviado.
// ======================================================

async function leerSolicitud(req) {
  const tipo = req.headers["content-type"] || "";

  if (!tipo.includes("multipart/form-data")) {
    throw new Error("Formato de solicitud no compatible.");
  }

  const formulario = formidable({
    maxFileSize: 10 * 1024 * 1024,
    maxTotalFileSize: 10 * 1024 * 1024,
    maxFiles: 1,
    maxFields: 5,
  });

  const [campos, archivos] = await formulario.parse(req);

  const archivo =
    archivos.cv?.[0] ||
    archivos["cv-file"]?.[0] ||
    null;

  return {
    oferta: Array.isArray(campos.oferta)
      ? campos.oferta[0]
      : campos.oferta,
    archivo,
  };
}

// ======================================================
// 4. EXTRAER TEXTO DEL CV
// Admite archivos PDF y DOCX.
// ======================================================

async function extraerTextoCV(archivo) {
  if (!archivo) {
    throw new Error("Selecciona tu currículum en PDF o DOCX.");
  }

  const nombre = archivo.originalFilename || "";
  const extension = nombre.split(".").pop().toLowerCase();
  const contenido = await readFile(archivo.filepath);

  if (extension === "pdf") {
    // CanvasFactory le da a la librería lo que necesita en el servidor.
    const lector = new PDFParse({ data: contenido, CanvasFactory });

    try {
      const resultado = await lector.getText();
      return resultado.text;
    } finally {
      await lector.destroy();
    }
  }

  if (extension === "docx") {
    const resultado = await mammoth.extractRawText({
      buffer: contenido,
    });

    return resultado.value;
  }

  throw new Error("El archivo debe ser PDF o DOCX.");
}

// ======================================================
// 5. LIMPIAR CARACTERES INVÁLIDOS  (NUEVO)
// Los PDF y los modelos de IA a veces entregan caracteres
// de control invisibles. Word NO los acepta dentro de un
// archivo .docx y por eso mostraba el error al abrirlo.
// Esta función los elimina antes de crear el Word.
// ======================================================

function limpiarCaracteresInvalidos(texto) {
  return texto.replace(
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,
    ""
  );
}

// ======================================================
// 6. LIMPIAR TEXTO MARKDOWN
// Elimina marcas de títulos, viñetas y escapes innecesarios.
// Conserva las marcas de negrita y cursiva para dar formato.
// ======================================================

function limpiarLinea(texto) {
  return texto
    .trim()
    .replace(/^#{1,6}\s+/, "")
    .replace(/^[-*•]\s+/, "")
    .replace(/\\([*_#-])/g, "$1")
    .trim();
}

// ======================================================
// 7. CONVERTIR TEXTO EN FORMATO WORD
// Convierte **negrita** y _cursiva_ en estilos reales.
// ======================================================

function crearTextoConFormato(texto, tamano = 20) {
  const partes = texto
    .split(/(\*\*.*?\*\*|\*.*?\*|_.*?_)/g)
    .filter(Boolean);

  return partes.map((parte) => {
    const negrita = /^\*\*.*\*\*$/.test(parte);
    const cursiva =
      (/^\*.*\*$/.test(parte) && !negrita) ||
      /^_.*_$/.test(parte);

    let contenido = parte;

    if (negrita) {
      contenido = contenido.slice(2, -2);
    } else if (cursiva) {
      contenido = contenido.slice(1, -1);
    }

    return new TextRun({
      text: contenido,
      bold: negrita,
      italics: cursiva,
      font: "Arial",
      size: tamano,
    });
  });
}

// ======================================================
// 8. RECONOCER ENCABEZADOS
// Identifica secciones habituales de un currículum.
// También reconoce títulos escritos con # en Markdown.
// ======================================================

function esEncabezado(texto) {
  const titulo = texto
    .replace(/:$/, "")
    .replace(/\*\*/g, "")
    .replace(/_/g, "")
    .trim()
    .toLocaleLowerCase("es");

  const secciones = [
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

  return secciones.includes(titulo);
}

// ======================================================
// 9. CREAR PÁRRAFOS CON ESTILO HARVARD
// Centra el encabezado, destaca las secciones y genera
// viñetas reales en Word.
// ======================================================

function crearParrafosCV(texto) {
  const lineas = texto.replace(/\r\n/g, "\n").split("\n");
  const parrafos = [];

  let numeroLineaEncabezado = 0;
  let comenzoContenido = false;

  for (const linea of lineas) {
    const original = linea.trim();

    // Ignorar líneas vacías y separadores Markdown.
    if (!original || /^\\?---+$/.test(original)) {
      continue;
    }

    const esTituloMarkdown = /^#{1,6}\s+/.test(original);
    const esVineta = /^[-*•]\s+/.test(original);

    const contenido = limpiarLinea(original);

    if (!contenido) {
      continue;
    }

    const esTitulo = esTituloMarkdown || esEncabezado(contenido);

    // Nombre del candidato: primer párrafo.
    if (parrafos.length === 0) {
      parrafos.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: contenido.replace(/\*\*/g, "").replace(/_/g, ""),
              bold: true,
              font: "Arial",
              size: 30,
              color: "17365D",
            }),
          ],
          spacing: { after: 100 },
          keepNext: true,
        })
      );

      continue;
    }

    // Encabezados de sección en azul y mayúsculas.
    if (esTitulo) {
      comenzoContenido = true;

      parrafos.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: contenido
                .replace(/\*\*/g, "")
                .replace(/_/g, "")
                .toLocaleUpperCase("es"),
              bold: true,
              font: "Arial",
              size: 22,
              color: "17365D",
            }),
          ],
          spacing: { before: 220, after: 90 },
          keepNext: true,
        })
      );

      continue;
    }

    // Centrar especialidades y datos de contacto
    // hasta que comienza la primera sección del CV.
    if (!comenzoContenido && numeroLineaEncabezado < 3) {
      numeroLineaEncabezado++;

      parrafos.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: crearTextoConFormato(
            contenido,
            numeroLineaEncabezado === 1 ? 19 : 17
          ),
          spacing: { after: 40 },
          keepNext: true,
        })
      );

      continue;
    }

    // Cargos, empresas, fechas, ubicación y contenido.
    //
    // CORRECCIÓN IMPORTANTE:
    // Antes se usaba  bullet: { indent: 300 }  y eso generaba
    // un XML inválido (era la causa de que Word no abriera
    // el archivo). Ahora las viñetas usan "numbering", que es
    // la forma correcta. La definición "vinetas" se crea en
    // la sección 10, dentro de crearWord().
    parrafos.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: crearTextoConFormato(contenido, 20),
        numbering: esVineta
          ? { reference: "vinetas", level: 0 }
          : undefined,
        spacing: { after: esVineta ? 45 : 75 },
      })
    );
  }

  return parrafos;
}

// ======================================================
// 10. GENERAR EL ARCHIVO WORD
// Usa los párrafos formateados y devuelve el DOCX en Base64.
// ======================================================

async function crearWord(texto) {
  // Se eliminan los caracteres que Word no acepta.
  const textoLimpio = limpiarCaracteresInvalidos(texto);
  const parrafos = crearParrafosCV(textoLimpio);

  const documento = new Document({
    // Definición de las viñetas (el punto • de cada lista).
    numbering: {
      config: [
        {
          reference: "vinetas",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 500, hanging: 250 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {},
        children: parrafos,
      },
    ],
  });

  const buffer = await Packer.toBuffer(documento);

  return buffer.toString("base64");
}

// ======================================================
// 11. CONTROLADOR PRINCIPAL
// Valida la solicitud, adapta el CV con Groq y devuelve
// el texto adaptado junto con el archivo Word.
// ======================================================

export default async function handler(req, res) {
  // Solo se admiten solicitudes POST.
  if (req.method !== "POST") {
    return res.status(405).json({
      error: "Utiliza el método POST.",
    });
  }

  try {
    // Leer la oferta laboral y el archivo adjunto.
    const { oferta, archivo } = await leerSolicitud(req);

    if (typeof oferta !== "string" || !oferta.trim()) {
      return res.status(400).json({
        error: "Pega la oferta laboral antes de continuar.",
      });
    }

    // Extraer y validar el contenido del currículum.
    const cv = await extraerTextoCV(archivo);

    if (!cv.trim()) {
      return res.status(400).json({
        error:
          "No encontramos texto en el CV. Comprueba que el PDF no sea una imagen escaneada.",
      });
    }

    if (cv.length > 30000 || oferta.length > 15000) {
      return res.status(413).json({
        error: "El CV o la oferta son demasiado extensos.",
      });
    }

    // Instrucciones para adaptar el currículum sin inventar datos.
    const instrucciones = `
Adapta el siguiente currículum a la oferta laboral.

Reglas:
- Escribe en español.
- Utiliza únicamente información respaldada por el CV original.
- No inventes experiencia, estudios, habilidades ni resultados.
- Destaca las competencias y experiencias más relevantes para la oferta.
- Conserva todos los cargos relevantes, sus empresas, fechas y ubicaciones disponibles.
- Usa el nombre del candidato como primera línea.
- Escribe las especialidades debajo del nombre.
- Usa títulos de sección precedidos por ##.
- Utiliza **negrita** para los nombres de cargos y empresas.
- Utiliza _cursiva_ para fechas cuando corresponda.
- Usa viñetas que comiencen con - para las responsabilidades.
- No incluyas separadores como ---.
- Devuelve solamente el currículum adaptado.

CURRÍCULUM ORIGINAL:
${cv}

OFERTA LABORAL:
${oferta}
`;

    // Solicitar la adaptación al modelo de Groq.
    const respuesta = await groq.chat.completions.create({
      model: "openai/gpt-oss-120b",
      messages: [
        {
          role: "user",
          content: instrucciones,
        },
      ],
      temperature: 0.3,
      max_completion_tokens: 6000,
    });

    const resultado = respuesta.choices[0]?.message?.content;

    if (!resultado || !resultado.trim()) {
      throw new Error("Groq devolvió una respuesta vacía.");
    }

    // Generar el Word a partir del texto adaptado.
    const docxBase64 = await crearWord(resultado);

    // Enviar el resultado y el archivo a la página web.
    return res.status(200).json({
      resultado,
      docxBase64,
    });
  } catch (error) {
    // Registrar el detalle real en la terminal del servidor.
    console.error("Error al adaptar el CV con Groq:", error);

    // Devolver un mensaje comprensible al usuario.
    return res.status(500).json({
      error:
        "No pudimos procesar el CV. Comprueba la configuración de Groq e inténtalo nuevamente.",
    });
  }
}