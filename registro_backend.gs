/**
 * Pega este código en tu proyecto de Apps Script (el mismo que ya tienes
 * desplegado como Web App). Si ya tienes un doPost(e), no crees uno nuevo:
 * copia el "if (action === 'registrar_mesero')" de adentro y pégalo dentro
 * de tu doPost existente, adaptando el enrutamiento.
 */

// ID de la carpeta de Drive donde se guardarán las fotos de perfil.
// Créala una vez en Drive, ábrela, y copia el ID de la URL
// (drive.google.com/drive/folders/ESTE_ID)
const CARPETA_FOTOS_ID = "1RBhcYyPQW8VciYR8Ic-nA1AkZMOMVhaj";

// Nombre de la hoja donde vive la tabla de trabajadores
const HOJA_TRABAJADORES = "Trabajadores";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === "registrar_mesero") {
      return registrarMesero(body);
    }

    return respuesta({ ok: false, error: "Acción no reconocida" });
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  }
}

function registrarMesero(body) {
  const { nombre, apellido, telefono, edad, dni, foto_base64 } = body;

  if (!nombre || !apellido || !telefono || !edad || !dni || !foto_base64) {
    return respuesta({ ok: false, error: "Faltan campos obligatorios" });
  }

  // Genera un código único para el mesero, ej: SC042
  const codigo = generarCodigo(nombre, apellido);

  // Guarda la foto en Drive y obtiene su URL pública
  const fotoUrl = guardarFotoEnDrive(foto_base64, codigo);

  // Agrega la fila a la hoja de trabajadores
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
  hoja.appendRow([
    new Date(),
    codigo,
    nombre,
    apellido,
    telefono,
    edad,
    dni,
    fotoUrl
  ]);

  return respuesta({ ok: true, codigo: codigo, foto_url: fotoUrl });
}

function guardarFotoEnDrive(base64Data, codigo) {
  // base64Data viene como "data:image/jpeg;base64,AAAA..."
  const partes = base64Data.split(",");
  const meta = partes[0]; // "data:image/jpeg;base64"
  const datos = partes[1];

  const tipoMime = meta.match(/data:(.*);base64/)[1];
  const bytes = Utilities.base64Decode(datos);
  const blob = Utilities.newBlob(bytes, tipoMime, codigo + ".jpg");

  const carpeta = DriveApp.getFolderById(CARPETA_FOTOS_ID);
  const archivo = carpeta.createFile(blob);

  // Hace el archivo visible por link (no listado públicamente)
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // URL directa para usar en <img src="...">
  return "https://drive.google.com/uc?export=view&id=" + archivo.getId();
}

function generarCodigo(nombre, apellido) {
  const iniciales = (nombre[0] + apellido[0]).toUpperCase();
  const numero = Math.floor(100 + Math.random() * 900); // 3 dígitos
  return iniciales + numero;
}

function respuesta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
