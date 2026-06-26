/**
 * FUNCIÓN PRINCIPAL
 */
function doPost(e) {
  try {
    // ID de la carpeta donde se guardarán los archivos independientes por UES
    const DESTINO_FOLDER_ID = "1bI7b1TID85ZdwIsYFtL-EPfWrIYIQMm6";
    
    // 1. Datos del formulario
    var nombre   = e.parameter.nombre;
    var email    = e.parameter.email;
    var telefono = e.parameter.telefono;
    var plantel  = e.parameter.plantel; 
    var carrera  = e.parameter.area;
    var fecha    = new Date();

    // --- ACCIÓN 1: Guardar en el Excel Maestro (Hoja General) ---
    var ssPrincipal = SpreadsheetApp.getActiveSpreadsheet();
    var sheetGeneral = ssPrincipal.getSheetByName("Registros_General") || ssPrincipal.getSheets()[0];
    sheetGeneral.appendRow([fecha, nombre, email, telefono, plantel, carrera]);

    // --- ACCIÓN 2: Guardar en Archivo Independiente en la carpeta de Drive ---
    guardarEnArchivoIndividual(DESTINO_FOLDER_ID, plantel, [fecha, nombre, email, telefono, plantel, carrera]);

    // --- ACCIÓN 3: Envío de correo electrónico con diseño mejorado ---
    enviarCorreoConAdjunto(nombre, email, carrera);
    
    return ContentService.createTextOutput("Success")
                         .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString())
                         .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Busca o crea un archivo de Excel para la UES en la carpeta de Drive
 */
function guardarEnArchivoIndividual(folderId, nombreArchivo, datos) {
  var carpeta = DriveApp.getFolderById(folderId);
  var archivos = carpeta.getFilesByName(nombreArchivo);
  var archivoExcel;

  if (archivos.hasNext()) {
    var file = archivos.next();
    archivoExcel = SpreadsheetApp.openById(file.getId());
  } else {
    // Si no existe, crea uno nuevo
    archivoExcel = SpreadsheetApp.create(nombreArchivo);
    var file = DriveApp.getFileById(archivoExcel.getId());
    carpeta.addFile(file);
    DriveApp.getRootFolder().removeFile(file); 
    
    var sheet = archivoExcel.getSheets()[0];
    sheet.appendRow(["FECHA", "NOMBRE", "EMAIL", "TELÉFONO", "PLANTEL", "CARRERA"]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#691c32").setFontColor("white");
  }

  archivoExcel.getSheets()[0].appendRow(datos);
}

/**
 * FUNCIÓN DE ENVÍO (Diseño mejorado solicitado)
 */
/**
 * FUNCIÓN DE ENVÍO (Diseño mejorado con Footer Actualizado)
 */
function enviarCorreoConAdjunto(nombre, destinatario, carrera) {
  const folderId = "1AryieRud3yPuoMK8yxrUoDNMpmirYTh0";
  const carpeta = DriveApp.getFolderById(folderId);
  const logoUrl = "https://raw.githubusercontent.com/ismaelumb/MAPA-UMB/refs/heads/main/Logotipo%20UMB%20-%20copia.png";
  
  const archivos = carpeta.getFilesByName(carrera + ".png"); 
  
  const cuerpoHtml = `
    <div style="background-color: #f4f4f4; padding: 20px; font-family: 'Segoe UI', Arial, sans-serif; color: #333;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
        
        <div style="background-color: #691c32; height: 10px;"></div>

        <div style="padding: 30px 20px; text-align: center;">
          <img src="${logoUrl}" alt="UMB Logo" style="height: 70px; width: auto;">
        </div>

        <div style="padding: 0 40px 40px 40px;">
          <h2 style="color: #691c32; font-size: 24px; margin-bottom: 15px; text-align: center; font-weight: 700;">
            ¡Hola, ${nombre}!
          </h2>
          
          <div style="color: #444; font-size: 15px; line-height: 1.7; text-align: center; margin-bottom: 25px;">
            <p>Es un gusto saludarte y darte la bienvenida a este primer acercamiento con la <strong>Universidad Mexiquense del Bicentenario</strong>.</p>
            <p style="font-style: italic; color: #691c32;">"Independencia y Libertad de Pensamiento"</p>
            <p>Sabemos que elegir tu carrera es una de las decisiones más importantes de tu vida. En la UMB, nos comprometemos a brindarte una educación pública de excelencia, con programas diseñados para que te conviertas en el profesional que el Estado de México y el país necesitan.</p>
          </div>
          <div style="margin: 20px 0; background-color: #fdf8f9; border: 1px solid #f2e1e5; border-radius: 10px; padding: 20px; text-align: center;">
            <span style="color: #bc955c; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">Has mostrado interés en:</span>
            <div style="color: #691c32; font-size: 19px; font-weight: bold; margin-top: 8px;">${carrera}</div>
          </div>

          <p style="font-size: 14px; line-height: 1.6; text-align: justify; color: #666; margin-bottom: 30px;">
            Como parte de tu proceso, te adjuntamos la ficha técnica de este programa académico. En ella encontrarás los objetivos de la carrera, las competencias que desarrollarás y el campo laboral donde podrás destacar.
          </p>

          <div style="text-align: center;">
            <a href="https://umb.edomex.gob.mx/" style="background-color: #bc955c; color: #ffffff; padding: 14px 30px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 13px; display: inline-block;">
              VISITAR PORTAL DE ADMISIONES
            </a>
          </div>
        </div>

        <div style="background-color: #2c2c2c; color: #ffffff; padding: 30px 20px; text-align: center;">
          <p style="margin: 0 0 10px 0; font-size: 13px; font-weight: bold; color: #bc955c; letter-spacing: 1px;">UNIVERSIDAD MEXIQUENSE DEL BICENTENARIO</p>
          
          <p style="margin: 5px 0; font-size: 11px; color: #cccccc; line-height: 1.4;">
            Rectoría | Carretera México - Toluca km 43.5, Bo San Miguel,<br>
            Ocoyoacac, Estado de México, C.P. 52740
          </p>
          
          <p style="margin: 5px 0; font-size: 11px; color: #cccccc;">
            Teléfonos: 55 98 13 66 92 y 55 98 13 66 93
          </p>
          
          <p style="margin: 5px 0; font-size: 11px;">
            <a href="mailto:convocatoria@umb.mx" style="color: #bc955c; text-decoration: none;">convocatoria@umb.mx</a>
          </p>

          <div style="font-size: 10px; opacity: 0.4; border-top: 1px solid rgba(255,255,255,0.1); padding-top: 15px; margin-top: 15px;">
            © 2026 UMB - Todos los derechos reservados.
          </div>
        </div>
      </div>
    </div>
  `;

  const opcionesEmail = {
    to: destinatario,
    subject: "Información Importante: " + carrera + " - UMB",
    htmlBody: cuerpoHtml
  };

  if (archivos.hasNext()) {
    opcionesEmail.attachments = [archivos.next().getBlob()];
  }

  MailApp.sendEmail(opcionesEmail);
}
