function enviarCorreoNotificacion(e) {
  try {
    Logger.log("Iniciando ejecución del script...");
    // Validar si hay datos en la respuesta
    if (!e || !e.namedValues) {
      Logger.log("No se encontraron datos en la respuesta del formulario.");
      return;
    }
    
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ultimaFila = hoja.getLastRow();
    Logger.log("Última fila con datos: " + ultimaFila);
    // Obtener los datos del formulario
    var rangoDatos = hoja.getRange(ultimaFila, 2, 1, hoja.getLastColumn() - 1);
    var datos = rangoDatos.getValues()[0];
    var encabezados = hoja.getRange(1, 2, 1, hoja.getLastColumn() - 1).getValues()[0];
    
    Logger.log("Encabezados completos: " + JSON.stringify(encabezados));
    Logger.log("Datos completos: " + JSON.stringify(datos));
    
    // Función para formatear fechas
    function formatearFecha(fecha) {
      if (fecha instanceof Date) {
        var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
        return fecha.getDate() + " de " + meses[fecha.getMonth()] + " de " + fecha.getFullYear();
      }
      return fecha; // Si no es una fecha, devolver el valor tal cual
    }
    
    // Construir el cuerpo del correo
    var cuerpoCorreo = "<p>Buen día,</p>";
    cuerpoCorreo += "<p><strong>Se ha recibido un nuevo reporte.</strong></p>";
    cuerpoCorreo += "<ul>";
    for (var i = 0; i < encabezados.length; i++) {
      if (datos[i] !== "") { // Excluir campos vacíos
        var valor = datos[i];
        
        // Verificar si es una fecha para formatearla
        if (encabezados[i] === "Fecha de solicitud" || encabezados[i] === "Fecha de ocurrencia del hecho vital (Nacimiento o Defunción)") {
          valor = formatearFecha(valor);
        }

        cuerpoCorreo += "<li><strong>" + encabezados[i] + ":</strong> " + valor + "</li>";
      }
    }
    cuerpoCorreo += "</ul>";
    
    cuerpoCorreo += "<p><strong>Foscal</strong><br>Inspirados por la vida</p>";
    
    Logger.log("Cuerpo del correo generado: " + cuerpoCorreo);
    
    // Enviar el correo a los encargados
    var destinatarios = "datamanager.investigaciones@foscal.com.co,foscalestudiosclinicos@gmail.com";
    var asunto = "Nuevo reporte: formulario Solicitud pérdida Nacido Vivo o Defunción";
    
    MailApp.sendEmail({
      to: destinatarios,
      subject: asunto,
      htmlBody: cuerpoCorreo
    });
    
    Logger.log("Correo enviado exitosamente a: " + destinatarios);
  } catch (error) {
    Logger.log("Error en el script: " + error.toString());
  }
}
