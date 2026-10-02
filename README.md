Manual de Usuario y Guía de Operaciones — SPD Farmacia Villegas
Este manual recopila todos los procesos, flujos de trabajo y soluciones a dudas habituales para el manejo diario de la aplicación de Sistemas Personalizados de Dosificación (SPD).
1. Acceso al Sistema y Roles
La aplicación cuenta con dos perfiles de usuario diferenciados con permisos específicos:
Perfil Administrador (Farmacéutico): Control total del sistema. Permite validar altas y bajas, gestionar el catálogo de pacientes, procesar pedidos definitivos con escaneo DataMatrix, desbloquear medicamentos y administrar usuarios/permisos.
Perfil Enfermería: Diseñado for la operativa externa o en planta. Permite proponer altas de nuevos pacientes, proponer bajas, seleccionar medicamentos para pedidos semanales y reportar o resolver incidencias.
2. Gestión de Pacientes y Altas
Alta de un Paciente (Enfermería)
Accede a la sección ➕ ALTAS desde la barra superior.
Rellena los campos obligatorios: Nombre Completo y Código CIP. El campo de Referencia (Ref) lo asignará posteriormente la farmacia.
Añade los medicamentos del tratamiento en la tabla inferior.
Pulsa 📤 Enviar Propuesta de Alta. El formulario se borrará y vaciará automáticamente al pulsar el botón.
Podrás ver el estado de tu propuesta ("Pendiente", "Validada" o "Rechazada") en el listado inferior.
Alta Directa o Carga Masiva (Farmacia)
Alta Directa: Desde la pestaña de Administrador, introduce los datos del paciente, su referencia, CIP y su tabla de medicamentos. Al pulsar el botón de alta directa, el formulario se reiniciará por completo.
Carga Masiva: Descarga la plantilla de Excel oficial desde la aplicación, rellena las filas y súbela para dar de alta a todos de golpe.
3. Gestión de Tratamientos y Pedidos Semanales
Selección de Pedidos (Enfermería)
Accede a 📦 PROPUESTA.
Marca las casillas de selección de los medicamentos que necesites pedir y ajusta el número de cajas/elementos.
Pulsa 🚀 Solicitar Pedido Definitivo.
Validación y Escaneo DataMatrix (Farmacia)
Accede a 🛒 PEDIDOS.
Haz clic en la columna DataMatrix y escanea con el lector el código bidimensional de la caja.
El sistema descodificará automáticamente el Lote y la Caducidad.
Imprime el Albarán de Entrega (PDF) para firmarlo junto a enfermería.
Pulsa 📌 Actualizar y Limpiar para registrar la fecha de última entrega y vaciar la lista.
4. Gestión de Bajas y Devoluciones
Accede a ➖ BAJAS y selecciona al paciente.
SIN Devolución: El paciente se da de baja directamente.
CON Devolución: Se abre el proceso para escanear los medicamentos sobrantes devueltos. El sistema generará un Albarán de Devolución en PDF.
5. Control de Incidencias
Creación: Se marca la incidencia desde la ficha del paciente (falta de receta, posología dudosa, etc.).
Resolución: Las incidencias aparecen en el panel de incidencias. Enfermería puede marcarlas como "Resueltas" y el Administrador podrá validarlas definitivamente.
6. Resolución de Dudas y Preguntas Frecuentes (FAQ)
🔴 El archivo Excel no se actualiza o da error al guardar
Causa: Tienes abierto el archivo Tratamientos_Por_Paciente.xlsx en Microsoft Excel.
Solución: Cierra Excel por completo para liberar el bloqueo de Windows.
🔴 La enfermera no puede entrar desde su móvil fuera de la farmacia
Causa: El túnel web de Cloudflare se ha cerrado o cambiado.
Solución: Asegúrate de tener activa la consola del túnel en la unidad Z: con el comando correspondiente y usa la dirección web configurada.
