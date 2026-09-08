"""Condiciones de uso que el usuario debe aceptar antes del primer inicio."""

VERSION_CONDICIONES = "2026-09-05"

TEXTO_CONDICIONES = """CONDICIONES DE USO Y AVISO DE SEGURIDAD DE SiPP

Version de las condiciones: 2026-09-05

1. OBJETO Y ALCANCE
SiPP es un sistema de apoyo para la gestion de personal, asistencia, planilla,
permisos, vacaciones, contratos, liquidaciones, reportes y documentos laborales.
El programa organiza informacion proporcionada por la empresa y no sustituye
la revision profesional, contable, laboral o legal que corresponda.

Estas condiciones regulan el uso del programa por parte de la empresa, sus
administradores y las personas autorizadas por ella.

2. RESPONSABILIDAD DE LA EMPRESA
La empresa usuaria es responsable de:
- Proporcionar informacion exacta, completa y actualizada.
- Verificar los calculos, reportes y documentos antes de utilizarlos.
- Cumplir la legislacion laboral, fiscal, de privacidad y de proteccion de datos
  aplicable a su actividad y jurisdiccion.
- Mantener actualizados los datos de empleados, salarios, horarios, permisos y
  demas registros necesarios para sus procesos.
- Designar administradores y usuarios confiables.

SiPP facilita la organizacion de informacion, pero la empresa conserva la
responsabilidad final sobre las decisiones, pagos, declaraciones, contratos,
notificaciones y documentos que produzca o utilice.

3. CUENTAS, CREDENCIALES Y PERMISOS
Cada persona debe utilizar su propia cuenta. Las credenciales son personales y
no deben compartirse. La empresa debe asignar a cada usuario solamente los
permisos que necesite para realizar su trabajo.

Los administradores pueden gestionar configuraciones, usuarios y operaciones
sensibles. La empresa debe revisar periodicamente los administradores activos,
retirar accesos innecesarios y comunicar cualquier uso no autorizado.

4. INFORMACION Y DATOS DE LA EMPRESA
La empresa es responsable de la legitimidad de los datos que registre, importe,
consulte, modifique o elimine mediante SiPP. Debe contar con las autorizaciones
necesarias para tratar datos personales y documentos de sus empleados.

La empresa debe evitar registrar contrasenas, claves secretas o informacion que
no sea necesaria para sus procesos laborales. Los archivos adjuntos deben ser
legales, pertinentes y almacenados de acuerdo con sus politicas internas.

5. SERVIDOR Y CONECTIVIDAD
SiPP puede conectarse al servidor PostgreSQL configurado por la empresa. La
empresa es responsable de proteger ese servidor, sus usuarios, sus contrasenas,
su firewall, sus copias de seguridad y los permisos de red.

No se debe exponer PostgreSQL directamente a internet sin controles adecuados.
La empresa debe utilizar conexiones cifradas cuando el servidor sea remoto y
proporcionar a SiPP un usuario de base de datos con permisos limitados.

6. AUTORIZACION DE EQUIPOS
Una computadora nueva puede solicitar autorizacion antes de utilizar SiPP.
El codigo de autorizacion es temporal, de un solo uso y puede caducar. La
empresa puede rechazar una solicitud o retirar el acceso a un equipo.

La autorizacion de un equipo no reemplaza la autenticacion de usuarios ni
concede permisos adicionales dentro del sistema.

7. COPIAS DE SEGURIDAD
La empresa debe realizar copias de seguridad periodicas y comprobar que pueden
restaurarse. Una copia no verificada no debe considerarse suficiente para
recuperar la informacion ante una falla, perdida, ataque o error humano.

Las copias deben protegerse contra acceso no autorizado y conservarse en una
ubicacion distinta cuando sea posible.

8. USOS PROHIBIDOS
No esta permitido:
- Acceder a informacion de otra empresa o persona sin autorizacion.
- Compartir cuentas o utilizar credenciales ajenas.
- Intentar evadir controles de acceso o autorizacion.
- Alterar, extraer o destruir informacion sin permiso.
- Introducir malware o usar SiPP para actividades ilegales.
- Modificar el programa para ocultar acciones o evadir auditorias.
- Utilizar informacion personal para fines distintos de los autorizados.

9. REGISTROS Y SEGURIDAD
SiPP puede registrar inicios de sesion, intentos fallidos, bloqueos, cambios y
acciones administrativas para proteger el sistema y facilitar investigaciones.
La empresa debe limitar el acceso a esos registros y conservarlos conforme a
sus obligaciones legales y politicas internas.

10. DISPONIBILIDAD Y LIMITACIONES
El funcionamiento puede depender de la disponibilidad de Windows, PostgreSQL,
la red, internet, servicios de correo, respaldos y otros componentes externos.
Una interrupcion de esos servicios puede impedir temporalmente el acceso o el
envio de notificaciones.

Antes de aplicar cambios importantes, la empresa debe contar con un respaldo
reciente y revisar el resultado de las operaciones.

11. ACEPTACION
Al marcar "Acepto las condiciones" y continuar, el usuario confirma que ha
leido este documento, que actua con autorizacion de la empresa y que comprende
sus responsabilidades. La aplicacion registra la version aceptada, la fecha,
el usuario de Windows y la identificacion del equipo para dejar constancia de
la aceptacion.

Si el usuario no acepta estas condiciones, debe cerrar SiPP y no podra utilizar
la aplicacion.
"""
