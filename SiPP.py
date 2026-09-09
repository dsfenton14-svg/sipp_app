import customtkinter as ctk
from tkinter import messagebox, Menu, filedialog, ttk
from PIL import Image
from io import BytesIO
import os
import sys
import shutil
import subprocess
import tkinter as tk
import ctypes
import calendar
from datetime import datetime, timedelta, timezone
import logging
import json
import threading
import queue
import tempfile
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen, urlretrieve
from xml.sax.saxutils import escape
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image as PdfImage, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from horarios_laborales import LIMITES_DIARIOS, dia_cuenta_como_pagado, feriados_panama
from condiciones_uso import TEXTO_CONDICIONES, VERSION_CONDICIONES
from autorizacion_codigo import (
    ClienteAutorizacion,
    ErrorAutorizacion,
    cargar_autorizacion_local,
    guardar_autorizacion_local,
    cargar_aceptacion_condiciones,
    guardar_aceptacion_condiciones,
    obtener_datos_equipo,
)

datos_empresa = {"nombre": "", "ruc": "", "direccion": "", "telefono": "", "logo": ""}
VERSION_APLICACION = "2.3.6"
URL_MANIFIESTO_ACTUALIZACION = ""
URL_RELEASES_GITHUB = "https://api.github.com/repos/dsfenton14-svg/sipp_app/releases/latest"
COLA_UI = queue.Queue()


def programar_en_ui(funcion):
    COLA_UI.put(funcion)


def obtener_horas_ordinarias_configuradas(horarios):
    """Devuelve el limite diario configurado, conservando el valor anterior si no hay horario."""
    if not horarios:
        return 8.0
    horario = horarios[0]
    tipo_jornada = "DIURNA"
    if len(horario) >= 4:
        tipo_jornada = str(horario[2] or "DIURNA").strip().upper()
    return LIMITES_DIARIOS.get(tipo_jornada, 8.0)


def ajustar_toplevel_a_pantalla(top_level, ancho, alto, margen_x=100, margen_y=100, min_ancho=320, min_alto=220):
    try:
        ancho_pantalla = top_level.winfo_screenwidth()
        alto_pantalla = top_level.winfo_screenheight()
    except Exception:
        ancho_pantalla = 1920
        alto_pantalla = 1080

    ancho_ajustado = max(min_ancho, min(int(ancho), max(min_ancho, ancho_pantalla - margen_x)))
    alto_ajustado = max(min_alto, min(int(alto), max(min_alto, alto_pantalla - margen_y)))
    x = max(0, (ancho_pantalla - ancho_ajustado) // 2)
    y = max(0, (alto_pantalla - alto_ajustado) // 2)
    top_level.geometry(f"{ancho_ajustado}x{alto_ajustado}+{x}+{y}")
    return ancho_ajustado, alto_ajustado


def construir_ruta_guardado_segura(ruta_sugerida, nombre_default="documento.xlsx"):
    """Devuelve una ruta segura y escribible; si la ruta elegida no es accesible, usa una carpeta local de exportación."""
    if not ruta_sugerida:
        return os.path.join(os.getcwd(), nombre_default)

    ruta_absoluta = os.path.abspath(ruta_sugerida)
    carpeta = os.path.dirname(ruta_absoluta) or os.getcwd()
    nombre_archivo = os.path.basename(ruta_absoluta) or nombre_default

    try:
        if os.path.isdir(carpeta) and os.access(carpeta, os.W_OK):
            return ruta_absoluta
    except OSError:
        pass

    try:
        os.makedirs(carpeta, exist_ok=True)
        if os.access(carpeta, os.W_OK):
            return ruta_absoluta
    except OSError:
        pass

    carpeta_fallback = os.path.join(os.getcwd(), "exportaciones")
    os.makedirs(carpeta_fallback, exist_ok=True)
    return os.path.join(carpeta_fallback, nombre_archivo)


def ejecutar_en_segundo_plano(funcion, al_terminar, al_fallar=None):
    """Ejecuta trabajo bloqueante fuera del hilo de Tkinter."""
    def trabajador():
        try:
            resultado = funcion()
        except Exception as exc:
            if al_fallar:
                al_fallar(exc)
            return
        al_terminar(resultado)

    threading.Thread(target=trabajador, daemon=True).start()

# DateEntry personalizado (evita problemas de compatibilidad con tkcalendar)
class DateEntry(ctk.CTkEntry):
    def __init__(self, parent, *args, **kwargs):
        # Extraer parámetros personalizados
        self.width = kwargs.pop('width', 20)
        self.date_pattern = kwargs.pop('date_pattern', "yyyy-MM-dd")
        self.year = kwargs.pop('year', datetime.now().year)
        self.month = kwargs.pop('month', datetime.now().month)
        self.day = kwargs.pop('day', datetime.now().day)
        # Ignorar parámetros de tkcalendar que no se usan
        kwargs.pop('background', None)
        kwargs.pop('foreground', None)
        kwargs.pop('borderwidth', None)
        kwargs.pop('font', None)
        
        super().__init__(parent, *args, **kwargs)
        self._set_initial_date()
    
    def _set_initial_date(self):
        """Establece la fecha inicial en el entry"""
        fecha = datetime(self.year, self.month, self.day)
        self.set_date(fecha)
    
    def set_date(self, value):
        """Establece la fecha en el entry"""
        self.delete(0, tk.END)
        if isinstance(value, datetime):
            self.insert(0, value.strftime("%d/%m/%Y"))
        else:
            self.insert(0, str(value))
    
    def get_date(self):
        """Obtiene la fecha como objeto datetime"""
        fecha_str = self.get().strip()
        if not fecha_str:
            return datetime.now().date()
        
        # Intentar parse con diferentes formatos
        formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
        for fmt in formatos:
            try:
                return datetime.strptime(fecha_str, fmt).date()
            except ValueError:
                continue
        return datetime.now().date()

    def open_calendar(self):
        """Abre un calendario sencillo para seleccionar una fecha."""
        fecha_actual = self.get_date()
        ventana = ctk.CTkToplevel(self)
        ventana.title("Seleccionar fecha")
        ajustar_toplevel_a_pantalla(ventana, 280, 300, margen_x=120, margen_y=120, min_ancho=260, min_alto=260)
        ventana.resizable(False, False)
        ventana.transient(self.winfo_toplevel())
        ventana.grab_set()

        estado = {"year": fecha_actual.year, "month": fecha_actual.month}
        encabezado = ctk.CTkFrame(ventana, fg_color="transparent")
        encabezado.pack(fill="x", padx=10, pady=(10, 4))
        titulo = ctk.CTkLabel(encabezado, font=ctk.CTkFont(size=14, weight="bold"))
        titulo.pack(side="left", expand=True)
        ctk.CTkButton(encabezado, text="<", width=28, command=lambda: cambiar_mes(-1)).pack(side="left", padx=2)
        ctk.CTkButton(encabezado, text=">", width=28, command=lambda: cambiar_mes(1)).pack(side="left", padx=2)

        dias_frame = ctk.CTkFrame(ventana, fg_color="transparent")
        dias_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        def seleccionar_dia(dia):
            self.set_date(datetime(estado["year"], estado["month"], dia))
            self.event_generate("<<DateEntrySelected>>")
            ventana.destroy()

        def mostrar_mes():
            for widget in dias_frame.winfo_children():
                widget.destroy()
            titulo.configure(text=f"{calendar.month_name[estado['month']]} {estado['year']}")
            for columna, nombre_dia in enumerate(("Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do")):
                ctk.CTkLabel(dias_frame, text=nombre_dia, width=32).grid(row=0, column=columna, pady=3)
            for fila, semana in enumerate(calendar.monthcalendar(estado["year"], estado["month"]), start=1):
                for columna, dia in enumerate(semana):
                    if dia:
                        ctk.CTkButton(
                            dias_frame, text=str(dia), width=32, height=28,
                            command=lambda dia=dia: seleccionar_dia(dia),
                        ).grid(row=fila, column=columna, padx=1, pady=1)

        def cambiar_mes(direccion):
            nuevo_mes = estado["month"] + direccion
            if nuevo_mes < 1:
                estado["month"], estado["year"] = 12, estado["year"] - 1
            elif nuevo_mes > 12:
                estado["month"], estado["year"] = 1, estado["year"] + 1
            else:
                estado["month"] = nuevo_mes
            mostrar_mes()

        mostrar_mes()

try:
    from customtkinterthemes import theme_manager
except ModuleNotFoundError:
    class _ThemeManagerFallback:
        @staticmethod
        def get_font_path():
            return None
        @staticmethod
        def get(name):
            return "blue"
    theme_manager = _ThemeManagerFallback()

try:
    import descarga
except ModuleNotFoundError:
    descarga = None

try:
    import datos_Sipp as db
except ModuleNotFoundError:
    db = None

try:
    import respaldo_sipp
except ModuleNotFoundError:
    respaldo_sipp = None

try:
    import validaciones_legales as val_legal
except ModuleNotFoundError:
    val_legal = None

try:
    import ediciones as edict
except ModuleNotFoundError:
    edict = None

try:
    from lector_biometrico import ErrorConector, crear_conector
except ModuleNotFoundError:
    ErrorConector = Exception
    crear_conector = None

# Keep originals to avoid double-wrapping
_original_tk_after = getattr(tk.Misc, 'after', None)
_original_tk_after_cancel = getattr(tk.Misc, 'after_cancel', None)


def _wrap_tk_after():
    """Wrap tk.Misc.after and after_cancel to track scheduled callbacks per toplevel."""
    global _original_tk_after, _original_tk_after_cancel
    if not _original_tk_after:
        return

    def after(self_widget, ms, func=None, *args):
        # Call original after
        try:
            if func is None:
                return _original_tk_after(self_widget, ms)
            after_id = _original_tk_after(self_widget, ms, func, *args)
            try:
                toplevel = self_widget.winfo_toplevel()
                if not hasattr(toplevel, '_after_ids'):
                    toplevel._after_ids = set()
                toplevel._after_ids.add(after_id)
            except Exception:
                pass
            return after_id
        except Exception:
            return _original_tk_after(self_widget, ms, func, *args)

    def after_cancel(self_widget, id_):
        try:
            _original_tk_after_cancel(self_widget, id_)
        finally:
            try:
                toplevel = self_widget.winfo_toplevel()
                if hasattr(toplevel, '_after_ids') and id_ in toplevel._after_ids:
                    toplevel._after_ids.discard(id_)
            except Exception:
                pass

    tk.Misc.after = after
    tk.Misc.after_cancel = after_cancel


# Initialize wrapping once
_wrap_tk_after()



def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def boton_buscar_profesional(master, command=None, width=96, text="Buscar", fg_color=None, hover_color=None, height=32):
    opciones = {
        "text": text,
        "width": width,
        "height": height,
        "corner_radius": 10,
        "font": ctk.CTkFont(size=12, weight="bold"),
    }
    if command is not None:
        opciones["command"] = command
    if fg_color is not None:
        opciones["fg_color"] = fg_color
    if hover_color is not None:
        opciones["hover_color"] = hover_color
    return ctk.CTkButton(master, **opciones)


TRADUCCIONES = {
    "English": {
        # Menú y navegación
        "Archivo": "File",
        "Usuario": "User",
        "Crear Usuario": "Create User",
        "Crear Administrador": "Create Administrator",
        "Salir": "Exit",
        "editar": "edit",
        "Experiencia": "Experience",
        "Temas": "Themes",
        "Configuracion": "Settings",
        "Datos de la empresa": "Company Data",
        "Idioma": "Language",
        "API Lector": "Reader API",
        "Notificaciones": "Notifications",
        "Tamaño de texto y opciones visuales": "Text Size and Visual Options",
        "Horarios": "Schedules",
        "Seguridad": "Security",
        "Revisión": "Review",
        "Respaldo de seguridad": "Backup",
        "Reportes": "Reports",
        "Tardanzas y Ausencias": "Lateness and Absences",
        "Ayuda": "Help",
        "Información": "Information",
        "Buscar actualizaciones": "Check for Updates",
        "SiPP": "SiPP",
        "Sistema de Planilla Profesional": "Professional Payroll System",
        "Gestión Personal": "Personnel Management",
        "Gestión Planilla": "Payroll Management",
        "Gestión de Tiempo": "Time Management",
        "Descuentos y Préstamos": "Discounts and Loans",
        "Deducciones y Retenciones": "Deductions and Withholdings",
        "Remuneraciones": "Remunerations",
        "Ausencias y Permisos": "Absences and Leave",
        "Vacaciones": "Vacations",
        "Contratos de Trabajo": "Employment Contracts",
        "Liquidaciones": "Settlements",
        "Cartas": "Letters",
        "Cartas de trabajo": "Employment Letters",
        "Manuales": "Manuals",
        "Panel Principal": "Main Panel",
        "Administra personal, planilla, marcaciones y procesos clave desde un solo lugar.": "Manage personnel, payroll, time tracking and key processes in one place.",
        "Personal y expedientes centralizados": "Centralized personnel and records",
        "Cortes, marcaciones y aprobaciones de horas extras": "Cutoffs, time tracking and overtime approvals",
        "Accesos rápidos para planilla y contratos": "Quick access to payroll and contracts",
        "Ir a Gestión de Tiempo": "Go to Time Management",
        "Ver Planilla": "View Payroll",
        # Acciones genéricas usadas en toda la aplicación
        "Guardar": "Save",
        "Guardar cambios": "Save Changes",
        "Cancelar": "Cancel",
        "Cerrar": "Close",
        "Buscar": "Search",
        "Eliminar": "Delete",
        "Editar": "Edit",
        "Agregar": "Add",
        "Añadir": "Add",
        "Actualizar": "Update",
        "Nuevo": "New",
        "Nueva": "New",
        "Aceptar": "Accept",
        "Continuar": "Continue",
        "Aplicar": "Apply",
        "Restaurar": "Restore",
        "Descargar": "Download",
        "Exportar": "Export",
        "Imprimir": "Print",
        "Adjuntar": "Attach",
        "Seleccionar": "Select",
        "Examinar": "Browse",
        "Limpiar": "Clear",
        "Anular": "Cancel Entry",
        "Confirmar": "Confirm",
        "Ver detalles": "View Details",
        "Acceder": "Open",
        "Volver": "Back",
        "Siguiente": "Next",
        "Anterior": "Previous",
        # Campos comunes de formularios
        "Nombre": "Name",
        "Apellido": "Last Name",
        "Correo": "Email",
        "Dirección": "Address",
        "Teléfono": "Phone",
        "Puesto": "Position",
        "Departamento": "Department",
        "Salario": "Salary",
        "Fecha de ingreso": "Hire Date",
        "Fecha de nacimiento": "Date of Birth",
        "DUI": "National ID",
        "Cargo": "Job Title",
        "Estado": "Status",
        "Tipo": "Type",
        "Monto": "Amount",
        "Motivo": "Reason",
        "Descripción": "Description",
        "Fecha": "Date",
        "Hora": "Time",
        "Entrada": "Check-in",
        "Salida": "Check-out",
        "Jornada": "Work Shift",
        "Documento": "Document",
        "Todos": "All",
        # Estados y valores comunes
        "Activo": "Active",
        "Inactivo": "Inactive",
        "Finalizado": "Terminated",
        "Suspendido": "Suspended",
        "Pendiente": "Pending",
        "Aprobado": "Approved",
        "Rechazado": "Rejected",
        "Masculino": "Male",
        "Femenino": "Female",
        "Indefinido": "Indefinite",
        "Definido": "Fixed-term",
        # Mensajes / títulos de diálogos comunes
        "Error": "Error",
        "Advertencia": "Warning",
        "Aviso": "Notice",
        "Éxito": "Success",
        "Validación": "Validation",
        "Confirmación": "Confirmation",
        # Textos específicos de cada módulo (login, personal, planilla, tiempo, descuentos,
        # deducciones, remuneraciones, permisos, vacaciones, contratos, liquidaciones, cartas,
        # configuración, seguridad, reportes)
        "*Aviso legal: las deducciones de nómina se aplican conforme a la normativa vigente y a la política interna de retenciones, incluyendo CSS, ISSS e ISR.": "*Legal notice: payroll deductions are applied according to current regulations and the internal withholding policy, including CSS, ISSS and income tax.",
        "0 registros": "0 records",
        "Abrir": "Open",
        "Acceso al sistema": "System Access",
        "Activos": "Active",
        "Actualizaciones": "Updates",
        "Actualizaciones de SiPP": "SiPP Updates",
        "Actualizar reporte": "Refresh Report",
        "Además de políticas de deducción empresarial, La empresa garantiza transparencia en los cálculos según normativa laboral vigente.": "In addition to company deduction policies, the company guarantees transparency in calculations according to current labor regulations.",
        "Adjuntar archivo": "Attach File",
        "Administra empleados, departamentos y datos laborales desde una vista centralizada.": "Manage employees, departments and employment data from a centralized view.",
        "Administra la vigencia, jornada y documentación contractual del personal.": "Manage the validity, work shift and contractual documentation of staff.",
        "Administre descuentos, préstamos y deducciones del personal": "Manage staff discounts, loans and deductions",
        "Administre las nóminas de su empresa de forma rápida y segura": "Manage your company's payroll quickly and securely",
        "Agregar horario": "Add Schedule",
        "Ajusta la escala de la interfaz para mejorar la lectura.": "Adjust the interface scale to improve readability.",
        "Alto contraste": "High Contrast",
        "Anulados": "Cancelled",
        "Análisis de Planilla": "Payroll Analysis",
        "Análisis de planilla": "Payroll analysis",
        "Apariencia:": "Appearance:",
        "Aprobacion Exitosa": "Approval Successful",
        "Aprobadas": "Approved",
        "Aprobar": "Approve",
        "Aprueba o rechaza solicitudes de vacaciones pendientes.": "Approve or reject pending vacation requests.",
        "Aquí solo queda la interfaz: búsqueda, resumen y tablero visual, sin cálculos ni conexión a base de datos.": "Only the interface remains here: search, summary and visual dashboard, without calculations or database connection.",
        "Archivo CSV/JSON, SQLite o URL REST": "CSV/JSON file, SQLite or REST URL",
        "Archivo PDF adjunto\nEl documento se abrirá con el visor predeterminado del sistema.": "Attached PDF file\nThe document will open with the system's default viewer.",
        "Archivo adjunto disponible.\nUse 'Ver archivo' para abrirlo.": "Attached file available.\nUse 'View file' to open it.",
        "Ausencias\n0": "Absences\n0",
        "Automático": "Automatic",
        "Autorización de acceso": "Access Authorization",
        "Añadir Marcación": "Add Time Entry",
        "Añadir Personal": "Add Employee",
        "Aún no hay remuneraciones registradas.": "No remunerations have been registered yet.",
        "Aún no se ha configurado una dirección de actualización.": "No update address has been configured yet.",
        "Busca el archivo original para asociarlo a este empleado.": "Search for the original file to associate it with this employee.",
        "Busca, filtra y consulta el detalle completo de cada liquidación guardada.": "Search, filter and view the full detail of each saved settlement.",
        "Buscar Registro": "Search Record",
        "Buscar actualización": "Check for Update",
        "Buscar archivo y guardar": "Search File and Save",
        "Buscar deducción": "Search Deduction",
        "Buscar empleado": "Search Employee",
        "Buscar empleado...": "Search employee...",
        "Buscar por DUI o nombre": "Search by ID or name",
        "Buscar por nombre o DUI": "Search by name or ID",
        "Buscar por nombre o DUI...": "Search by name or ID...",
        "Buscar...": "Search...",
        "Buscar:": "Search:",
        "Calcula, revisa y conserva el historial de finiquitos del personal.": "Calculate, review and keep the history of staff settlements.",
        "Calcular": "Calculate",
        "Calcular décimo": "Calculate Bonus",
        "Cambiar Estado": "Change Status",
        "Cambiar Estado del Descuento": "Change Discount Status",
        "Cargando registros...": "Loading records...",
        "Cargar": "Load",
        "Click derecho sobre una fila para editar o eliminar.": "Right-click on a row to edit or delete.",
        "Completa todos los campos principales para mantener la ficha del empleado consistente.": "Complete all main fields to keep the employee record consistent.",
        "Complete el período que desea enviar a aprobación.": "Fill in the period you want to submit for approval.",
        "Complete los datos del descuento o préstamo": "Fill in the discount or loan details",
        "Comprobante de pago": "Payment Receipt",
        "Comprobantes": "Receipts",
        "Comprobantes de pago": "Payment Receipts",
        "Comprueba la sintaxis de los archivos Python sin ejecutar la aplicación.": "Check the syntax of the Python files without running the application.",
        "Concepto:": "Item:",
        "Conector de lector biométrico": "Biometric Reader Connector",
        "Consejo": "Tip",
        "Consulta SQLite:": "SQLite Query:",
        "Consulta el importe proporcional o completo de cada empleado según el período seleccionado.": "Check the proportional or full amount for each employee based on the selected period.",
        "Consulta marcaciones por corte, busca empleados y revisa jornadas registradas.": "Check time entries by cutoff, search employees and review recorded work shifts.",
        "Consulta saldos, revisa períodos y registra nuevas solicitudes del personal.": "Check balances, review periods and register new staff requests.",
        "Consulta y revisa cortes guardados por rango": "View and review saved cutoffs by date range",
        "Contraseña": "Password",
        "Contrato cargado. Puedes actualizar sus datos o consultar el expediente laboral.": "Contract loaded. You can update its details or view the employment record.",
        "Control del Día": "Day Control",
        "Copia de seguridad": "Backup",
        "Copias disponibles": "Available Backups",
        "Corte Activo:": "Active Cutoff:",
        "Corte activo con fechas invalidas": "Active cutoff with invalid dates",
        "Corte de Planilla": "Payroll Cutoff",
        "Corte en solo lectura: sin edicion ni aprobaciones.": "Read-only cutoff: no editing or approvals.",
        "Corte realizado:": "Cutoff performed:",
        "Corte:": "Cutoff:",
        "Corte\nSin corte activo": "Cutoff\nNo active cutoff",
        "Cortes de Fechas": "Date Cutoffs",
        "Crea copias de la base de datos y los documentos, y restaura la última si algo sale mal.": "Create backups of the database and documents, and restore the latest one if something goes wrong.",
        "Crea los horarios laborales que se utilizarán en la empresa.": "Create the work schedules that will be used in the company.",
        "Crear copia ahora": "Create Backup Now",
        "Crear usuario": "Create User",
        "Cálculo del décimo tercer mes": "13th Month Bonus Calculation",
        "Código": "Code",
        "Código de seguridad": "Security Code",
        "Código generado": "Generated Code",
        "Decision de revision": "Review Decision",
        "Deducciones": "Deductions",
        "Departamento:": "Department:",
        "Depto:": "Dept.:",
        "Descarga el comprobante individual con el detalle del pago y las deducciones.": "Download the individual receipt with the payment and deduction details.",
        "Descargando actualización...": "Downloading update...",
        "Descargar Excel": "Download Excel",
        "Descargar PDF": "Download PDF",
        "Descripción breve": "Brief Description",
        "Desde:": "From:",
        "Detalles del Permiso/Ausencia": "Leave/Absence Details",
        "Detalles del registro": "Record Details",
        "Direccion": "Address",
        "Dirección de actualización:": "Update Address:",
        "Documento contractual guardado en la base de datos.": "Contract document saved in the database.",
        "Documento:": "Document:",
        "Dpto": "Dept.",
        "Décimo 3er mes": "13th Month Bonus",
        "Décimo tercer mes": "13th Month Bonus",
        "Días": "Days",
        "Editar Marcacion": "Edit Time Entry",
        "Editar Marcacion del Dia": "Edit Day's Time Entry",
        "Editar empleado": "Edit Employee",
        "Ej: 500 o 500.50": "E.g.: 500 or 500.50",
        "Ej: préstamo extraordinario": "E.g.: extraordinary loan",
        "Ejecutar revisión": "Run Review",
        "El contrato es elaborado y firmado por la empresa. La aplicación solo almacena su archivo y lo vincula al expediente.": "The contract is drafted and signed by the company. The application only stores its file and links it to the record.",
        "El corte seleccionado no tiene registros; mostrando los disponibles.": "The selected cutoff has no records; showing the available ones.",
        "El identificador recibido debe ser el DUI del empleado registrado en SiPP.": "The identifier received must be the national ID of the employee registered in SiPP.",
        "Elige una opción para abrirla a pantalla completa.": "Choose an option to open it in full screen.",
        "Eliminar datos": "Delete Data",
        "Eliminar empleado": "Delete Employee",
        "Eliminar seleccionada": "Delete Selected",
        "Eliminar seleccionado": "Delete Selected",
        "Empleado": "Employee",
        "Empleado:": "Employee:",
        "Encuentra personal por nombre, documento o reporte asociado.": "Find staff by name, document or associated report.",
        "Entidad / Acreedor": "Entity / Creditor",
        "Entrada:": "Check-in:",
        "Enviar solicitud": "Submit Request",
        "Estado:": "Status:",
        "Este empleado aún no tiene un contrato registrado.": "This employee does not have a registered contract yet.",
        "Expedientes laborales": "Employment Records",
        "Fecha Desde:": "Date From:",
        "Fecha Fin": "End Date",
        "Fecha Hasta:": "Date To:",
        "Fecha Inicio": "Start Date",
        "Fecha de aplicación": "Application Date",
        "Forma de conexión:": "Connection Type:",
        "Fórmula": "Formula",
        "Genera una carta laboral profesional para cada empleado registrado.": "Generate a professional employment letter for each registered employee.",
        "Generar carta PDF": "Generate PDF Letter",
        "Gestiona estados de pago y registra los pagos recientes de personal.": "Manage payment statuses and record recent staff payments.",
        "Gestión de Tiempos": "Time Management",
        "Gestión de Vacaciones": "Vacation Management",
        "Gestión empresarial segura": "Secure Business Management",
        "Grabar": "Save",
        "Guarda el contrato emitido por la empresa y consulta el historial completo del personal.": "Save the contract issued by the company and check the full staff history.",
        "Guardar Cambios": "Save Changes",
        "Guardar contrato": "Save Contract",
        "Guardar decisión": "Save Decision",
        "Guardar horarios": "Save Schedules",
        "Guardar liquidación": "Save Settlement",
        "Guardar usuario": "Save User",
        "Género": "Gender",
        "Hasta:": "To:",
        "Hora de entrada": "Check-in Time",
        "Hora de salida": "Check-out Time",
        "Horarios laborales": "Work Schedules",
        "Horarios laborales de la empresa": "Company Work Schedules",
        "Importe total": "Total Amount",
        "Imprimir comprobante": "Print Receipt",
        "Indica la dirección publicada por el administrador de SiPP.": "Enter the address published by the SiPP administrator.",
        "Información de la aplicación": "Application Information",
        "Informe integral del expediente": "Full Record Report",
        "Ingresar": "Log In",
        "Ingrese el código maestro para administrar usuarios": "Enter the master code to manage users",
        "Ingrese el código maestro para crear usuarios": "Enter the master code to create users",
        "Ingrese su contraseña": "Enter your password",
        "Ingrese su usuario": "Enter your username",
        "Ingrese un nuevo concepto de descuento o ajuste salarial": "Enter a new discount concept or salary adjustment",
        "Ingreso": "Hire Date",
        "Inicie sesión para continuar": "Log in to continue",
        "Jornada:": "Work Shift:",
        "La vista quedó limpia y sin cálculos ocultos.": "The view is now clean, with no hidden calculations.",
        "Liquidaciones registradas": "Registered Settlements",
        "Liquidación eliminada. Selecciona otra para consultar su detalle.": "Settlement deleted. Select another one to view its details.",
        "Listado de empleados": "Employee List",
        "Listo para configurar.": "Ready to configure.",
        "Logo:": "Logo:",
        "Los cambios se aplican a las ventanas nuevas.": "Changes apply to new windows.",
        "Los horarios guardados aquí serán los que se utilizarán posteriormente en la empresa.": "The schedules saved here will be the ones used later in the company.",
        "Marcaciones Aprobadas": "Approved Time Entries",
        "Modo:": "Mode:",
        "Monto (B/.)": "Amount (B/.)",
        "Monto Total": "Total Amount",
        "Mostrando todos los registros de planilla.": "Showing all payroll records.",
        "Motivo (opcional)": "Reason (optional)",
        "Motivo (opcional):": "Reason (optional):",
        "Nacimiento": "Date of Birth",
        "Ningún contrato guardado": "No contract saved",
        "No hay corte activo": "No active cutoff",
        "No hay corte activo. Active un corte para ver marcaciones.": "No active cutoff. Activate a cutoff to view time entries.",
        "No hay cortes disponibles": "No cutoffs available",
        "No hay cortes para mostrar marcaciones.": "No cutoffs available to show time entries.",
        "No hay registros guardados. Usa la búsqueda para filtrar por nombre o DUI.": "No records saved. Use the search to filter by name or ID.",
        "No hay reportes disponibles todavía. Usa el buscador para filtrar por nombre, tipo o estado.": "No reports available yet. Use the search to filter by name, type or status.",
        "No hay solicitudes de vacaciones aprobadas.": "There are no approved vacation requests.",
        "No se encontraron empleados con ese criterio.": "No employees found matching that criteria.",
        "No se encontró el módulo de conectores.": "The connectors module was not found.",
        "No se encontró el módulo de respaldo.": "The backup module was not found.",
        "No se ha seleccionado un logo": "No logo selected",
        "No se pudo interpretar el rango del corte activo.": "Could not interpret the active cutoff's date range.",
        "No se pudo leer el rango del corte seleccionado.": "Could not read the selected cutoff's date range.",
        "No se pudo previsualizar la imagen.\nUse 'Ver archivo' para abrirla con el visor del sistema.": "The image could not be previewed.\nUse 'View file' to open it with the system viewer.",
        "Nombre Personal": "Employee Name",
        "Nombre de la entidad": "Entity Name",
        "Nombre de usuario": "Username",
        "Nombre o DUI": "Name or ID",
        "Nombre o Dui": "Name or ID",
        "Nueva liquidación": "New Settlement",
        "Nueva solicitud de vacaciones": "New Vacation Request",
        "Nuevo / actualizar contrato": "New / Update Contract",
        "Nuevo Corte": "New Cutoff",
        "Nuevo ajuste": "New Adjustment",
        "Nuevo descuento autorizado": "New Authorized Discount",
        "Nuevo valor:": "New Value:",
        "Número de identificación": "ID Number",
        "Observaciones": "Remarks",
        "Observaciones (opcional)": "Remarks (optional)",
        "Observaciones:": "Remarks:",
        "Opcional": "Optional",
        "Opciones visuales restauradas.": "Visual options restored.",
        "Otros descuentos autorizados": "Other Authorized Discounts",
        "Pagado": "Paid",
        "Pagados": "Paid",
        "Pago neto": "Net Pay",
        "Permisos y Ausencias": "Leave and Absences",
        "Personal": "Staff",
        "Período:": "Period:",
        "Presentación guiada: qué hace cada parte de SiPP": "Guided presentation: what each part of SiPP does",
        "Probar conexión": "Test Connection",
        "Progresivo": "Progressive",
        "Rangos, aportes y conceptos complementarios": "Ranges, contributions and additional items",
        "Realizar Corte": "Perform Cutoff",
        "Rechazar": "Reject",
        "Reducir movimiento": "Reduce Motion",
        "Registrar Remuneración": "Register Remuneration",
        "Registro manual profesional de entrada y salida del personal": "Professional manual check-in/check-out record for staff",
        "Registro profesional": "Professional Record",
        "Registro rápido de acceso seguro.": "Quick secure access registration.",
        "Registros guardados": "Saved Records",
        "Regla aplicable:": "Applicable Rule:",
        "Reiniciar": "Reset",
        "Reporte de descuentos por corte cerrado": "Discount Report by Closed Cutoff",
        "Reportes de Descuentos": "Discount Reports",
        "Restauración completada. Reinicia SiPP para ver los datos restaurados.": "Restore complete. Restart SiPP to see the restored data.",
        "Restaurar la última": "Restore the Latest",
        "Restaurar seleccionada": "Restore Selected",
        "Resumen del corte activo según las marcaciones registradas.": "Summary of the active cutoff based on recorded time entries.",
        "Resumen por empleado": "Summary by Employee",
        "Retenciones legales y descuentos autorizados de nómina": "Legal withholdings and authorized payroll discounts",
        "Revisar solicitudes": "Review Requests",
        "Revision de Horas Extras": "Overtime Review",
        "Revisión de la aplicación": "Application Review",
        "Revisión de solicitudes": "Request Review",
        "Revisión de solicitudes de vacaciones": "Vacation Request Review",
        "Ruta o URL:": "Path or URL:",
        "Salario:": "Salary:",
        "Salida:": "Check-out:",
        "Seguro Educativo": "Educational Insurance",
        "Seguro educativo": "Educational insurance",
        "Selecciona a la persona cuyo expediente deseas gestionar.": "Select the person whose record you want to manage.",
        "Selecciona un empleado": "Select an employee",
        "Selecciona un empleado para comenzar.": "Select an employee to get started.",
        "Selecciona un empleado para consultar o registrar su contrato laboral.": "Select an employee to view or register their employment contract.",
        "Selecciona un empleado para consultar su contrato.": "Select an employee to view their contract.",
        "Selecciona un empleado para preparar su liquidación.": "Select an employee to prepare their settlement.",
        "Selecciona un empleado y calcula la liquidación.": "Select an employee and calculate the settlement.",
        "Selecciona un período para calcular los importes.": "Select a period to calculate the amounts.",
        "Selecciona un registro para imprimirlo.": "Select a record to print it.",
        "Selecciona un registro para ver información detallada.": "Select a record to view detailed information.",
        "Selecciona una copia para restaurar.": "Select a backup to restore.",
        "Selecciona una liquidación para ver todos sus conceptos.": "Select a settlement to view all its items.",
        "Selecciona una opción para ocupar toda la pantalla de contenido. Cada módulo se abrirá aquí con un botón X para volver.": "Select an option to fill the entire content screen. Each module will open here with an X button to go back.",
        "Seleccionar fecha": "Select Date",
        "Seleccione los filtros para preparar la descarga.": "Select the filters to prepare the download.",
        "Seleccione los filtros y presione Cargar.": "Select the filters and press Load.",
        "Seleccione un usuario para eliminarlo": "Select a user to delete",
        "SiPP | Control de personal y planilla": "SiPP | Personnel and Payroll Control",
        "Siguiente ▶": "Next ▶",
        "Sin archivo seleccionado": "No file selected",
        "Sin contrato": "No Contract",
        "Sincronizar": "Sync",
        "Solicitar": "Request",
        "Supervisa horas extras, filtra empleados y gestiona aprobaciones por corte activo.": "Monitor overtime, filter employees and manage approvals by active cutoff.",
        "Tamaño del texto:": "Text Size:",
        "Tardanzas\n0": "Lateness\n0",
        "Telefono": "Phone",
        "Tema:": "Theme:",
        "Tipo de Retención": "Withholding Type",
        "Tipo de día": "Day Type",
        "Tipo de remuneración": "Remuneration Type",
        "Tipo:": "Type:",
        "Token REST:": "REST Token:",
        "Tolerancia (min):": "Tolerance (min):",
        "Total bruto": "Gross Total",
        "Total de empleados: 0": "Total employees: 0",
        "Total movimientos": "Total Movements",
        "URL del manifiesto JSON": "JSON Manifest URL",
        "Use formato 00:00 y seleccione AM/PM para evitar errores de captura.": "Use the 00:00 format and select AM/PM to avoid input errors.",
        "Usuarios del sistema": "System Users",
        "Usuarios registrados": "Registered Users",
        "Vacaciones aprobadas": "Approved Vacations",
        "Vacaciones:": "Vacations:",
        "Valor actual:": "Current Value:",
        "Valor diario": "Daily Value",
        "Ver Detalles": "View Details",
        "Ver Marcaciones": "View Time Entries",
        "Ver archivo": "View File",
        "Ver informe integral": "View Full Report",
        "Vista de planilla": "Payroll View",
        "Vista previa": "Preview",
        "Visualiza ausencias y permisos guardados en un formato profesional.": "View saved absences and leave in a professional format.",
        "© 2026 Sistema SiPP": "© 2026 SiPP System",
        "↻ Estado": "↻ Status",
        "◀ Anterior": "◀ Previous",
        "✏️ Editar": "✏️ Edit",
        "✓ Guardar": "✓ Save",
        "✕ Cancelar": "✕ Cancel",
        "👤 Empleado": "👤 Employee",
        "💰 Gestión de Descuentos y Préstamos": "💰 Discounts and Loans Management",
        "📁 No hay archivos adjuntos": "📁 No attached files",
        "📋 Gestión de Planilla": "📋 Payroll Management",
        "📥 Descargar": "📥 Download",
        "🗑️ Eliminar": "🗑️ Delete",
        "🗑️ Limpiar": "🗑️ Clear",
        # Títulos de diálogos (messageboxes) - Correcciones de traducción
        "Usuario Guardado": "User Saved",
        "CODIGO DE USUARIO": "USER CODE",
        "Usuario Eliminado": "User Deleted",
        "Actualizado": "Updated",
        "Exito": "Success",
        "Excel": "Excel",
        "Frontend": "Frontend",
        "Solo lectura": "Read-only",
        "Sin horas extras": "No Overtime",
        "Guardado": "Saved",
        "Registro no encontrado": "Record not found",
        "Detalles": "Details",
        "Empleado no encontrado": "Employee not found",
        "Archivo subido": "File uploaded",
        "Solicitud registrada": "Request registered",
        "Liquidación guardada": "Settlement saved",
        "Aviso de asistencia": "Attendance Notice",
        "Horario": "Schedule",
        "Avertencia": "Warning",
        "Registro existente": "Existing Record",
        "Sin resultados": "No Results",
        "Datos invalidos": "Invalid Data",
        "Sin cambios": "No Changes",
        "Error de Monto": "Amount Error",
        "Duplicado": "Duplicate",
        "Selección": "Selection",
        "Atención": "Attention",
        "Reporte": "Report",
        "Logo": "Logo",
        "Campos vacíos": "Empty Fields",
        "Error de sistema": "System Error",
        "Error al eliminar": "Delete Error",
        "Error al guardar": "Save Error",
        "Décimo tercer mes": "13th Month Bonus",
        "Comprobante": "Receipt",
        "Aprobado": "Approved",
        "Rechazado": "Rejected",
        "Informe": "Report",
        "Contratos": "Contracts",
        "Documento": "Document",
        "Solicitud": "Request",
        "Vacaciones": "Vacations",
        "Cartas de trabajo": "Work Letters",
        "Validación": "Validation",
        # Strings para cartas de trabajo
        "CARTA DE TRABAJO": "WORK LETTER",
        "A quien corresponda:": "To Whom It May Concern:",
        "Por medio de la presente hacemos constar que el señor(a) ": "By this letter we certify that ",
        " con documento de identidad ": " with identification document ",
        " labora o laboró en nuestra organización desempeñándose como ": " works or worked in our organization performing as ",
        " en el departamento de ": " in the department of ",
        "El vínculo laboral inició el ": "The work relationship started on ",
        " y registra un tiempo de servicio de ": " and has a service time of ",
        ". Durante este período ha demostrado responsabilidad, compromiso, disposición y un desempeño profesional satisfactorio en las funciones asignadas.": ". During this period has demonstrated responsibility, commitment, willingness and satisfactory professional performance in the assigned functions.",
        "Extendemos la presente carta a solicitud del interesado para los fines que estime convenientes. Certificamos que la información consignada es verdadera según nuestros registros laborales.": "We extend this letter at the request of the interested party for the purposes they deem appropriate. We certify that the information provided is true according to our employment records.",
        "Atentamente,": "Sincerely,",
        "Emitida el ": "Issued on ",
        # Mensajes de bases de datos y conexiones
        "Bienvenido a SiPP": "Welcome to SiPP",
        "Inicio con exito": "Startup successful",
        "Error al Crear Base de Datos": "Database Creation Error",
        "No se pudo crear/acceder a 'SiPP-2':": "Could not create/access 'SiPP-2':",
        "Error al conectar a la base de datos:": "Error connecting to database:",
        "Teléfono": "Phone",
        "Dirección": "Address",
        "Dirección no especificada": "Address not specified",
        "Teléfono no especificado": "Phone not specified",
        "No especificado": "Not specified",
        "Reporte de descuentos aplicados": "Report of discounts applied",
        "RUC": "RUC",
        "COMPROBANTE DE PAGO": "PAYMENT RECEIPT",
    }
}

RUTA_CONFIG_IDIOMA = os.path.join(os.path.abspath("."), "config_idioma.json")


def cargar_idioma_guardado():
    """Lee el idioma guardado en disco de una ejecución anterior; por defecto Español."""
    try:
        with open(RUTA_CONFIG_IDIOMA, "r", encoding="utf-8") as archivo:
            idioma = json.load(archivo).get("idioma", "Español")
        return idioma if idioma == "Español" or idioma in TRADUCCIONES else "Español"
    except (OSError, ValueError):
        return "Español"


def guardar_idioma_en_disco(idioma):
    """Guarda el idioma seleccionado para recordarlo la próxima vez que se abra SiPP."""
    try:
        with open(RUTA_CONFIG_IDIOMA, "w", encoding="utf-8") as archivo:
            json.dump({"idioma": idioma}, archivo)
    except OSError:
        logging.exception("No se pudo guardar el idioma seleccionado en disco")


IDIOMA_APLICACION = {"actual": cargar_idioma_guardado()}


def establecer_idioma_aplicacion(idioma):
    """Fija el idioma activo global usado por el traductor de la interfaz."""
    IDIOMA_APLICACION["actual"] = idioma if idioma in TRADUCCIONES else "Español"


def idioma_aplicacion_actual():
    return IDIOMA_APLICACION["actual"]


def _traducir(idioma, texto):
    """Traduce un texto de la interfaz al idioma indicado; si no hay traducción, devuelve el original."""
    return TRADUCCIONES.get(idioma, {}).get(texto, texto)


def _tr(texto):
    """Traduce un texto usando el idioma global actual (usado por los widgets parcheados)."""
    if not isinstance(texto, str):
        return texto
    return _traducir(IDIOMA_APLICACION["actual"], texto)


def _tr_lista(valores):
    if not isinstance(valores, (list, tuple)):
        return valores
    return [_tr(valor) if isinstance(valor, str) else valor for valor in valores]


def _envolver_texto(funcion_original, claves):
    """Envuelve un método para traducir uno o más kwargs de texto antes de llamarlo."""
    def envoltura(self, *args, **kwargs):
        for clave in claves:
            if clave in kwargs and isinstance(kwargs[clave], str):
                kwargs[clave] = _tr(kwargs[clave])
        return funcion_original(self, *args, **kwargs)
    return envoltura


def _envolver_lista(funcion_original, clave):
    def envoltura(self, *args, **kwargs):
        if clave in kwargs:
            kwargs[clave] = _tr_lista(kwargs[clave])
        return funcion_original(self, *args, **kwargs)
    return envoltura


def _envolver_titulo(funcion_original):
    def envoltura(self, *args, **kwargs):
        if args and isinstance(args[0], str):
            args = (_tr(args[0]),) + args[1:]
        return funcion_original(self, *args, **kwargs)
    return envoltura


def _envolver_messagebox(funcion_original):
    def envoltura(*args, **kwargs):
        args = list(args)
        if len(args) >= 1 and isinstance(args[0], str):
            args[0] = _tr(args[0])
        if len(args) >= 2 and isinstance(args[1], str):
            args[1] = _tr(args[1])
        if isinstance(kwargs.get("title"), str):
            kwargs["title"] = _tr(kwargs["title"])
        if isinstance(kwargs.get("message"), str):
            kwargs["message"] = _tr(kwargs["message"])
        return funcion_original(*args, **kwargs)
    return envoltura


def _activar_traduccion_automatica_widgets():
    """Parchea los widgets usados en toda la app para traducir su texto según el idioma activo.

    No modifica datos provenientes de la base de datos: solo se traducen las cadenas
    que coinciden exactamente con una entrada del diccionario TRADUCCIONES; cualquier
    otro texto (nombres, DUI, montos, etc.) pasa sin cambios.
    """
    for clase in (ctk.CTkLabel, ctk.CTkButton, ctk.CTkCheckBox, ctk.CTkRadioButton):
        clase.__init__ = _envolver_texto(clase.__init__, ("text",))
        clase.configure = _envolver_texto(clase.configure, ("text",))

    ctk.CTkEntry.__init__ = _envolver_texto(ctk.CTkEntry.__init__, ("placeholder_text",))
    ctk.CTkEntry.configure = _envolver_texto(ctk.CTkEntry.configure, ("placeholder_text",))

    ctk.CTkComboBox.__init__ = _envolver_lista(ctk.CTkComboBox.__init__, "values")
    ctk.CTkComboBox.configure = _envolver_lista(ctk.CTkComboBox.configure, "values")

    Menu.add_command = _envolver_texto(Menu.add_command, ("label",))
    Menu.add_cascade = _envolver_texto(Menu.add_cascade, ("label",))

    ttk.Treeview.heading = _envolver_texto(ttk.Treeview.heading, ("text",))

    tk.Wm.title = _envolver_titulo(tk.Wm.title)

    for nombre_funcion in ("showinfo", "showwarning", "showerror", "askyesno", "askquestion", "askokcancel"):
        setattr(messagebox, nombre_funcion, _envolver_messagebox(getattr(messagebox, nombre_funcion)))


_activar_traduccion_automatica_widgets()


class SiPP(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.after(50, self._procesar_cola_ui)
        # Configurar logging para registrar errores y trazas en fichero
        try:
            log_path = os.path.join(os.path.abspath('.'), 'sipp.log')
            logging.basicConfig(filename=log_path, level=logging.DEBUG,
                                format='%(asctime)s %(levelname)s: %(message)s')
            import traceback
            def _report_callback_exception(exc, val, tb):
                logging.error('Exception in Tk callback', exc_info=(exc, val, tb))
                # Also print to console for developer convenience
                traceback.print_exception(exc, val, tb)
            self.report_callback_exception = _report_callback_exception
        except Exception:
            pass
        # Ensure we have a place to track scheduled after callbacks
        self._after_ids = set()
        self._postgresql_credentials = None
        self._inicio_cancelado = False
        self.title("SiPP")
        if not self.solicitar_aceptacion_condiciones():
            self._inicio_cancelado = True
            return
        if not self.solicitar_autorizacion_codigo():
            self._inicio_cancelado = True
            return
        if not self.solicitar_acceso_postgresql():
            self._inicio_cancelado = True
            return
        if db is not None and hasattr(db, 'basedatos'):
            try:
                db.basedatos()
            except Exception:
                pass
        self.vista()
        self.tema()
        self.ventana_Login()
        self.icono()

    def _procesar_cola_ui(self):
        try:
            while True:
                COLA_UI.get_nowait()()
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(50, self._procesar_cola_ui)

    def solicitar_aceptacion_condiciones(self):
        if cargar_aceptacion_condiciones() == VERSION_CONDICIONES:
            return True
        try:
            cliente = ClienteAutorizacion()
        except ErrorAutorizacion as exc:
            messagebox.showerror(
                "Condiciones de uso",
                "No se pueden registrar las condiciones porque no hay conexion con el servidor de SiPP.\n\n"
                f"Detalle: {exc}",
                parent=self,
            )
            return False

        resultado = {"aceptado": False, "en_curso": False}
        top_level = ctk.CTkToplevel(self)
        top_level.title("Condiciones de uso de SiPP")
        ajustar_toplevel_a_pantalla(top_level, 700, 620, margen_x=70, margen_y=70, min_ancho=520, min_alto=420)
        top_level.resizable(True, True)
        top_level.transient(self)
        top_level.grab_set()
        top_level.protocol("WM_DELETE_WINDOW", top_level.destroy)

        contenedor = ctk.CTkFrame(top_level, corner_radius=14)
        contenedor.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            contenedor,
            text="Condiciones de uso y aviso de seguridad",
            font=ctk.CTkFont(size=21, weight="bold"),
        ).pack(pady=(16, 2))
        ctk.CTkLabel(
            contenedor,
            text=f"Version {VERSION_CONDICIONES}. Debe leer y aceptar este documento para continuar.",
            wraplength=620,
        ).pack(pady=(0, 10))
        texto = ctk.CTkTextbox(contenedor, wrap="word", height=260)
        texto.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        texto.insert("1.0", TEXTO_CONDICIONES)
        texto.configure(state="disabled")
        acepto = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            contenedor,
            text="He leido y acepto las condiciones de uso.",
            variable=acepto,
        ).pack(anchor="w", padx=20, pady=(0, 8))
        mensaje = ctk.CTkLabel(contenedor, text="", wraplength=620, justify="left")
        mensaje.pack(fill="x", padx=20, pady=(0, 8))

        def registrar():
            if not acepto.get():
                mensaje.configure(text="Debe marcar la casilla de aceptación para continuar.")
                return
            if resultado["en_curso"]:
                return
            resultado["en_curso"] = True
            boton_aceptar.configure(state="disabled")
            mensaje.configure(text="Registrando la aceptación...")
            ejecutar_en_segundo_plano(
                lambda: cliente.registrar_aceptacion_condiciones(VERSION_CONDICIONES, TEXTO_CONDICIONES),
                lambda respuesta: programar_en_ui(lambda: aceptar_registrado(respuesta)),
                lambda exc: programar_en_ui(lambda: fallo_aceptacion(exc)),
            )

        def aceptar_registrado(respuesta):
            if str(respuesta.get("status", "")).upper() != "REGISTRADA":
                fallo_aceptacion(RuntimeError("El servidor no confirmó la aceptación."))
                return
            guardar_aceptacion_condiciones(VERSION_CONDICIONES)
            resultado["aceptado"] = True
            mensaje.configure(text="Aceptación registrada. Continuando...")
            top_level.after(300, top_level.destroy)

        def fallo_aceptacion(exc):
            resultado["en_curso"] = False
            boton_aceptar.configure(state="normal")
            mensaje.configure(
                text="No se pudo registrar la aceptación. Revise su conexión a internet e inténtelo nuevamente."
            )
            logging.warning("Fallo al registrar condiciones: %s", exc)

        botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkButton(
            botones,
            text="No acepto",
            width=120,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(side="left")
        boton_aceptar = ctk.CTkButton(botones, text="Aceptar y continuar", width=170, command=registrar)
        boton_aceptar.pack(side="right")
        top_level.wait_window()
        return resultado["aceptado"]

    def solicitar_autorizacion_codigo(self):
        if cargar_autorizacion_local():
            return True
        try:
            cliente = ClienteAutorizacion()
        except ErrorAutorizacion as exc:
            messagebox.showerror(
                "Autorizacion",
                "No se puede iniciar la autorizacion. Verifique que haya internet y que el servidor de SiPP este configurado.\n\n"
                f"Detalle: {exc}",
                parent=self,
            )
            return False

        estado = {"autorizado": False, "request_id": "", "expires_at": None, "ocupado": False}
        top_level = ctk.CTkToplevel(self)
        top_level.title("Autorizacion de SiPP")
        ajustar_toplevel_a_pantalla(top_level, 520, 440, margen_x=70, margen_y=70, min_ancho=420, min_alto=360)
        top_level.resizable(False, False)
        top_level.transient(self)
        top_level.grab_set()
        top_level.protocol("WM_DELETE_WINDOW", top_level.destroy)

        contenedor = ctk.CTkFrame(top_level, corner_radius=14)
        contenedor.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            contenedor,
            text="Autorizacion de SiPP",
            font=ctk.CTkFont(size=21, weight="bold"),
        ).pack(pady=(18, 4))
        ctk.CTkLabel(
            contenedor,
            text="Esta computadora necesita autorizacion para continuar.",
            wraplength=400,
        ).pack(pady=(0, 4))
        ctk.CTkLabel(
            contenedor,
            text=f"Equipo: {obtener_datos_equipo()['device_name']}",
            anchor="w",
        ).pack(fill="x", padx=24, pady=(0, 10))
        mensaje = ctk.CTkLabel(contenedor, text="Solicitando codigo...", wraplength=410, justify="left")
        mensaje.pack(fill="x", padx=24, pady=(0, 8))
        entrada_codigo = ctk.CTkEntry(contenedor, width=230, height=38, justify="center", placeholder_text="Codigo de autorizacion")
        entrada_codigo.pack(pady=4)
        contador = ctk.CTkLabel(contenedor, text="", text_color=("#475569", "#cbd5e1"))
        contador.pack(pady=(2, 8))

        def actualizar_contador():
            if not top_level.winfo_exists() or estado["autorizado"]:
                return
            expiracion = estado["expires_at"]
            if expiracion:
                restante = max(0, int((expiracion - datetime.now(timezone.utc)).total_seconds()))
                contador.configure(text=f"El codigo caduca en {restante // 60:02d}:{restante % 60:02d}")
                if restante == 0:
                    mensaje.configure(text="El codigo caduco. Solicite uno nuevo.")
            top_level.after(1000, actualizar_contador)

        def fallo(exc):
            estado["ocupado"] = False
            mensaje.configure(
                text="No hay conexion con el servidor de autorizacion.\n"
                "Revise internet y vuelva a solicitar el codigo."
            )
            logging.warning("Fallo de autorizacion: %s", exc)
            boton_solicitar.configure(state="normal")
            boton_validar.configure(state="normal")

        def recibir_solicitud(resultado):
            estado["ocupado"] = False
            estado["request_id"] = str(resultado["request_id"])
            try:
                estado["expires_at"] = datetime.fromisoformat(str(resultado["expires_at"]).replace("Z", "+00:00"))
            except (KeyError, ValueError, TypeError):
                estado["expires_at"] = None
            codigo_local = str(resultado.get("code", "")).strip()
            if codigo_local:
                mensaje.configure(text=f"Codigo local de autorizacion: {codigo_local}")
            else:
                mensaje.configure(text="Se envio un codigo al correo del administrador.")
            boton_solicitar.configure(state="normal")
            boton_validar.configure(state="normal")

        def solicitar_codigo():
            if estado["ocupado"]:
                return
            estado["ocupado"] = True
            boton_solicitar.configure(state="disabled")
            boton_validar.configure(state="disabled")
            mensaje.configure(text="Solicitando codigo...")
            ejecutar_en_segundo_plano(
                cliente.solicitar_codigo,
                lambda resultado: programar_en_ui(lambda: recibir_solicitud(resultado)),
                lambda exc: programar_en_ui(lambda: fallo(exc)),
            )

        def validar_codigo():
            codigo = entrada_codigo.get().strip()
            if not estado["request_id"]:
                mensaje.configure(text="Primero solicite un codigo.")
                return
            if not codigo:
                mensaje.configure(text="Ingrese el codigo recibido.")
                return
            if estado["ocupado"]:
                return
            estado["ocupado"] = True
            boton_solicitar.configure(state="disabled")
            boton_validar.configure(state="disabled")
            mensaje.configure(text="Validando codigo...")
            ejecutar_en_segundo_plano(
                lambda: cliente.validar_codigo(estado["request_id"], codigo),
                lambda resultado: programar_en_ui(lambda: autorizar(resultado)),
                lambda exc: programar_en_ui(lambda: fallo(exc)),
            )

        def autorizar(resultado):
            estado["ocupado"] = False
            if str(resultado.get("status", "")).upper() != "APROBADO":
                mensaje.configure(text=str(resultado.get("message", "El codigo no fue aprobado.")))
                boton_solicitar.configure(state="normal")
                boton_validar.configure(state="normal")
                return
            guardar_autorizacion_local()
            estado["autorizado"] = True
            mensaje.configure(text="Equipo autorizado. Continuando...")
            top_level.after(300, top_level.destroy)

        botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        botones.pack(fill="x", padx=24, pady=(2, 12))
        boton_solicitar = ctk.CTkButton(botones, text="Solicitar nuevo codigo", width=180, command=solicitar_codigo)
        boton_solicitar.pack(side="left")
        boton_validar = ctk.CTkButton(botones, text="Validar", width=100, command=validar_codigo)
        boton_validar.pack(side="right")
        entrada_codigo.bind("<Return>", lambda event: validar_codigo())
        top_level.after(100, solicitar_codigo)
        top_level.after(1000, actualizar_contador)
        top_level.wait_window()
        return estado["autorizado"]

    def solicitar_acceso_postgresql(self):
        if self._postgresql_credentials is not None:
            return True

        import config_seguridad

        def validar_conexion(host, puerto, base_datos, nombre_usuario, clave):
            import psycopg

            sslmode = "disable" if host in ("localhost", "127.0.0.1") else config_seguridad.DB_SSLMODE
            with psycopg.connect(
                host=host,
            port=puerto,
            dbname=base_datos,
                user=nombre_usuario,
                password=clave,
                sslmode=sslmode,
                connect_timeout=5,
            ) as conexion:
                conexion.execute("SELECT 1")

        credenciales_guardadas = config_seguridad.cargar_credenciales()
        if not credenciales_guardadas:
            usuario_entorno = os.getenv("SIPP_DB_USER", "").strip()
            contraseña_entorno = os.getenv("SIPP_DB_PASSWORD", "")
            if usuario_entorno and contraseña_entorno:
                credenciales_guardadas = {
                    "usuario": usuario_entorno,
                    "contraseña": contraseña_entorno,
                    "host": os.getenv("SIPP_DB_HOST", config_seguridad.DB_HOST),
                    "puerto": os.getenv("SIPP_DB_PORT", config_seguridad.DB_PORT),
                    "base_datos": os.getenv("SIPP_DB_NAME", config_seguridad.DB_NAME),
                }
        if credenciales_guardadas:
            try:
                validar_conexion(
                    credenciales_guardadas["host"],
                    credenciales_guardadas["puerto"],
                    credenciales_guardadas["base_datos"],
                    credenciales_guardadas["usuario"],
                    credenciales_guardadas["contraseña"],
                )
            except Exception:
                credenciales_guardadas = None
            else:
                config_seguridad.DB_HOST = credenciales_guardadas["host"]
                config_seguridad.DB_PORT = credenciales_guardadas["puerto"]
                config_seguridad.DB_NAME = credenciales_guardadas["base_datos"]
                config_seguridad.DB_USER = credenciales_guardadas["usuario"]
                config_seguridad.DB_PASSWORD = credenciales_guardadas["contraseña"]
                if db is not None:
                    db.DB_HOST = config_seguridad.DB_HOST
                    db.DB_PORT = config_seguridad.DB_PORT
                    db.DB_NAME = config_seguridad.DB_NAME
                    db.DB_USER = config_seguridad.DB_USER
                    db.DB_PASSWORD = credenciales_guardadas["contraseña"]
                if respaldo_sipp is not None:
                    respaldo_sipp.HOST = config_seguridad.DB_HOST
                    respaldo_sipp.PUERTO = config_seguridad.DB_PORT
                    respaldo_sipp.BASE_DATOS = config_seguridad.DB_NAME
                    respaldo_sipp.USUARIO = config_seguridad.DB_USER
                self._postgresql_credentials = credenciales_guardadas
                return True

        resultado = {"aceptado": False}
        top_level = ctk.CTkToplevel(self)
        top_level.title("Conexión a PostgreSQL")
        top_level.geometry("500x430")
        top_level.update_idletasks()
        ancho_pantalla = top_level.winfo_screenwidth()
        alto_pantalla = top_level.winfo_screenheight()
        x = max(0, (ancho_pantalla - 500) // 2)
        y = max(0, (alto_pantalla - 430) // 2)
        top_level.geometry(f"500x430+{x}+{y}")
        top_level.resizable(False, False)
        top_level.transient(self)
        top_level.grab_set()

        contenedor = ctk.CTkFrame(top_level, corner_radius=14)
        contenedor.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            contenedor,
            text="Conectar con PostgreSQL",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(18, 12))
        host = ctk.CTkEntry(contenedor, width=320, placeholder_text="Servidor PostgreSQL")
        host.pack(pady=6)
        host.insert(0, (credenciales_guardadas or {}).get("host", config_seguridad.DB_HOST))
        puerto = ctk.CTkEntry(contenedor, width=320, placeholder_text="Puerto PostgreSQL")
        puerto.pack(pady=6)
        puerto.insert(0, str((credenciales_guardadas or {}).get("puerto", config_seguridad.DB_PORT)))
        base_datos = ctk.CTkEntry(contenedor, width=320, placeholder_text="Nombre de la base de datos")
        base_datos.pack(pady=6)
        base_datos.insert(0, (credenciales_guardadas or {}).get("base_datos", config_seguridad.DB_NAME))
        usuario = ctk.CTkEntry(contenedor, width=320, placeholder_text="Usuario de PostgreSQL")
        usuario.pack(pady=6)
        usuario_configurado = (credenciales_guardadas or {}).get("usuario") or os.getenv("SIPP_DB_USER", "").strip()
        usuario.insert(0, "postgres" if not usuario_configurado or usuario_configurado.startswith("@@") else usuario_configurado)
        contraseña = ctk.CTkEntry(
            contenedor,
            width=320,
            placeholder_text="Contraseña de PostgreSQL",
            show="*",
        )
        contraseña.pack(pady=6)
        contraseña.insert(0, os.getenv("SIPP_DB_PASSWORD", ""))
        mensaje = ctk.CTkLabel(contenedor, text="", text_color="#dc2626")
        mensaje.pack(pady=(2, 4))

        def aceptar():
            nombre_host = host.get().strip()
            numero_puerto = puerto.get().strip()
            nombre_base_datos = base_datos.get().strip()
            nombre_usuario = usuario.get().strip()
            clave = contraseña.get()
            if not nombre_host or not numero_puerto or not nombre_base_datos or not nombre_usuario or not clave:
                mensaje.configure(text="Completa todos los datos de conexión.")
                return
            try:
                validar_conexion(nombre_host, numero_puerto, nombre_base_datos, nombre_usuario, clave)
                config_seguridad.DB_HOST = nombre_host
                config_seguridad.DB_PORT = numero_puerto
                config_seguridad.DB_NAME = nombre_base_datos
                config_seguridad.DB_USER = nombre_usuario
                config_seguridad.DB_PASSWORD = clave
                if db is not None:
                    db.DB_HOST = nombre_host
                    db.DB_PORT = numero_puerto
                    db.DB_NAME = nombre_base_datos
                    db.DB_USER = nombre_usuario
                    db.DB_PASSWORD = clave
                if respaldo_sipp is not None:
                    respaldo_sipp.HOST = nombre_host
                    respaldo_sipp.PUERTO = numero_puerto
                    respaldo_sipp.BASE_DATOS = nombre_base_datos
                    respaldo_sipp.USUARIO = nombre_usuario
                self._postgresql_credentials = {
                    "usuario": nombre_usuario,
                    "contraseña": clave,
                    "host": nombre_host,
                    "puerto": numero_puerto,
                    "base_datos": nombre_base_datos,
                }
                config_seguridad.guardar_credenciales(
                    nombre_usuario,
                    clave,
                    nombre_host,
                    numero_puerto,
                    nombre_base_datos,
                )
                resultado["aceptado"] = True
                top_level.destroy()
            except Exception as exc:
                mensaje.configure(text=f"Usuario o contraseña incorrectos, o no se puede conectar a PostgreSQL: {exc}")

        acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
        acciones.pack(fill="x", padx=28, pady=(6, 12))
        ctk.CTkButton(acciones, text="Conectar", width=130, command=aceptar).pack(side="left")
        ctk.CTkButton(
            acciones,
            text="Cancelar",
            width=110,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(side="right")
        contraseña.bind("<Return>", lambda event: aceptar())
        top_level.protocol("WM_DELETE_WINDOW", top_level.destroy)
        top_level.wait_window()
        return resultado["aceptado"]

    def tema(self):
        if edict is None or not hasattr(edict, 'Tema'):
            ctk.set_default_color_theme("blue")
            return
        try:
            datos = edict.Tema()
            aparen = datos.obtener_tema()
            for i in aparen:
                for x in i:
                    aparen = x
            try:
                font_dir = theme_manager.get_font_path()
                if font_dir and font_dir.exists():
                    FR_PRIVATE = 0x10
                    for f in font_dir.iterdir():
                        if f.suffix.lower() in (".ttf", ".otf"):
                            try:
                                ctypes.windll.gdi32.AddFontResourceExW(str(f), FR_PRIVATE, 0)
                            except Exception:
                                pass
            except Exception:
                pass

            try:
                apariencia = theme_manager.get(str(aparen))
                ctk.set_default_color_theme(str(apariencia))
            except FileNotFoundError:
                ctk.set_default_color_theme("blue")
            except Exception:
                ctk.set_default_color_theme("blue")
        except Exception:
            ctk.set_default_color_theme("blue")

    def vista(self):
        if edict is None or not hasattr(edict, 'Tema'):
            try:
                ctk.set_appearance_mode("system")
            except Exception:
                pass
            return
        try:
            datos = edict.Tema()
            tema = datos.obtener_vista()
            for i in tema:
                for x in i:
                    tema = x
            ctk.set_appearance_mode(str(tema))
        except Exception:
            try:
                ctk.set_appearance_mode("system")
            except Exception:
                pass
        
    def icono(self):
        icono = resource_path("imagenes/icono1.ico")
        self.iconbitmap(icono)


    def ventana_Login(self):

        ancho_pantalla = self.winfo_screenwidth()
        alto_pantalla = self.winfo_screenheight()
        ancho_ventana = min(900, max(720, ancho_pantalla - 80))
        alto_ventana = min(480, max(420, alto_pantalla - 80))
        ventana = self.posicionar_ventana(ancho_ventana, alto_ventana)
        self.geometry(f"{ancho_ventana}x{alto_ventana}+{ventana[0]}+{ventana[1]}")
        self.minsize(min(720, max(320, ancho_pantalla - 40)), min(420, max(300, alto_pantalla - 80)))
        self.frameventana_principal = ctk.CTkFrame(
            self,
            corner_radius=20,
            fg_color=("#edf6fb", "#101b2a"),
            border_width=1,
            border_color=("#b9d6e5", "#274068")
        )
        self.frameventana_principal.pack(fill="both", expand=True)
        self.resizable(False, False)

        # Panel izquierdo con identidad visual
        frame1 = ctk.CTkFrame(
            self.frameventana_principal,
            corner_radius=18,
            fg_color=("#d9edf7", "#17263d"),
            width=420,
            height=480,
            border_width=1
        )
        frame1.grid(row=0, column=0, padx=(18, 10), pady=18, sticky="nsew")
        frame1.grid_propagate(False)

        try:
            imagen = ctk.CTkImage(
                light_image=Image.open(resource_path("imagenes/icono1.png")),
                dark_image=Image.open(resource_path("imagenes/icono1.png")),
                size=(180, 180)
            )
        except Exception:
            imagen = None

        if imagen is not None:
            label_imagen = ctk.CTkLabel(frame1, image=imagen, text="", width=180, height=180)
            label_imagen.place(relx=0.5, rely=0.20, anchor="center")

        ctk.CTkLabel(
            frame1,
            text="SiPP",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=("#0f172a", "#e8eef5")
        ).place(relx=0.5, rely=0.48, anchor="center")

        ctk.CTkLabel(
            frame1,
            text="Sistema de Planilla Profesional",
            font=ctk.CTkFont(size=12),
            text_color=("#435468", "#b7c2ce")
        ).place(relx=0.5, rely=0.56, anchor="center")

        ctk.CTkLabel(
            frame1,
            text="Gestión empresarial segura",
            font=ctk.CTkFont(size=10),
            text_color=("#526071", "#a1a9b8")
        ).place(relx=0.5, rely=0.64, anchor="center")

        # Línea decorativa inferior
        ctk.CTkFrame(frame1, height=2, corner_radius=1, fg_color=("#2f798c", "#48bdec")).place(relx=0.10, rely=0.73, relwidth=0.80)

        # Panel derecho con formulario
        frame2 = ctk.CTkFrame(
            self.frameventana_principal,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            width=430,
            height=480,
            border_width=1,
            border_color=("#ccd7de", "#2b3b4d")
        )
        frame2.grid(row=0, column=1, padx=(10, 18), pady=18, sticky="nsew")
        frame2.grid_propagate(False)

        ctk.CTkLabel(
            frame2,
            text="Acceso al sistema",
            font=ctk.CTkFont(size=26, weight="bold"),
            anchor="w"
        ).place(x=40, y=34)

        ctk.CTkLabel(
            frame2,
            text="Inicie sesión para continuar",
            font=ctk.CTkFont(size=11),
            text_color=("#526071", "#a1a9b8")
        ).place(x=40, y=82)

        ctk.CTkLabel(frame2, text="Usuario", font=ctk.CTkFont(size=13, weight="bold"), anchor="w").place(x=40, y=130)
        usuario = ctk.CTkEntry(frame2, placeholder_text="Ingrese su usuario", width=300, height=38, font=ctk.CTkFont(size=13))
        usuario.place(x=40, y=160)
        usuario.focus_set()

        ctk.CTkLabel(frame2, text="Contraseña", font=ctk.CTkFont(size=13, weight="bold"), anchor="w").place(x=40, y=220)
        contenedor_contraseña = ctk.CTkFrame(frame2, fg_color="transparent", width=300, height=38)
        contenedor_contraseña.place(x=40, y=250)
        contenedor_contraseña.grid_propagate(False)
        contraseña = ctk.CTkEntry(contenedor_contraseña, placeholder_text="Ingrese su contraseña", show="*", width=300, height=38, font=ctk.CTkFont(size=13))
        contraseña.place(x=0, y=0)
        contraseña_visible = {"valor": False}

        def alternar_contraseña():
            contraseña_visible["valor"] = not contraseña_visible["valor"]
            contraseña.configure(show="" if contraseña_visible["valor"] else "*")
            boton_ver_contraseña.configure(
                fg_color=("#dbeafe", "#274c77") if contraseña_visible["valor"] else "transparent"
            )

        boton_ver_contraseña = ctk.CTkButton(
            contenedor_contraseña,
            text="👁",
            width=30,
            height=30,
            font=ctk.CTkFont(size=15),
            fg_color="transparent",
            hover_color=("#dbeafe", "#274c77"),
            text_color=("#475569", "#cbd5e1"),
            command=alternar_contraseña,
        )
        boton_ver_contraseña.place(relx=1.0, rely=0.5, anchor="e", x=-4, y=0)

        ctk.CTkButton(
            frame2,
            text="Ingresar",
            width=300,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=lambda: self.iniciar_sesion(usuario.get(), contraseña.get())
        ).place(x=40, y=330)
        contraseña.bind("<Return>", lambda event: self.iniciar_sesion(usuario.get(), contraseña.get()))

        ctk.CTkLabel(
            frame2,
            text="© 2026 Sistema SiPP",
            font=ctk.CTkFont(size=10),
            text_color=("#526071", "#a1a9b8")
        ).place(x=40, y=398)

    def entrada_ventana_pincipal(self):
        self.frameventana_principal.destroy()
        self.ventana_principal_obj = ventana_principal()
        self.ventana_principal_obj.idioma_actual = idioma_aplicacion_actual()
        self.ventana_principal_obj.ventana_principal_entrada(self)

    def iniciar_sesion(self, usuario, contraseña):
        usuario = str(usuario or "").strip()
        if not usuario or not contraseña:
            messagebox.showwarning("Acceso", "Ingrese usuario y contraseña.", parent=self)
            return
        try:
            datos_usuario = db.ConexionDB_datos_usuarios().autenticar(usuario, contraseña) if db else None
        except Exception as exc:
            logging.exception("No se pudo autenticar al usuario")
            messagebox.showerror(
                "Sin conexion",
                "No se pudo conectar con la base de datos.\n"
                "Revise su conexion a internet o la red del servidor e intentelo nuevamente.",
                parent=self,
            )
            return
        if not datos_usuario:
            messagebox.showerror("Acceso", "Usuario o contraseña incorrectos.", parent=self)
            return
        self.usuario_autenticado = datos_usuario
        self.entrada_ventana_pincipal()

    def entrada_ventana_pincipal_enter(self, event):
        self.frameventana_principal.destroy()
        self.ventana_principal_obj = ventana_principal()
        self.ventana_principal_obj.idioma_actual = idioma_aplicacion_actual()
        self.ventana_principal_obj.ventana_principal_entrada(self)

    def tamaño_pantalla(self):
        ancho = self.winfo_screenwidth()
        alto = self.winfo_screenheight()
        return ancho, alto
    
    def posicionar_ventana(self, ancho_ventana=900, alto_ventana=480):
        ancho, alto = self.tamaño_pantalla()
        x = max(0, (ancho - ancho_ventana) // 2)
        y = max(0, (alto - alto_ventana) // 2)
        return x, y



class ventana_principal:

    def t(self, texto):
        return _traducir(getattr(self, "idioma_actual", "Español"), texto)

    @staticmethod
    def _es_administrador(sipp):
        usuario = getattr(sipp, "usuario_autenticado", None) or {}
        return str(usuario.get("rol", "")).upper() == "ADMIN"

    def tamaño_pantalla_ancho(self, app):
        ancho = app.winfo_screenwidth()
        return ancho

    def tamaño_pantalla_alto(self, app):
        alto = app.winfo_screenheight()
        return alto

    def _ancho_sidebar(self, app):
        return round(self.tamaño_pantalla_ancho(app) * 0.19)

    def _crear_area_contenido(self, sipp):
        ancho_sidebar = self._ancho_sidebar(sipp)
        # Best-effort cleanup: release modal grabs and cancel scheduled callbacks
        try:
            root = sipp.winfo_toplevel()
            # If some widget currently has the grab, release it
            try:
                grabbed = root.grab_current()
                if grabbed and hasattr(grabbed, 'grab_release'):
                    try:
                        grabbed.grab_release()
                    except Exception:
                        pass
                # also attempt to call grab_release on root itself
                try:
                    if hasattr(root, 'grab_release'):
                        root.grab_release()
                except Exception:
                    pass
            except Exception:
                pass

            # Release grabs and cancel after callbacks on any Toplevel children
            for w in root.winfo_children():
                try:
                    # release grab if present
                    if hasattr(w, 'grab_release'):
                        try:
                            w.grab_release()
                        except Exception:
                            pass

                    # cancel tracked after callbacks (wrapper stores ids on _after_ids)
                    if hasattr(w, '_after_ids') and isinstance(w._after_ids, set):
                        for aid in list(w._after_ids):
                            try:
                                w.after_cancel(aid)
                            except Exception:
                                pass
                        try:
                            w._after_ids.clear()
                        except Exception:
                            pass
                    # try to force widget to process pending idle tasks
                    try:
                        if hasattr(w, 'update_idletasks'):
                            w.update_idletasks()
                    except Exception:
                        pass
                    # If a Toplevel/window was left open, destroy it to avoid lingering grabs
                    try:
                        if isinstance(w, (tk.Toplevel, ctk.CTkToplevel)):
                            try:
                                w.destroy()
                            except Exception:
                                pass
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass
        # Also clean any toplevels registered centrally on the root
        try:
            reg = getattr(root, '_registered_toplevels', set())
            for t in list(reg):
                try:
                    if getattr(t, 'winfo_exists', lambda: False)() :
                        try:
                            t.grab_release()
                        except Exception:
                            pass
                        try:
                            t.destroy()
                        except Exception:
                            pass
                except Exception:
                    pass
            try:
                if hasattr(root, '_registered_toplevels'):
                    root._registered_toplevels.clear()
            except Exception:
                pass
        except Exception:
            pass

        # If this controller stored a current_view, clear it when recreating the content area
        try:
            if hasattr(self, 'current_view'):
                self.current_view = None
        except Exception:
            pass

        if hasattr(self, "frame2") and self.frame2.winfo_exists():
            self.frame2.destroy()

        ancho_contenido = max(300, self.tamaño_pantalla_ancho(sipp) - ancho_sidebar)
        self.frame2 = ctk.CTkFrame(
            sipp,
            width=ancho_contenido,
            corner_radius=0,
            fg_color=("#f5f7fb", "#111827"),
        )
        self.frame2.place(x=ancho_sidebar, y=0, relheight=1.0)
        return self.frame2

    def _crear_boton_lateral(self, parent, texto, comando):
        return ctk.CTkButton(
            parent,
            text=texto,
            height=42,
            corner_radius=12,
            anchor="w",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#e9eef8", "#1f2937"),
            hover_color=("#d8e3f5", "#374151"),
            text_color=("#0f172a", "#f8fafc"),
            command=comando,
        )

    def ajustar_toplevel_a_pantalla(self, top_level, ancho, alto, margen_x=80, margen_y=100):
        try:
            ancho_pantalla = top_level.winfo_screenwidth()
            alto_pantalla = top_level.winfo_screenheight()
        except Exception:
            ancho_pantalla = 1920
            alto_pantalla = 1080

        ancho_ajustado = min(ancho, max(320, ancho_pantalla - margen_x))
        alto_ajustado = min(alto, max(260, alto_pantalla - margen_y))
        top_level.geometry(f"{ancho_ajustado}x{alto_ajustado}")
        return ancho_ajustado, alto_ajustado

    def ventana_principal_entrada(self, Sipp):
        ancho_pantalla = Sipp.winfo_screenwidth()
        alto_pantalla = Sipp.winfo_screenheight()
        Sipp.minsize(min(1024, max(760, ancho_pantalla - 40)), min(680, max(420, alto_pantalla - 100)))
        Sipp.maxsize(ancho_pantalla, alto_pantalla)
        Sipp.state("zoomed")
        Sipp.resizable(True, True)
        self.menu(Sipp)
        self.frameprincipal = ctk.CTkFrame(Sipp)
        self.frameprincipal.pack(fill="both", expand=True)
        self.plataformas(self.frameprincipal)
        Sipp.bind("<Control-Alt-l>", self.abrir_limpiar)
        Sipp.bind("<Control-Shift-g>", self.abrir_log)
        Sipp.bind("<Control-Alt-e>", self.abrir_eliminar_administradores)
        Sipp.after(1200, lambda: self.revisar_notificaciones_asistencia(Sipp))

    def abrir_eliminar_administradores(self, event=None):
        sipp = self.frameprincipal.winfo_toplevel()
        if not self._es_administrador(sipp):
            messagebox.showerror(
                "Acceso denegado",
                "Solo un administrador puede eliminar administradores.",
                parent=sipp,
            )
            return "break"

        top_level = ctk.CTkToplevel(sipp)
        top_level.title("Eliminar administradores")
        ajustar_toplevel_a_pantalla(
            top_level,
            560,
            480,
            margen_x=80,
            margen_y=80,
            min_ancho=460,
            min_alto=380,
        )
        top_level.resizable(False, True)
        top_level.transient(sipp)
        top_level.grab_set()

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            frame,
            text="Eliminar administradores",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=22, pady=(20, 4))
        ctk.CTkLabel(
            frame,
            text="@@22 está protegido y no se puede eliminar.",
            anchor="w",
            text_color=("#b45309", "#fbbf24"),
        ).pack(fill="x", padx=22, pady=(0, 14))

        lista = ctk.CTkScrollableFrame(frame, corner_radius=12)
        lista.pack(fill="both", expand=True, padx=22, pady=(0, 12))

        def cargar_administradores():
            for widget in lista.winfo_children():
                widget.destroy()
            try:
                administradores = db.ConexionDB_datos_usuarios().obtener_administradores()
            except Exception as exc:
                ctk.CTkLabel(
                    lista,
                    text=f"No se pudieron cargar los administradores:\n{exc}",
                    anchor="w",
                    justify="left",
                ).pack(fill="x", padx=12, pady=12)
                return
            if not administradores:
                ctk.CTkLabel(lista, text="No hay administradores registrados.").pack(pady=18)
                return
            for nombre in administradores:
                fila = ctk.CTkFrame(lista, fg_color="transparent")
                fila.pack(fill="x", padx=8, pady=6)
                ctk.CTkLabel(fila, text=str(nombre), anchor="w").pack(side="left", fill="x", expand=True)
                protegido = str(nombre).casefold() == "@@22"

                def eliminar(nombre_administrador=nombre):
                    if str(nombre_administrador).casefold() == "@@22":
                        messagebox.showwarning(
                            "Administrador protegido",
                            "El administrador @@22 no se puede eliminar.",
                            parent=top_level,
                        )
                        return
                    confirmar = messagebox.askyesno(
                        "Confirmar eliminación",
                        f"¿Deseas eliminar al administrador {nombre_administrador}?",
                        parent=top_level,
                    )
                    if not confirmar:
                        return
                    try:
                        eliminado = db.ConexionDB_datos_usuarios().eliminar_administrador(nombre_administrador)
                        if not eliminado:
                            raise RuntimeError("El administrador no existe o está protegido.")
                    except Exception as exc:
                        messagebox.showerror("Eliminar administrador", str(exc), parent=top_level)
                        return
                    cargar_administradores()

                ctk.CTkButton(
                    fila,
                    text="Protegido" if protegido else "Eliminar",
                    width=110,
                    state="disabled" if protegido else "normal",
                    fg_color="#64748b" if protegido else "#b91c1c",
                    hover_color="#991b1b",
                    command=eliminar,
                ).pack(side="right")

        cargar_administradores()
        ctk.CTkButton(
            frame,
            text="Cerrar",
            width=120,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(anchor="e", padx=22, pady=(0, 16))


    def abrir_limpiar(self, event=None):
        sipp = self.frameprincipal.winfo_toplevel()
        if not self._es_administrador(sipp):
            messagebox.showerror("Acceso denegado", "Solo un administrador puede limpiar la base de datos.", parent=sipp)
            return "break"
        top_level = ctk.CTkToplevel(self.frameprincipal.winfo_toplevel())
        top_level.title("Limpiar")
        self.ajustar_toplevel_a_pantalla(top_level, 560, 300, margen_x=160, margen_y=160)
        top_level.transient(self.frameprincipal.winfo_toplevel())
        top_level.focus_force()

        contenedor = ctk.CTkFrame(top_level, corner_radius=16, border_width=1)
        contenedor.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            contenedor,
            text="Limpiar datos",
            font=ctk.CTkFont(size=24, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(22, 8))
        ctk.CTkLabel(
            contenedor,
            text="Elimina empleados, planillas, liquidaciones y demás datos operativos.",
            font=ctk.CTkFont(size=12),
            anchor="w",
            wraplength=460,
        ).pack(fill="x", padx=20, pady=(0, 6))
        ctk.CTkLabel(
            contenedor,
            text="Se conservará el usuario de acceso y la configuración del sistema.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8"),
            anchor="w",
            wraplength=460,
        ).pack(fill="x", padx=20, pady=(0, 18))

        estado_limpieza = ctk.CTkLabel(contenedor, text="", anchor="w")
        estado_limpieza.pack(fill="x", padx=20, pady=(0, 10))
        progreso_limpieza = ctk.CTkProgressBar(contenedor, mode="indeterminate", height=8)
        progreso_limpieza.pack(fill="x", padx=20, pady=(0, 10))
        progreso_limpieza.stop()
        acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
        acciones.pack(fill="x", padx=20, pady=(0, 16))
        boton_limpiar = None

        def limpiar_base_datos():
            confirmado = messagebox.askyesno(
                "Confirmar limpieza",
                "Se eliminarán todos los datos operativos. Esta acción no se puede deshacer.\n\n¿Deseas continuar?",
                parent=top_level,
            )
            if not confirmado:
                return
            confirmado_de_nuevo = messagebox.askyesno(
                "Confirmación final",
                "Confirma nuevamente que deseas dejar la base de datos vacía.",
                parent=top_level,
            )
            if not confirmado_de_nuevo:
                return
            progreso_limpieza.start()
            boton_limpiar.configure(state="disabled")

            def limpieza_terminada(resultado):
                progreso_limpieza.stop()
                boton_limpiar.configure(state="normal")
                estado_limpieza.configure(text=f"Limpieza completada: {resultado['cantidad_tablas']} tablas vaciadas.")
                messagebox.showinfo("Limpiar", "La base de datos quedó lista para volver a usarse.", parent=top_level)

            def limpieza_fallida(exc):
                progreso_limpieza.stop()
                boton_limpiar.configure(state="normal")
                logging.exception("No se pudieron limpiar los datos operativos")
                estado_limpieza.configure(text="No se pudo completar la limpieza.")
                messagebox.showerror("Limpiar", f"No se pudo limpiar la base de datos:\n{exc}", parent=top_level)

            def ejecutar_limpieza():
                if respaldo_sipp is None:
                    raise RuntimeError("No está disponible el módulo de respaldos.")
                ruta_respaldo = respaldo_sipp.crear_copia_seguridad()
                cantidad_tablas = db.limpiar_datos_operativos() if db else 0
                usuario = (getattr(sipp, "usuario_autenticado", None) or {}).get("nombre", "desconocido")
                db.ConexionDB_datos_usuarios().registrar_auditoria(
                    usuario, "LIMPIEZA_DATOS", f"Respaldo previo: {os.path.basename(ruta_respaldo)}"
                )
                return {"cantidad_tablas": cantidad_tablas}

            ejecutar_en_segundo_plano(
                ejecutar_limpieza,
                lambda resultado: top_level.after(0, lambda: limpieza_terminada(resultado)),
                lambda exc: top_level.after(0, lambda: limpieza_fallida(exc)),
            )

        boton_limpiar = ctk.CTkButton(
            acciones,
            text="Limpiar base de datos",
            width=190,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=limpiar_base_datos,
        )
        boton_limpiar.pack(side="left")
        ctk.CTkButton(
            acciones,
            text="Cerrar",
            width=110,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(side="right")
        return "break"

    def abrir_log(self, event=None):
        top_level = ctk.CTkToplevel(self.frameprincipal.winfo_toplevel())
        top_level.title("Log")
        self.ajustar_toplevel_a_pantalla(top_level, 900, 600, margen_x=120, margen_y=120)
        top_level.transient(self.frameprincipal.winfo_toplevel())

        contenedor = ctk.CTkFrame(top_level, corner_radius=16, border_width=1)
        contenedor.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            contenedor,
            text="Registro de actividad",
            font=ctk.CTkFont(size=24, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(18, 2))
        ctk.CTkLabel(
            contenedor,
            text="Los eventos se actualizan automáticamente y están agrupados por día.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 12))

        area_log = ctk.CTkFrame(contenedor, fg_color="transparent")
        area_log.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        area_log.grid_rowconfigure(0, weight=1)
        area_log.grid_columnconfigure(0, weight=1)
        visor_log = tk.Text(
            area_log,
            wrap="none",
            state="disabled",
            font=("Consolas", 10),
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="#e2e8f0",
        )
        visor_log.grid(row=0, column=0, sticky="nsew")
        scroll_vertical = ttk.Scrollbar(area_log, orient="vertical", command=visor_log.yview)
        scroll_vertical.grid(row=0, column=1, sticky="ns")
        scroll_horizontal = ttk.Scrollbar(area_log, orient="horizontal", command=visor_log.xview)
        scroll_horizontal.grid(row=1, column=0, sticky="ew")
        visor_log.configure(
            yscrollcommand=scroll_vertical.set,
            xscrollcommand=scroll_horizontal.set,
        )

        ruta_log = os.path.join(os.path.abspath("."), "sipp.log")
        ultimo_contenido = {"valor": None}

        def contenido_por_dia(contenido):
            grupos = []
            indice_dia = {}
            for linea in contenido.splitlines():
                dia = linea[:10] if len(linea) >= 10 and linea[4] == "-" and linea[7] == "-" else "Sin fecha"
                if dia not in indice_dia:
                    indice_dia[dia] = len(grupos)
                    grupos.append([dia, []])
                grupos[indice_dia[dia]][1].append(linea)
            return "\n\n".join(
                f"===== {dia} =====\n" + "\n".join(lineas)
                for dia, lineas in grupos
            )

        def actualizar_log():
            if not top_level.winfo_exists():
                return
            try:
                with open(ruta_log, "r", encoding="utf-8", errors="replace") as archivo:
                    contenido = archivo.read()
                if contenido != ultimo_contenido["valor"]:
                    visor_log.configure(state="normal")
                    visor_log.delete("1.0", tk.END)
                    visor_log.insert("1.0", contenido_por_dia(contenido) or "No hay eventos registrados.")
                    visor_log.configure(state="disabled")
                    visor_log.see(tk.END)
                    ultimo_contenido["valor"] = contenido
            except OSError as exc:
                visor_log.configure(state="normal")
                visor_log.delete("1.0", tk.END)
                visor_log.insert("1.0", f"No se pudo leer el archivo de log:\n{exc}")
                visor_log.configure(state="disabled")
            top_level.after(1000, actualizar_log)

        acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
        acciones.pack(fill="x", padx=20, pady=(0, 14))
        ctk.CTkButton(acciones, text="Actualizar", width=120, command=actualizar_log).pack(side="left")
        ctk.CTkButton(acciones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        actualizar_log()
        top_level.focus_force()
        return "break"

    def generar_notificacion_asistencia(self, sipp):
        if not getattr(sipp, "notificaciones_activas", True) or not getattr(sipp, "avisos_marcaciones", True):
            return
        if getattr(sipp, "_notificacion_asistencia_fecha", None) == datetime.now().date():
            return

        hoy = datetime.now().date()
        if hoy.weekday() >= 5 or hoy in feriados_panama(hoy.year):
            return
        horarios = getattr(sipp, "horarios_laborales", [])
        if not horarios:
            return
        horario = horarios[0]
        try:
            hora_inicio = datetime.strptime(str(horario[0]).strip().upper(), "%I:%M %p")
        except ValueError:
            try:
                hora_inicio = datetime.strptime(str(horario[0]).strip(), "%H:%M")
            except ValueError:
                return
        tolerancia = 0
        if len(horario) >= 4:
            try:
                tolerancia = max(0, int(horario[3]))
            except (TypeError, ValueError):
                pass
        limite_entrada = hora_inicio.replace(year=hoy.year, month=hoy.month, day=hoy.day) + timedelta(minutes=tolerancia)
        if datetime.now() <= limite_entrada:
            return

        try:
            corte = db.cortes_planilla().obtener_corte_activo()
            if not corte:
                return
            empleados = db.gestion_empleado().obtener_empleados()
            marcaciones = db.marcaciones().obtener_marcaciones_por_rango(hoy, hoy)
            marcaciones_por_dui = {}
            for marcacion in marcaciones:
                dui = str(marcacion[2] or "").strip()
                marcaciones_por_dui[dui] = marcacion

            tardanzas = []
            ausencias = []
            for empleado in empleados:
                if len(empleado) <= 2:
                    continue
                nombre = str(empleado[1] or "Sin nombre").strip()
                dui = str(empleado[2] or "").strip()
                marcacion = marcaciones_por_dui.get(dui)
                if not marcacion:
                    ausencias.append(f"- {nombre} ({dui})")
                    continue
                estado = str(marcacion[6] if len(marcacion) > 6 else "JORNADA LABORAL").strip().upper()
                if estado in {"AUSENCIA", "DÍA LIBRE", "DIA LIBRE"}:
                    continue
                try:
                    entrada = None
                    texto_entrada = str(marcacion[3] or "").strip().upper()
                    for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                        try:
                            entrada = datetime.strptime(texto_entrada, formato)
                            break
                        except ValueError:
                            pass
                    entrada_hoy = datetime.combine(hoy, entrada.time()) if entrada else None
                    if entrada_hoy and entrada_hoy > limite_entrada:
                        minutos = int((entrada_hoy - limite_entrada).total_seconds() // 60)
                        tardanzas.append(f"- {nombre} ({minutos} min)")
                except (TypeError, ValueError):
                    logging.exception("No se pudo revisar la marcación de %s", dui)

            sipp._notificacion_asistencia_fecha = hoy
            mensajes = []
            if tardanzas:
                mensajes.append("Llegaron tarde:\n" + "\n".join(tardanzas[:15]))
            if ausencias:
                mensajes.append("No tienen marcación:\n" + "\n".join(ausencias[:15]))
            if mensajes:
                extra = "\n\nHay más registros." if len(tardanzas) + len(ausencias) > 15 else ""
                return "\n\n".join(mensajes) + extra
        except Exception as exc:
            logging.exception("No se pudieron generar notificaciones de asistencia")

        return None

    def revisar_notificaciones_asistencia(self, sipp):
        def mostrar_aviso(mensaje):
            if mensaje:
                messagebox.showwarning("Aviso de asistencia", mensaje, parent=sipp)

        ejecutar_en_segundo_plano(
            lambda: self.generar_notificacion_asistencia(sipp),
            lambda resultado: sipp.after(0, lambda: mostrar_aviso(resultado)),
            lambda exc: logging.exception("No se pudieron generar notificaciones de asistencia", exc_info=exc),
        )
        

        
    def menu(self, sipp):

        menu = Menu(sipp)
        sipp.config(menu=menu)  
        menu_archivo = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("Archivo"), menu=menu_archivo)
        #menu_archivo.add_command(label="Reiniciar", font=ctk.CTkFont(size=16), command=lambda: self.plataformas(sipp))
        menu_archivo.add_command(label=self.t("Usuario"), font=ctk.CTkFont(size=12), command=lambda: self.idenficar_usuario(sipp))
        menu_archivo.add_command(label=self.t("Crear Usuario"), font=ctk.CTkFont(size=12), command=lambda: self.crear_usuarios(sipp))
        menu_archivo.add_command(label=self.t("Crear Administrador"), font=ctk.CTkFont(size=12), command=lambda: self.crear_administrador(sipp))
        menu_archivo.add_separator()
        menu_archivo.add_command(label=self.t("Salir"), font=ctk.CTkFont(size=12), command=sipp.quit)
        
        #edicciones 
        menu_editar = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("editar"), menu=menu_editar)
        menu_editar.add_command(label=self.t("Experiencia"), font=ctk.CTkFont(size=12),
                                command=lambda : self.editar_apariencia(sipp))
        menu_editar.add_command(label=self.t("Temas"), font=ctk.CTkFont(size=12), 
                                command= lambda : self.editar_tema(sipp))

        menu_configuracion = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("Configuracion"), menu=menu_configuracion)
        menu_configuracion.add_command(
            label=self.t("Datos de la empresa"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configuracion_nueva(sipp),
        )
        menu_configuracion.add_command(
            label=self.t("Idioma"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_idioma(sipp),
        )
        menu_configuracion.add_command(
            label=self.t("API Lector"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_app(sipp),
        )
        menu_configuracion.add_command(
            label=self.t("Notificaciones"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_notificaciones(sipp),
        )
        menu_configuracion.add_command(
            label=self.t("Tamaño de texto y opciones visuales"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_opciones_visuales(sipp),
        )
        menu_configuracion.add_command(
            label=self.t("Horarios"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_horarios(sipp),
        )

        menu_seguridad = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("Seguridad"), menu=menu_seguridad)
        menu_seguridad.add_command(
            label=self.t("Revisión"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.revisar_aplicacion(sipp),
        )
        menu_seguridad.add_command(
            label=self.t("Respaldo de seguridad"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.configurar_copia_seguridad(sipp),
        )

        menu_reportes = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("Reportes"), menu=menu_reportes)
        menu_reportes.add_command(
            label=self.t("Tardanzas y Ausencias"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.reporte_tardanzas_ausencias(sipp),
        )

        menu_ayuda = Menu(menu, tearoff=0)
        menu.add_cascade(label=self.t("Ayuda"), menu=menu_ayuda)
        menu_ayuda.add_command(
            label=self.t("Información"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.mostrar_informacion_aplicacion(sipp),
        )
        menu_ayuda.add_command(
            label=self.t("Buscar actualizaciones"),
            font=ctk.CTkFont(size=12),
            command=lambda: self.actualizar_aplicacion(sipp),
        )


    def actualizar_aplicacion(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 650, 430)
        top_level.title("Actualizaciones de SiPP")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Actualizaciones", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=22, pady=(22, 5))
        ctk.CTkLabel(frame, text=f"Versión instalada: {VERSION_APLICACION}", font=ctk.CTkFont(size=12), anchor="w").pack(fill="x", padx=22, pady=(0, 14))

        estado = ctk.CTkLabel(frame, text="La aplicación consultará automáticamente el Release oficial de GitHub.", anchor="w", justify="left", wraplength=570, text_color=("#526071", "#cbd5e1"))
        estado.pack(fill="x", padx=22, pady=(16, 12))
        progreso_actualizacion = ctk.CTkProgressBar(frame, mode="indeterminate", height=8)
        progreso_actualizacion.pack(fill="x", padx=22, pady=(0, 12))
        progreso_actualizacion.stop()

        def consultar_actualizacion():
            boton_buscar.configure(state="disabled")
            progreso_actualizacion.start()
            estado.configure(text="Buscando actualizaciones...")

            def consultar_y_descargar():
                solicitud = Request(
                    URL_RELEASES_GITHUB,
                    headers={
                        "Accept": "application/vnd.github+json",
                        "User-Agent": "SiPP-Updater",
                    },
                )
                with urlopen(solicitud, timeout=10) as respuesta:
                    release = json.loads(respuesta.read().decode("utf-8"))
                etiqueta = str(release.get("tag_name", "")).strip()
                version_nueva = etiqueta.removeprefix("v")
                notas = str(release.get("body", "Sin notas de versión.")).strip() or "Sin notas de versión."
                activos = release.get("assets", [])
                nombre_instalador = f"SiPP-Setup-{etiqueta}.exe"
                activo = next(
                    (elemento for elemento in activos if elemento.get("name") == nombre_instalador),
                    None,
                )
                if not version_nueva or not activo:
                    raise ValueError("El Release no contiene un instalador válido de SiPP.")

                def version_tuple(version):
                    return tuple(int(parte) for parte in version.split(".")[:3])

                if version_tuple(version_nueva) <= version_tuple(VERSION_APLICACION):
                    return {"estado": f"Ya tienes la versión {VERSION_APLICACION}.\n\n{notas}"}
                return {
                    "estado": "disponible",
                    "version": version_nueva,
                    "notas": notas,
                    "url": activo["browser_download_url"],
                    "nombre": nombre_instalador,
                }

            def consulta_terminada(resultado):
                if resultado.get("estado") != "disponible":
                    progreso_actualizacion.stop()
                    boton_buscar.configure(state="normal")
                    estado.configure(text=resultado["estado"])
                    return
                confirmar = messagebox.askyesno(
                    "Actualización disponible",
                    f"Nueva versión: {resultado['version']}\n\n{resultado['notas']}\n\n¿Deseas descargarla e instalarla?",
                    parent=top_level,
                )
                if not confirmar:
                    progreso_actualizacion.stop()
                    boton_buscar.configure(state="normal")
                    estado.configure(text="Actualización cancelada.")
                    return
                estado.configure(text="Descargando actualización...")
                ejecutar_en_segundo_plano(
                    lambda: urlretrieve(resultado["url"], os.path.join(tempfile.gettempdir(), resultado["nombre"])),
                    lambda ruta: top_level.after(0, lambda: descarga_terminada(ruta)),
                    lambda exc: top_level.after(0, lambda: descarga_fallida(exc)),
                )

            def descarga_terminada(resultado):
                nombre_temporal = resultado[0] if isinstance(resultado, tuple) else resultado
                progreso_actualizacion.stop()
                estado.configure(
                    text=(
                        "Actualización descargada. SiPP se cerrará para iniciar el instalador.\n\n"
                        "Si Windows muestra SmartScreen, pulsa «Más información» y después "
                        "«Ejecutar de todas formas» únicamente si verificaste que el archivo "
                        "proviene del Release oficial de GitHub."
                    )
                )
                top_level.after(800, lambda: (os.startfile(nombre_temporal), sipp.destroy()))

            def descarga_fallida(exc):
                progreso_actualizacion.stop()
                boton_buscar.configure(state="normal")
                estado.configure(text=f"No se pudo descargar la actualización:\n{exc}")

            def consulta_fallida(exc):
                progreso_actualizacion.stop()
                boton_buscar.configure(state="normal")
                estado.configure(text=f"No se pudo consultar la actualización:\n{exc}")

            ejecutar_en_segundo_plano(
                consultar_y_descargar,
                lambda resultado: top_level.after(0, lambda: consulta_terminada(resultado)),
                lambda exc: top_level.after(0, lambda: consulta_fallida(exc)),
            )

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(fill="x", padx=22, pady=(0, 20))
        boton_buscar = ctk.CTkButton(botones, text="Buscar actualización", width=160, fg_color="#2563eb", hover_color="#1d4ed8", command=consultar_actualizacion)
        boton_buscar.pack(side="left")
        ctk.CTkButton(botones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")


    def mostrar_informacion_aplicacion(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 620, 420)
        top_level.title("Información de la aplicación")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()

        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text="Información de la aplicación",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=22, pady=(22, 16))

        informacion = (
            "SiPP es una aplicación para la gestión integral del personal y los procesos laborales de la empresa.\n\n"
            "Permite administrar empleados, marcaciones, horarios, planillas, deducciones, reportes, vacaciones y liquidaciones desde un solo lugar.\n\n"
            "Su finalidad es organizar la información laboral, facilitar el control de asistencia y apoyar la gestión administrativa de la empresa.\n\n"
            "Creada por: Dariel Fenton A.A\n"
            "Correo: darielfenton0@gmail.com"
        )
        contenido = tk.Text(
            frame,
            wrap="word",
            height=11,
            font=("Segoe UI", 11),
            bg="#f8fafc",
            fg="#1e293b",
            relief="flat",
            padx=14,
            pady=12,
        )
        contenido.insert("1.0", informacion)
        contenido.configure(state="disabled")
        contenido.pack(fill="both", expand=True, padx=22, pady=(0, 16))

        ctk.CTkButton(
            frame,
            text="Cerrar",
            width=120,
            height=34,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(anchor="e", padx=22, pady=(0, 18))


    def reporte_tardanzas_ausencias(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 820, 560)
        top_level.title("Tardanzas y Ausencias")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Tardanzas y Ausencias", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 4))
        ctk.CTkLabel(frame, text="Resumen del corte activo según las marcaciones registradas.", font=ctk.CTkFont(size=12), anchor="w").pack(fill="x", padx=20, pady=(0, 14))

        resumen = ctk.CTkFrame(frame, fg_color="transparent")
        resumen.pack(fill="x", padx=20, pady=(0, 12))
        for columna in range(3):
            resumen.grid_columnconfigure(columna, weight=1)
        etiqueta_tardanzas = ctk.CTkLabel(resumen, text="Tardanzas\n0", height=62, corner_radius=10, fg_color=("#fef3c7", "#422006"), text_color=("#92400e", "#fde68a"), font=ctk.CTkFont(size=14, weight="bold"))
        etiqueta_tardanzas.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        etiqueta_ausencias = ctk.CTkLabel(resumen, text="Ausencias\n0", height=62, corner_radius=10, fg_color=("#fee2e2", "#450a0a"), text_color=("#b91c1c", "#fecaca"), font=ctk.CTkFont(size=14, weight="bold"))
        etiqueta_ausencias.grid(row=0, column=1, padx=8, sticky="ew")
        etiqueta_corte = ctk.CTkLabel(resumen, text="Corte\nSin corte activo", height=62, corner_radius=10, fg_color=("#e0f2fe", "#082f49"), text_color=("#0369a1", "#bae6fd"), font=ctk.CTkFont(size=12, weight="bold"))
        etiqueta_corte.grid(row=0, column=2, padx=(8, 0), sticky="ew")

        tabla = ctk.CTkFrame(frame, fg_color="transparent")
        tabla.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        tabla.grid_rowconfigure(0, weight=1)
        tabla.grid_columnconfigure(0, weight=1)
        detalle = ttk.Treeview(tabla, columns=("Nombre", "DUI", "Fecha", "Tipo", "Entrada", "Salida", "Tardanza"), show="headings", height=13)
        for columna, ancho in (("Nombre", 170), ("DUI", 105), ("Fecha", 95), ("Tipo", 120), ("Entrada", 100), ("Salida", 100), ("Tardanza", 85)):
            detalle.heading(columna, text=columna)
            detalle.column(columna, width=ancho, anchor="center")
        detalle.grid(row=0, column=0, sticky="nsew")
        scroll_detalle = ttk.Scrollbar(tabla, orient="vertical", command=detalle.yview)
        scroll_detalle.grid(row=0, column=1, sticky="ns")
        detalle.configure(yscrollcommand=scroll_detalle.set)

        def convertir_hora(valor):
            texto = str(valor or "").strip().upper()
            for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                try:
                    return datetime.strptime(texto, formato)
                except ValueError:
                    continue
            return None

        def cargar_reporte():
            for item in detalle.get_children():
                detalle.delete(item)
            tardanzas = 0
            ausencias = 0
            corte = db.cortes_planilla().obtener_corte_activo()
            if not corte:
                etiqueta_tardanzas.configure(text="Tardanzas\n0")
                etiqueta_ausencias.configure(text="Ausencias\n0")
                etiqueta_corte.configure(text="Corte\nSin corte activo")
                return

            etiqueta_corte.configure(text=f"Corte\n#{corte[0]}: {corte[1]} a {corte[2]}")
            horario_configurado = getattr(self, "horarios_laborales", [])
            hora_inicio = "08:00 AM"
            tolerancia_minutos = 0
            if horario_configurado and len(horario_configurado[0]) >= 1:
                primer_horario = horario_configurado[0]
                hora_inicio = primer_horario[0]
                if len(primer_horario) >= 4:
                    try:
                        tolerancia_minutos = max(0, int(primer_horario[3]))
                    except (TypeError, ValueError):
                        tolerancia_minutos = 0
            inicio_programado = convertir_hora(hora_inicio) or datetime.strptime("08:00 AM", "%I:%M %p")
            inicio_con_tolerancia = inicio_programado + timedelta(minutes=tolerancia_minutos)

            try:
                marcaciones = db.marcaciones().obtener_marcaciones_por_rango(corte[1], corte[2])
                empleados = db.gestion_empleado().obtener_empleados()
                marcaciones_por_persona_fecha = set()
                for marcacion in marcaciones:
                    estado_dia = str(marcacion[6] if len(marcacion) > 6 else "JORNADA LABORAL").strip().upper()
                    if estado_dia in {"DÍA LIBRE", "DIA LIBRE"}:
                        continue
                    fecha = marcacion[5]
                    fecha_mostrar = fecha.strftime("%d/%m/%Y") if hasattr(fecha, "strftime") else str(fecha)
                    fecha_clave = fecha.date() if hasattr(fecha, "date") else fecha
                    if isinstance(fecha_clave, str):
                        fecha_clave = datetime.strptime(fecha_clave[:10], "%Y-%m-%d").date()
                    dui = str(marcacion[2] or "").strip()
                    marcaciones_por_persona_fecha.add((dui, fecha_clave))
                    tipo = "Ausencia" if estado_dia == "AUSENCIA" else ""
                    if estado_dia == "AUSENCIA":
                        ausencias += 1
                    else:
                        entrada = convertir_hora(marcacion[3])
                        minutos_tardanza = max(0, int((entrada - inicio_con_tolerancia).total_seconds() // 60)) if entrada else 0
                        if entrada and entrada > inicio_con_tolerancia:
                            tardanzas += 1
                            tipo = "Tardanza"
                        else:
                            minutos_tardanza = 0
                    if tipo:
                        detalle.insert("", "end", values=(marcacion[1], marcacion[2], fecha_mostrar, tipo, marcacion[3], marcacion[4], f"{minutos_tardanza} min" if minutos_tardanza else "-"))
                fecha_inicio = corte[1].date() if hasattr(corte[1], "date") else corte[1]
                fecha_fin = corte[2].date() if hasattr(corte[2], "date") else corte[2]
                if isinstance(fecha_inicio, str):
                    fecha_inicio = datetime.strptime(fecha_inicio[:10], "%Y-%m-%d").date()
                if isinstance(fecha_fin, str):
                    fecha_fin = datetime.strptime(fecha_fin[:10], "%Y-%m-%d").date()
                for empleado in empleados:
                    if len(empleado) <= 2:
                        continue
                    dui = str(empleado[2] or "").strip()
                    fecha = fecha_inicio
                    while fecha <= fecha_fin:
                        if fecha.weekday() < 5 and fecha not in feriados_panama(fecha.year) and (dui, fecha) not in marcaciones_por_persona_fecha:
                            ausencias += 1
                            detalle.insert("", "end", values=(empleado[1], dui, fecha.strftime("%d/%m/%Y"), "Ausencia", "-", "-", "-"))
                        fecha += timedelta(days=1)
            except Exception as exc:
                logging.exception("No se pudo generar el reporte de tardanzas y ausencias")
                messagebox.showerror("Reporte", f"No se pudo generar el reporte: {exc}")

            etiqueta_tardanzas.configure(text=f"Tardanzas\n{tardanzas}")
            etiqueta_ausencias.configure(text=f"Ausencias\n{ausencias}")

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkButton(botones, text="Actualizar reporte", width=150, fg_color="#2563eb", hover_color="#1d4ed8", command=cargar_reporte).pack(side="left")
        ctk.CTkButton(botones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        cargar_reporte()


    def configurar_horarios(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(top_level, 760, 600, margen_x=120, margen_y=120, min_ancho=520, min_alto=420)
        top_level.title("Horarios laborales")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text="Horarios laborales de la empresa",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(22, 6))
        ctk.CTkLabel(
            frame,
            text="Crea los horarios laborales que se utilizarán en la empresa.",
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=20, pady=(0, 18))

        formulario = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#f1f5f9", "#1f2937"))
        formulario.pack(fill="x", padx=20, pady=(0, 12))
        formulario.grid_columnconfigure(1, weight=1)
        formulario.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(formulario, text="Entrada:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="w")
        entrada_horario = ctk.CTkEntry(formulario, placeholder_text="08:00 AM", height=32)
        entrada_horario.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        ctk.CTkLabel(formulario, text="Salida:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(16, 8), pady=12, sticky="w")
        salida_horario = ctk.CTkEntry(formulario, placeholder_text="04:00 PM", height=32)
        salida_horario.grid(row=0, column=3, padx=(8, 12), pady=12, sticky="ew")

        ctk.CTkLabel(formulario, text="Jornada:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=0, padx=(12, 8), pady=8, sticky="w")
        tipo_horario = ctk.CTkComboBox(formulario, values=["DIURNA", "NOCTURNA", "MIXTA"], state="readonly", height=32)
        tipo_horario.set("DIURNA")
        tipo_horario.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkLabel(formulario, text="Tolerancia (min):", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=2, padx=(16, 8), pady=8, sticky="w")
        tolerancia_horario = ctk.CTkEntry(formulario, placeholder_text="0", height=32)
        tolerancia_horario.insert(0, "0")
        tolerancia_horario.grid(row=1, column=3, padx=(8, 12), pady=8, sticky="ew")

        tabla = ctk.CTkFrame(frame, fg_color="transparent")
        tabla.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        tabla.grid_rowconfigure(0, weight=1)
        tabla.grid_columnconfigure(0, weight=1)
        tabla_horarios = ttk.Treeview(tabla, columns=("Entrada", "Salida", "Jornada", "Tolerancia"), show="headings", height=10)
        for columna in ("Entrada", "Salida", "Jornada", "Tolerancia"):
            tabla_horarios.heading(columna, text=columna)
            tabla_horarios.column(columna, anchor="center", width=150)
        tabla_horarios.grid(row=0, column=0, sticky="nsew")
        scroll_horarios = ttk.Scrollbar(tabla, orient="vertical", command=tabla_horarios.yview)
        scroll_horarios.grid(row=0, column=1, sticky="ns")
        tabla_horarios.configure(yscrollcommand=scroll_horarios.set)

        horarios_iniciales = getattr(self, "horarios_laborales", [])
        for horario in horarios_iniciales:
            valores_horario = tuple(horario[:4]) if len(horario) >= 4 else tuple(horario)
            tabla_horarios.insert("", "end", values=valores_horario)

        def _validar_hora(valor, formato):
            try:
                datetime.strptime(valor.upper(), formato)
                return True
            except ValueError:
                return False

        def agregar_horario():
            entrada = entrada_horario.get().strip()
            salida = salida_horario.get().strip()
            tipo = tipo_horario.get().strip().upper()
            tolerancia = tolerancia_horario.get().strip()
            formatos_horarios = ("%I:%M %p", "%H:%M")
            try:
                if not any(_validar_hora(entrada, formato) for formato in formatos_horarios):
                    raise ValueError
                if not any(_validar_hora(salida, formato) for formato in formatos_horarios):
                    raise ValueError
                if int(tolerancia) < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Horario", "Use formatos como 08:00 AM, 04:00 PM o 08:00.")
                return
            tabla_horarios.insert("", "end", values=(entrada, salida, tipo, tolerancia))

        def eliminar_horario():
            seleccionados = tabla_horarios.selection()
            for item in seleccionados:
                tabla_horarios.delete(item)

        acciones_horario = ctk.CTkFrame(frame, fg_color="transparent")
        acciones_horario.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkButton(acciones_horario, text="Agregar horario", width=140, fg_color="#0f766e", hover_color="#115e59", command=agregar_horario).pack(side="left")
        ctk.CTkButton(acciones_horario, text="Eliminar seleccionado", width=160, fg_color="#dc2626", hover_color="#b91c1c", command=eliminar_horario).pack(side="left", padx=(10, 0))

        ctk.CTkLabel(
            frame,
            text="Los horarios guardados aquí serán los que se utilizarán posteriormente en la empresa.",
            font=ctk.CTkFont(size=11),
            text_color=("#64748b", "#cbd5e1"),
            anchor="w",
        ).pack(fill="x", padx=20)
        botones_horarios = ctk.CTkFrame(frame, fg_color="transparent")
        botones_horarios.pack(fill="x", padx=20, pady=12)
        def guardar_horarios():
            self.horarios_laborales = [
                tabla_horarios.item(item, "values")
                for item in tabla_horarios.get_children()
            ]
            top_level.destroy()

        ctk.CTkButton(
            botones_horarios,
            text="Guardar horarios",
            width=140,
            height=34,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=guardar_horarios,
        ).pack(side="right", padx=(10, 10))
        ctk.CTkButton(
            botones_horarios,
            text="Cerrar",
            width=120,
            height=34,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(side="right")

    def revisar_aplicacion(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        top_level.withdraw()
        self.ajustar_toplevel_a_pantalla(top_level, 820, 560)
        top_level.title("Revisión de la aplicación")
        top_level.transient(sipp)
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Revisión de la aplicación", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 6))
        ctk.CTkLabel(frame, text="Comprueba la sintaxis de los archivos Python sin ejecutar la aplicación.", font=ctk.CTkFont(size=12), anchor="w").pack(fill="x", padx=20, pady=(0, 12))

        estado_revision = ctk.CTkLabel(frame, text="Preparando revisión...", anchor="w")
        estado_revision.pack(fill="x", padx=20, pady=(0, 4))
        progreso_revision = ctk.CTkProgressBar(frame, mode="indeterminate", height=8)
        progreso_revision.pack(fill="x", padx=20, pady=(0, 12))
        progreso_revision.set(0)

        area = ctk.CTkFrame(frame, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        area.grid_rowconfigure(0, weight=1)
        area.grid_columnconfigure(0, weight=1)
        informe = tk.Text(area, wrap="word", state="disabled", font=("Consolas", 10), bg="#0f172a", fg="#e2e8f0", insertbackground="#e2e8f0")
        informe.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(area, orient="vertical", command=informe.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        informe.configure(yscrollcommand=scroll.set)

        ultimo_informe_revision = ""
        revision_en_curso = False
        resultados_revision = queue.Queue()

        def generar_informe_revision(actualizar_estado):
            raiz = os.path.abspath(os.path.dirname(__file__))
            excluidas = {
                ".git",
                ".hg",
                ".svn",
                "__pycache__",
                "installer",
                "site-packages",
                "node_modules",
            }
            archivos = []
            errores = []
            import_problemas = []
            duplicados = []
            check_db = {"ok": False, "mensaje": "No verificado"}

            def es_directorio_excluido(nombre):
                nombre_normalizado = nombre.casefold()
                return (
                    nombre_normalizado in excluidas
                    or nombre_normalizado in {"venv", ".venv"}
                    or nombre_normalizado.startswith(".venv-")
                    or nombre_normalizado.startswith("venv-")
            )

            try:
                import ast
                import importlib.util
            except Exception as exc:
                errores.append(("__internals__", exc))

            for directorio, nombres, archivos_en_directorio in os.walk(raiz):
                nombres[:] = [nombre for nombre in nombres if not es_directorio_excluido(nombre)]
                for nombre_archivo in archivos_en_directorio:
                    if nombre_archivo.endswith(".py"):
                        archivos.append(os.path.join(directorio, nombre_archivo))

            actualizar_estado(f"Analizando archivos Python (0 de {len(archivos)})...")

            sys_path_original = list(sys.path)
            sys.path.insert(0, raiz)
            try:
                archivos_ordenados = sorted(archivos)
                for indice, archivo in enumerate(archivos_ordenados, start=1):
                    actualizar_estado(f"Analizando archivos Python ({indice} de {len(archivos_ordenados)})...")
                    try:
                        with open(archivo, "rb") as contenido:
                            fuente = contenido.read()
                        compile(fuente, archivo, "exec")
                        try:
                            arbol = ast.parse(fuente.decode("utf-8", errors="ignore"), filename=archivo)
                            nombres_por_bloque = {}
                            for nodo in ast.walk(arbol):
                                if isinstance(nodo, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                                    nombres_bloque = {}
                                    nombres_por_bloque[id(nodo)] = nombres_bloque
                                    for elemento in nodo.body:
                                        if isinstance(elemento, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                                            nombre = elemento.name
                                            anterior = nombres_bloque.get(nombre)
                                            if anterior is not None:
                                                duplicados.append(
                                                    (
                                                        os.path.relpath(archivo, raiz),
                                                        nombre,
                                                        getattr(anterior, "lineno", "?"),
                                                        getattr(elemento, "lineno", "?"),
                                                    )
                                                )
                                            else:
                                                nombres_bloque[nombre] = elemento
                            for nodo in ast.walk(arbol):
                                if isinstance(nodo, ast.Import):
                                    for alias in nodo.names:
                                        nombre = alias.name.split(".")[0]
                                        if nombre and nombre not in {"__future__"}:
                                            try:
                                                if importlib.util.find_spec(nombre) is None:
                                                    import_problemas.append((os.path.relpath(archivo, raiz), f"Import no encontrado: {nombre}"))
                                            except (ImportError, ValueError, AttributeError):
                                                import_problemas.append((os.path.relpath(archivo, raiz), f"Import no resoluble: {nombre}"))
                                elif isinstance(nodo, ast.ImportFrom):
                                    if nodo.level:
                                        continue
                                    modulo = (nodo.module or "").split(".")[0]
                                    if modulo and modulo not in {"typing", "math", "os", "sys", "json", "datetime", "tkinter", "customtkinter", "PIL", "openpyxl", "psycopg"}:
                                        try:
                                            if importlib.util.find_spec(modulo) is None:
                                                import_problemas.append((os.path.relpath(archivo, raiz), f"Módulo importado no encontrado: {modulo}"))
                                        except (ImportError, ValueError, AttributeError):
                                            import_problemas.append((os.path.relpath(archivo, raiz), f"Módulo importado no resoluble: {modulo}"))
                        except SyntaxError as exc:
                            errores.append((os.path.relpath(archivo, raiz), exc))
                    except SyntaxError as exc:
                        errores.append((os.path.relpath(archivo, raiz), exc))
                    except (OSError, UnicodeError) as exc:
                        errores.append((os.path.relpath(archivo, raiz), exc))

                actualizar_estado("Verificando conexión con la base de datos...")
                try:
                    import psycopg
                    conexion_revision = psycopg.connect(
                        host="localhost",
                        user="postgres",
                        password="12sql@?",
                        port="5432",
                        dbname="SiPP-2",
                        connect_timeout=3,
                        options="-c statement_timeout=5000 -c lock_timeout=2000",
                    )
                    try:
                        with conexion_revision.cursor() as cursor:
                            cursor.execute("SELECT 1")
                        check_db["ok"] = True
                        check_db["mensaje"] = "Conexión a la base de datos: OK"
                    finally:
                        conexion_revision.close()
                except Exception as exc:
                    check_db["mensaje"] = f"Conexión a la base de datos: FALLA ({exc})"
            finally:
                sys.path[:] = sys_path_original

            lineas = [
                "INFORME DE REVISIÓN DE LA APLICACIÓN",
                "=" * 44,
                f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                f"Archivos Python revisados: {len(archivos)}",
                f"Errores de sintaxis: {len(errores)}",
                f"Problemas de importación: {len(import_problemas)}",
                f"Declaraciones duplicadas: {len(duplicados)}",
                f"Estado de base de datos: {'OK' if check_db['ok'] else 'REVISION'}",
                "",
            ]
            lineas.append(check_db["mensaje"])
            lineas.append("")

            if errores or import_problemas or duplicados:
                lineas.append("DETALLE DE PROBLEMAS")
                lineas.append("-" * 22)
                for archivo, error in errores:
                    ruta_relativa = os.path.relpath(os.path.join(raiz, archivo), raiz) if not os.path.isabs(archivo) else os.path.relpath(archivo, raiz)
                    linea = getattr(error, "lineno", "?")
                    columna = getattr(error, "offset", "?")
                    detalle = getattr(error, "msg", str(error))
                    lineas.append(f"Archivo: {ruta_relativa}")
                    lineas.append(f"Línea: {linea} | Columna: {columna}")
                    lineas.append(f"Problema: {detalle}")
                    lineas.append("")
                for archivo, detalle in import_problemas:
                    lineas.append(f"Archivo: {archivo}")
                    lineas.append(f"Problema: {detalle}")
                    lineas.append("")
                for archivo, nombre, linea_anterior, linea_actual in duplicados:
                    lineas.append(f"Archivo: {archivo}")
                    lineas.append(
                        f"Declaración duplicada: {nombre} (líneas {linea_anterior} y {linea_actual})"
                    )
                    lineas.append("")
            else:
                lineas.append("RESULTADO: No se encontraron errores de sintaxis ni problemas de importación.")
                lineas.append("La revisión estática terminó correctamente.")

            return "\n".join(lineas)

        def mostrar_informe_revision(texto_informe):
            nonlocal ultimo_informe_revision
            ultimo_informe_revision = texto_informe
            informe.configure(state="normal")
            informe.delete("1.0", tk.END)
            informe.insert("1.0", texto_informe)
            informe.configure(state="disabled")

        def mostrar_error_revision(exc):
            mostrar_informe_revision(
                "INFORME DE REVISIÓN DE LA APLICACIÓN\n"
                "============================================\n\n"
                f"No se pudo completar la revisión:\n{exc}"
            )

        def ejecutar_revision():
            nonlocal revision_en_curso
            if revision_en_curso:
                return
            revision_en_curso = True
            progreso_revision.start()
            estado_revision.configure(text="Iniciando revisión...")
            informe.configure(state="normal")
            informe.delete("1.0", tk.END)
            informe.insert("1.0", "Revisión en curso...\nLa ventana seguirá disponible mientras termina.")
            informe.configure(state="disabled")

            def trabajador_revision():
                try:
                    def publicar_estado(mensaje):
                        resultados_revision.put(("estado", mensaje))

                    resultados_revision.put(("resultado", True, generar_informe_revision(publicar_estado)))
                except Exception as exc:
                    resultados_revision.put(("resultado", False, exc))

            threading.Thread(target=trabajador_revision, daemon=True).start()

            def revisar_resultado():
                nonlocal revision_en_curso
                try:
                    evento = resultados_revision.get_nowait()
                except queue.Empty:
                    if top_level.winfo_exists():
                        top_level.after(100, revisar_resultado)
                    return
                if evento[0] == "estado":
                    estado_revision.configure(text=evento[1])
                    if top_level.winfo_exists():
                        top_level.after(100, revisar_resultado)
                    return
                _, resultado_ok, resultado = evento
                revision_en_curso = False
                progreso_revision.stop()
                estado_revision.configure(text="Revisión finalizada.")
                if resultado_ok:
                    mostrar_informe_revision(resultado)
                else:
                    mostrar_error_revision(resultado)

            top_level.after(100, revisar_resultado)

        def guardar_informe_revision():
            if not ultimo_informe_revision:
                messagebox.showwarning("Revisión", "Primero ejecuta la revisión para generar el informe.", parent=top_level)
                return
            ruta = filedialog.asksaveasfilename(
                parent=top_level,
                title="Guardar informe de revisión",
                defaultextension=".txt",
                filetypes=[("Archivo de texto", "*.txt"), ("Todos los archivos", "*.*")],
                initialfile=f"reporte_revision_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            )
            if not ruta:
                return
            try:
                with open(ruta, "w", encoding="utf-8") as archivo:
                    archivo.write(ultimo_informe_revision)
                messagebox.showinfo("Revisión", f"Informe guardado en:\n{ruta}", parent=top_level)
            except OSError as exc:
                messagebox.showerror("Revisión", f"No se pudo guardar el archivo:\n{exc}", parent=top_level)

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkButton(botones, text="Ejecutar revisión", width=150, fg_color="#2563eb", hover_color="#1d4ed8", command=ejecutar_revision).pack(side="left")
        ctk.CTkButton(botones, text="Guardar informe", width=150, fg_color="#0f766e", hover_color="#115e59", command=guardar_informe_revision).pack(side="left", padx=(10, 0))
        ctk.CTkButton(botones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        top_level.update_idletasks()
        top_level.deiconify()
        top_level.lift()
        top_level.focus_force()
        top_level.grab_set()
        top_level.after(50, ejecutar_revision)

    def configurar_copia_seguridad(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 620, 560)
        top_level.title("Copia de seguridad")
        top_level.resizable(False, False)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Copia de seguridad", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(24, 12))
        ctk.CTkLabel(
            frame,
            text="Crea copias de la base de datos y los documentos, y restaura la última si algo sale mal.",
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(fill="x", padx=20, pady=(0, 10))

        lista_frame = ctk.CTkFrame(frame, fg_color=("#f1f5f9", "#1f2937"), corner_radius=10)
        lista_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        ctk.CTkLabel(lista_frame, text="Copias disponibles", font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(fill="x", padx=14, pady=(10, 4))
        selector_copias = ctk.CTkComboBox(lista_frame, values=["(sin copias)"], state="readonly")
        selector_copias.pack(fill="x", padx=14, pady=(0, 12))

        estado = ctk.CTkLabel(frame, text="", anchor="w", justify="left", wraplength=560, text_color=("#526071", "#cbd5e1"))
        estado.pack(fill="x", padx=20, pady=(0, 8))
        progreso_respaldo = ctk.CTkProgressBar(frame, mode="indeterminate", height=8)
        progreso_respaldo.pack(fill="x", padx=20, pady=(0, 10))
        progreso_respaldo.stop()

        def refrescar_lista_copias():
            if respaldo_sipp is None:
                return
            copias = respaldo_sipp.listar_copias()
            if not copias:
                selector_copias.configure(values=["(sin copias)"])
                selector_copias.set("(sin copias)")
                return
            nombres = [os.path.basename(ruta) for ruta in copias]
            selector_copias.configure(values=nombres)
            selector_copias.set(nombres[0])

        def ruta_copia_seleccionada():
            nombre = selector_copias.get()
            if not nombre or nombre == "(sin copias)":
                return None
            return os.path.join(respaldo_sipp.CARPETA_RESPALDOS, nombre)

        def crear_copia():
            if respaldo_sipp is None:
                estado.configure(text="No se encontró el módulo de respaldo.")
                return
            estado.configure(text="Creando copia de seguridad...")
            progreso_respaldo.start()
            boton_crear.configure(state="disabled")

            def copia_terminada(ruta):
                progreso_respaldo.stop()
                boton_crear.configure(state="normal")
                estado.configure(text=f"Copia creada correctamente: {os.path.basename(ruta)}")
                refrescar_lista_copias()

            def copia_fallida(exc):
                progreso_respaldo.stop()
                boton_crear.configure(state="normal")
                logging.exception("No se pudo crear la copia de seguridad")
                estado.configure(text=f"No se pudo crear la copia: {exc}")

            ejecutar_en_segundo_plano(
                respaldo_sipp.crear_copia_seguridad,
                lambda ruta: top_level.after(0, lambda: copia_terminada(ruta)),
                lambda exc: top_level.after(0, lambda: copia_fallida(exc)),
            )

        def restaurar(ruta):
            if not ruta:
                estado.configure(text="Selecciona una copia para restaurar.")
                return
            confirmar = messagebox.askyesno(
                "Restaurar copia de seguridad",
                "Esto reemplazará los datos actuales por los de la copia seleccionada. ¿Deseas continuar?",
                parent=top_level,
            )
            if not confirmar:
                return
            estado.configure(text="Restaurando copia de seguridad...")
            progreso_respaldo.start()
            boton_restaurar_seleccionada.configure(state="disabled")
            boton_restaurar_ultima.configure(state="disabled")

            def restauracion_terminada(_resultado):
                progreso_respaldo.stop()
                boton_restaurar_seleccionada.configure(state="normal")
                boton_restaurar_ultima.configure(state="normal")
                estado.configure(text="Restauración completada. Reinicia SiPP para ver los datos restaurados.")

            def restauracion_fallida(exc):
                progreso_respaldo.stop()
                boton_restaurar_seleccionada.configure(state="normal")
                boton_restaurar_ultima.configure(state="normal")
                logging.exception("No se pudo restaurar la copia de seguridad")
                estado.configure(text=f"No se pudo restaurar: {exc}")

            ejecutar_en_segundo_plano(
                lambda: respaldo_sipp.restaurar_copia_seguridad(ruta),
                lambda resultado: top_level.after(0, lambda: restauracion_terminada(resultado)),
                lambda exc: top_level.after(0, lambda: restauracion_fallida(exc)),
            )

        def restaurar_seleccionada():
            restaurar(ruta_copia_seleccionada())

        def restaurar_ultima():
            restaurar(respaldo_sipp.obtener_ultima_copia() if respaldo_sipp else None)

        refrescar_lista_copias()

        botones_arriba = ctk.CTkFrame(frame, fg_color="transparent")
        botones_arriba.pack(fill="x", padx=20, pady=(0, 6))
        boton_crear = ctk.CTkButton(botones_arriba, text="Crear copia ahora", width=160, fg_color="#2563eb", hover_color="#1d4ed8", command=crear_copia)
        boton_crear.pack(side="left")

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=20, pady=20)
        ctk.CTkButton(botones, text="Cerrar", width=120, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        boton_restaurar_seleccionada = ctk.CTkButton(botones, text="Restaurar seleccionada", width=170, fg_color="#b45309", hover_color="#92400e", command=restaurar_seleccionada)
        boton_restaurar_seleccionada.pack(side="right", padx=(0, 10))
        boton_restaurar_ultima = ctk.CTkButton(botones, text="Restaurar la última", width=150, fg_color="#dc2626", hover_color="#b91c1c", command=restaurar_ultima)
        boton_restaurar_ultima.pack(side="right", padx=(0, 10))


    def configuracion_nueva(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 620, 520)
        top_level.title("Datos de la empresa")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(frame, text="Datos de la empresa", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(22, 18))

        formulario = ctk.CTkFrame(frame, fg_color="transparent")
        formulario.pack(fill="both", expand=True, padx=20)
        formulario.grid_columnconfigure(1, weight=1)
        campos = {}
        for fila, (clave, texto) in enumerate((("nombre", "Nombre"), ("ruc", "RUC"), ("direccion", "Dirección"), ("telefono", "Teléfono"))):
            ctk.CTkLabel(formulario, text=f"{texto}:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(row=fila, column=0, padx=(0, 16), pady=8, sticky="w")
            campos[clave] = ctk.CTkEntry(formulario, height=34)
            campos[clave].grid(row=fila, column=1, pady=8, sticky="ew")

        logo_seleccionado = ctk.StringVar(value="")
        try:
            empresa_guardada = db.empresa().obtener_empresa() if db and hasattr(db, "empresa") else None
            if empresa_guardada:
                for indice, clave in enumerate(("nombre", "ruc", "direccion", "telefono"), start=1):
                    campos[clave].insert(0, empresa_guardada[indice] or "")
                logo_seleccionado.set(empresa_guardada[5] or "")
        except Exception:
            logging.exception("No se pudieron cargar los datos de la empresa")
        ctk.CTkLabel(formulario, text="Logo:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(row=4, column=0, padx=(0, 16), pady=8, sticky="w")
        logo_frame = ctk.CTkFrame(formulario, fg_color="transparent")
        logo_frame.grid(row=4, column=1, pady=8, sticky="ew")
        logo_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(logo_frame, textvariable=logo_seleccionado, height=34).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        vista_logo = ctk.CTkLabel(
            formulario,
            text="No se ha seleccionado un logo",
            width=220,
            height=100,
            corner_radius=10,
            fg_color=("#f1f5f9", "#1f2937"),
            text_color=("#64748b", "#cbd5e1"),
        )
        vista_logo.grid(row=5, column=1, pady=(4, 8), sticky="w")

        def cargar_vista_logo():
            ruta_logo = filedialog.askopenfilename(
                title="Seleccionar logo",
                filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.ico"), ("Todos los archivos", "*.*")],
            )
            if not ruta_logo:
                return
            try:
                imagen_logo = Image.open(ruta_logo).convert("RGBA")
                imagen_logo.thumbnail((220, 100), Image.Resampling.LANCZOS)
                imagen_ctk = ctk.CTkImage(
                    light_image=imagen_logo,
                    dark_image=imagen_logo,
                    size=imagen_logo.size,
                )
                vista_logo.configure(image=imagen_ctk, text="")
                vista_logo.image = imagen_ctk
                logo_seleccionado.set(ruta_logo)
            except (OSError, ValueError) as exc:
                messagebox.showerror("Logo", f"No se pudo cargar la imagen: {exc}")

        ctk.CTkButton(
            logo_frame,
            text="Examinar",
            width=100,
            height=34,
            command=cargar_vista_logo,
        ).grid(row=0, column=1)

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=20)
        ctk.CTkButton(botones, text="Cerrar", width=120, height=34, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")

        def guardar_datos_empresa():
            global datos_empresa
            datos_empresa = {clave: campos[clave].get().strip() for clave in campos}
            datos_empresa["logo"] = logo_seleccionado.get().strip()
            try:
                db.empresa().guardar_empresa(
                    datos_empresa["nombre"],
                    datos_empresa["ruc"],
                    datos_empresa["direccion"],
                    datos_empresa["telefono"],
                    datos_empresa["logo"],
                )
            except Exception as exc:
                logging.exception("No se pudieron guardar los datos de la empresa")
                messagebox.showerror("Datos de la empresa", f"No se pudieron guardar: {exc}", parent=top_level)
                return
            mensaje = "Datos de la empresa actualizados correctamente." if empresa_guardada else "Datos de la empresa guardados correctamente."
            messagebox.showinfo("Datos de la empresa", mensaje, parent=top_level)
            top_level.destroy()

        texto_guardar = "Actualizar" if empresa_guardada else "Guardar"
        ctk.CTkButton(botones, text=texto_guardar, width=120, height=34, fg_color="#2563eb", hover_color="#1d4ed8", command=guardar_datos_empresa).pack(side="right", padx=(0, 10))

        if empresa_guardada:
            def eliminar_datos_empresa():
                global datos_empresa
                confirmar = messagebox.askyesno(
                    "Eliminar datos de la empresa",
                    "¿Está seguro de eliminar los datos guardados de la empresa?",
                    parent=top_level,
                )
                if not confirmar:
                    return
                try:
                    if db.empresa().eliminar_empresa():
                        datos_empresa = {"nombre": "", "ruc": "", "direccion": "", "telefono": "", "logo": ""}
                        messagebox.showinfo("Datos de la empresa", "Los datos fueron eliminados correctamente.", parent=top_level)
                        top_level.destroy()
                    else:
                        messagebox.showwarning("Datos de la empresa", "No se encontraron datos para eliminar.", parent=top_level)
                except Exception as exc:
                    logging.exception("No se pudieron eliminar los datos de la empresa")
                    messagebox.showerror("Datos de la empresa", f"No se pudieron eliminar: {exc}", parent=top_level)

            ctk.CTkButton(
                botones,
                text="Eliminar datos",
                width=130,
                height=34,
                fg_color="#dc2626",
                hover_color="#b91c1c",
                command=eliminar_datos_empresa,
            ).pack(side="left")



    def configurar_idioma(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 420, 240)
        top_level.title("Idioma")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text="Idioma",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(22, 14))

        idioma_actual = getattr(self, "idioma_actual", "Español")
        selector_idioma = ctk.CTkComboBox(
            frame,
            values=["Español", "English"],
            width=220,
            state="readonly",
        )
        selector_idioma.set(idioma_actual)
        selector_idioma.pack(padx=20, pady=(0, 18), anchor="w")

        def guardar_idioma():
            self.idioma_actual = selector_idioma.get()
            establecer_idioma_aplicacion(self.idioma_actual)
            guardar_idioma_en_disco(self.idioma_actual)
            top_level.destroy()
            if hasattr(self, "frameprincipal") and self.frameprincipal.winfo_exists():
                self.frameprincipal.destroy()
            self.ventana_principal_entrada(sipp)

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=20, pady=16)
        ctk.CTkButton(
            botones,
            text="Cancelar",
            width=110,
            fg_color="#64748b",
            hover_color="#475569",
            command=top_level.destroy,
        ).pack(side="right")
        ctk.CTkButton(
            botones,
            text="Guardar",
            width=110,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=guardar_idioma,
        ).pack(side="right", padx=(0, 10))


    def configurar_app(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 720, 620, margen_x=100, margen_y=100)
        top_level.title("API Lector")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(
            top_level,
            corner_radius=18,
            fg_color=("#ffffff", "#111b27"),
            border_width=1,
        )
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(frame, text="Conector de lector biométrico", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 4))
        ctk.CTkLabel(frame, text="El identificador recibido debe ser el DUI del empleado registrado en SiPP.", font=ctk.CTkFont(size=12), anchor="w", text_color=("#526071", "#a1a9b8")).pack(fill="x", padx=20, pady=(0, 14))

        formulario = ctk.CTkFrame(frame, fg_color="transparent")
        formulario.pack(fill="x", padx=20)
        formulario.grid_columnconfigure(1, weight=1)
        formulario.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(formulario, text="Forma de conexión:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(0, 10), pady=7, sticky="w")
        selector_tipo = ctk.CTkComboBox(formulario, values=["CSV", "JSON", "REST", "SQLite"], state="readonly", width=180)
        selector_tipo.set("CSV")
        selector_tipo.grid(row=0, column=1, padx=(0, 14), pady=7, sticky="ew")

        ctk.CTkLabel(formulario, text="Ruta o URL:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(0, 10), pady=7, sticky="w")
        entrada_origen = ctk.CTkEntry(formulario, placeholder_text="Archivo CSV/JSON, SQLite o URL REST")
        entrada_origen.grid(row=0, column=3, pady=7, sticky="ew")

        ctk.CTkLabel(formulario, text="Token REST:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=0, padx=(0, 10), pady=7, sticky="w")
        entrada_token = ctk.CTkEntry(formulario, placeholder_text="Opcional", show="*")
        entrada_token.grid(row=1, column=1, padx=(0, 14), pady=7, sticky="ew")

        ctk.CTkLabel(formulario, text="Consulta SQLite:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=2, padx=(0, 10), pady=7, sticky="w")
        entrada_consulta = ctk.CTkEntry(formulario, placeholder_text="SELECT * FROM marcaciones")
        entrada_consulta.grid(row=1, column=3, pady=7, sticky="ew")

        columnas = (("Identificador DUI", "dui"), ("Fecha", "fecha"), ("Hora", "hora"), ("Tipo", "tipo"))
        controles_columnas = {}
        for indice, (texto, clave) in enumerate(columnas, start=2):
            fila = indice // 2
            columna = (indice % 2) * 2
            ctk.CTkLabel(formulario, text=f"Columna {texto}:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=fila, column=columna, padx=(0, 10), pady=7, sticky="w")
            controles_columnas[clave] = ctk.CTkEntry(formulario)
            controles_columnas[clave].grid(row=fila, column=columna + 1, padx=(0, 14), pady=7, sticky="ew")

        controles_columnas["dui"].insert(0, "dui")
        controles_columnas["fecha"].insert(0, "fecha")
        controles_columnas["hora"].insert(0, "hora")
        controles_columnas["tipo"].insert(0, "tipo")

        estado = ctk.CTkLabel(frame, text="Listo para configurar.", anchor="w", justify="left", wraplength=650, text_color=("#526071", "#cbd5e1"))
        estado.pack(fill="x", padx=20, pady=(14, 10))

        def obtener_configuracion():
            tipo = selector_tipo.get().strip()
            comun = {
                "columna_id": controles_columnas["dui"].get().strip(),
                "columna_fecha": controles_columnas["fecha"].get().strip(),
                "columna_hora": controles_columnas["hora"].get().strip(),
                "columna_tipo": controles_columnas["tipo"].get().strip(),
            }
            if tipo in {"CSV", "JSON"}:
                return tipo, {"ruta": entrada_origen.get().strip(), **comun}
            if tipo == "REST":
                return tipo, {"url": entrada_origen.get().strip(), "token": entrada_token.get().strip(), **comun}
            return tipo, {"ruta": entrada_origen.get().strip(), "consulta": entrada_consulta.get().strip(), **comun}

        def probar_conexion():
            if crear_conector is None:
                estado.configure(text="No se encontró el módulo de conectores.")
                return
            try:
                tipo, configuracion = obtener_configuracion()
                estado.configure(text=crear_conector(tipo, **configuracion).probar_conexion())
            except ErrorConector as exc:
                estado.configure(text=f"Error de conexión: {exc}")
            except Exception as exc:
                estado.configure(text=f"Error inesperado: {exc}")

        def sincronizar():
            if crear_conector is None:
                estado.configure(text="No se encontró el módulo de conectores.")
                return
            try:
                tipo, configuracion = obtener_configuracion()
                registros = crear_conector(tipo, **configuracion).obtener_marcaciones()
                nuevos = actualizados = rechazados = duplicados = 0
                for registro in registros:
                    empleado = db.gestion_empleado().buscar_empleado(registro.identificador)
                    if not empleado or str(empleado[2] or "").strip() != registro.identificador:
                        rechazados += 1
                        continue
                    dui = str(empleado[2]).strip()
                    existente = db.marcaciones.buscar_marcacion(dui, registro.fecha)
                    tipo_marcacion = registro.tipo.upper()
                    es_salida = tipo_marcacion in {"SALIDA", "OUT", "CHECK-OUT", "CHECKOUT"}
                    if existente:
                        if es_salida:
                            actualizado = db.marcaciones().actualizar_marcacion(dui, registro.fecha, existente[2], registro.hora, empleado[1])
                            actualizados += 1 if actualizado else 0
                        else:
                            duplicados += 1
                        continue
                    if db.marcaciones.insertar_marcacion(empleado[1], dui, registro.hora, registro.hora, registro.fecha):
                        nuevos += 1
                estado.configure(text=f"Sincronización terminada. Nuevas: {nuevos} | Actualizadas: {actualizados} | Duplicadas: {duplicados} | Rechazadas: {rechazados}")
            except ErrorConector as exc:
                estado.configure(text=f"No se pudo sincronizar: {exc}")
            except Exception as exc:
                logging.exception("No se pudo sincronizar el lector biométrico")
                estado.configure(text=f"Error durante la sincronización: {exc}")

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 18))
        ctk.CTkButton(botones, text="Probar conexión", width=140, fg_color="#2563eb", hover_color="#1d4ed8", command=probar_conexion).pack(side="left")
        ctk.CTkButton(botones, text="Sincronizar", width=130, fg_color="#0f766e", hover_color="#115e59", command=sincronizar).pack(side="left", padx=(10, 0))
        ctk.CTkButton(botones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")


    def configurar_notificaciones(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 480, 330)
        top_level.title("Notificaciones")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Notificaciones", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(22, 16))

        notificaciones = {
            "notificaciones_activas": ctk.BooleanVar(value=getattr(self, "notificaciones_activas", True)),
            "avisos_planilla": ctk.BooleanVar(value=getattr(self, "avisos_planilla", True)),
            "avisos_marcaciones": ctk.BooleanVar(value=getattr(self, "avisos_marcaciones", True)),
        }
        for texto, clave in (
            ("Activar notificaciones", "notificaciones_activas"),
            ("Avisos de planilla", "avisos_planilla"),
            ("Avisos de marcaciones", "avisos_marcaciones"),
        ):
            ctk.CTkCheckBox(frame, text=texto, variable=notificaciones[clave]).pack(anchor="w", padx=24, pady=7)

        def guardar_notificaciones():
            for clave, variable in notificaciones.items():
                setattr(self, clave, variable.get())
            top_level.destroy()

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=20, pady=16)
        ctk.CTkButton(botones, text="Cancelar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        ctk.CTkButton(botones, text="Guardar", width=110, fg_color="#2563eb", hover_color="#1d4ed8", command=guardar_notificaciones).pack(side="right", padx=(0, 10))


    def configurar_opciones_visuales(self, sipp):
        top_level = ctk.CTkToplevel(sipp)
        self.ajustar_toplevel_a_pantalla(top_level, 560, 500)
        top_level.title("Tamaño de texto y opciones visuales")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Tamaño de texto y opciones visuales", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(22, 16))

        ctk.CTkLabel(frame, text="Ajusta la escala de la interfaz para mejorar la lectura.", font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w").pack(fill="x", padx=24, pady=(0, 12))
        tamaño_actual = getattr(self, "tamaño_texto", "Mediano")
        selector_tamaño = ctk.CTkComboBox(frame, values=["Pequeño", "Mediano", "Grande"], width=220, state="readonly")
        selector_tamaño.set(tamaño_actual)
        fila_tamaño = ctk.CTkFrame(frame, fg_color="transparent")
        fila_tamaño.pack(fill="x", padx=24, pady=(0, 14))
        ctk.CTkLabel(fila_tamaño, text="Tamaño del texto:", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(side="left")
        selector_tamaño.pack(in_=fila_tamaño, side="right")
        alto_contraste = ctk.BooleanVar(value=getattr(self, "alto_contraste", False))
        reducir_movimiento = ctk.BooleanVar(value=getattr(self, "reducir_movimiento", False))
        ctk.CTkCheckBox(frame, text="Alto contraste", variable=alto_contraste).pack(anchor="w", padx=24, pady=7)
        ctk.CTkCheckBox(frame, text="Reducir movimiento", variable=reducir_movimiento).pack(anchor="w", padx=24, pady=7)

        vista_previa = ctk.CTkFrame(frame, corner_radius=10, fg_color=("#f1f5f9", "#1f2937"))
        vista_previa.pack(fill="x", padx=24, pady=(14, 10))
        ctk.CTkLabel(vista_previa, text="Vista previa", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(fill="x", padx=14, pady=(10, 3))
        texto_previo = ctk.CTkLabel(vista_previa, text="SiPP | Control de personal y planilla", anchor="w")
        texto_previo.pack(fill="x", padx=14, pady=(0, 10))

        escala_por_tamaño = {"Pequeño": 0.90, "Mediano": 1.0, "Grande": 1.15}

        def aplicar_opciones():
            tamaño = selector_tamaño.get()
            escala = escala_por_tamaño.get(tamaño, 1.0)
            ctk.set_widget_scaling(escala)
            ctk.set_window_scaling(escala)
            self.tamaño_texto = tamaño
            self.escala_visual = escala
            self.alto_contraste = alto_contraste.get()
            self.reducir_movimiento = reducir_movimiento.get()
            color = ("#0f172a", "#ffffff") if self.alto_contraste else ("#1e293b", "#e2e8f0")
            texto_previo.configure(text_color=color, font=ctk.CTkFont(size=13 if tamaño == "Grande" else 11, weight="bold" if self.alto_contraste else "normal"))
            estado.configure(text=f"Aplicado: texto {tamaño.lower()} | contraste {'alto' if self.alto_contraste else 'normal'}.")

        def restaurar_opciones():
            selector_tamaño.set("Mediano")
            alto_contraste.set(False)
            reducir_movimiento.set(False)
            ctk.set_widget_scaling(1.0)
            ctk.set_window_scaling(1.0)
            self.tamaño_texto = "Mediano"
            self.escala_visual = 1.0
            self.alto_contraste = False
            self.reducir_movimiento = False
            texto_previo.configure(text_color=("#1e293b", "#e2e8f0"), font=ctk.CTkFont(size=11))
            estado.configure(text="Opciones visuales restauradas.")

        estado = ctk.CTkLabel(frame, text="Los cambios se aplican a las ventanas nuevas.", anchor="w", text_color=("#526071", "#cbd5e1"))
        estado.pack(fill="x", padx=24, pady=(0, 8))

        def guardar_opciones():
            aplicar_opciones()
            top_level.destroy()

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=20, pady=16)
        ctk.CTkButton(botones, text="Restaurar", width=105, fg_color="#b45309", hover_color="#92400e", command=restaurar_opciones).pack(side="left")
        ctk.CTkButton(botones, text="Aplicar", width=100, fg_color="#0f766e", hover_color="#115e59", command=aplicar_opciones).pack(side="left", padx=(8, 0))
        ctk.CTkButton(botones, text="Cancelar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        ctk.CTkButton(botones, text="Guardar", width=110, fg_color="#2563eb", hover_color="#1d4ed8", command=guardar_opciones).pack(side="right", padx=(0, 10))


    def crear_administrador(self, sipp):
        if not self._es_administrador(sipp):
            messagebox.showerror("Acceso denegado", "Solo un administrador puede crear administradores.", parent=sipp)
            return

        top_level = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(top_level, 560, 650, margen_x=80, margen_y=80, min_ancho=480, min_alto=560)
        top_level.title("Crear Administrador")
        top_level.resizable(False, True)
        top_level.transient(sipp)
        top_level.grab_set()

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Crear administrador", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(18, 4))
        ctk.CTkLabel(frame, text="El nuevo usuario tendrá permisos completos.", anchor="w").pack(fill="x", padx=20, pady=(0, 14))

        ctk.CTkLabel(frame, text="Código maestro de seguridad", anchor="w").pack(fill="x", padx=20)
        entrada_codigo_seguridad = ctk.CTkEntry(
            frame,
            height=34,
            show="*",
            placeholder_text="Código del administrador actual",
        )
        entrada_codigo_seguridad.pack(fill="x", padx=20, pady=(4, 10))

        ctk.CTkLabel(frame, text="Nombre de usuario", anchor="w").pack(fill="x", padx=20)
        entrada_usuario = ctk.CTkEntry(frame, height=34, placeholder_text="Ej: administrador2")
        entrada_usuario.pack(fill="x", padx=20, pady=(4, 10))

        ctk.CTkLabel(frame, text="Contraseña", anchor="w").pack(fill="x", padx=20)
        entrada_contraseña = ctk.CTkEntry(frame, height=34, show="*", placeholder_text="Contraseña segura")
        entrada_contraseña.pack(fill="x", padx=20, pady=(4, 10))

        ctk.CTkLabel(frame, text="Confirmar contraseña", anchor="w").pack(fill="x", padx=20)
        confirmar_contraseña = ctk.CTkEntry(frame, height=34, show="*", placeholder_text="Repita la contraseña")
        confirmar_contraseña.pack(fill="x", padx=20, pady=(4, 14))
        codigo_nuevo = db.generar_codigo_8_digitos()
        ctk.CTkLabel(frame, text="Código de seguridad del nuevo administrador", anchor="w").pack(fill="x", padx=20)
        ctk.CTkLabel(
            frame,
            text=codigo_nuevo,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#2563eb",
        ).pack(fill="x", padx=20, pady=(4, 14))

        def guardar():
            codigo_seguridad = entrada_codigo_seguridad.get().strip()
            nombre = entrada_usuario.get().strip()
            contraseña = entrada_contraseña.get()
            confirmacion = confirmar_contraseña.get()
            if not codigo_seguridad or not nombre or not contraseña or not confirmacion:
                messagebox.showwarning("Crear administrador", "Complete todos los campos.", parent=top_level)
                return
            if codigo_seguridad != str(db._clave__de__acceso()):
                messagebox.showwarning("Crear administrador", "El código maestro de seguridad es incorrecto.", parent=top_level)
                return
            if contraseña != confirmacion:
                messagebox.showwarning("Crear administrador", "Las contraseñas no coinciden.", parent=top_level)
                return
            try:
                conexion = db.ConexionDB_datos_usuarios()
                if any(str(usuario[0]).strip().casefold() == nombre.casefold() for usuario in conexion.obtener_usuarios()):
                    messagebox.showwarning("Crear administrador", "Ese nombre de usuario ya existe.", parent=top_level)
                    return
                conexion.insertar_administrador(nombre, contraseña, codigo_nuevo)
            except ValueError as exc:
                messagebox.showwarning("Crear administrador", str(exc), parent=top_level)
                return
            except Exception as exc:
                logging.exception("No se pudo crear el administrador")
                messagebox.showerror("Crear administrador", f"No se pudo crear el administrador:\n{exc}", parent=top_level)
                return
            messagebox.showinfo(
                "Crear administrador",
                f"El administrador fue creado correctamente.\n\nCódigo de seguridad: {codigo_nuevo}",
                parent=sipp,
            )
            top_level.destroy()

        botones = ctk.CTkFrame(frame, fg_color="transparent")
        botones.pack(side="bottom", fill="x", padx=20, pady=(0, 14))
        ctk.CTkButton(botones, text="Cancelar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="left")
        ctk.CTkButton(botones, text="Crear", width=110, command=guardar).pack(side="right")
        entrada_usuario.focus_set()

    def crear_usuarios(self, sipp):
        if not self._es_administrador(sipp):
            messagebox.showerror("Acceso denegado", "Solo un administrador puede gestionar usuarios.", parent=sipp)
            return
        top_level = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(top_level, 420, 210, margen_x=120, margen_y=100, min_ancho=340, min_alto=190)
        top_level.title("Autorización de acceso")
        top_level.resizable(False, False)
        top_level.transient(sipp)
        top_level.grab_set()
        top_level.lift()
        top_level.focus_force()
        top_level.attributes("-topmost", True)
        top_level.after(10, lambda: top_level.attributes("-topmost", False))
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(frame, text="Código de seguridad", font=ctk.CTkFont(size=24, weight="bold"), anchor="w").place(x=20, y=18)
        ctk.CTkLabel(frame, text="Ingrese el código maestro para crear usuarios", font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w").place(x=20, y=58)

        ctk.CTkLabel(frame, text="Código", font=ctk.CTkFont(size=13, weight="bold")).place(x=20, y=100)
        entrada_contrasena = ctk.CTkEntry(frame, font=ctk.CTkFont(size=15), width=250, height=40, justify="center", show="*")
        entrada_contrasena.place(x=120, y=94)
        entrada_contrasena.focus_set()
        entrada_contrasena.bind("<Return>", command=lambda event: self.top_crear_usuario(sipp, entrada_contrasena.get().strip(), top_level))

        boton = ctk.CTkButton(frame, text="Continuar", width=140, height=36, fg_color="#2563eb", hover_color="#1d4ed8",
                              command=lambda: self.top_crear_usuario(sipp, entrada_contrasena.get().strip(), top_level))
        boton.place(x=230, y=150)


    def top_crear_usuario(self, sipp, contrasena, top_level):
        
        if self._es_administrador(sipp):
            top_level.destroy()
            top_level = ctk.CTkToplevel(sipp)
            ajustar_toplevel_a_pantalla(top_level, 520, 460, margen_x=120, margen_y=120, min_ancho=420, min_alto=330)
            top_level.title("Crear Usuario")
            icono = "imagenes/icono1.ico"
            top_level.resizable(False, False)
            top_level.transient(sipp)
            top_level.grab_set()
            top_level.lift()
            top_level.focus_force()
            top_level.attributes("-topmost", True)
            top_level.after(10, lambda: top_level.attributes("-topmost", False))
            top_level.configure(fg_color=("#eef8fd", "#101b2a"))

            codigo = db.generar_codigo_8_digitos()

            frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#101b2a"), border_width=1)
            frame.pack(fill="both", expand=True, padx=10, pady=10)

            panel_titulo = ctk.CTkFrame(frame, corner_radius=14, fg_color=("#e0f2fe", "#1e293b"), width=500, height=76)
            panel_titulo.place(x=10, y=10)
            ctk.CTkLabel(panel_titulo, text="Crear usuario", font=ctk.CTkFont(size=20, weight="bold"), text_color="#0f172a").place(x=16, y=12)
            ctk.CTkLabel(panel_titulo, text="Registro rápido de acceso seguro.", font=ctk.CTkFont(size=10), text_color="#334155").place(x=16, y=44)

            imagen_masculina = ctk.CTkImage(
                light_image=Image.open(resource_path("imagenes/hombre.png")),
                dark_image=Image.open(resource_path("imagenes/hombre.png")),
                size=(64, 64)
            )
            imagen_femenina = ctk.CTkImage(
                light_image=Image.open(resource_path("imagenes/mujer.png")),
                dark_image=Image.open(resource_path("imagenes/mujer.png")),
                size=(64, 64)
            )
            top_level.imagen_masculina = imagen_masculina
            top_level.imagen_femenina = imagen_femenina

            ctk.CTkLabel(frame, text="Nombre de usuario", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=20, y=100)
            entrada_usuario = ctk.CTkEntry(frame, width=360, height=34, justify="center", placeholder_text="Ej: marvin2025")
            entrada_usuario.place(x=20, y=130)

            ctk.CTkLabel(frame, text="Género", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=20, y=176)
            genero_var = ctk.StringVar(value="Masculino")
            combo_genero = ctk.CTkComboBox(frame, width=160, values=["Masculino", "Femenino"], justify="center", variable=genero_var)
            combo_genero.place(x=20, y=204)

            icono_marco = ctk.CTkFrame(frame, corner_radius=26, fg_color=("#f8fafc", "#1f2937"), width=76, height=76)
            icono_marco.place(x=456, y=196, anchor="n")
            icono_genero = ctk.CTkLabel(icono_marco, image=imagen_masculina, text="", fg_color="transparent")
            icono_genero.place(relx=0.5, rely=0.5, anchor="center")

            def actualizar_icono(event=None):
                seleccion = genero_var.get().strip().lower()
                icono_genero.configure(image=top_level.imagen_femenina if seleccion.startswith("f") else top_level.imagen_masculina)

            genero_var.trace_add("write", lambda *_: actualizar_icono())
            combo_genero.bind("<<ComboboxSelected>>", actualizar_icono)
            actualizar_icono()

            ctk.CTkLabel(frame, text="Contraseña", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=20, y=250)
            entrada_contraseña = ctk.CTkEntry(frame, width=360, height=34, justify="center", show="*", placeholder_text="Ej: ash@ask?12")
            entrada_contraseña.place(x=20, y=280)

            ctk.CTkLabel(frame, text="Código generado", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=20, y=330)
            ctk.CTkLabel(frame, text=codigo, font=ctk.CTkFont(size=12, weight="bold"), text_color="#2563eb").place(x=20, y=360)

            botones_panel = ctk.CTkFrame(frame, fg_color="transparent", width=460, height=44)
            botones_panel.place(x=20, y=391)

            ctk.CTkButton(botones_panel, text="Cancelar", width=140, height=36, fg_color="#64748b", hover_color="#475569",
                           command=top_level.destroy).place(x=0, y=4)
            ctk.CTkButton(botones_panel, text="Guardar usuario", width=160, height=36,
                                        fg_color="#2f855a", hover_color="#276749",
                                        command=lambda: self.guardar_usuario(entrada_usuario.get().strip(), combo_genero.get().title(), entrada_contraseña.get(), codigo, top_level)).place(x=300, y=4)
            
            try:
                base_datos = db.ConexionDB_datos_usuarios()
                base_datos.crear_tabla_usuarios()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear la tabla de usuarios: {e}")
        else:
            messagebox.showerror("Error", "Codigo Incorrecto.")

        

    def guardar_usuario(self, nombre, genero, contraseña,codigo, top):
        conexion = db.ConexionDB_datos_usuarios()
        if nombre == "" or genero == "" or contraseña == "":
            messagebox.showerror("Error", "Por favor, complete todos los campos.")
            return
        if conexion:
            try:
                conexion.insertar_usuario(nombre, genero, contraseña, codigo)
            except ValueError as exc:
                messagebox.showwarning("Avertencia", str(exc))
                return
            messagebox.showinfo("Usuario Guardado", "El usuario ha sido guardado exitosamente.")
            messagebox.showinfo("CODIGO DE USUARIO", f"{codigo}")
        else:
            messagebox.showerror("Error", "No se pudo guardar el usuario. Intente nuevamente.")
        
        top.destroy()

    def idenficar_usuario(self, sipp):
        if not self._es_administrador(sipp):
            messagebox.showerror("Acceso denegado", "Solo un administrador puede gestionar usuarios.", parent=sipp)
            return
        top_level = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(top_level, 420, 210, margen_x=120, margen_y=100, min_ancho=340, min_alto=190)
        top_level.title("Autorización de acceso")
        top_level.resizable(False, False)
        top_level.grab_set()
        top_level.configure(fg_color=("#eef8fd", "#101b2a"))

        frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(frame, text="Código de seguridad", font=ctk.CTkFont(size=24, weight="bold"), anchor="w").place(x=20, y=18)
        ctk.CTkLabel(frame, text="Ingrese el código maestro para administrar usuarios", font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w").place(x=20, y=58)

        ctk.CTkLabel(frame, text="Código", font=ctk.CTkFont(size=13, weight="bold")).place(x=20, y=100)
        entrada_contrasena = ctk.CTkEntry(frame, font=ctk.CTkFont(size=15), width=250, height=40, justify="center", show="*")
        entrada_contrasena.place(x=120, y=94)
        entrada_contrasena.focus_set()
        entrada_contrasena.bind("<Return>", command=lambda event: self.top_level_usarios_creados(sipp, entrada_contrasena.get().strip(), top_level))

        boton = ctk.CTkButton(frame, text="Continuar", width=140, height=36, fg_color="#2563eb", hover_color="#1d4ed8",
                              command=lambda: self.top_level_usarios_creados(sipp, entrada_contrasena.get().strip(), top_level))
        boton.place(x=230, y=150)


    def top_level_usarios_creados(self, sipp, contrasena, top_level):
        if self._es_administrador(sipp):
            usuarios = [
                usuario for usuario in db.ConexionDB_datos_usuarios().obtener_usuarios()
                if str(usuario[0]).strip() not in {"@@22", "@@34"}
            ]
            if usuarios:
                top_level.destroy()
                top_level = ctk.CTkToplevel(sipp)
                ajustar_toplevel_a_pantalla(top_level, 420, 410, margen_x=120, margen_y=120, min_ancho=340, min_alto=300)
                top_level.title("Usuarios del sistema")
                top_level.resizable(False, False)
                top_level.transient(sipp)
                top_level.grab_set()
                top_level.lift()
                top_level.focus_force()
                top_level.configure(fg_color=("#eef8fd", "#101b2a"))

                frame = ctk.CTkFrame(top_level, corner_radius=18, fg_color=("#ffffff", "#111b27"), border_width=1)
                frame.pack(fill="both", expand=True, padx=16, pady=16)

                ctk.CTkLabel(frame, text="Usuarios registrados", font=ctk.CTkFont(size=24, weight="bold"), anchor="w").place(x=20, y=18)
                ctk.CTkLabel(frame, text="Seleccione un usuario para eliminarlo", font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w").place(x=20, y=58)

                self.variable_usuario = ctk.StringVar()
                usuarios_frame = ctk.CTkScrollableFrame(frame, width=380, height=220, fg_color=("#f8fafc", "#111827"))
                usuarios_frame.place(relx=0.5, y=100, anchor="n")

                imagen_hombre = ctk.CTkImage(
                    light_image=Image.open(resource_path("imagenes/hombre.png")),
                    dark_image=Image.open(resource_path("imagenes/hombre.png")),
                    size=(28, 28)
                )
                imagen_mujer = ctk.CTkImage(
                    light_image=Image.open(resource_path("imagenes/mujer.png")),
                    dark_image=Image.open(resource_path("imagenes/mujer.png")),
                    size=(28, 28)
                )

                for nombre, genero, *_ in usuarios:
                    item_frame = ctk.CTkFrame(usuarios_frame, fg_color=("#eef2ff", "#1f2937"), corner_radius=14)
                    item_frame.pack(fill="x", pady=8, padx=10)
                    item_frame.configure(height=48)

                    icono_genero = imagen_mujer if genero.strip().lower().startswith("f") else imagen_hombre
                    ctk.CTkLabel(item_frame, image=icono_genero, text="", width=32, height=32).place(x=10, rely=0.5, anchor="w")

                    self.radio_seleccion = ctk.CTkRadioButton(
                        item_frame,
                        variable=self.variable_usuario,
                        value=nombre,
                        text=f"{nombre}",
                        width=260,
                        height=30,
                        font=ctk.CTkFont(size=13),
                        command=lambda usuario=nombre: self.eliminar_usuario(usuario, top_level)
                    )
                    self.radio_seleccion.place(relx=0.5, rely=0.5, anchor="center")

                ctk.CTkButton(frame, text="Cerrar", width=100, height=34, fg_color="#64748b", command=top_level.destroy).place(x=250, y=340)
            else:
                messagebox.showerror("Error", "No existe ningun usuario creado.")
                return
        else:
            messagebox.showerror("Error", "Codigo Incorrecto.")
    
    def eliminar_usuario(self, nombre, top):
        if str(nombre).strip() in {"@@22", "@@34"}:
            messagebox.showwarning("Aviso", "Este usuario no se puede eliminar.", parent=top)
            return
        conexion = db.ConexionDB_datos_usuarios()
        if messagebox.askyesno("Aviso", "Seguro que deseas eliminar"):
            conexion.eliminar_usuario(nombre)
        else:
            return
        if conexion:
            messagebox.showinfo("Usuario Eliminado", "El usuario ha sido eliminado exitosamente.")
        else:
            messagebox.showerror("Error", "No se pudo eliminar el usuario. Intente nuevamente.")
        top.destroy()


    def editar_apariencia(self, sipp):
        datos = edict.Tema()
        aparen = datos.obtener_tema()
        for i in aparen:
            for x in i:
                aparen = x
        top_level = ctk.CTkToplevel(sipp)
        top_level.geometry("400x50")
        top_level.title(f"Editar la Apariencia                      ....{aparen}")
        top_level.grab_set()
        top_level.resizable(False, False)

        frame = ctk.CTkFrame(top_level, height=50)
        frame.place(relx=0.00, rely=0.03, relwidth=1, relheight=1)

        label = ctk.CTkLabel(frame, text="Apariencia: ", font=ctk.CTkFont(size=16))
        label.place(relx=0.03, rely=0.2)
        apariencias = ["sky", "breeze", "cherry", "blue", "green", "coffee", "extreme", "rime", 
                       "lavender", "violet", "patina", "midnigth"]
        combo_apariencia = ctk.CTkComboBox(frame, values=apariencias, font=ctk.CTkFont(size=16))
        combo_apariencia.place(relx=0.3, rely=0.2)
        boton = ctk.CTkButton(frame, text="Actualizar", font=ctk.CTkFont(size=16), width=100, 
                              command= lambda : self.actualizar_apariencia(combo_apariencia.get(), top_level, sipp))
        boton.place(relx=0.7, rely=0.2)

        combo_apariencia.bind("<KeyRelease>", command= lambda event: self.buscar_apariencia(str(combo_apariencia.get()), apariencias))
#correguir errores de compatibilidad 
    def buscar_apariencia(self, combo, datos):

        texto = combo.lower()

        # Filtrar resultados
        # prepare services
        ded_serv = Deducciones()
        rem_serv = db.registro_remuneracion()

        resultados = []

        for item in datos:
            if texto in item.lower():
                resultados.append(item)

        # Actualizar valores
        combo.configure(values=resultados)

        # Abrir dropdown automáticamente
        if resultados:
            combo._dropdown_menu.open()


    def actualizar_apariencia(self, apariencias, top_level, sipp):
        datos = edict.Tema()
        datos.limpiar()
        datos.insertar(apariencias)
        if datos:
            aparen = datos.obtener_tema()
            for i in aparen:
                for x in i:
                    aparen = x

        
            apariencia = theme_manager.get(str(aparen))
            ctk.set_default_color_theme(str(apariencia))
            messagebox.showinfo("Actualizado", f"{aparen}, actualizado correctamente. ")
            top_level.destroy()
            self.frameprincipal.destroy()
            self.ventana_principal_entrada(sipp)

        else:
            messagebox.showerror("Error", "Ocurrio un error al Actualizar. ")


    def editar_tema(self, sipp):
        datos = edict.Tema()
        tem = datos.obtener_vista()
        for i in tem:
            for x in i:
                tem = x
        top_level = ctk.CTkToplevel(sipp)
        top_level.geometry("400x50")
        top_level.title(f"Editar la tema                      ....{tem}")
        top_level.grab_set()
        top_level.resizable(False, False)

        frame = ctk.CTkFrame(top_level, height=50)
        frame.place(relx=0.00, rely=0.03, relwidth=1, relheight=1)

        label = ctk.CTkLabel(frame, text="Tema: ", font=ctk.CTkFont(size=16))
        label.place(relx=0.03, rely=0.2)
        apariencias = ["system", "dark"]
        combo_apariencia = ctk.CTkComboBox(frame, values=apariencias, font=ctk.CTkFont(size=16))
        combo_apariencia.place(relx=0.3, rely=0.2)
        boton = ctk.CTkButton(frame, text="Actualizar", font=ctk.CTkFont(size=16), width=100, 
                              command= lambda : self.actualizar_vista(combo_apariencia.get(), top_level, sipp))
        boton.place(relx=0.7, rely=0.2)

    def actualizar_vista(self, vista, top_level, sipp):
        datos = edict.Tema()
        datos.limpiar_vista()
        datos.insert_vista(vista)
        if datos:
            tema = datos.obtener_vista()
            for i in tema:
                for x in i:
                    tema = x

        
            
            ctk.set_appearance_mode(str(tema))
            messagebox.showinfo("Actualizado", f"{tema}, actualizado correctamente. ")
            top_level.destroy()
        else:
            messagebox.showerror("Error", "Ocurrio un error al Actualizar. ")

    def plataformas(self, sipp):
        self.ancho_pantalla = self.tamaño_pantalla_ancho(sipp)
        self.alto_pantalla = self.tamaño_pantalla_alto(sipp)

        if hasattr(self, "sidebar") and self.sidebar.winfo_exists():
            self.sidebar.destroy()

        ancho_sidebar = self._ancho_sidebar(sipp)
        self.sidebar = ctk.CTkFrame(
            sipp,
            width=ancho_sidebar,
            corner_radius=0,
            fg_color=("#e7edf8", "#0f172a"),
        )
        self.sidebar.place(x=0, y=0, relheight=1.0)

        cabecera_sidebar = ctk.CTkFrame(
            self.sidebar,
            corner_radius=0,
            fg_color=("#d7e3f6", "#111827"),
            height=180,
        )
        cabecera_sidebar.pack(fill="x")
        cabecera_sidebar.pack_propagate(False)

        self._sidebar_logo_img = ctk.CTkImage(
            light_image=Image.open(resource_path("imagenes/icono1.png")),
            dark_image=Image.open(resource_path("imagenes/icono1.png")),
            size=(96, 96),
        )
        ctk.CTkLabel(cabecera_sidebar, image=self._sidebar_logo_img, text="").pack(pady=(18, 8))
        ctk.CTkLabel(
            cabecera_sidebar,
            text=self.t("SiPP"),
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=("#0f172a", "#f8fafc"),
        ).pack()
        ctk.CTkLabel(
            cabecera_sidebar,
            text=self.t("Sistema de Planilla Profesional"),
            font=ctk.CTkFont(size=11),
            text_color=("#475569", "#cbd5e1"),
        ).pack(pady=(4, 0))

        nav = ctk.CTkScrollableFrame(
            self.sidebar,
            corner_radius=0,
            fg_color="transparent",
        )
        nav.pack(fill="both", expand=True, padx=12, pady=(12, 14))

        botones = [
            ("Gestión Personal", lambda: self.gestion_personal(sipp)),
            ("Gestión Planilla", lambda: self.gestion_planilla(sipp)),
            ("Gestión de Tiempo", lambda: self.gestion_de_tiempo(sipp)),
            ("Descuentos y Préstamos", lambda: self.gestion_descuento(sipp)),
            ("Deducciones y Retenciones", lambda: self.gestion_deducciones(sipp)),
            ("Remuneraciones", lambda: self.remuneraciones(sipp)),
            ("Ausencias y Permisos", lambda: self.permisos_y_ausensias(sipp)),
            ("Vacaciones", lambda: self.vacaciones(sipp)),
            ("Expedientes laborales", lambda: self.contratos(sipp)),
            ("Liquidaciones", lambda: self.liquidaciones(sipp)),
            ("Cartas", lambda: self.cartas(sipp)),
            ("Manuales", lambda: self.manuales(sipp)),
        ]

        for texto, comando in botones:
            boton = self._crear_boton_lateral(nav, self.t(texto), comando)
            if comando is None:
                boton.configure(state="disabled")
            boton.pack(fill="x", pady=(0, 10))

        contenido = self._crear_area_contenido(sipp)
        self.fondo_inicio(contenido)

    def gestion_personal(self, sipp):
        if getattr(self, 'current_view', None) == 'gestion_personal' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'gestion_personal'
        self._crear_area_contenido(sipp)
        gestionar = Gestion_Personal()
        gestionar.ventana_gestion_personal(self.frame2)

    def gestion_planilla(self, sipp):
        if getattr(self, 'current_view', None) == 'gestion_planilla' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'gestion_planilla'
        self._crear_area_contenido(sipp)
        gestionar = Gestion_Planilla()
        gestionar.ventana_gestion_planilla(self.frame2)

    def gestion_de_tiempo(self, sipp):
        if getattr(self, 'current_view', None) == 'gestion_tiempo' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'gestion_tiempo'
        self._crear_area_contenido(sipp)
        gestionar = Gestion_de_tiempo()
        gestionar.ventana_gestion_de_tiempo(self.frame2)        

    def gestion_descuento(self, sipp):
        if getattr(self, 'current_view', None) == 'gestion_descuento' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'gestion_descuento'
        self._crear_area_contenido(sipp)
        gestionar = Gestion_Descuentos()
        gestionar.ventana_gestion_prestamo(self.frame2, sipp)

    def gestion_deducciones(self, sipp):
        if getattr(self, 'current_view', None) == 'deducciones' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'deducciones'
        self._crear_area_contenido(sipp)
        gestionar = Deducciones()
        gestionar.info_deducciones(self.frame2)

    def permisos_y_ausensias(self, sipp):
        # If already on this view and frame2 exists, destroy it first to force a clean reopen
        if getattr(self, 'current_view', None) == 'permisos_ausencias' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            try:
                self.frame2.destroy()
            except Exception:
                pass
        self.current_view = 'permisos_ausencias'
        self._crear_area_contenido(sipp)
        gestionar = Permisos_Ausencias()
        gestionar.info_Per_Aus(self.frame2)

    def remuneraciones(self, sipp):
        if getattr(self, 'current_view', None) == 'remuneraciones' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        
        # Limpiar instancia anterior si existe
        if hasattr(self, '_remuneraciones_instance') and self._remuneraciones_instance:
            try:
                self._remuneraciones_instance.limpiar()
            except Exception:
                pass
            self._remuneraciones_instance = None
        
        self.current_view = 'remuneraciones'
        self._crear_area_contenido(sipp)
        
        # Crear nueva instancia y guardarla
        self._remuneraciones_instance = Remuneraciones()
        self._remuneraciones_instance.info_remuneraciones(self.frame2)

    def vacaciones(self, sipp):
        if getattr(self, 'current_view', None) == 'vacaciones' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'vacaciones'
        self._crear_area_contenido(sipp)
        gestionar = Vacaciones()
        gestionar.info_vacaciones(self.frame2)

    def contratos(self, sipp):
        if getattr(self, 'current_view', None) == 'contratos' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'contratos'
        self._crear_area_contenido(sipp)
        gestionar = Contrato_trabajo()
        gestionar.info_contrato_trabajo(self.frame2)


    def liquidaciones(self, sipp):
        if getattr(self, 'current_view', None) == 'liquidaciones' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'liquidaciones'
        self._crear_area_contenido(sipp)
        gestionar = Liquidaciones()
        gestionar.info_liquidaciones(self.frame2)

    def cartas(self, sipp):
        if getattr(self, 'current_view', None) == 'cartas' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'cartas'
        self._crear_area_contenido(sipp)
        gestionar = Cartas_de_trabajo()
        gestionar.info_cartas(self.frame2)

    def manuales(self, sipp):
        if getattr(self, 'current_view', None) == 'manuales' and hasattr(self, 'frame2') and self.frame2.winfo_exists():
            return
        self.current_view = 'manuales'
        self._crear_area_contenido(sipp)
        gestionar = Manuales()
        gestionar.info_manuales(self.frame2)

    def fondo_inicio(self, sipp):
        for widget in sipp.winfo_children():
            widget.destroy()

        hero = ctk.CTkFrame(
            sipp,
            corner_radius=20,
            fg_color=("#eef4ff", "#111827"),
            border_width=1,
            border_color=("#d8e3f5", "#334155"),
        )
        hero.place(relx=0.03, rely=0.05, relwidth=0.94, relheight=0.88)

        panel_texto = ctk.CTkFrame(hero, fg_color="transparent")
        panel_texto.place(relx=0.05, rely=0.10, relwidth=0.44, relheight=0.78)

        ctk.CTkLabel(
            panel_texto,
            text=self.t("Panel Principal"),
            font=ctk.CTkFont(size=34, weight="bold"),
            text_color=("#0f172a", "#f8fafc"),
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            panel_texto,
            text=self.t("Administra personal, planilla, marcaciones y procesos clave desde un solo lugar."),
            font=ctk.CTkFont(size=15),
            text_color=("#475569", "#cbd5e1"),
            justify="left",
            wraplength=420,
            anchor="w",
        ).pack(fill="x", pady=(14, 24))

        resumen = [
            "Personal y expedientes centralizados",
            "Cortes, marcaciones y aprobaciones de horas extras",
            "Accesos rápidos para planilla y contratos",
        ]
        for texto in resumen:
            item = ctk.CTkFrame(panel_texto, fg_color="transparent")
            item.pack(fill="x", pady=(0, 12))
            ctk.CTkLabel(item, text="•", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#0f766e", "#5eead4")).pack(side="left")
            ctk.CTkLabel(item, text=self.t(texto), font=ctk.CTkFont(size=14), text_color=("#1e293b", "#e5e7eb"), anchor="w").pack(side="left", padx=(10, 0))

        accesos = ctk.CTkFrame(panel_texto, fg_color="transparent")
        accesos.pack(fill="x", pady=(28, 0))
        boton_buscar_profesional(accesos, text=self.t("Ir a Gestión de Tiempo"), command=lambda: self.gestion_de_tiempo(self.frameprincipal), width=180, height=36).pack(side="left", padx=(0, 10))
        boton_buscar_profesional(accesos, text=self.t("Ver Planilla"), command=lambda: self.gestion_planilla(self.frameprincipal), width=150, height=36).pack(side="left")

        panel_imagen = ctk.CTkFrame(hero, fg_color="transparent")
        panel_imagen.place(relx=0.53, rely=0.07, relwidth=0.42, relheight=0.82)
        self._fondo_inicio_img = ctk.CTkImage(
            light_image=Image.open(resource_path("imagenes/admin.png")),
            dark_image=Image.open(resource_path("imagenes/admin.png")),
            size=(520, 520),
        )
        ctk.CTkLabel(panel_imagen, image=self._fondo_inicio_img, text="").pack(fill="both", expand=True)



class Gestion_Personal:

    def ventana_gestion_personal(self, frame):
        self.frame_padre = frame
        for widget in frame.winfo_children():
            widget.destroy()

        cabecera = ctk.CTkFrame(frame, fg_color="transparent")
        cabecera.place(relx=0.03, rely=0.03, relwidth=0.94, relheight=0.12)

        ctk.CTkLabel(cabecera, text="Gestión Personal", font=ctk.CTkFont(size=28, weight="bold"), anchor="w").place(relx=0.00, rely=0.02)
        ctk.CTkLabel(cabecera, text="Administra empleados, departamentos y datos laborales desde una vista centralizada.", font=ctk.CTkFont(size=12), text_color="gray75", anchor="w").place(relx=0.00, rely=0.52)

        frame = ctk.CTkFrame(frame, corner_radius=20, border_width=1, border_color="#3f3f46")
        frame.place(relx=0.02, rely=0.18, relwidth=0.96, relheight=0.78)

        barra_acciones = ctk.CTkFrame(frame, fg_color="transparent")
        barra_acciones.place(relx=0.03, rely=0.02, relwidth=0.94, relheight=0.12)
        barra_acciones.grid_columnconfigure(0, weight=0)
        barra_acciones.grid_columnconfigure(1, weight=0)
        barra_acciones.grid_columnconfigure(2, weight=0)
        barra_acciones.grid_columnconfigure(3, weight=0)
        barra_acciones.grid_columnconfigure(4, weight=1)

        self.combo_departamento = ctk.CTkComboBox(
            barra_acciones,
            width=220,
            height=34,
            justify="center",
            state="readonly",
            command=self._filtrar_por_departamento,
        )
        self._actualizar_departamentos_combo()
        self.combo_departamento.grid(row=0, column=0, padx=(18, 10), pady=16, sticky="w")

        self.entrada_buscar = ctk.CTkEntry(
            barra_acciones,
            placeholder_text="Buscar por nombre o DUI",
            width=260,
            height=34,
            font=ctk.CTkFont(size=14),
            justify="left",
        )
        self.entrada_buscar.grid(row=0, column=1, padx=(0, 10), pady=16, sticky="w")
        self.entrada_buscar.bind("<KeyRelease>", self.buscar_por_escritura)

        boton_agregar = boton_buscar_profesional(
            barra_acciones,
            text="Añadir Personal",
            width=140,
            height=34,
            command=lambda: self.agregar_personal(frame),
            fg_color="#0F766E",
            hover_color="#115E59",
        )
        boton_agregar.grid(row=0, column=2, padx=(0, 10), pady=16, sticky="w")

        boton_excel = boton_buscar_profesional(
            barra_acciones,
            text="Excel",
            width=110,
            height=34,
            command=self.exportar_personal_excel,
            fg_color="#10b981",
            hover_color="#059669",
        )
        boton_excel.grid(row=0, column=3, padx=(0, 12), pady=16, sticky="w")

        self.label_total_personal = ctk.CTkLabel(
            barra_acciones,
            text="Total de empleados: 0",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="gray75",
            anchor="e",
        )
        self.label_total_personal.grid(row=0, column=4, padx=(12, 18), pady=16, sticky="e")

        frame_tabla = ctk.CTkFrame(frame, corner_radius=18, border_width=1, border_color="#3f3f46")
        frame_tabla.place(relx=0.03, rely=0.30, relwidth=0.94, relheight=0.66)

        ctk.CTkLabel(
            frame_tabla,
            text="Listado de empleados",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w",
        ).place(relx=0.02, rely=0.03)

        ctk.CTkLabel(
            frame_tabla,
            text="Click derecho sobre una fila para editar o eliminar.",
            font=ctk.CTkFont(size=11),
            text_color="gray70",
            anchor="w",
        ).place(relx=0.02, rely=0.10)

        contenedor_tabla = ctk.CTkFrame(frame_tabla, fg_color="transparent")
        contenedor_tabla.place(relx=0.02, rely=0.18, relwidth=0.96, relheight=0.78)

        columnas = ("ID", "Nombre", "DUI", "Correo", "Direccion", "Telefono", "Puesto", "Departamento", "Salario", "Ingreso  ", "Nacimiento")

        self.tabla_personal = ttk.Treeview(contenedor_tabla, columns=columnas, show="headings", style="Personal.Treeview")
        self.tabla_personal.grid(row=0, column=0, sticky="nsew")

        self.tabla_personal.heading("ID", text="ID")
        self.tabla_personal.heading("Nombre", text="Nombre")
        self.tabla_personal.heading("DUI", text="DUI")
        self.tabla_personal.heading("Correo", text="Correo")
        self.tabla_personal.heading("Direccion", text="Direccion")
        self.tabla_personal.heading("Telefono", text="Telefono")
        self.tabla_personal.heading("Puesto", text="Puesto")
        self.tabla_personal.heading("Departamento", text="Dpto")
        self.tabla_personal.heading("Salario", text="Salario")
        self.tabla_personal.heading("Ingreso  ", text="Ingreso  ")
        self.tabla_personal.heading("Nacimiento", text="Nacimiento")

        for columna in columnas:
            self.tabla_personal.column(columna, anchor="center", width=120, stretch=True)

        scrollbar_vertical = ttk.Scrollbar(contenedor_tabla, orient="vertical", command=self.tabla_personal.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        self.tabla_personal.configure(yscrollcommand=scrollbar_vertical.set)

        scrollbar_horizontal = ttk.Scrollbar(contenedor_tabla, orient="horizontal", command=self.tabla_personal.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.tabla_personal.configure(xscrollcommand=scrollbar_horizontal.set)

        contenedor_tabla.grid_rowconfigure(0, weight=1)
        contenedor_tabla.grid_columnconfigure(0, weight=1)

        def _ajustar_columnas_personal(event, tabla=self.tabla_personal, columnas=columnas):
            total_width = event.width
            scrollbar_space = 18
            usable = max(total_width - scrollbar_space, 100)
            col_width = max(int(usable / max(len(columnas), 1)), 50)
            for col in columnas:
                tabla.column(col, width=col_width)

        contenedor_tabla.bind("<Configure>", _ajustar_columnas_personal)

        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure(
            "Personal.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background="#0f766e",
            foreground="white",
            relief="flat",
            padding=8,
        )
        estilo.configure(
            "Personal.Treeview",
            font=("Segoe UI", 10),
            background="#f8fafc",
            foreground="#0f172a",
            rowheight=34,
            fieldbackground="#f8fafc",
            borderwidth=0,
        )
        estilo.map("Personal.Treeview", background=[("selected", "#cfe8e5")], foreground=[("selected", "#0f172a")])

        self.tabla_personal.tag_configure("oddrow", background="#f8fafc")
        self.tabla_personal.tag_configure("evenrow", background="#eef6f5")
        self.cargar_datos()
        self.crear_menu(contenedor_tabla)
        self.entrada_buscar.focus_set()

    def _actualizar_departamentos_combo(self):
        if not hasattr(self, "combo_departamento") or self.combo_departamento is None:
            return

        try:
            empleados = db.gestion_empleado().obtener_empleados() if db and db.gestion_empleado else []
        except Exception:
            empleados = []

        departamentos = []
        for empleado in empleados:
            if len(empleado) > 7:
                depto = str(empleado[7]).strip()
                if depto and depto not in departamentos:
                    departamentos.append(depto)

        departamentos = sorted(departamentos)
        valores = ["Todos"] + departamentos
        actual = self.combo_departamento.get() if self.combo_departamento.winfo_exists() else "Todos"

        self.combo_departamento.configure(values=valores)
        if actual in valores:
            self.combo_departamento.set(actual)
        else:
            self.combo_departamento.set("Todos")

    def _filtrar_por_departamento(self, valor=None):
        texto_busqueda = self.entrada_buscar.get().strip().lower()
        departamento = self.combo_departamento.get() if hasattr(self, "combo_departamento") else "Todos"
        self.cargar_datos(texto_busqueda, departamento)

    def buscar_por_escritura(self, event=None):
        texto_busqueda = self.entrada_buscar.get().strip().lower()
        departamento = self.combo_departamento.get() if hasattr(self, "combo_departamento") else "Todos"
        self.cargar_datos(texto_busqueda, departamento)

    def crear_menu(self, frame):
        self.menu = tk.Menu(
            frame,
            tearoff=0,
            bg="#f8fafc",
            fg="#0f172a",
            activebackground="#d1fae5",
            activeforeground="#0f172a",
            bd=1,
            relief="solid",
        )
        self.menu.configure(font=("Segoe UI", 11))
        self.menu.add_command(
            label="Editar empleado",
            font=("Segoe UI", 11),
            command=lambda: self.editar_personal(self.frame_padre),
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="Eliminar empleado",
            font=("Segoe UI", 11),
            command=self.eliminar,
        )
        self.tabla_personal.bind("<Button-3>", self.mostrar_menu)

    def mostrar_menu(self, event):
        item = self.tabla_personal.identify_row(event.y)
        if not item:
            return

        self.tabla_personal.selection_set(item)
        self.tabla_personal.focus(item)
        self.menu.tk_popup(event.x_root, event.y_root)

    def exportar_personal_excel(self):
        """Exporta el listado completo de empleados a un archivo Excel profesional."""
        try:
            empleados = db.gestion_empleado().obtener_empleados() if db and db.gestion_empleado else []
            if not empleados:
                messagebox.showwarning("Exportar personal", "No hay empleados para exportar.")
                return

            empresa = datos_empresa or {}
            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

            wb = Workbook()
            ws = wb.active
            ws.title = "Personal"

            for letra in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
                ws.column_dimensions[letra].width = 16
            ws.column_dimensions["A"].width = 8
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 15
            ws.column_dimensions["D"].width = 26
            ws.column_dimensions["E"].width = 28
            ws.column_dimensions["F"].width = 16
            ws.column_dimensions["G"].width = 18
            ws.column_dimensions["H"].width = 18
            ws.column_dimensions["I"].width = 14
            ws.column_dimensions["J"].width = 14
            ws.column_dimensions["K"].width = 14

            color_primario = "0F766E"
            color_fondo_totales = "ECFDF5"
            color_texto_titulo = "FFFFFF"

            font_titulo = Font(name="Segoe UI", size=16, bold=True, color=color_texto_titulo)
            font_encabezado = Font(name="Segoe UI", size=10, bold=True, color=color_texto_titulo)
            font_normal = Font(name="Segoe UI", size=9)
            font_total = Font(name="Segoe UI", size=10, bold=True)

            fill_primario = PatternFill(start_color=color_primario, end_color=color_primario, fill_type="solid")
            fill_total = PatternFill(start_color=color_fondo_totales, end_color=color_fondo_totales, fill_type="solid")

            alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            alignment_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
            border_thin = Border(
                left=Side(style="thin", color="D0D7DE"),
                right=Side(style="thin", color="D0D7DE"),
                top=Side(style="thin", color="D0D7DE"),
                bottom=Side(style="thin", color="D0D7DE"),
            )

            fila = 1
            ws.merge_cells(f"A{fila}:K{fila}")
            ws[f"A{fila}"] = "REPORTE GENERAL DE PERSONAL"
            ws[f"A{fila}"].font = font_titulo
            ws[f"A{fila}"].fill = fill_primario
            ws[f"A{fila}"].alignment = alignment_center
            ws.row_dimensions[fila].height = 26

            fila += 1
            ws.merge_cells(f"A{fila}:K{fila}")
            ws[f"A{fila}"] = str(empresa.get("nombre") or "Empresa")
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=12, bold=True)
            ws[f"A{fila}"].alignment = alignment_center

            fila += 1
            ws.merge_cells(f"A{fila}:K{fila}")
            ws[f"A{fila}"] = (
                f"RUC: {empresa.get('ruc', '')} | Dirección: {empresa.get('direccion', '')} | Teléfono: {empresa.get('telefono', '')}"
            )
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=9, color="525252")
            ws[f"A{fila}"].alignment = alignment_center

            fila += 1
            ws.merge_cells(f"A{fila}:K{fila}")
            ws[f"A{fila}"] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=9, color="525252")
            ws[f"A{fila}"].alignment = alignment_left

            fila += 2
            encabezados = ["ID", "Nombre", "DUI", "Correo", "Dirección", "Teléfono", "Puesto", "Departamento", "Salario", "Ingreso", "Nacimiento"]
            for col_idx, titulo in enumerate(encabezados, start=1):
                celda = ws.cell(row=fila, column=col_idx)
                celda.value = titulo
                celda.font = font_encabezado
                celda.fill = fill_primario
                celda.alignment = alignment_center
                celda.border = border_thin

            ws.row_dimensions[fila].height = 22
            fila += 1

            for empleado in empleados:
                fila_empleado = list(empleado[:11])
                if len(fila_empleado) < 11:
                    fila_empleado.extend([""] * (11 - len(fila_empleado)))

                for col_idx, valor in enumerate(fila_empleado, start=1):
                    celda = ws.cell(row=fila, column=col_idx)
                    celda.value = valor
                    celda.font = font_normal
                    celda.border = border_thin
                    if col_idx == 1:
                        celda.alignment = alignment_center
                    elif col_idx in {2, 4, 5, 6, 7, 8, 9, 10, 11}:
                        celda.alignment = alignment_left
                    else:
                        celda.alignment = alignment_center

                fila += 1

            fila_total = fila
            ws.merge_cells(f"A{fila_total}:J{fila_total}")
            ws[f"A{fila_total}"] = f"TOTAL DE EMPLEADOS: {len(empleados)}"
            ws[f"A{fila_total}"].font = font_total
            ws[f"A{fila_total}"].fill = fill_total
            ws[f"A{fila_total}"].alignment = alignment_center
            ws[f"A{fila_total}"].border = border_thin
            ws[f"K{fila_total}"] = len(empleados)
            ws[f"K{fila_total}"].font = font_total
            ws[f"K{fila_total}"].fill = fill_total
            ws[f"K{fila_total}"].alignment = alignment_center
            ws[f"K{fila_total}"].border = border_thin

            ws.freeze_panes = "A7"

            ruta = filedialog.asksaveasfilename(
                title="Guardar personal en Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("Todos", "*.*")],
                initialfile=f"personal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            )

            if not ruta:
                return

            carpeta_base = os.path.dirname(os.path.abspath(ruta)) or os.getcwd()
            try:
                if not os.path.isdir(carpeta_base) or not os.access(carpeta_base, os.W_OK):
                    raise PermissionError
                wb.save(ruta)
                ruta_guardada = ruta
            except PermissionError:
                carpeta_fallback = os.path.join(os.getcwd(), "exportaciones")
                os.makedirs(carpeta_fallback, exist_ok=True)
                nombre_archivo = os.path.basename(ruta) or f"personal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ruta_guardada = os.path.join(carpeta_fallback, nombre_archivo)
                wb.save(ruta_guardada)
                messagebox.showwarning(
                    "Ruta no escribible",
                    "La carpeta seleccionada no permite guardar archivos. Se guardó el Excel en la carpeta local de exportaciones.",
                    parent=self.frame_padre if hasattr(self, "frame_padre") else None,
                )

            ruta_abs = os.path.abspath(ruta_guardada)
            messagebox.showinfo(
                "Exportación completada",
                f"El reporte de personal se guardó en:\n{ruta_abs}",
                parent=self.frame_padre if hasattr(self, "frame_padre") else None,
            )
            if sys.platform.startswith("win"):
                os.startfile(ruta_abs)

        except Exception as exc:
            logging.exception("Error al exportar el personal a Excel")
            messagebox.showerror("Error al Exportar", f"No se pudo exportar el personal:\n{exc}")

    def cargar_datos(self, texto_busqueda="", departamento="Todos"):
        self._actualizar_departamentos_combo()

        for fila in self.tabla_personal.get_children():
            self.tabla_personal.delete(fila)

        datos = db.gestion_empleado().obtener_empleados()
        texto_busqueda = texto_busqueda.strip().lower()
        departamento_filtro = (departamento or "Todos").strip()
        total_visibles = 0

        for indice, fila in enumerate(datos):
            nombre = str(fila[1]).lower() if len(fila) > 1 else ""
            dui = str(fila[2]).lower() if len(fila) > 2 else ""
            depto = str(fila[7]).strip() if len(fila) > 7 else ""

            coincide_texto = (not texto_busqueda) or (texto_busqueda in nombre) or (texto_busqueda in dui)
            coincide_departamento = (departamento_filtro == "Todos") or (depto == departamento_filtro)

            if coincide_texto and coincide_departamento:
                tag = "evenrow" if indice % 2 == 0 else "oddrow"
                self.tabla_personal.insert("", tk.END, values=fila, tags=(tag,))
                total_visibles += 1

        if hasattr(self, "label_total_personal") and self.label_total_personal.winfo_exists():
            self.label_total_personal.configure(text=f"Total de empleados: {total_visibles}")

    def eliminar(self):
        seleccion = self.tabla_personal.focus()
        if not seleccion:
            messagebox.showwarning("Aviso", "Seleccione un empleado para eliminar.")
            return

        datos = self.tabla_personal.item(seleccion, "values")
        if not datos:
            messagebox.showwarning("Aviso", "Seleccione un empleado para eliminar.")
            return

        valores = datos
        id_registro = valores[0]
        if not messagebox.askyesno("Eliminar", f"¿Deseas eliminar a {valores[1]}?"):
            return

        data = db.gestion_empleado().eliminar_empleado(id_registro)
        if data:
            messagebox.showinfo("Exito", f"{valores[1]} eliminado exitosamente")
            departamento_actual = self.combo_departamento.get() if hasattr(self, "combo_departamento") else "Todos"
            self.cargar_datos(self.entrada_buscar.get(), departamento_actual)
        else:
            messagebox.showerror("Error", "No se pudo eliminar el empleado")

    def _parsear_fecha_formulario(self, valor):
        if hasattr(valor, "strftime"):
            return valor

        texto = str(valor or "").strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
        return None

    def _crear_formulario_personal_modal(self, parent, titulo_modal, subtitulo, texto_accion, comando_accion, comando_secundario):
        ancho = 760
        alto = 700
        x = (parent.winfo_screenwidth() - ancho) // 2
        y = (parent.winfo_screenheight() - alto) // 2

        top_level = ctk.CTkToplevel(parent)
        top_level.geometry(f"{ancho}x{alto}+{x}+{y}")
        top_level.grab_set()
        top_level.title(titulo_modal)
        top_level.resizable(False, False)

        contenedor = ctk.CTkFrame(top_level, corner_radius=18, border_width=1, border_color="#3f3f46")
        contenedor.place(relx=0.03, rely=0.04, relwidth=0.94, relheight=0.92)

        ctk.CTkLabel(
            contenedor,
            text=titulo_modal,
            font=ctk.CTkFont(size=24, weight="bold"),
            anchor="w",
        ).place(relx=0.04, rely=0.04)
        ctk.CTkLabel(
            contenedor,
            text=subtitulo,
            font=ctk.CTkFont(size=12),
            text_color="gray75",
            anchor="w",
        ).place(relx=0.04, rely=0.10)

        tarjeta = ctk.CTkFrame(contenedor, corner_radius=14, border_width=1, border_color="#334155")
        tarjeta.place(relx=0.04, rely=0.16, relwidth=0.92, relheight=0.66)
        tarjeta.grid_columnconfigure(0, weight=0)
        tarjeta.grid_columnconfigure(1, weight=1)
        tarjeta.grid_columnconfigure(2, weight=0)
        tarjeta.grid_columnconfigure(3, weight=1)

        campos = [
            ("Nombre Completo", "nombre"),
            ("Número de Identificación", "dui"),
            ("Correo Electrónico", "correo"),
            ("Teléfono", "telefono"),
            ("Dirección", "direccion"),
            ("Cargo", "cargo"),
            ("Fecha de Ingreso", "ingreso"),
            ("Fecha de Nacimiento", "nacimiento"),
            ("Departamento", "departamento"),
            ("Salario", "salario"),
        ]

        controles = {}
        cargos = self.cargo()
        departamentos = self.departamento()
        salarios = self.salario()

        for indice, (texto_label, clave) in enumerate(campos):
            fila = indice // 2
            bloque = (indice % 2) * 2

            ctk.CTkLabel(
                tarjeta,
                text=f"{texto_label}:",
                font=ctk.CTkFont(size=13, weight="bold"),
                anchor="w",
            ).grid(row=fila, column=bloque, sticky="w", padx=(18, 10), pady=(16 if fila == 0 else 10, 0))

            if clave in ("ingreso", "nacimiento"):
                control = DateEntry(tarjeta, font=("Segoe UI", 12), width=18, date_pattern="dd/MM/yyyy")
                control.grid(row=fila, column=bloque + 1, sticky="ew", padx=(0, 18), pady=(16 if fila == 0 else 10, 0))
            elif clave == "cargo":
                control = ctk.CTkComboBox(tarjeta, values=cargos, width=260)
                control.grid(row=fila, column=bloque + 1, sticky="ew", padx=(0, 18), pady=(16 if fila == 0 else 10, 0))
            elif clave == "departamento":
                control = ctk.CTkComboBox(tarjeta, values=departamentos, width=260)
                control.grid(row=fila, column=bloque + 1, sticky="ew", padx=(0, 18), pady=(16 if fila == 0 else 10, 0))
            elif clave == "salario":
                control = ctk.CTkComboBox(tarjeta, values=salarios, width=260)
                control.grid(row=fila, column=bloque + 1, sticky="ew", padx=(0, 18), pady=(16 if fila == 0 else 10, 0))
            else:
                control = ctk.CTkEntry(tarjeta, width=260, height=34, justify="left", font=ctk.CTkFont(size=13))
                control.grid(row=fila, column=bloque + 1, sticky="ew", padx=(0, 18), pady=(16 if fila == 0 else 10, 0))

            controles[clave] = control

        panel_info = ctk.CTkFrame(contenedor, corner_radius=12, border_width=1, border_color="#334155")
        panel_info.place(relx=0.04, rely=0.84, relwidth=0.50, relheight=0.10)
        ctk.CTkLabel(
            panel_info,
            text="Consejo",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).place(relx=0.04, rely=0.18)
        ctk.CTkLabel(
            panel_info,
            text="Completa todos los campos principales para mantener la ficha del empleado consistente.",
            font=ctk.CTkFont(size=11),
            text_color="gray75",
            wraplength=300,
            justify="left",
            anchor="w",
        ).place(relx=0.04, rely=0.48)

        barra_botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        barra_botones.place(relx=0.58, rely=0.84, relwidth=0.38, relheight=0.10)
        boton_secundario = boton_buscar_profesional(
            barra_botones,
            text="Limpiar",
            command=comando_secundario,
            width=120,
            height=36,
            fg_color="#4b5563",
            hover_color="#374151",
        )
        boton_secundario.pack(side="left", padx=(0, 10), pady=10)
        boton_primario = boton_buscar_profesional(
            barra_botones,
            text=texto_accion,
            command=comando_accion,
            width=150,
            height=36,
            fg_color="#0F766E",
            hover_color="#115E59",
        )
        boton_primario.pack(side="left", pady=10)

        if "nombre" in controles:
            controles["nombre"].focus_set()

        return top_level, controles


    def agregar_personal(self, sipp):
        parent = sipp if sipp is not None else getattr(self, "frame_padre", None)
        if parent is None:
            parent = self.tabla_personal.winfo_toplevel()
        top_level = None
        controles = None

        def limpiar_formulario():
            if top_level and top_level.winfo_exists():
                top_level.destroy()
            self.agregar_personal(parent)

        def guardar_formulario():
            self.guardar_personal(
                controles["nombre"].get().title(),
                controles["dui"].get(),
                controles["correo"].get(),
                controles["direccion"].get(),
                controles["telefono"].get(),
                controles["cargo"].get(),
                controles["departamento"].get(),
                controles["salario"].get(),
                controles["ingreso"].get(),
                controles["nacimiento"].get(),
                parent,
                top_level,
            )

        top_level, controles = self._crear_formulario_personal_modal(
            parent,
            "Añadir Nuevo Personal",
            "Crea una ficha completa de empleado con información laboral y personal.",
            "Guardar",
            guardar_formulario,
            limpiar_formulario,
        )
        
    def guardar_personal(self, Nombre, Dui, Correo, Direccion, Telefono, Puesto, Departamento, Salario, Ingreso, Nacimiento, sipp, top_level, id_empleado=None, editar=False):
        datos = db.gestion_empleado()
        if Nombre == "" or Dui == "" or Correo == "" or Direccion == "" or Telefono == "":
            messagebox.showwarning("Aviso", "No pueden existir campos vacios")
            return
        if not datos:
            messagebox.showerror("Error", f"Error con base de dato: {Nombre}, no se ha podido guardar empleado.")
            return

        if editar and id_empleado is not None:
            actualizado = datos.actualizar_empleado(id_empleado, Nombre, Dui, Correo, Direccion, Telefono,
                                                    Puesto, Departamento, Salario, Ingreso, Nacimiento)
            if actualizado:
                messagebox.showinfo("Exito", f"El empleado {Nombre} ha sido actualizado correctamente.")
                top_level.destroy()
                departamento_actual = self.combo_departamento.get() if hasattr(self, "combo_departamento") else "Todos"
                self.cargar_datos(self.entrada_buscar.get(), departamento_actual)
            else:
                messagebox.showerror("Error", "No se pudo actualizar el empleado.")
            return

        datos.insertar_empleados(Nombre, Dui, Correo, Direccion, Telefono, Puesto, Departamento, Salario, Ingreso, Nacimiento)
        messagebox.showinfo("Exito", f"El empleado {Nombre}, ha sido ingresado con exito.")
        if messagebox.askyesno("Añadir Nuevo Personal", "Deseas añadir un nuevo personal"):
            self.limpiar(sipp, top_level)
        else:
            top_level.destroy()

        departamento_actual = self.combo_departamento.get() if hasattr(self, "combo_departamento") else "Todos"
        self.cargar_datos(self.entrada_buscar.get(), departamento_actual)

    def limpiar(self, sipp, top_level):
         top_level.destroy()
         self.agregar_personal(sipp)

    def editar_personal(self, sipp):
        if not self.tabla_personal.get_children():
            messagebox.showerror("Error", "La tabla no tiene datos.")
            return

        seleccion = self.tabla_personal.focus()
        if not seleccion:
            messagebox.showwarning("Aviso", "Seleccione un empleado para editar.")
            return

        datos = self.tabla_personal.item(seleccion, "values")
        if not datos:
            messagebox.showwarning("Aviso", "Seleccione un empleado para editar.")
            return

        parent = sipp if sipp is not None else getattr(self, "frame_padre", None)
        if parent is None:
            parent = self.tabla_personal.winfo_toplevel()
        valores = datos
        id_registro = valores[0]
        date = db.gestion_empleado().obtener_empleado_id(id_registro)

        top_level = None
        controles = None

        def limpiar_formulario():
            if top_level and top_level.winfo_exists():
                top_level.destroy()
            self.editar_personal(parent)

        def actualizar_formulario():
            self.guardar_personal(
                controles["nombre"].get(),
                controles["dui"].get(),
                controles["correo"].get(),
                controles["direccion"].get(),
                controles["telefono"].get(),
                controles["cargo"].get(),
                controles["departamento"].get(),
                controles["salario"].get(),
                controles["ingreso"].get(),
                controles["nacimiento"].get(),
                parent,
                top_level,
                id_empleado=id_registro,
                editar=True,
            )

        top_level, controles = self._crear_formulario_personal_modal(
            parent,
            "Editar Personal",
            "Actualiza la ficha del empleado y conserva la información laboral consistente.",
            "Actualizar",
            actualizar_formulario,
            limpiar_formulario,
        )

        if date is not None:
            controles["nombre"].insert(0, date[1] if len(date) > 1 else "")
            controles["dui"].insert(0, date[2] if len(date) > 2 else "")
            controles["correo"].insert(0, date[3] if len(date) > 3 else "")
            controles["direccion"].insert(0, date[4] if len(date) > 4 else "")
            controles["telefono"].insert(0, date[5] if len(date) > 5 else "")
            if len(date) > 6:
                controles["cargo"].set(date[6])
            if len(date) > 7:
                controles["departamento"].set(date[7])
            if len(date) > 8:
                controles["salario"].set(date[8])
            fecha_ingreso = self._parsear_fecha_formulario(date[9] if len(date) > 9 else "")
            fecha_nacimiento = self._parsear_fecha_formulario(date[10] if len(date) > 10 else "")
            if fecha_ingreso:
                controles["ingreso"].set_date(fecha_ingreso)
            if fecha_nacimiento:
                controles["nacimiento"].set_date(fecha_nacimiento)
        else:
            messagebox.showerror("Error", "Ocurrio un Error al cargar datos")

    def departamento(self):
        departamentos = [
    "Gerencia General / Dirección",
    "Administración",
    "Recursos Humanos (RRHH)",
    "Contabilidad y Finanzas",
    "Ventas / Comercial",
    "Marketing y Publicidad",
    "Atención al Cliente / Servicio al Cliente",
    "Operaciones / Producción",
    "Logística y Almacén",
    "Compras",
    "Tecnología / Sistemas (IT)",
    "Legal / Jurídico",
    "Seguridad",
    "Mantenimiento",
    "Investigación y Desarrollo (I+D)"
]
        return departamentos

    def cargo(self):
        cargos = [
    "Director General (CEO) / Gerente General", "Subgerente General", "Asistente de Gerencia",
    "Secretario(a) Ejecutivo(a)", "Administrador(a)", "Jefe de Administración",
    "Asistente Administrativo", "Recepcionista", "Archivista",

    "Gerente de Recursos Humanos", "Reclutador(a)", "Analista de Recursos Humanos",
    "Encargado(a) de Nómina", "Coordinador(a) de Capacitación",

    "Director Financiero (CFO)", "Contador(a) General", "Auxiliar Contable",
    "Auditor(a)", "Analista Financiero", "Tesorero(a)",
    "Encargado(a) de Cuentas por Cobrar", "Encargado(a) de Cuentas por Pagar", "Cajero(a)",

    "Gerente de Ventas", "Supervisor(a) de Ventas", "Ejecutivo(a) de Ventas",
    "Vendedor(a)", "Asesor Comercial",

    "Gerente de Marketing", "Community Manager", "Diseñador(a) Gráfico",
    "Creador(a) de Contenido", "Especialista en Publicidad Digital",

    "Representante de Atención al Cliente", "Operador(a) de Call Center",
    "Supervisor(a) de Atención al Cliente",

    "Gerente de Operaciones (COO)", "Supervisor(a) de Producción", "Operario(a)",
    "Encargado(a) de Planta", "Inspector(a) de Control de Calidad", "Técnico(a) de Producción",

    "Jefe de Logística", "Coordinador(a) de Almacén", "Encargado(a) de Inventario",
    "Bodeguero(a)", "Despachador(a)", "Montacarguista",
    "Conductor / Chofer", "Repartidor(a)",

    "Jefe de Compras", "Analista de Compras", "Encargado(a) de Proveedores",

    "Director de Tecnología (CTO)", "Jefe de Sistemas", "Técnico de Soporte IT",
    "Administrador(a) de Redes", "Programador(a) / Desarrollador(a)", "Analista de Sistemas",
    "Encargado(a) de Ciberseguridad",

    "Abogado(a) Corporativo", "Asesor(a) Legal",

    "Jefe de Seguridad", "Guardia de Seguridad", "Supervisor(a) de Seguridad",

    "Jefe de Mantenimiento", "Técnico(a) de Mantenimiento", "Electricista",
    "Plomero", "Mecánico",

    "Director(a) de Innovación", "Investigador(a)", "Analista de Desarrollo"
]
        return cargos

    def salario(self):
        salarios_mensuales_panama = [
    "B/. 320.00",
    "B/. 350.00",
    "B/. 636.80",
    "B/. 700.00",
    "B/. 800.00",
    "B/. 900.00",
    "B/. 1,000.00",
    "B/. 1,100.00",
    "B/. 1,200.00",
    "B/. 1,400.00",
    "B/. 1,600.00",
    "B/. 1,800.00",
    "B/. 2,000.00",
    "B/. 2,300.00",
    "B/. 2,500.00",
    "B/. 2,800.00",
    "B/. 3,000.00",
    "B/. 3,500.00"
]
        return salarios_mensuales_panama
        


class Gestion_Planilla:

    def __init__(self):
        self.tabla_planilla = None
        self.entrada_buscar = None
        self.total_empleados_label = None
        self.total_registros_label = None
        self.total_neto_label = None
        self.actualizacion_label = None
        self.etiqueta_corte_activo = None

    def ventana_gestion_planilla(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

        frame.configure(fg_color="transparent")

        header = ctk.CTkFrame(
            frame,
            corner_radius=24,
            fg_color=("#f8fafc", "#0f172a"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
        )
        header.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.22)

        # Línea decorativa superior
        ctk.CTkFrame(
            header,
            height=3,
            fg_color=("#94a3b8", "#64748b"),
            corner_radius=2,
        ).place(relx=0.0, rely=0.0, relwidth=1.0)

        # Título principal con icono
        ctk.CTkLabel(
            header,
            text="📋 Gestión de Planilla",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=("#334155", "#cbd5e1"),
            anchor="w",
        ).place(relx=0.03, rely=0.15)

        ctk.CTkLabel(
            header,
            text="Administre las nóminas de su empresa de forma rápida y segura",
            font=ctk.CTkFont(size=13),
            text_color="#94a3b8",
            anchor="w",
            wraplength=700,
            justify="left",
        ).place(relx=0.03, rely=0.60)

        ctk.CTkButton(
            header,
            text="Décimo 3er mes",
            width=145,
            height=34,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#0f766e",
            hover_color="#115e59",
            command=self.toplevel_decimo_tercer_mes,
        ).place(relx=0.97, rely=0.18, anchor="e")

        panel = ctk.CTkFrame(
            frame,
            corner_radius=24,
            fg_color=("#f8fafc", "#0f172a"),
            border_width=1,
            border_color=("#e2e8f0", "#334155"),
        )
        panel.place(relx=0.02, rely=0.28, relwidth=0.96, relheight=0.62)

        ctk.CTkLabel(
            panel,
            text="Vista de planilla",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).place(relx=0.04, rely=0.08)

        ctk.CTkLabel(
            panel,
            text="Aquí solo queda la interfaz: búsqueda, resumen y tablero visual, sin cálculos ni conexión a base de datos.",
            font=ctk.CTkFont(size=13),
            text_color="#94a3b8",
            wraplength=700,
            justify="left",
            anchor="w",
        ).place(relx=0.04, rely=0.20)

        # Barra de herramientas con búsqueda y botones
        toolbar = ctk.CTkFrame(panel, fg_color="transparent")
        toolbar.place(relx=0.04, rely=0.30, relwidth=0.92, relheight=0.08)

        self.combo_departamento_planilla = ctk.CTkComboBox(
            toolbar,
            values=["Todos"],
            width=150,
            height=32,
            font=ctk.CTkFont(size=11),
            state="readonly",
            command=lambda valor: self.cargar_planilla(
                self.entrada_buscar.get(), valor
            ),
        )
        self.combo_departamento_planilla.set("Todos")
        self.combo_departamento_planilla.place(relx=0.0, rely=0.0, relwidth=0.20, relheight=1.0)

        self.entrada_buscar = ctk.CTkEntry(
            toolbar,
            placeholder_text="Buscar...",
            font=ctk.CTkFont(size=11),
            height=32,
        )
        self.entrada_buscar.place(relx=0.21, rely=0.0, relwidth=0.25, relheight=1.0)
        self.entrada_buscar.bind(
            "<KeyRelease>",
            lambda event: self.cargar_planilla(
                self.entrada_buscar.get(), self.combo_departamento_planilla.get()
            ),
        )

        self.combo_corte_planilla = ctk.CTkComboBox(
            toolbar,
            values=["Corte activo"],
            width=190,
            height=32,
            font=ctk.CTkFont(size=11),
            state="readonly",
            command=lambda valor: self.cargar_planilla(
                self.entrada_buscar.get(),
                self.combo_departamento_planilla.get(),
                valor,
            ),
        )
        self.combo_corte_planilla.set("Corte activo")
        self.combo_corte_planilla.place(relx=0.47, rely=0.0, relwidth=0.22, relheight=1.0)

        # Botón Corte de Planilla
        ctk.CTkButton(
            toolbar,
            text="Corte de Planilla",
            width=100,
            height=32,
            font=ctk.CTkFont(size=11),
            fg_color="#8b5cf6",
            hover_color="#7c3aed",
            command=self.toplevel_corte_planilla,
        ).place(relx=0.70, rely=0.0, relwidth=0.10, relheight=1.0)

        ctk.CTkButton(
            toolbar,
            text="Comprobantes",
            width=100,
            height=32,
            font=ctk.CTkFont(size=11),
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.toplevel_comprobantes_pago,
        ).place(relx=0.81, rely=0.0, relwidth=0.09, relheight=1.0)

        # Botón Excel
        ctk.CTkButton(
            toolbar,
            text="Excel",
            width=80,
            height=32,
            font=ctk.CTkFont(size=11),
            fg_color="#10b981",
            hover_color="#059669",
            command=self.exportar_planilla_excel,
        ).place(relx=0.91, rely=0.0, relwidth=0.09, relheight=1.0)

        # Frame para la tabla
        tabla_frame = ctk.CTkFrame(panel, corner_radius=18, border_width=1, border_color="#334155")
        tabla_frame.place(relx=0.04, rely=0.40, relwidth=0.92, relheight=0.46)

        # Definir las columnas de la tabla
        columnas = (
            "ID", "Nombre Completo", "DUI", "Cargo", "Salario",
            "Por Horas", "Horas Trabajadas", "Días", "Total por Hora", "Horas Extras",
            "Total Extras", "Remuneraciones", "Subtotal", "CSS", "Seg. Educativo",
            "ISR", "Otros descuentos", "Total Deducciones", "Salario Neto"
        )

        # Crear Treeview con scroll horizontal y vertical
        tree_frame = ctk.CTkFrame(tabla_frame, fg_color="transparent")
        tree_frame.place(relx=0.04, rely=0.08, relwidth=0.92, relheight=0.88)

        # Crear Treeview
        self.tabla_planilla = ttk.Treeview(
            tree_frame,
            columns=columnas,
            height=15,
            show="headings"
        )

        # Definir encabezados y ancho de columnas
        anchos_columnas = {
            "ID": 50,
            "Nombre Completo": 150,
            "DUI": 90,
            "Cargo": 100,
            "Salario": 100,
            "Por Horas": 80,
            "Horas Trabajadas": 110,
            "Días": 60,
            "Total por Hora": 100,
            "Horas Extras": 100,
            "Total Extras": 100,
            "Remuneraciones": 110,
            "Subtotal": 90,
            "CSS": 80,
            "Seg. Educativo": 110,
            "ISR": 80,
            "Otros descuentos": 120,
            "Total Deducciones": 120,
            "Salario Neto": 100
        }
        for col in columnas:
            ancho = anchos_columnas.get(col, 80)
            self.tabla_planilla.column(col, width=ancho, anchor="center")
            self.tabla_planilla.heading(col, text=col, anchor="center")

        # Scroll vertical
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tabla_planilla.yview)
        self.tabla_planilla.configure(yscroll=scroll_y.set)

        # Scroll horizontal
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tabla_planilla.xview)
        self.tabla_planilla.configure(xscroll=scroll_x.set)

        # Colocar elementos
        self.tabla_planilla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Aplicar estilo a la tabla
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#1e293b",
            foreground="#e2e8f0",
            fieldbackground="#1e293b",
            rowheight=25,
            font=("Segoe UI", 9)
        )
        style.configure(
            "Treeview.Heading",
            background="#334155",
            foreground="#e2e8f0",
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview", background=[("selected", "#3b82f6")])

        self.cargar_planilla()

    def cargar_planilla(self, filtro=None, departamento="Todos", corte_seleccionado=None):
        if self.tabla_planilla is None:
            return

        for item in self.tabla_planilla.get_children():
            self.tabla_planilla.delete(item)

        try:
            empleados = db.gestion_empleado().obtener_empleados()
        except Exception as exc:
            logging.exception("No se pudo cargar el personal en la planilla")
            messagebox.showerror("Error", f"No se pudo cargar el personal: {exc}")
            return

        cortes_mapa = {}
        corte_activo = None
        valor_inicial = None
        try:
            cortes_guardados = db.cortes_planilla().obtener_cortes_planilla()
            corte_activo = db.cortes_planilla().obtener_corte_activo()

            def formatear_fecha_corte(fecha):
                if hasattr(fecha, "strftime"):
                    return fecha.strftime("%d/%m/%Y")
                try:
                    return datetime.strptime(str(fecha)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
                except ValueError:
                    return str(fecha)

            valores_cortes = []
            for corte in sorted(cortes_guardados or [], key=lambda dato: dato[0], reverse=True):
                corte_id, fecha_inicio, fecha_fin, estado = corte[:4]
                texto_corte = (
                    f"Corte {corte_id}: {formatear_fecha_corte(fecha_inicio)} "
                    f"a {formatear_fecha_corte(fecha_fin)} ({estado})"
                )
                valores_cortes.append(texto_corte)
                cortes_mapa[texto_corte] = (fecha_inicio, fecha_fin, estado, corte_id)

            if corte_seleccionado in cortes_mapa:
                valor_inicial = corte_seleccionado
            elif corte_activo:
                for texto_corte, datos_corte in cortes_mapa.items():
                    if datos_corte[3] == corte_activo[0]:
                        valor_inicial = texto_corte
                        break
            if valor_inicial is None and valores_cortes:
                valor_inicial = valores_cortes[0]

            if hasattr(self, "combo_corte_planilla"):
                self.combo_corte_planilla.configure(values=valores_cortes or ["Sin cortes"])
                self.combo_corte_planilla.set(valor_inicial or "Sin cortes")
                for indice, texto_corte in enumerate(valores_cortes):
                    estado = str(cortes_mapa[texto_corte][2] or "").strip().upper()
                    color_estado = {
                        "ACTIVO": "#16a34a",
                        "CERRADO": "#dc2626",
                    }.get(estado, "#64748b")
                    self.combo_corte_planilla._dropdown_menu.entryconfigure(
                        indice,
                        foreground=color_estado,
                        activeforeground=color_estado,
                    )
                estado_seleccionado = (
                    str(cortes_mapa.get(valor_inicial, (None, None, "", None))[2] or "")
                    .strip()
                    .upper()
                )
                color_seleccionado = {
                    "ACTIVO": "#16a34a",
                    "CERRADO": "#dc2626",
                }.get(estado_seleccionado, "#64748b")
                self.combo_corte_planilla.configure(text_color=color_seleccionado)
        except Exception as exc:
            logging.exception("No se pudieron cargar los cortes de planilla")
            messagebox.showerror("Error", f"No se pudieron cargar los cortes: {exc}")

        corte_consultado = cortes_mapa.get(valor_inicial) if cortes_mapa else None

        departamentos = sorted(
            {
                str(empleado[7] or "").strip()
                for empleado in empleados
                if str(empleado[7] or "").strip()
            }
        )
        valores_departamentos = ["Todos"] + departamentos
        if hasattr(self, "combo_departamento_planilla"):
            self.combo_departamento_planilla.configure(values=valores_departamentos)
            departamento_actual = str(departamento or "Todos").strip()
            if departamento_actual not in valores_departamentos:
                departamento_actual = "Todos"
            self.combo_departamento_planilla.set(departamento_actual)

        dias_por_dui = {}
        horas_extras_por_dui = {}
        horas_ordinarias_diarias = obtener_horas_ordinarias_configuradas(getattr(self, "horarios_laborales", []))
        otros_descuentos_por_dui = {}
        remuneraciones_por_nombre = {}
        planilla_completa = False
        try:
            registros_remuneraciones = db.registro_remuneracion().obtener_remuneraciones_detalladas()
            tipos_remuneracion = {
                "SALARIO",
                "VACACION", "VACACIONES",
                "REMUNERACION", "REMUNERACIONES",
                "BONIFICACION", "BONIFICACIONES",
            }

            def convertir_fecha_remuneracion(fecha):
                if hasattr(fecha, "date"):
                    return fecha.date()
                if hasattr(fecha, "year") and hasattr(fecha, "month") and hasattr(fecha, "day"):
                    return fecha
                texto = str(fecha or "").strip()
                for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                    try:
                        return datetime.strptime(texto[:10], formato).date()
                    except ValueError:
                        continue
                return None

            rango_remuneraciones = None
            if corte_consultado:
                fecha_inicio_remuneracion = convertir_fecha_remuneracion(corte_consultado[0])
                fecha_fin_remuneracion = convertir_fecha_remuneracion(corte_consultado[1])
                if fecha_inicio_remuneracion and fecha_fin_remuneracion:
                    rango_remuneraciones = (
                        fecha_inicio_remuneracion,
                        fecha_fin_remuneracion,
                    )

            for registro in registros_remuneraciones:
                if len(registro) < 8:
                    continue
                if rango_remuneraciones:
                    fecha_registro = convertir_fecha_remuneracion(registro[5])
                    if not fecha_registro or not (
                        rango_remuneraciones[0] <= fecha_registro <= rango_remuneraciones[1]
                    ):
                        continue
                identificador_remuneracion = str(registro[2] or "").strip().casefold()
                tipo_remuneracion = str(registro[3] or "").strip().upper()
                if identificador_remuneracion and tipo_remuneracion in tipos_remuneracion:
                    remuneraciones_por_nombre[identificador_remuneracion] = (
                        remuneraciones_por_nombre.get(identificador_remuneracion, 0.0)
                        + float(registro[4] or 0)
                    )
        except (TypeError, ValueError) as exc:
            logging.exception("No se pudieron calcular las remuneraciones del personal")
            messagebox.showerror("Error", f"No se pudieron calcular las remuneraciones: {exc}")

        try:
            if corte_consultado:
                fecha_inicio_corte, fecha_fin_corte, estado_corte, _ = corte_consultado
                if hasattr(fecha_fin_corte, "date"):
                    fecha_fin_corte = fecha_fin_corte.date()
                elif isinstance(fecha_fin_corte, str):
                    fecha_fin_corte = datetime.strptime(fecha_fin_corte[:10], "%Y-%m-%d").date()
                planilla_completa = (
                    str(estado_corte or "").strip().upper() == "CERRADO"
                    or datetime.now().date() >= fecha_fin_corte
                )
                marcaciones = db.marcaciones().obtener_marcaciones_por_rango(
                    fecha_inicio_corte, corte_consultado[1]
                )
                revisiones = db.marcaciones().obtener_revisiones_marcaciones_rango(
                    fecha_inicio_corte, corte_consultado[1]
                )
                for marcacion in marcaciones:
                    dui_marcacion = str(marcacion[2]).strip()
                    estado_dia = str(marcacion[6] if len(marcacion) > 6 else "JORNADA LABORAL").strip().upper()
                    if estado_dia in {"DÍA LIBRE", "DIA LIBRE", "AUSENCIA"}:
                        continue
                    if dia_cuenta_como_pagado(estado_dia):
                        dias_por_dui[dui_marcacion] = dias_por_dui.get(dui_marcacion, 0) + 1
                    else:
                        continue

                    if len(marcacion) < 5:
                        continue
                    hora_entrada = str(marcacion[3] or "").strip().upper()
                    hora_salida = str(marcacion[4] or "").strip().upper()
                    inicio = None
                    fin = None
                    for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                        if inicio is None:
                            try:
                                inicio = datetime.strptime(hora_entrada, formato)
                            except ValueError:
                                pass
                        if fin is None:
                            try:
                                fin = datetime.strptime(hora_salida, formato)
                            except ValueError:
                                pass
                    if not inicio or not fin:
                        continue
                    if fin < inicio:
                        fin += timedelta(days=1)

                    horas_extra = max(0.0, (fin - inicio).total_seconds() / 3600 - horas_ordinarias_diarias)
                    fecha_marcacion = marcacion[5]
                    revision = revisiones.get((dui_marcacion, fecha_marcacion), ("PENDIENTE", ""))[0]
                    if horas_extra > 0 and revision == "APROBADA":
                        horas_extras_por_dui[dui_marcacion] = (
                            horas_extras_por_dui.get(dui_marcacion, 0.0) + horas_extra
                        )
        except Exception as exc:
            logging.exception("No se pudieron consultar los días laborados")
            messagebox.showerror("Error", f"No se pudieron consultar las marcaciones: {exc}")

        if corte_consultado and str(corte_consultado[2] or "").strip().upper() == "CERRADO":
            try:
                fecha_inicio_descuentos = corte_consultado[0]
                fecha_fin_descuentos = corte_consultado[1]
                if hasattr(fecha_inicio_descuentos, "date"):
                    fecha_inicio_descuentos = fecha_inicio_descuentos.date()
                if hasattr(fecha_fin_descuentos, "date"):
                    fecha_fin_descuentos = fecha_fin_descuentos.date()
                servicio_descuentos = db.descuentos()
                otros_descuentos_por_dui = servicio_descuentos.obtener_totales_por_rango(
                    fecha_inicio_descuentos,
                    fecha_fin_descuentos,
                )
            except Exception as exc:
                logging.exception("No se pudieron consultar otros descuentos")
                messagebox.showerror("Error", f"No se pudieron consultar otros descuentos: {exc}")

        criterio = (filtro or "").strip().lower()
        departamento_filtro = str(departamento or "Todos").strip().casefold()
        for empleado in empleados:
            id_empleado, nombre, dui = empleado[0], empleado[1], empleado[2]
            cargo, salario = empleado[6], empleado[8]
            departamento_empleado = str(empleado[7] or "").strip().casefold()

            if criterio and criterio not in str(nombre).lower() and criterio not in str(dui).lower():
                continue
            if departamento_filtro != "todos" and departamento_empleado != departamento_filtro:
                continue

            try:
                salario_periodo = float(
                    str(salario).replace(",", "").replace("B/.", "").strip()
                )
                salario_mostrar = f"{salario_periodo:.2f}"
                valor_por_hora = salario_periodo / 15 / 8
                salario_por_hora = f"{valor_por_hora:.2f}"
                horas_trabajadas = horas_ordinarias_diarias
                dias_laborados = dias_por_dui.get(str(dui).strip(), 0)
                total_por_hora = f"{valor_por_hora * horas_trabajadas * dias_laborados:.2f}"
                horas_extras = horas_extras_por_dui.get(str(dui).strip(), 0.0)
                total_extras = f"{valor_por_hora * horas_extras * 1.5:.2f}"
                remuneraciones = remuneraciones_por_nombre.get(str(dui).strip().casefold(), 0.0)
                remuneraciones_mostrar = f"{remuneraciones:.2f}"
                otros_descuentos = otros_descuentos_por_dui.get(str(dui).strip(), 0.0)
                subtotal = (
                    valor_por_hora * horas_trabajadas * dias_laborados
                    + valor_por_hora * horas_extras * 1.5
                    + remuneraciones
                )
                deducciones = Deducciones().calcular_deducciones_nomina(
                    subtotal,
                    otros_descuentos=otros_descuentos,
                    salario_mensual=salario_periodo,
                    dias_pagados=dias_laborados,
                )
                css = deducciones["css"]
                seguro_educativo = deducciones["seguro_educativo"]
                isr = deducciones["isr"]
                salario_neto = deducciones["salario_neto"]
                subtotal_mostrar = f"{subtotal:.2f}"
                css_mostrar = f"{css:.2f}"
                seguro_educativo_mostrar = f"{seguro_educativo:.2f}"
                isr_mostrar = f"{isr:.2f}"
                otros_descuentos_mostrar = f"{otros_descuentos:.2f}"
                total_deducciones_mostrar = f"{deducciones['total_deducciones']:.2f}"
                salario_neto_mostrar = f"{salario_neto:.2f}"
            except (TypeError, ValueError):
                salario_mostrar = str(salario or "")
                salario_por_hora = ""
                horas_trabajadas = horas_ordinarias_diarias
                dias_laborados = dias_por_dui.get(str(dui).strip(), 0)
                total_por_hora = ""
                horas_extras = 0.0
                total_extras = ""
                remuneraciones_mostrar = ""
                subtotal_mostrar = ""
                css_mostrar = ""
                seguro_educativo_mostrar = ""
                isr_mostrar = ""
                otros_descuentos_mostrar = ""
                total_deducciones_mostrar = ""
                salario_neto_mostrar = ""

            if not planilla_completa:
                isr_mostrar = "0.00"

            valores = [
                id_empleado,
                nombre or "",
                dui or "",
                cargo or "",
                salario_mostrar,
                salario_por_hora,
                str(horas_trabajadas),
                str(dias_laborados),
                total_por_hora,
                f"{horas_extras:.2f}",
                total_extras,
                remuneraciones_mostrar,
                subtotal_mostrar,
                css_mostrar,
                seguro_educativo_mostrar,
                isr_mostrar,
                otros_descuentos_mostrar,
                total_deducciones_mostrar,
                salario_neto_mostrar,
            ]
            self.tabla_planilla.insert("", "end", values=valores)

    def exportar_planilla_excel(self):
        """Exporta la planilla actual a un archivo Excel con formato profesional."""
        try:
            items = self.tabla_planilla.get_children()
            if not items:
                messagebox.showwarning("Exportar Planilla", "No hay datos para exportar en la tabla.")
                return

            empresa = datos_empresa or {}
            corte_actual = ""
            try:
                combo_corte = getattr(self, "combo_corte_planilla", None)
                if combo_corte:
                    corte_actual = combo_corte.get() or "Sin corte especificado"
            except Exception:
                pass

            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

            def parse_numero(valor):
                if valor is None or valor == "":
                    return None
                texto = str(valor).strip().replace(",", "").replace("$", "").replace(" ", "")
                if texto in {"", "-", "--", "—"}:
                    return None
                try:
                    return float(texto)
                except ValueError:
                    return None

            wb = Workbook()
            ws = wb.active
            ws.title = "Planilla"

            for letra in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]:
                ws.column_dimensions[letra].width = 12
            ws.column_dimensions["A"].width = 5
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 14
            ws.column_dimensions["D"].width = 18
            ws.column_dimensions["E"].width = 13
            ws.column_dimensions["F"].width = 11
            ws.column_dimensions["G"].width = 9
            ws.column_dimensions["H"].width = 8
            ws.column_dimensions["I"].width = 12
            ws.column_dimensions["J"].width = 11
            ws.column_dimensions["K"].width = 12
            ws.column_dimensions["L"].width = 15
            ws.column_dimensions["M"].width = 12
            ws.column_dimensions["N"].width = 10
            ws.column_dimensions["O"].width = 12
            ws.column_dimensions["P"].width = 10
            ws.column_dimensions["Q"].width = 13
            ws.column_dimensions["R"].width = 15
            ws.column_dimensions["S"].width = 14

            color_primario = "0F766E"
            color_fondo_totales = "ECFDF5"
            color_texto_titulo = "FFFFFF"

            font_titulo = Font(name="Segoe UI", size=16, bold=True, color=color_texto_titulo)
            font_encabezado = Font(name="Segoe UI", size=10, bold=True, color=color_texto_titulo)
            font_normal = Font(name="Segoe UI", size=9)
            font_total = Font(name="Segoe UI", size=10, bold=True)

            fill_primario = PatternFill(start_color=color_primario, end_color=color_primario, fill_type="solid")
            fill_total = PatternFill(start_color=color_fondo_totales, end_color=color_fondo_totales, fill_type="solid")

            alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            alignment_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

            border_thin = Border(
                left=Side(style="thin", color="D0D7DE"),
                right=Side(style="thin", color="D0D7DE"),
                top=Side(style="thin", color="D0D7DE"),
                bottom=Side(style="thin", color="D0D7DE"),
            )

            # Tabla real de la planilla: 19 columnas visibles en la UI
            encabezados = [
                "#", "Empleado", "DUI", "Puesto", "Salario", "Sal/Hora", "Horas", "Días",
                "Total Hora", "Horas Extras", "Total Extras", "Remuneraciones", "Subtotal",
                "CSS", "Seg. Educativo", "ISR", "Otros desc.", "Total Deducciones", "Salario Neto"
            ]

            # Columnas monetarias exactas de la planilla (a partir de la columna 5 = Salario)
            columnas_monetarias = {5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19}
            total_monetario = {columna: 0.0 for columna in columnas_monetarias}

            fila = 1
            ws.merge_cells(f"A{fila}:S{fila}")
            ws[f"A{fila}"] = "PLANILLA DE PAGO"
            ws[f"A{fila}"].font = font_titulo
            ws[f"A{fila}"].fill = fill_primario
            ws[f"A{fila}"].alignment = alignment_center
            ws.row_dimensions[fila].height = 24

            fila += 1
            ws.merge_cells(f"A{fila}:S{fila}")
            ws[f"A{fila}"] = str(empresa.get("nombre") or "Empresa")
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=12, bold=True)
            ws[f"A{fila}"].alignment = alignment_center

            fila += 1
            ws.merge_cells(f"A{fila}:S{fila}")
            ws[f"A{fila}"] = (
                f"RUC: {empresa.get('ruc', '')} | Dirección: {empresa.get('direccion', '')} | "
                f"Teléfono: {empresa.get('telefono', '')}"
            )
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=9, color="525252")
            ws[f"A{fila}"].alignment = alignment_center

            fila += 1
            ws.merge_cells(f"A{fila}:S{fila}")
            ws[f"A{fila}"] = f"Corte: {corte_actual}"
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=10, bold=True, color="0F766E")
            ws[f"A{fila}"].alignment = alignment_left

            fila += 1
            ws.merge_cells(f"A{fila}:S{fila}")
            ws[f"A{fila}"] = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ws[f"A{fila}"].font = Font(name="Segoe UI", size=9, color="525252")
            ws[f"A{fila}"].alignment = alignment_left

            fila += 2
            for col_idx, titulo in enumerate(encabezados, start=1):
                celda = ws.cell(row=fila, column=col_idx)
                celda.value = titulo
                celda.font = font_encabezado
                celda.fill = fill_primario
                celda.alignment = alignment_center
                celda.border = border_thin

            ws.row_dimensions[fila].height = 22
            fila += 1

            total_empleados = 0
            for item in items:
                valores = list(self.tabla_planilla.item(item, "values") or [])
                if len(valores) < 19:
                    continue
                total_empleados += 1
                fila_valores = [total_empleados] + valores[1:]

                for col_idx, valor in enumerate(fila_valores, start=1):
                    celda = ws.cell(row=fila, column=col_idx)
                    celda.value = valor
                    celda.font = font_normal
                    celda.border = border_thin
                    if col_idx == 1:
                        celda.alignment = alignment_center
                    elif col_idx <= 4:
                        celda.alignment = alignment_left
                    else:
                        celda.alignment = alignment_right

                    if col_idx in columnas_monetarias:
                        numero = parse_numero(valor)
                        if numero is not None:
                            celda.number_format = "#,##0.00"
                            total_monetario[col_idx] += numero
                fila += 1

            fila_total = fila
            ws.merge_cells(f"A{fila_total}:D{fila_total}")
            ws[f"A{fila_total}"] = f"TOTAL ({total_empleados} empleados)"
            ws[f"A{fila_total}"].font = font_total
            ws[f"A{fila_total}"].fill = fill_total
            ws[f"A{fila_total}"].alignment = alignment_center
            ws[f"A{fila_total}"].border = border_thin

            for col_idx in range(5, 20):
                celda = ws.cell(row=fila_total, column=col_idx)
                if col_idx in columnas_monetarias:
                    celda.value = total_monetario[col_idx]
                    celda.font = font_total
                    celda.fill = fill_total
                    celda.alignment = alignment_right
                    celda.number_format = "#,##0.00"
                else:
                    celda.value = None
                celda.border = border_thin

            ws.freeze_panes = "A7"

            ruta = filedialog.asksaveasfilename(
                title="Guardar Planilla en Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("Todos", "*.*")],
                initialfile=f"planilla_{corte_actual.replace(':', '-').replace(' ', '_')}.xlsx"
            )

            if ruta:
                ruta_segura = construir_ruta_guardado_segura(
                    ruta,
                    nombre_default=f"planilla_{corte_actual.replace(':', '-').replace(' ', '_')}.xlsx",
                )
                try:
                    wb.save(ruta_segura)
                except PermissionError:
                    ruta_fallback = os.path.join(os.getcwd(), "exportaciones", os.path.basename(ruta_segura))
                    os.makedirs(os.path.dirname(ruta_fallback), exist_ok=True)
                    wb.save(ruta_fallback)
                    ruta_segura = ruta_fallback
                    messagebox.showwarning(
                        "Ruta no escribible",
                        "La carpeta seleccionada no permite guardar archivos. Se guardó la exportación en la carpeta local de exportaciones.",
                        parent=self.ventana_principal if hasattr(self, "ventana_principal") else None,
                    )

                ruta_abs = os.path.abspath(ruta_segura)
                messagebox.showinfo(
                    "Planilla Exportada",
                    f"Planilla exportada correctamente en:\n{ruta_abs}",
                    parent=self.ventana_principal if hasattr(self, "ventana_principal") else None,
                )
                if sys.platform.startswith("win"):
                    os.startfile(ruta_abs)

        except Exception as e:
            logging.exception("Error al exportar planilla a Excel")
            messagebox.showerror("Error al Exportar", f"No se pudo exportar la planilla:\n{e}")

    def exportar_planilla_csv(self):
        messagebox.showinfo("Frontend", "La exportación quedó desactivada en la versión frontal.")

    def toplevel_analisis_Planilla(self, sipp):
        toplevel = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(toplevel, 560, 260, margen_x=120, margen_y=120, min_ancho=380, min_alto=200)
        toplevel.grab_set()
        toplevel.title("Análisis de Planilla")
        contenedor = ctk.CTkFrame(toplevel, corner_radius=18, border_width=1, border_color="#334155")
        contenedor.place(relx=0.04, rely=0.04, relwidth=0.92, relheight=0.92)
        ctk.CTkLabel(
            contenedor,
            text="Análisis de planilla",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).place(relx=0.05, rely=0.08)
        ctk.CTkLabel(
            contenedor,
            text="La vista quedó limpia y sin cálculos ocultos.",
            font=ctk.CTkFont(size=13),
            text_color="#94a3b8",
            wraplength=420,
            justify="left",
            anchor="w",
        ).place(relx=0.05, rely=0.22)
        ctk.CTkButton(
            contenedor,
            text="Cerrar",
            width=120,
            height=36,
            fg_color="#4b5563",
            hover_color="#374151",
            command=toplevel.destroy,
        ).place(relx=0.40, rely=0.72)

    def toplevel_decimo_tercer_mes(self):
        top_level = ctk.CTkToplevel()
        top_level.title("Décimo tercer mes")
        ajustar_toplevel_a_pantalla(top_level, 1040, 650, margen_x=80, margen_y=80, min_ancho=720, min_alto=500)
        top_level.grab_set()
        top_level.resizable(False, False)
        top_level.configure(fg_color=("#f8fafc", "#0f172a"))

        contenedor = ctk.CTkFrame(top_level, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        contenedor.pack(fill="both", expand=True, padx=14, pady=14)
        ctk.CTkLabel(contenedor, text="Cálculo del décimo tercer mes", font=ctk.CTkFont(size=24, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(16, 2))
        ctk.CTkLabel(contenedor, text="Consulta el importe proporcional o completo de cada empleado según el período seleccionado.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=18, pady=(0, 12))

        controles = ctk.CTkFrame(contenedor, corner_radius=12, fg_color=("#f8fafc", "#1f2937"))
        controles.pack(fill="x", padx=18, pady=(0, 12))
        controles.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(controles, text="Período:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(12, 8), pady=10)
        año = datetime.now().year
        periodos = {
            f"Diciembre {año - 1} - Abril {año}": (datetime(año - 1, 12, 16).date(), datetime(año, 4, 15).date()),
            f"Abril - Agosto {año}": (datetime(año, 4, 16).date(), datetime(año, 8, 15).date()),
            f"Agosto - Diciembre {año}": (datetime(año, 8, 16).date(), datetime(año, 12, 15).date()),
        }
        selector_periodo = ctk.CTkComboBox(controles, values=list(periodos), state="readonly", width=240, height=34)
        selector_periodo.set(list(periodos)[0])
        selector_periodo.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="w")
        ctk.CTkLabel(controles, text="Buscar:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(0, 8), pady=10)
        entrada_busqueda = ctk.CTkEntry(controles, placeholder_text="Nombre o DUI", height=34)
        entrada_busqueda.grid(row=0, column=3, padx=(0, 12), pady=10, sticky="ew")
        ctk.CTkLabel(controles, text="Departamento:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=4, padx=(0, 8), pady=10)
        empleados = db.gestion_empleado().obtener_empleados() if db else []
        servicio_decimo = None
        try:
            servicio_decimo = db.decimoXIII() if db and hasattr(db, "decimoXIII") else None
        except Exception as exc:
            logging.exception("No se pudo inicializar la tabla decimoxiii")
            resumen_texto_db = f"No se pudo conectar con decimoxiii: {exc}"
        else:
            resumen_texto_db = ""
        departamentos = sorted({str(empleado[7] or "").strip() for empleado in empleados if len(empleado) > 7 and str(empleado[7] or "").strip()})
        selector_departamento = ctk.CTkComboBox(controles, values=["Todos"] + departamentos, state="readonly", width=170, height=34)
        selector_departamento.set("Todos")
        selector_departamento.grid(row=0, column=5, padx=(0, 12), pady=10)

        resumen = ctk.CTkLabel(contenedor, text="Selecciona un período para calcular los importes.", font=ctk.CTkFont(size=13), anchor="w")
        resumen.pack(fill="x", padx=18, pady=(0, 8))
        area_tabla = ctk.CTkFrame(contenedor, fg_color="transparent")
        area_tabla.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        area_tabla.grid_rowconfigure(0, weight=1)
        area_tabla.grid_columnconfigure(0, weight=1)
        columnas = ("Empleado", "DUI", "Departamento", "Días", "Salario computable", "Décimo bruto")
        tabla = ttk.Treeview(area_tabla, columns=columnas, show="headings")
        for columna, ancho in (("Empleado", 180), ("DUI", 110), ("Departamento", 180), ("Días", 70), ("Salario computable", 150), ("Décimo bruto", 130)):
            tabla.heading(columna, text=columna)
            tabla.column(columna, width=ancho, minwidth=60, anchor="w", stretch=True)
        tabla.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(area_tabla, orient="vertical", command=tabla.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        tabla.configure(yscrollcommand=scroll.set)

        def convertir_fecha(valor):
            if hasattr(valor, "date"):
                return valor.date()
            if hasattr(valor, "year"):
                return valor
            texto = str(valor or "").strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto[:10], formato).date()
                except ValueError:
                    continue
            return None

        def calcular_decimo(event=None):
            for item in tabla.get_children():
                tabla.delete(item)
            inicio, fin = periodos[selector_periodo.get()]
            if datetime.now().date() < fin:
                resumen.configure(text=f"El período aún no se cumple. Finaliza el {fin:%d/%m/%Y}.")
                boton_comprobante.pack_forget()
                if event is None:
                    messagebox.showwarning(
                        "Décimo tercer mes",
                        f"No se puede calcular este período porque todavía no ha finalizado.\n\nFecha de cierre: {fin:%d/%m/%Y}",
                        parent=top_level,
                    )
                return
            criterio = entrada_busqueda.get().strip().casefold()
            departamento = selector_departamento.get().strip()
            total = 0.0
            cantidad = 0
            for empleado in empleados:
                nombre = str(empleado[1] or "")
                dui = str(empleado[2] or "")
                departamento_empleado = str(empleado[7] or "").strip() if len(empleado) > 7 else "Sin departamento"
                if criterio and criterio not in nombre.casefold() and criterio not in dui.casefold():
                    continue
                if departamento != "Todos" and departamento_empleado != departamento:
                    continue
                fecha_ingreso = convertir_fecha(empleado[9]) if len(empleado) > 9 else None
                inicio_computable = max(inicio, fecha_ingreso) if fecha_ingreso else inicio
                if inicio_computable > fin:
                    continue
                dias = (fin - inicio_computable).days + 1
                try:
                    salario = float(str(empleado[8] or 0).replace("B/.", "").replace(",", "").strip())
                except (TypeError, ValueError):
                    salario = 0.0
                salario_computable = salario * dias / 360
                decimo = salario_computable / 3
                total += decimo
                cantidad += 1
                tabla.insert("", "end", values=(nombre, dui, departamento_empleado, dias, f"B/. {salario_computable:,.2f}", f"B/. {decimo:,.2f}"))
                if servicio_decimo:
                    try:
                        servicio_decimo.guardar_decimo(
                            empleado[0], nombre, dui, departamento_empleado,
                            inicio, fin, dias, salario_computable, decimo,
                        )
                    except Exception:
                        logging.exception("No se pudo guardar el décimo de %s", dui)
            texto_resumen = f"{cantidad} empleado(s) | Total décimo bruto: B/. {total:,.2f} | Período: {inicio:%d/%m/%Y} al {fin:%d/%m/%Y}"
            resumen.configure(text=f"{texto_resumen}\n{resumen_texto_db}" if resumen_texto_db else texto_resumen)
            if cantidad:
                boton_comprobante.pack(side="left", padx=(10, 0))
            else:
                boton_comprobante.pack_forget()

        def generar_comprobante_decimo():
            registros = [tabla.item(item, "values") for item in tabla.get_children()]
            if not registros:
                messagebox.showwarning("Décimo tercer mes", "Primero calcula un período con empleados.", parent=top_level)
                return
            inicio, fin = periodos[selector_periodo.get()]
            ruta = filedialog.asksaveasfilename(
                parent=top_level,
                title="Guardar comprobante del décimo",
                defaultextension=".pdf",
                filetypes=[("Documento PDF", "*.pdf")],
                initialfile="comprobante_decimo_tercer_mes.pdf",
            )
            if not ruta:
                return
            try:
                estilos = getSampleStyleSheet()
                documento = SimpleDocTemplate(ruta, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
                estilo_titulo = estilos["Title"].clone("TituloDecimo")
                estilo_titulo.alignment = 1
                contenido = [
                    Paragraph(_tr("COMPROBANTE DE PAGO"), estilo_titulo),
                    Spacer(1, 8),
                    Paragraph(_tr("Décimo tercer mes"), estilos["Heading2"]),
                    Paragraph(escape(f"{_tr('Período')}: {selector_periodo.get()} ({inicio:%d/%m/%Y} al {fin:%d/%m/%Y})"), estilos["Normal"]),
                    Paragraph(escape(f"{_tr('Departamento')}: {selector_departamento.get()}"), estilos["Normal"]),
                    Spacer(1, 12),
                ]
                encabezados = [_tr("Empleado"), _tr("DUI"), _tr("Departamento"), _tr("Días"), _tr("Salario computable"), _tr("Décimo bruto")]
                datos_pdf = [[Paragraph(f"<b>{escape(valor)}</b>", estilos["Normal"]) for valor in encabezados]]
                datos_pdf.extend([[Paragraph(escape(str(valor or "-")), estilos["Normal"]) for valor in registro] for registro in registros])
                tabla_pdf = Table(datos_pdf, colWidths=[1.45 * inch, 1.0 * inch, 1.35 * inch, 0.55 * inch, 1.25 * inch, 1.1 * inch], repeatRows=1)
                tabla_pdf.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("PADDING", (0, 0), (-1, -1), 5),
                ]))
                documento.build(contenido + [tabla_pdf])
                ruta = os.path.abspath(ruta)
                if not os.path.isfile(ruta) or os.path.getsize(ruta) == 0:
                    raise OSError("El comprobante no se creó correctamente.")
                if sys.platform.startswith("win"):
                    os.startfile(ruta)
                messagebox.showinfo("Décimo tercer mes", f"Comprobante generado correctamente en:\n{ruta}", parent=top_level)
            except Exception as exc:
                messagebox.showerror("Décimo tercer mes", f"No se pudo generar el comprobante: {exc}", parent=top_level)

        acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
        acciones.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkButton(acciones, text="Calcular décimo", width=145, fg_color="#0f766e", hover_color="#115e59", command=lambda: calcular_decimo()).pack(side="left")
        boton_comprobante = ctk.CTkButton(acciones, text="Comprobante de pago", width=170, fg_color="#2563eb", hover_color="#1d4ed8", command=generar_comprobante_decimo)
        ctk.CTkButton(acciones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        selector_periodo.configure(command=calcular_decimo)
        selector_departamento.configure(command=calcular_decimo)
        entrada_busqueda.bind("<KeyRelease>", calcular_decimo)
        calcular_decimo(event="inicio")

    def toplevel_comprobantes_pago(self):
        parent = self.tabla_planilla.winfo_toplevel() if self.tabla_planilla is not None else None
        top_level = ctk.CTkToplevel(parent)
        top_level.title("Comprobantes de pago")
        ajustar_toplevel_a_pantalla(top_level, 900, 400, margen_x=80, margen_y=80, min_ancho=680, min_alto=320)
        top_level.grab_set()
        top_level.resizable(False, False)
        top_level.configure(fg_color=("#f8fafc", "#0f172a"))

        contenedor = ctk.CTkFrame(top_level, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        contenedor.pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkLabel(contenedor, text="Comprobantes de pago", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").pack(fill="x", padx=14, pady=(12, 2))
        ctk.CTkLabel(contenedor, text="Seleccione los filtros para preparar la descarga.", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=14, pady=(0, 10))

        filtros = ctk.CTkFrame(contenedor, corner_radius=12, fg_color=("#f8fafc", "#1f2937"))
        filtros.pack(fill="x", padx=14, pady=(0, 10))
        for columna in (1, 3, 5):
            filtros.grid_columnconfigure(columna, weight=1)
        ctk.CTkLabel(filtros, text="Corte:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(12, 8), pady=10)
        cortes_por_opcion = {"Todos": None}
        datos_cortes_por_opcion = {}
        opciones_corte = ["Todos"]
        corte_activo_opcion = None
        try:
            cortes_disponibles = db.cortes_planilla().obtener_cortes_planilla() if db else []
            for corte in sorted(cortes_disponibles, key=lambda dato: dato[0], reverse=True):
                if len(corte) < 4:
                    continue
                corte_id, fecha_inicio, fecha_fin, estado = corte[:4]
                inicio = fecha_inicio.strftime("%d/%m/%Y") if hasattr(fecha_inicio, "strftime") else str(fecha_inicio)
                fin = fecha_fin.strftime("%d/%m/%Y") if hasattr(fecha_fin, "strftime") else str(fecha_fin)
                opcion = f"Corte {corte_id}: {inicio} a {fin} ({estado})"
                opciones_corte.append(opcion)
                cortes_por_opcion[opcion] = int(corte_id)
                datos_cortes_por_opcion[opcion] = (fecha_inicio, fecha_fin)
                if str(estado or "").strip().upper() == "ACTIVO":
                    corte_activo_opcion = opcion
        except Exception:
            logging.exception("No se pudieron cargar los cortes para comprobantes")
        selector_corte = ctk.CTkComboBox(filtros, values=opciones_corte, state="readonly", height=34)
        selector_corte.set(corte_activo_opcion or "Todos")
        selector_corte.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")

        def cambiar_corte(valor):
            estado = str(valor.rsplit("(", 1)[-1].rstrip(")")).strip().upper()
            selector_corte.configure(text_color="#16a34a" if estado == "ACTIVO" else "#dc2626")

        selector_corte.configure(command=cambiar_corte)
        cambiar_corte(selector_corte.get())
        for indice, opcion in enumerate(opciones_corte[1:]):
            estado = str(opcion.rsplit("(", 1)[-1].rstrip(")")).strip().upper()
            color_estado = "#16a34a" if estado == "ACTIVO" else "#dc2626"
            selector_corte._dropdown_menu.entryconfigure(
                indice + 1,
                foreground=color_estado,
                activeforeground=color_estado,
            )
        ctk.CTkLabel(filtros, text="Departamento:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(0, 8), pady=10)
        empleados = db.gestion_empleado().obtener_empleados() if db else []
        empleados_por_dui = {str(empleado[2] or "").strip(): empleado for empleado in empleados if len(empleado) > 2}
        departamentos = sorted({str(empleado[7] or "").strip() for empleado in empleados if len(empleado) > 7 and str(empleado[7] or "").strip()})
        selector_departamento = ctk.CTkComboBox(filtros, values=["Todos"] + departamentos, state="readonly", width=150, height=34)
        selector_departamento.set("Todos")
        selector_departamento.grid(row=0, column=3, padx=(0, 12), pady=10, sticky="ew")
        ctk.CTkLabel(filtros, text="Modo:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=4, padx=(0, 8), pady=10)
        selector_modo = ctk.CTkComboBox(filtros, values=["Todos", "Individual"], state="readonly", width=120, height=34)
        selector_modo.set("Todos")
        selector_modo.grid(row=0, column=5, padx=(0, 12), pady=10, sticky="ew")
        entrada_individual = ctk.CTkEntry(filtros, placeholder_text="Nombre o DUI", height=34)
        entrada_individual.grid(row=0, column=6, padx=(0, 10), pady=10, sticky="ew")
        entrada_individual.grid_remove()
        filtros.grid_columnconfigure(6, weight=1)

        def cambiar_modo(valor):
            if valor == "Individual":
                entrada_individual.grid()
            else:
                entrada_individual.grid_remove()

        selector_modo.configure(command=cambiar_modo)
        boton_cargar = ctk.CTkButton(
            filtros,
            text="Cargar",
            width=100,
            height=34,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
        )
        boton_cargar.grid(row=0, column=7, padx=(0, 12), pady=10)
        filtros.grid_columnconfigure(7, weight=0)

        resultados = ctk.CTkFrame(contenedor, fg_color="transparent")
        resultados.pack(fill="x", padx=14, pady=(0, 10))
        etiqueta_resultado = ctk.CTkLabel(
            resultados,
            text="Seleccione los filtros y presione Cargar.",
            height=82,
            corner_radius=12,
            fg_color=("#f1f5f9", "#1f2937"),
            text_color=("#64748b", "#cbd5e1"),
            font=ctk.CTkFont(size=12),
            wraplength=850,
        )
        etiqueta_resultado.pack(fill="both", expand=True)
        columnas_planilla = (
            "ID", "Nombre Completo", "DUI", "Cargo", "Salario", "Por Horas",
            "Horas Trabajadas", "Días", "Total por Hora", "Horas Extras",
            "Total Extras", "Remuneraciones", "Subtotal", "CSS", "Seg. Educativo",
            "ISR", "Otros descuentos", "Total Deducciones", "Salario Neto"
        )
        registros_cargados = []
        referencia_cargada = ""

        def cargar_comprobantes():
            nonlocal registros_cargados, referencia_cargada
            try:
                departamento = selector_departamento.get().strip()
                modo = selector_modo.get().strip()
                criterio_individual = entrada_individual.get().strip().casefold()
                referencia_cargada = selector_corte.get()
                corte_seleccionado = None if selector_corte.get() == "Todos" else selector_corte.get()
                self.cargar_planilla("", departamento, corte_seleccionado)
                registros = [
                    self.tabla_planilla.item(item, "values")
                    for item in self.tabla_planilla.get_children()
                ] if self.tabla_planilla is not None else []
                registros_cargados = []
                for registro in registros:
                    if modo == "Individual":
                        texto_busqueda = f"{registro[1] or ''} {registro[2] or ''}".casefold()
                        if not criterio_individual or criterio_individual not in texto_busqueda:
                            continue
                    registros_cargados.append(registro)
            except Exception as exc:
                registros_cargados = []
                etiqueta_resultado.configure(text=f"No se pudieron cargar los registros:\n\n{exc}")
                logging.exception("No se pudieron cargar los comprobantes")
                return
            referencia = f"{referencia_cargada} | Departamento: {departamento} | Modo: {modo}"
            if modo == "Individual":
                referencia += f" | Persona: {entrada_individual.get().strip() or 'sin especificar'}"
            etiqueta_resultado.configure(
                text=f"Cargado correctamente.\n\n{referencia}\n\nRegistros preparados: {len(registros_cargados)}"
                if registros_cargados
                else f"No se encontraron registros.\n\n{referencia}"
            )

        def descargar_comprobantes():
            if not registros_cargados:
                messagebox.showwarning(
                    "Comprobantes",
                    "No hay registros cargados. Presiona Cargar y selecciona un corte que tenga datos en la planilla.",
                    parent=top_level,
                )
                return
            ruta = filedialog.asksaveasfilename(
                parent=top_level,
                title="Descargar comprobantes",
                defaultextension=".pdf",
                filetypes=[("Documento PDF", "*.pdf")],
                initialfile="comprobante_pago.pdf",
            )
            if not ruta:
                return
            try:
                estilos = getSampleStyleSheet()
                documento = SimpleDocTemplate(ruta, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
                empresa = datos_empresa
                fecha_inicio, fecha_fin = datos_cortes_por_opcion.get(referencia_cargada, (None, None))
                if hasattr(fecha_inicio, "date"):
                    fecha_inicio = fecha_inicio.date()
                if hasattr(fecha_fin, "date"):
                    fecha_fin = fecha_fin.date()
                meses = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")
                if fecha_inicio and hasattr(fecha_inicio, "month"):
                    periodo = f"{'Primera' if fecha_inicio.day <= 15 else 'Segunda'} quincena de {meses[fecha_inicio.month - 1]} de {fecha_inicio.year}"
                else:
                    periodo = "Período no especificado"
                estilo_pequeno = estilos["Normal"].clone("DetallePequeno")
                estilo_pequeno.fontSize = 8
                estilo_pequeno.leading = 10
                estilo_titulo = estilos["Title"].clone("TituloComprobante")
                estilo_titulo.fontSize = 17
                estilo_titulo.leading = 20
                estilo_titulo.alignment = 1
                estilo_seccion = estilos["Heading2"].clone("SeccionComprobante")
                estilo_seccion.fontSize = 10
                estilo_seccion.leading = 12
                nombres_datos = tuple(_tr(nombre) for nombre in (
                    "ID", "Nombre", "DUI", "Cargo", "Salario", "Por horas",
                    "Horas trabajadas", "Días", "Total por hora", "Horas extras",
                    "Total extras", "Remuneraciones", "Subtotal", "CSS",
                    "Seguro educativo", "ISR", "Otros descuentos",
                    "Total deducciones", "Salario neto",
                ))
                indices_dinero = {4, 5, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18}

                def mostrar_valor(indice, valor):
                    texto = str(valor or "").strip()
                    if indice in indices_dinero and texto and not texto.startswith("B/."):
                        try:
                            return f"B/. {float(texto.replace(',', '')):,.2f}"
                        except (TypeError, ValueError):
                            pass
                    return texto or "-"

                def crear_tabla_detalle(titulo, campos, columnas=2):
                    filas = [[
                        Paragraph(f"<b>{_tr('Concepto')}</b>", estilo_pequeno),
                        Paragraph(f"<b>{_tr('Detalle')}</b>", estilo_pequeno),
                    ]]
                    filas.extend([
                        [Paragraph(escape(nombres_datos[indice]), estilo_pequeno), Paragraph(escape(mostrar_valor(indice, registro[indice])), estilo_pequeno)]
                        for indice in campos
                    ])
                    return Table(filas, colWidths=[1.45 * inch, 2.0 * inch])

                contenido = []
                for indice, registro in enumerate(registros_cargados):
                    contenido.extend([Paragraph(_tr("COMPROBANTE DE PAGO"), estilo_titulo), Spacer(1, 5)])
                    logo = empresa.get("logo") or ""
                    logo_pdf = PdfImage(logo, width=0.75 * inch, height=0.75 * inch) if logo and os.path.isfile(logo) else ""
                    empleado = empleados_por_dui.get(str(registro[2] or "").strip())
                    departamento_empleado = (
                        str(empleado[7] or "").strip()
                        if empleado and len(empleado) > 7 and str(empleado[7] or "").strip()
                        else _tr("Sin departamento")
                    )
                    datos_empresa_pdf = [
                        [Paragraph(f"<b>{escape(empresa.get('nombre') or _tr('Nombre de la empresa'))}</b>", estilo_pequeno)],
                        [Paragraph(escape(f"RUC: {empresa.get('ruc') or _tr('No especificado')}"), estilo_pequeno)],
                        [Paragraph(escape(f"{_tr('Dirección')}: {empresa.get('direccion') or _tr('No especificada')}"), estilo_pequeno)],
                        [Paragraph(escape(f"{_tr('Teléfono')}: {empresa.get('telefono') or _tr('No especificado')}"), estilo_pequeno)],
                        [Paragraph(escape(f"{_tr('Período')}: {periodo}"), estilo_pequeno)],
                        [Paragraph(escape(f"{_tr('Departamento')}: {departamento_empleado}"), estilo_pequeno)],
                    ]
                    encabezado = Table([[logo_pdf, Table(datos_empresa_pdf, colWidths=[5.9 * inch])]], colWidths=[0.9 * inch, 5.9 * inch])
                    encabezado.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]))
                    contenido.extend([encabezado, Spacer(1, 7), Paragraph(escape(f"{_tr('Empleado')}: {registro[1]} | {_tr('DUI')}: {registro[2]}"), estilo_seccion), Spacer(1, 3)])
                    bloques = (
                        (_tr("Información del empleado"), (0, 1, 2, 3)),
                        (_tr("Asistencia e ingresos"), (4, 5, 6, 7, 8, 9, 10, 11, 12)),
                        (_tr("Deducciones y pago neto"), (13, 14, 15, 16, 17, 18)),
                    )
                    for titulo, campos in bloques:
                        contenido.append(Paragraph(titulo, estilo_seccion))
                        tablas = [crear_tabla_detalle(titulo, campos)]
                        tablas[0].setStyle(TableStyle([
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("PADDING", (0, 0), (-1, -1), 3),
                        ]))
                        contenido.extend([tablas[0], Spacer(1, 5)])
                    if indice < len(registros_cargados) - 1:
                        contenido.append(PageBreak())
                documento.build(contenido)
                ruta = os.path.abspath(ruta)
                if not os.path.isfile(ruta) or os.path.getsize(ruta) == 0:
                    raise OSError("El archivo PDF no se creó correctamente.")
            except Exception as exc:
                messagebox.showerror("Comprobantes", f"No se pudo descargar el archivo: {exc}", parent=top_level)
                return

            try:
                if sys.platform.startswith("win"):
                    os.startfile(ruta)
            except OSError:
                logging.exception("El comprobante se creó, pero no se pudo abrir automáticamente")
            messagebox.showinfo("Comprobantes", f"Archivo descargado correctamente en:\n{ruta}", parent=top_level)

        botones_resultados = ctk.CTkFrame(contenedor, fg_color="transparent")
        botones_resultados.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkButton(botones_resultados, text="Descargar PDF", width=130, fg_color="#0f766e", hover_color="#115e59", command=descargar_comprobantes).pack(side="left")
        ctk.CTkButton(botones_resultados, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        entrada_individual.bind("<Return>", lambda event: cargar_comprobantes())
        boton_cargar.configure(command=cargar_comprobantes)
        return

        ctk.CTkLabel(filtros, text="Buscar:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=0, padx=(12, 8), pady=10)
        entrada_busqueda = ctk.CTkEntry(filtros, placeholder_text="Nombre o DUI", height=34)
        entrada_busqueda.grid(row=1, column=1, columnspan=5, padx=(0, 12), pady=10, sticky="ew")

        area_tabla = ctk.CTkFrame(contenedor, fg_color="transparent")
        area_tabla.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        area_tabla.grid_rowconfigure(0, weight=1)
        area_tabla.grid_columnconfigure(0, weight=1)
        columnas = ("ID", "Empleado", "DUI", "Puesto", "Salario mensual", "Subtotal", "Deducciones", "Neto", "Fecha", "Corte")
        tabla = ttk.Treeview(area_tabla, columns=columnas, show="headings", selectmode="browse")
        for columna, ancho in (("ID", 45), ("Empleado", 180), ("DUI", 110), ("Puesto", 150), ("Salario mensual", 110), ("Subtotal", 110), ("Deducciones", 120), ("Neto", 110), ("Fecha", 100), ("Corte", 90)):
            tabla.heading(columna, text=columna)
            tabla.column(columna, width=ancho, minwidth=45, anchor="w", stretch=True)
        tabla.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(area_tabla, orient="vertical", command=tabla.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        tabla.configure(yscrollcommand=scroll.set)
        etiqueta_seleccion = ctk.CTkLabel(contenedor, text="Selecciona un registro para imprimirlo.", font=ctk.CTkFont(size=12), anchor="w")
        etiqueta_seleccion.pack(fill="x", padx=18, pady=(0, 10))

        registros = []
        empleados_por_dui = {str(empleado[2] or "").strip(): empleado for empleado in empleados if len(empleado) > 2}
        cortes_por_id = {}
        try:
            cortes_por_id = {
                int(corte[0]): corte
                for corte in db.cortes_planilla().obtener_cortes_planilla()
                if len(corte) >= 4
            }
        except Exception:
            cortes_por_id = {}

        def cargar_registros(event=None):
            criterio = entrada_busqueda.get().strip().casefold()
            departamento = selector_departamento.get().strip()
            corte_id = cortes_por_opcion.get(selector_corte.get())
            modo = selector_modo.get().strip()
            seleccionado = tabla.selection()[0] if modo == "Individual" and tabla.selection() else None
            for item in tabla.get_children():
                tabla.delete(item)
            for registro in registros:
                if corte_id is not None and str(registro[17]) != str(corte_id):
                    continue
                if seleccionado is not None and str(registro[0]) != seleccionado:
                    continue
                empleado = empleados_por_dui.get(str(registro[2] or "").strip())
                departamento_registro = str(empleado[7] or "").strip() if empleado and len(empleado) > 7 else "Sin departamento"
                coincide = not criterio or criterio in str(registro[1] or "").casefold() or criterio in str(registro[2] or "").casefold()
                if coincide and (departamento == "Todos" or departamento_registro == departamento):
                    corte = cortes_por_id.get(registro[17])
                    corte_texto = f"#{registro[17]}" if registro[17] else "Sin corte"
                    subtotal = float(registro[14] or 0) + float(registro[15] or 0)
                    tabla.insert("", "end", iid=str(registro[0]), values=(registro[0], registro[1], registro[2], registro[3], f"B/. {float(registro[4] or 0):,.2f}", f"B/. {subtotal:,.2f}", f"B/. {float(registro[14] or 0):,.2f}", f"B/. {float(registro[15] or 0):,.2f}", registro[16], corte_texto))

        def seleccionar_registro(event=None):
            seleccion = tabla.selection()
            if not seleccion:
                return
            registro = next((fila for fila in registros if str(fila[0]) == seleccion[0]), None)
            if registro:
                corte = cortes_por_id.get(registro[17])
                estado_corte = corte[3] if corte else "Sin corte"
                etiqueta_seleccion.configure(text=f"Seleccionado: {registro[1]} | DUI: {registro[2]} | Neto: B/. {float(registro[15] or 0):,.2f} | Corte: {estado_corte}")

        def imprimir_comprobante():
            seleccion = tabla.selection()
            if not seleccion:
                messagebox.showwarning("Comprobantes", "Selecciona un registro de planilla.", parent=top_level)
                return
            registro = next((fila for fila in registros if str(fila[0]) == seleccion[0]), None)
            if not registro:
                return
            corte = cortes_por_id.get(registro[17])
            if not corte or str(corte[3] or "").upper() != "CERRADO":
                messagebox.showwarning("Comprobante", "Solo se puede imprimir un comprobante de un corte cerrado.", parent=top_level)
                return
            ruta = filedialog.asksaveasfilename(parent=top_level, title="Guardar comprobante de pago", defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"comprobante_pago_{registro[0]}.pdf")
            if not ruta:
                return
            try:
                estilos = getSampleStyleSheet()
                documento = SimpleDocTemplate(ruta, pagesize=landscape(letter), rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
                empleado = empleados_por_dui.get(str(registro[2] or "").strip())
                salario_mensual = (
                    float(str(empleado[8] or 0).replace("B/.", "").replace(",", "").strip())
                    if empleado and len(empleado) > 8
                    else float(registro[4] or 0)
                )
                if hasattr(corte[1], "date"):
                    fecha_inicio = corte[1].date()
                else:
                    fecha_inicio = corte[1]
                if hasattr(corte[2], "date"):
                    fecha_fin = corte[2].date()
                else:
                    fecha_fin = corte[2]
                servicio_marcaciones = db.marcaciones()
                marcaciones_periodo = servicio_marcaciones.obtener_marcaciones_por_rango(fecha_inicio, fecha_fin)
                revisiones = servicio_marcaciones.obtener_revisiones_marcaciones_rango(fecha_inicio, fecha_fin)
                dias_laborados = 0
                horas_extras = 0.0
                detalle_diario = []
                limite_diario = obtener_horas_ordinarias_configuradas(getattr(self, "horarios_laborales", []))
                nombres_dias = ("Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom")
                for marcacion in marcaciones_periodo:
                    if str(marcacion[2] or "").strip() != str(registro[2] or "").strip():
                        continue
                    estado_dia = str(marcacion[6] if len(marcacion) > 6 else "JORNADA LABORAL").strip().upper()
                    if estado_dia in {"DÍA LIBRE", "DIA LIBRE", "AUSENCIA"}:
                        continue
                    if dia_cuenta_como_pagado(estado_dia):
                        dias_laborados += 1
                    fecha_marcacion = marcacion[5]
                    if hasattr(fecha_marcacion, "date"):
                        fecha_marcacion = fecha_marcacion.date()
                    if isinstance(fecha_marcacion, str):
                        fecha_marcacion = datetime.strptime(fecha_marcacion[:10], "%Y-%m-%d").date()
                    revision = revisiones.get((str(registro[2]).strip(), marcacion[5]), ("PENDIENTE", ""))[0]
                    inicio = fin = None
                    for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                        if inicio is None:
                            try:
                                inicio = datetime.strptime(str(marcacion[3] or "").strip().upper(), formato)
                            except ValueError:
                                pass
                        if fin is None:
                            try:
                                fin = datetime.strptime(str(marcacion[4] or "").strip().upper(), formato)
                            except ValueError:
                                pass
                    if inicio and fin:
                        if fin < inicio:
                            fin += timedelta(days=1)
                        horas_trabajadas_dia = round((fin - inicio).total_seconds() / 3600, 2)
                        horas_extra = max(0.0, horas_trabajadas_dia - limite_diario)
                        if revision == "APROBADA":
                            horas_extras += horas_extra
                    else:
                        horas_trabajadas_dia = 0.0
                        horas_extra = 0.0
                    feriado = feriados_panama(fecha_marcacion.year).get(fecha_marcacion, "")
                    pago_dia = salario_mensual / 30 if dia_cuenta_como_pagado(estado_dia) else 0.0
                    detalle_diario.append([
                        fecha_marcacion.strftime("%d/%m/%Y"),
                        nombres_dias[fecha_marcacion.weekday()],
                        "FERIADO: " + feriado if feriado else "Ordinario",
                        str(marcacion[3] or "-")[:11],
                        str(marcacion[4] or "-")[:11],
                        f"{horas_trabajadas_dia:.2f}",
                        f"{min(limite_diario, horas_trabajadas_dia):.2f}",
                        f"{horas_extra:.2f}" if revision == "APROBADA" else "0.00",
                        f"B/. {pago_dia:,.2f}",
                    ])
                salario_base = salario_mensual / 30 * dias_laborados
                total_horas_extras = salario_mensual / 30 / 8 * horas_extras * 1.5
                remuneraciones = 0.0
                tipos_remuneracion = {"SALARIO", "VACACION", "VACACIONES", "REMUNERACION", "REMUNERACIONES", "BONIFICACION", "BONIFICACIONES"}
                for remuneracion in db.registro_remuneracion().obtener_remuneraciones_detalladas(dui=registro[2], fecha_inicio=fecha_inicio, fecha_fin=fecha_fin):
                    identificador = str(remuneracion[2] or "").strip().casefold()
                    tipo = str(remuneracion[3] or "").strip().upper()
                    fecha_remuneracion = remuneracion[5]
                    if hasattr(fecha_remuneracion, "date"):
                        fecha_remuneracion = fecha_remuneracion.date()
                    if identificador == str(registro[2] or "").strip().casefold() and tipo in tipos_remuneracion:
                        remuneraciones += float(remuneracion[4] or 0)
                comisiones = 0.0
                bonificaciones = 0.0
                otros_ingresos = remuneraciones
                subtotal = salario_base + total_horas_extras + comisiones + bonificaciones + otros_ingresos
                otros_descuentos = db.descuentos().obtener_total_descuentos_por_rango(registro[2], fecha_inicio, fecha_fin)
                deducciones = Deducciones().calcular_deducciones_nomina(
                    subtotal,
                    otros_descuentos=otros_descuentos,
                    salario_mensual=salario_mensual,
                    dias_pagados=dias_laborados,
                )
                total_bruto = subtotal
                total_deducciones = deducciones["total_deducciones"]
                total_neto = deducciones["salario_neto"]
                periodo = f"{corte[1]:%d/%m/%Y} al {corte[2]:%d/%m/%Y}" if hasattr(corte[1], "strftime") and hasattr(corte[2], "strftime") else f"{corte[1]} al {corte[2]}"
                detalle = [[_tr("Fecha"), _tr("Día"), _tr("Condición"), _tr("Entrada"), _tr("Salida"), _tr("Horas"), _tr("Ordinarias"), _tr("Extras"), _tr("Pago día")]] + detalle_diario
                tabla_detalle = Table(detalle, colWidths=[52, 28, 145, 52, 52, 38, 50, 38, 58], repeatRows=1)
                tabla_detalle.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 6.2), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cbd5e1")), ("ALIGN", (3, 1), (-1, -1), "CENTER"), ("BACKGROUND", (2, 1), (2, -1), colors.HexColor("#f0fdfa")), ("PADDING", (0, 0), (-1, -1), 2.5)]))
                resumen = [[_tr("Salario base"), f"B/. {salario_base:,.2f}", _tr("Horas extra"), f"B/. {total_horas_extras:,.2f}", _tr("Deducciones"), f"B/. {total_deducciones:,.2f}", _tr("NETO"), f"B/. {total_neto:,.2f}"]]
                tabla_resumen = Table(resumen, colWidths=[65, 62, 62, 62, 62, 62, 35, 70])
                tabla_resumen.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ecfdf5")), ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 7), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")), ("ALIGN", (1, 0), (-1, 0), "RIGHT"), ("PADDING", (0, 0), (-1, -1), 4)]))
                encabezado = Paragraph(f"<b>{_tr('COMPROBANTE DE PAGO')}</b> | {registro[1]} | {_tr('DUI')}: {registro[2]} | {_tr('Puesto')}: {registro[3]} | {_tr('Corte')} #{registro[17]} | {periodo}", estilos["Normal"])
                documento.build([encabezado, Spacer(1, 5), tabla_detalle, Spacer(1, 5), tabla_resumen, Spacer(1, 3), Paragraph(_tr("Detalle generado desde las marcaciones y el registro de planilla de SiPP. Los feriados se identifican para revisión del recargo que corresponda."), estilos["Normal"])])
                messagebox.showinfo("Comprobantes", "Comprobante generado correctamente.", parent=top_level)
                if sys.platform.startswith("win"):
                    os.startfile(ruta)
            except Exception as exc:
                messagebox.showerror("Comprobantes", f"No se pudo generar el PDF: {exc}", parent=top_level)

        acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
        acciones.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkButton(acciones, text="Imprimir comprobante", width=170, fg_color="#0f766e", hover_color="#115e59", command=imprimir_comprobante).pack(side="left")
        ctk.CTkButton(acciones, text="Actualizar", width=110, fg_color="#2563eb", hover_color="#1d4ed8", command=cargar_registros).pack(side="left", padx=(10, 0))
        ctk.CTkButton(acciones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
        tabla.bind("<<TreeviewSelect>>", seleccionar_registro)
        entrada_busqueda.bind("<KeyRelease>", cargar_registros)
        selector_departamento.configure(command=cargar_registros)
        selector_corte.configure(command=cargar_registros)
        selector_modo.configure(command=cargar_registros)
        try:
            corte_visualizado = self.combo_corte_planilla.get() if self.combo_corte_planilla is not None else ""
            corte_id_visualizado = None
            for corte_id, fecha_inicio, fecha_fin, estado in db.cortes_planilla().obtener_cortes_planilla() if db else []:
                texto_corte = (
                    f"Corte {corte_id}: {fecha_inicio.strftime('%d/%m/%Y') if hasattr(fecha_inicio, 'strftime') else fecha_inicio} "
                    f"a {fecha_fin.strftime('%d/%m/%Y') if hasattr(fecha_fin, 'strftime') else fecha_fin} ({estado})"
                )
                if texto_corte == corte_visualizado or corte_visualizado.strip().lower().startswith(f"corte {corte_id}:".lower()):
                    corte_id_visualizado = corte_id
                    break
            if corte_id_visualizado is None and corte_visualizado.strip().lower() == "corte activo":
                corte_activo = db.cortes_planilla().obtener_corte_activo() if db else None
                corte_id_visualizado = corte_activo[0] if corte_activo else None
            servicio_planilla = db.gestion_planilla() if db else None
            registros_del_corte = []
            if servicio_planilla:
                registros = servicio_planilla.obtener_planilla()
                if corte_id_visualizado is not None:
                    registros_del_corte = servicio_planilla.obtener_planilla_por_corte(corte_id_visualizado)
                    if registros_del_corte:
                        registros = registros_del_corte
            if corte_id_visualizado is None:
                etiqueta_seleccion.configure(text="Mostrando todos los registros de planilla.")
            elif not registros_del_corte:
                etiqueta_seleccion.configure(text="El corte seleccionado no tiene registros; mostrando los disponibles.")
            cargar_registros()
        except Exception as exc:
            etiqueta_seleccion.configure(text=f"No se pudieron cargar los registros: {exc}")
        
    def toplevel_corte_planilla(self):
        toplevel = ctk.CTkToplevel()
        ajustar_toplevel_a_pantalla(toplevel, 600, 420, margen_x=140, margen_y=120, min_ancho=430, min_alto=300)
        toplevel.grab_set()
        toplevel.title("Corte de Planilla")
        toplevel.resizable(False, False)
        
        # Fondo principal
        toplevel.configure(fg_color=("#f8fafc", "#0f172a"))
        
        # Verificar si existe un corte activo
        corte_activo = db.cortes_planilla().obtener_corte_activo()
        hay_corte_activo = corte_activo is not None
        
        # Header
        header = ctk.CTkFrame(
            toplevel,
            corner_radius=16,
            fg_color=("#eef2ff", "#1e293b"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
        )
        header.place(relx=0.04, rely=0.04, relwidth=0.92, relheight=0.20)
        
        titulo_header = "📋 Realizar Corte de Planilla" if hay_corte_activo else "📋 Crear Corte de Planilla"
        subtitulo_header = "Complete el corte de nómina actual" if hay_corte_activo else "Establezca el rango de fechas para crear un nuevo corte"
        
        ctk.CTkLabel(
            header,
            text=titulo_header,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=("#1e293b", "#f1f5f9"),
            anchor="w",
        ).place(relx=0.05, rely=0.15)
        
        ctk.CTkLabel(
            header,
            text=subtitulo_header,
            font=ctk.CTkFont(size=11),
            text_color="#94a3b8",
            anchor="w",
        ).place(relx=0.05, rely=0.55)
        
        # Panel de contenido
        contenedor = ctk.CTkFrame(
            toplevel,
            corner_radius=16,
            fg_color=("#ffffff", "#1e293b"),
            border_width=1,
            border_color=("#e2e8f0", "#334155"),
        )
        contenedor.place(relx=0.04, rely=0.28, relwidth=0.92, relheight=0.55)
        
        # Frame para los campos de fecha
        frame_fechas = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_fechas.place(relx=0.08, rely=0.12, relwidth=0.84, relheight=0.70)
        
        if hay_corte_activo:
            # Mostrar detalles del corte activo
            corte_id, fecha_inicio_corte, fecha_fin_corte, estado_corte = corte_activo
            fecha_hoy = datetime.now().date()
            fecha_fin_date = fecha_fin_corte if hasattr(fecha_fin_corte, 'date') else (
                datetime.strptime(str(fecha_fin_corte), "%Y-%m-%d").date() if isinstance(fecha_fin_corte, str) else fecha_fin_corte
            )
            puede_realizar_corte = fecha_hoy >= fecha_fin_date
            
            ctk.CTkLabel(
                frame_fechas,
                text="Corte Activo:",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1e293b", "#e2e8f0"),
                anchor="w",
            ).place(relx=0.0, rely=0.0, relwidth=0.4, relheight=0.15)
            
            ctk.CTkLabel(
                frame_fechas,
                text=f"ID: {corte_id}",
                font=ctk.CTkFont(size=11),
                text_color=("#475569", "#cbd5e1"),
                anchor="w",
            ).place(relx=0.0, rely=0.18, relwidth=0.4, relheight=0.12)
            
            ctk.CTkLabel(
                frame_fechas,
                text="Desde:",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1e293b", "#e2e8f0"),
                anchor="w",
            ).place(relx=0.0, rely=0.32, relwidth=0.4, relheight=0.12)
            
            ctk.CTkLabel(
                frame_fechas,
                text=str(fecha_inicio_corte),
                font=ctk.CTkFont(size=11),
                text_color=("#0f766e", "#10b981"),
                anchor="w",
            ).place(relx=0.0, rely=0.44, relwidth=0.4, relheight=0.12)
            
            ctk.CTkLabel(
                frame_fechas,
                text="Hasta:",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1e293b", "#e2e8f0"),
                anchor="w",
            ).place(relx=0.6, rely=0.32, relwidth=0.4, relheight=0.12)
            
            color_fecha_fin = "#10b981" if puede_realizar_corte else "#ef4444"
            ctk.CTkLabel(
                frame_fechas,
                text=str(fecha_fin_corte),
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=color_fecha_fin,
                anchor="w",
            ).place(relx=0.6, rely=0.44, relwidth=0.4, relheight=0.12)
            
            if not puede_realizar_corte:
                ctk.CTkLabel(
                    frame_fechas,
                    text=f"⏳ Falta {(fecha_fin_date - fecha_hoy).days} día(s)",
                    font=ctk.CTkFont(size=10),
                    text_color="#ef4444",
                    anchor="w",
                ).place(relx=0.6, rely=0.58, relwidth=0.4, relheight=0.12)
        else:
            # Mostrar campos para crear nuevo corte
            ctk.CTkLabel(
                frame_fechas,
                text="Fecha Desde:",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1e293b", "#e2e8f0"),
                anchor="w",
            ).place(relx=0.0, rely=0.0, relwidth=0.4, relheight=0.15)
            
            entrada_desde = DateEntry(
                frame_fechas,
                width=25,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                font=ctk.CTkFont(size=11),
                year=datetime.now().year,
                month=datetime.now().month,
                day=datetime.now().day,
            )
            entrada_desde.place(relx=0.0, rely=0.25, relwidth=0.4, relheight=0.18)
            
            ctk.CTkLabel(
                frame_fechas,
                text="Fecha Hasta:",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1e293b", "#e2e8f0"),
                anchor="w",
            ).place(relx=0.6, rely=0.0, relwidth=0.4, relheight=0.15)
            
            entrada_hasta = DateEntry(
                frame_fechas,
                width=25,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                font=ctk.CTkFont(size=11),
                year=datetime.now().year,
                month=datetime.now().month,
                day=datetime.now().day,
            )
            entrada_hasta.place(relx=0.6, rely=0.25, relwidth=0.4, relheight=0.18)
        
        # Botones en la parte inferior
        frame_botones = ctk.CTkFrame(toplevel, fg_color="transparent")
        frame_botones.place(relx=0.04, rely=0.88, relwidth=0.92, relheight=0.08)
        
        if hay_corte_activo:
            # Modo: Realizar Corte (cerrar el corte activo)
            corte_id, fecha_inicio_corte, fecha_fin_corte, estado_corte = corte_activo
            fecha_hoy = datetime.now().date()
            fecha_fin_date = fecha_fin_corte if hasattr(fecha_fin_corte, 'date') else (
                datetime.strptime(str(fecha_fin_corte), "%Y-%m-%d").date() if isinstance(fecha_fin_corte, str) else fecha_fin_corte
            )
            puede_realizar_corte = fecha_hoy >= fecha_fin_date
            
            def realizar_corte_activo():
                if not puede_realizar_corte:
                    messagebox.showerror("Error", f"No puede realizar el corte. La fecha fin es {fecha_fin_date} y hoy es {fecha_hoy}")
                    return
                
                confirmar = messagebox.askyesno(
                    "Confirmar Cierre",
                    f"¿Está seguro de que desea cerrar el corte?\n\nDesde: {fecha_inicio_corte}\nHasta: {fecha_fin_corte}\n\nEsta acción no se puede deshacer."
                )
                if not confirmar:
                    return
                
                try:
                    resultado = db.cortes_planilla().cerrar_corte_activo(fecha_fin_date)
                    if resultado:
                        messagebox.showinfo("Éxito", f"✓ Corte de planilla realizado\n\nDesde: {fecha_inicio_corte}\nHasta: {fecha_fin_corte}\n\nAhora puede crear un nuevo corte.")
                        toplevel.destroy()
                    else:
                        messagebox.showerror("Error", "No se pudo cerrar el corte.")
                except Exception as exc:
                    messagebox.showerror("Error", f"Error al cerrar el corte: {exc}")
            
            ctk.CTkButton(
                frame_botones,
                text="Cancelar",
                width=120,
                height=36,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#4b5563",
                hover_color="#374151",
                text_color="white",
                command=toplevel.destroy,
            ).pack(side="left", padx=10, expand=True, fill="both")
            
            ctk.CTkButton(
                frame_botones,
                text="Realizar Corte",
                width=120,
                height=36,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#10b981" if puede_realizar_corte else "#6b7280",
                hover_color="#059669" if puede_realizar_corte else "#4b5563",
                text_color="white",
                state="normal" if puede_realizar_corte else "disabled",
                command=realizar_corte_activo if puede_realizar_corte else None,
            ).pack(side="left", padx=10, expand=True, fill="both")
        else:
            # Modo: Crear Nuevo Corte
            def crear_nuevo_corte():
                fecha_desde = entrada_desde.get_date()
                fecha_hasta = entrada_hasta.get_date()
                
                if fecha_desde > fecha_hasta:
                    messagebox.showerror("Error", "La fecha 'Desde' no puede ser mayor a la fecha 'Hasta'")
                    return
                
                try:
                    resultado = db.cortes_planilla().insertar_corte_planilla(fecha_desde, fecha_hasta)
                    if resultado:
                        messagebox.showinfo("Éxito", f"✓ Corte de planilla creado\n\nDesde: {fecha_desde}\nHasta: {fecha_hasta}\n\nAhora puede registrar marcaciones.")
                        toplevel.destroy()
                    else:
                        messagebox.showerror("Error", "No se pudo crear el corte. Verifique que no exista un corte activo.")
                except Exception as exc:
                    messagebox.showerror("Error", f"Error al crear el corte: {exc}")
            
            ctk.CTkButton(
                frame_botones,
                text="Cancelar",
                width=120,
                height=36,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#4b5563",
                hover_color="#374151",
                text_color="white",
                command=toplevel.destroy,
            ).pack(side="left", padx=10, expand=True, fill="both")
            
            ctk.CTkButton(
                frame_botones,
                text="Nuevo Corte",
                width=120,
                height=36,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#2563eb",
                hover_color="#1d4ed8",
                text_color="white",
                command=crear_nuevo_corte,
            ).pack(side="left", padx=10, expand=True, fill="both")






class Gestion_de_tiempo:

    def __init__(self):
        self.frame_vista = None

    def _crear_frame_vista(self, sipp):
        if getattr(self, "frame_vista", None) is not None:
            try:
                if self.frame_vista.winfo_exists():
                    self.frame_vista.destroy()
            except Exception:
                pass
        self.frame_vista = ctk.CTkFrame(sipp)
        self.frame_vista.place(relx=0.00, rely=0.2, relwidth=1.0, relheight=0.78)

    def ventana_gestion_de_tiempo(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

        frame_principal = ctk.CTkFrame(frame, fg_color="transparent")
        frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        def mostrar_tablero():
            for widget in frame_principal.winfo_children():
                widget.destroy()

            cabecera = ctk.CTkFrame(frame_principal, fg_color="transparent")
            cabecera.place(relx=0.03, rely=0.03, relwidth=0.94, relheight=0.10)

            ctk.CTkLabel(
                cabecera,
                text="Gestión de Tiempos",
                font=ctk.CTkFont(size=28, weight="bold"),
                anchor="w",
            ).place(relx=0.00, rely=0.08)
            ctk.CTkLabel(
                cabecera,
                text="Elige una opción para abrirla a pantalla completa.",
                font=ctk.CTkFont(size=12),
                text_color="gray75",
                anchor="w",
            ).place(relx=0.00, rely=0.62)

            panel_accesos = ctk.CTkFrame(frame_principal, corner_radius=18, border_width=1, border_color="#3f3f46")
            panel_accesos.place(relx=0.03, rely=0.16, relwidth=0.94, relheight=0.18)

            acciones = [
                ("Añadir Marcación", "Registra entradas y salidas manualmente.", lambda: abrir_seccion(lambda cont: self.añadir_marcacion_manual(cont))),
                ("Ver Marcaciones", "Consulta todas las marcaciones por corte.", lambda: abrir_seccion(lambda cont: self.ver_marcaciones(cont))),
                ("Buscar Registro", "Revisa y aprueba marcaciones por empleado.", lambda: abrir_seccion(lambda cont: self.ver_registro(cont))),
                ("Cortes de Fechas", "Explora el historial y descarga cortes.", lambda: abrir_seccion(lambda cont: self.buscar_cortes_fechas(cont))),
            ]

            ancho_tarjeta = 0.22
            espacios = [0.03, 0.265, 0.50, 0.735]
            for indice, (titulo, descripcion, comando) in enumerate(acciones):
                tarjeta = ctk.CTkFrame(panel_accesos, corner_radius=14, border_width=1, border_color="#334155")
                tarjeta.place(relx=espacios[indice], rely=0.14, relwidth=ancho_tarjeta, relheight=0.70)
                ctk.CTkLabel(tarjeta, text=titulo, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").place(relx=0.07, rely=0.14)
                ctk.CTkLabel(
                    tarjeta,
                    text=descripcion,
                    font=ctk.CTkFont(size=11),
                    text_color="gray75",
                    justify="left",
                    wraplength=180,
                    anchor="w",
                ).place(relx=0.07, rely=0.40)
                boton_buscar_profesional(tarjeta, text="Abrir", command=comando, width=96, height=32, fg_color="#0F766E", hover_color="#115E59").place(relx=0.07, rely=0.74)

            preview = ctk.CTkFrame(frame_principal, corner_radius=18, border_width=1, border_color="#3f3f46")
            preview.place(relx=0.03, rely=0.38, relwidth=0.94, relheight=0.56)
            ctk.CTkLabel(preview, text="Vista previa", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").place(relx=0.04, rely=0.06)
            ctk.CTkLabel(
                preview,
                text="Selecciona una opción para ocupar toda la pantalla de contenido. Cada módulo se abrirá aquí con un botón X para volver.",
                font=ctk.CTkFont(size=12),
                text_color="gray75",
                wraplength=520,
                justify="left",
                anchor="w",
            ).place(relx=0.04, rely=0.18)

        def abrir_seccion(comando):
            for widget in frame_principal.winfo_children():
                widget.destroy()

            shell = ctk.CTkFrame(frame_principal, corner_radius=0, fg_color=("#f5f7fb", "#111827"))
            shell.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)

            ctk.CTkButton(
                shell,
                text="X",
                width=36,
                height=36,
                corner_radius=10,
                fg_color="#4b5563",
                hover_color="#374151",
                command=mostrar_tablero,
            ).place(relx=0.98, rely=0.02, anchor="ne")

            contenido = ctk.CTkFrame(shell, fg_color="transparent")
            contenido.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)
            comando(contenido)

        mostrar_tablero()

        

    def añadir_marcacion_manual(self, sipp):
        self._crear_frame_vista(sipp)
        self.frame_vista.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)

        contenedor = ctk.CTkFrame(self.frame_vista, fg_color="transparent")
        contenedor.place(relx=0.01, rely=0.00, relwidth=0.98, relheight=1.0)

        cabecera = ctk.CTkFrame(contenedor, fg_color="transparent")
        cabecera.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=0.10)

        ctk.CTkLabel(
            cabecera,
            text="Añadir Marcación",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).place(relx=0.00, rely=0.05)

        ctk.CTkLabel(
            cabecera,
            text="Registro manual profesional de entrada y salida del personal",
            font=ctk.CTkFont(size=13),
            text_color="gray75",
        ).place(relx=0.00, rely=0.58)

        barra_busqueda = ctk.CTkFrame(contenedor, corner_radius=14, border_width=1, border_color="#4a4a4a")
        barra_busqueda.place(relx=0.00, rely=0.11, relwidth=1.0, relheight=0.12)
        barra_busqueda.grid_columnconfigure(0, weight=0)
        barra_busqueda.grid_columnconfigure(1, weight=0)

        self.entrada_busqueda_personal = ctk.CTkEntry(
            barra_busqueda,
            width=240,
            height=32,
            placeholder_text="Nombre o DUI",
            justify="left",
            font=ctk.CTkFont(size=13),
        )
        self.entrada_busqueda_personal.grid(row=0, column=0, padx=(16, 8), pady=14, sticky="w")

        boton_buscar = boton_buscar_profesional(barra_busqueda, command=lambda: self.añadir_manual(self.frame_vista), width=96)
        boton_buscar.grid(row=0, column=1, padx=(0, 16), pady=14, sticky="w")

        self.entrada_busqueda_personal.bind("<Return>", lambda event: self.añadir_manual(self.frame_vista))
        self.entrada_busqueda_personal.focus_set()

        
    def return_manual(self, event, sipp):
        self.añadir_manual(self, sipp)

    def formatear_fecha(self, fecha):
        if not fecha:
            return None
        if isinstance(fecha, datetime):
            return fecha.strftime("%Y-%m-%d")
        if isinstance(fecha, str):
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(fecha, formato).strftime("%Y-%m-%d")
                except ValueError:
                    continue
        return str(fecha)

    def añadir_manual(self, sipp):
        entrada_texto = self.entrada_busqueda_personal.get().strip()
        
        if not entrada_texto:
            messagebox.showerror("Campos vacíos", "Por favor ingrese un nombre o DUI del empleado")
            self.entrada_busqueda_personal.focus_set()
            return
        
        dato = db.gestion_empleado().buscar_empleado(entrada_texto.title())
        if dato is None:
            messagebox.showerror("Error", "No se encontró el empleado")
            self.entrada_busqueda_personal.delete(0, "end")
            return
        elif dato == "":
            messagebox.showerror("Error", "No se encontró el empleado")
            self.entrada_busqueda_personal.delete(0, "end")
            return

        self.entrada_busqueda_personal.delete(0, "end")
        nombre = dato[1]
        dui = dato[2]

        if hasattr(self, "frame_formulario_marcacion") and self.frame_formulario_marcacion.winfo_exists():
            self.frame_formulario_marcacion.destroy()

        self.frame_formulario_marcacion = ctk.CTkFrame(sipp, corner_radius=12, border_width=1, border_color="#3f3f3f")
        self.frame_formulario_marcacion.place(relx=0.01, rely=0.25, relwidth=0.98, relheight=0.72)

        frame = self.frame_formulario_marcacion

        header_form = ctk.CTkFrame(frame, fg_color="transparent")
        header_form.place(relx=0.02, rely=0.03, relwidth=0.96, relheight=0.14)
        label_nombre = ctk.CTkLabel(header_form, text=f"{nombre}", font=ctk.CTkFont(size=24, weight="bold"), anchor="w")
        label_nombre.place(relx=0.00, rely=0.05)
        label_dui = ctk.CTkLabel(header_form, text=f"DUI: {dui}", font=ctk.CTkFont(size=14), text_color="gray70", anchor="w")
        label_dui.place(relx=0.00, rely=0.62)

        tarjeta_horarios = ctk.CTkFrame(frame, corner_radius=10)
        tarjeta_horarios.place(relx=0.02, rely=0.20, relwidth=0.56, relheight=0.48)

        tarjeta_estado = ctk.CTkFrame(frame, corner_radius=10)
        tarjeta_estado.place(relx=0.60, rely=0.20, relwidth=0.38, relheight=0.48)

        label_nombre = ctk.CTkLabel(tarjeta_horarios, text="Entrada", font=ctk.CTkFont(size=16, weight="bold"))
        label_nombre.place(relx=0.04, rely=0.14)

        def validar_hora(text):
            if text == "":
                return True
            if not text.isdigit() or len(text) > 2:
                return False
            return 0 <= int(text) <= 12

        def validar_minuto(text):
            if text == "":
                return True
            return text.isdigit() and 0 <= int(text) <= 59

        def actualizar_estado_ausencia(valor):
            estado_dia = str(valor).strip().upper()
            if estado_dia in {"DÍA LIBRE", "DIA LIBRE", "AUSENCIA", "CERTIFICADO MÉDICO", "CERTIFICADO MEDICO", "INCAPACIDAD", "CONSTANCIA DE ASISTENCIA"}:
                for widget in [entrada_hora, entrada_minuto, salida_hora, salida_minuto, zona_entrada, zona_salida]:
                    widget.configure(state="disabled")
                    if widget in [entrada_hora, entrada_minuto, salida_hora, salida_minuto]:
                        widget.delete(0, "end")
                        widget.insert(0, "0")
                    else:
                        widget.set("AM")
            else:
                for widget in [entrada_hora, entrada_minuto, salida_hora, salida_minuto, zona_entrada, zona_salida]:
                    widget.configure(state="normal")
                    if widget in [entrada_hora, entrada_minuto, salida_hora, salida_minuto]:
                        widget.delete(0, "end")
                    else:
                        widget.set("")
                entrada_hora.focus_set()

        entrada_hora = ctk.CTkEntry(
            tarjeta_horarios,
            placeholder_text="00",
            width=30,
            justify="center",
            validate="key",
            validatecommand=(tarjeta_horarios.register(validar_hora), "%P"),
        )
        entrada_hora.place(relx=0.18, rely=0.14)
        label = ctk.CTkLabel(tarjeta_horarios, text=":")
        label.place(relx=0.29, rely=0.14)
        entrada_minuto = ctk.CTkEntry(
            tarjeta_horarios,
            placeholder_text="00",
            width=30,
            justify="center",
            validate="key",
            validatecommand=(tarjeta_horarios.register(validar_minuto), "%P"),
        )
        entrada_minuto.place(relx=0.33, rely=0.14)
        zona_entrada = ctk.CTkComboBox(tarjeta_horarios, values=["AM", "PM"], width=85)
        zona_entrada.place(relx=0.47, rely=0.14)



        label_nombre = ctk.CTkLabel(tarjeta_horarios, text="Salida", font=ctk.CTkFont(size=16, weight="bold"))
        label_nombre.place(relx=0.04, rely=0.42)

        salida_hora = ctk.CTkEntry(
            tarjeta_horarios,
            placeholder_text="00",
            width=30,
            justify="center",
            validate="key",
            validatecommand=(tarjeta_horarios.register(validar_hora), "%P"),
        )
        salida_hora.place(relx=0.18, rely=0.42)
        label2 = ctk.CTkLabel(tarjeta_horarios, text=":")
        label2.place(relx=0.29, rely=0.42)
        salida_minuto = ctk.CTkEntry(
            tarjeta_horarios,
            placeholder_text="00",
            width=30,
            justify="center",
            validate="key",
            validatecommand=(tarjeta_horarios.register(validar_minuto), "%P"),
        )
        salida_minuto.place(relx=0.33, rely=0.42)
        zona_salida = ctk.CTkComboBox(tarjeta_horarios, values=["AM", "PM"], width=85)
        zona_salida.place(relx=0.47, rely=0.42)

        
        ctk.CTkLabel(tarjeta_estado, text="Control del Día", font=ctk.CTkFont(size=17, weight="bold")).place(relx=0.08, rely=0.10)

        # ===== SECCIÓN DE FECHA MEJORADA =====
        ctk.CTkLabel(tarjeta_estado, text="Fecha", font=ctk.CTkFont(size=14, weight="bold")).place(relx=0.08, rely=0.28)
        
        # Entry para fecha con mejor visualización
        fecha = DateEntry(
            tarjeta_estado,
            width=18,
            font=("Segoe UI", 12),
            date_pattern="dd/MM/yyyy",
            justify="center",
        )
        fecha.place(relx=0.44, rely=0.28, relwidth=0.38)
        ctk.CTkButton(
            tarjeta_estado,
            text="📅",
            width=36,
            height=30,
            command=fecha.open_calendar,
        ).place(relx=0.84, rely=0.28)

        # ===== SECCIÓN DE ESTADO DEL DÍA =====
        ctk.CTkLabel(tarjeta_estado, text="Tipo de día", font=ctk.CTkFont(size=14, weight="bold")).place(relx=0.08, rely=0.58)

        respuesta = ctk.CTkComboBox(
            tarjeta_estado,
            values=["JORNADA LABORAL", "DÍA LIBRE", "AUSENCIA", "CERTIFICADO MÉDICO", "INCAPACIDAD", "CONSTANCIA DE ASISTENCIA"],
            width=170,
            command=actualizar_estado_ausencia,
            state="readonly",
        )
        respuesta.place(relx=0.52, rely=0.58)
        respuesta.set("JORNADA LABORAL")
        actualizar_estado_ausencia("JORNADA LABORAL")

        def actualizar_estado_boton(event=None):
            fecha_formateada = self.formatear_fecha(fecha.get())
            existe = db.marcaciones.buscar_marcacion(dui, fecha_formateada) if fecha_formateada else None
            if existe:
                boton_agregar.configure(text="Actualizar", fg_color="green")
            else:
                boton_agregar.configure(text="Guardar", fg_color="#1f6aa5")

        def ejecutar_accion():
            fecha_formateada = self.formatear_fecha(fecha.get())
            datos = self.condicion_marcacion(
                respuesta.get(),
                entrada_hora.get(),
                entrada_minuto.get(),
                salida_hora.get(),
                salida_minuto.get(),
                zona_entrada.get(),
                zona_salida.get(),
            )
            if boton_agregar.cget("text") == "Actualizar":
                operacion_exitosa = self.actualizar_marcacion_manual(nombre, dui, *datos, fecha_formateada)
            else:
                operacion_exitosa = self.guardar_marcacion_manual(nombre, dui, *datos, fecha_formateada)

            if operacion_exitosa:
                continuar = messagebox.askyesno(
                    "Continuar",
                    "¿Desea seguir trabajando con este usuario?"
                )
                if not continuar:
                    self.añadir_marcacion_manual(sipp.master)

        barra_acciones = ctk.CTkFrame(frame, fg_color="transparent")
        barra_acciones.place(relx=0.02, rely=0.74, relwidth=0.96, relheight=0.20)

        boton_agregar = ctk.CTkButton(barra_acciones, text="Guardar", width=130, height=36, fg_color="#0A9D58", command=ejecutar_accion)
        boton_agregar.place(relx=0.00, rely=0.30)

        ctk.CTkButton(
            barra_acciones,
            text="Limpiar",
            width=120,
            height=36,
            fg_color="#4a4a4a",
            command=lambda: self.añadir_marcacion_manual(sipp.master),
        ).place(relx=0.16, rely=0.30)

        fecha.bind("<<DateEntrySelected>>", actualizar_estado_boton)
        actualizar_estado_boton()

    def condicion_marcacion(self, estado_dia, entrada_h, entrada_m, salida_h, salida_m, entrada_zona="AM", salida_zona="AM"):
        def normalizar_hora(valor):
            texto = str(valor or "").strip()
            if texto == "":
                return "00"
            try:
                return f"{int(texto):02d}"
            except ValueError:
                return "00"

        def normalizar_minuto(valor):
            texto = str(valor or "").strip()
            if texto == "":
                return "00"
            try:
                return f"{int(texto):02d}"
            except ValueError:
                return "00"

        estado_normalizado = str(estado_dia or "JORNADA LABORAL").strip().upper()
        if estado_normalizado != "JORNADA LABORAL":
            return "00", "00", "00", "00", (entrada_zona or "AM").upper(), (salida_zona or "AM").upper(), estado_normalizado

        return (
            normalizar_hora(entrada_h),
            normalizar_minuto(entrada_m),
            normalizar_hora(salida_h),
            normalizar_minuto(salida_m),
            (entrada_zona or "AM").upper(),
            (salida_zona or "AM").upper(),
            estado_normalizado,
        )


    def guardar_marcacion_manual(self, nombre, dui, entrada_h, entrada_m, salida_h, salida_m, entrada_zona, salida_zona, estado_dia, fecha):
        try:
            fecha_formateada = self.formatear_fecha(fecha)
            entrada_hora = str(entrada_h or "00")
            entrada_minuto = str(entrada_m or "00")
            salida_hora = str(salida_h or "00")
            salida_minuto = str(salida_m or "00")

            entrada = f"{entrada_hora}:{entrada_minuto} {str(entrada_zona or 'AM').upper()}"
            salida = f"{salida_hora}:{salida_minuto} {str(salida_zona or 'AM').upper()}"

            insertado = db.marcaciones.insertar_marcacion(nombre, dui, entrada, salida, fecha_formateada, estado_dia)
            if insertado:
                return True
            else:
                messagebox.showwarning("Registro existente", "Ya existe una marcación para este empleado en la fecha seleccionada. Use Actualizar para modificarla.")
                return False
        except Exception as exc:
            messagebox.showerror("Error de sistema", f"No se pudo guardar la marcación: {exc}")
            return False

    def actualizar_marcacion_manual(self, nombre, dui, entrada_h, entrada_m, salida_h, salida_m, entrada_zona, salida_zona, estado_dia, fecha):
        try:
            fecha_formateada = self.formatear_fecha(fecha)
            entrada_hora = str(entrada_h or "00")
            entrada_minuto = str(entrada_m or "00")
            salida_hora = str(salida_h or "00")
            salida_minuto = str(salida_m or "00")

            entrada = f"{entrada_hora}:{entrada_minuto} {str(entrada_zona or 'AM').upper()}"
            salida = f"{salida_hora}:{salida_minuto} {str(salida_zona or 'AM').upper()}"

            actualizado = db.marcaciones().actualizar_marcacion(dui, fecha_formateada, entrada, salida, nombre, estado_dia)
            if actualizado:
                return True
            else:
                messagebox.showwarning("Sin resultados", "No existe una marcación para actualizar con los datos proporcionados.")
                return False
        except Exception as exc:
            messagebox.showerror("Error de sistema", f"No se pudo actualizar la marcación: {exc}")
            return False



#aqui se pueden ver las marcaciones de todo el personal dependiendo de la fecha 
    def ver_marcaciones(self, sipp):
        self._crear_frame_vista(sipp)
        self.frame_vista.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)

        header_frame = ctk.CTkFrame(self.frame_vista, fg_color="transparent")
        header_frame.place(relx=0.03, rely=0.03, relwidth=0.94, relheight=0.12)

        ctk.CTkLabel(
            header_frame,
            text="Ver Marcaciones",
            font=ctk.CTkFont(size=26, weight="bold"),
            anchor="w",
        ).place(relx=0.00, rely=0.02)
        ctk.CTkLabel(
            header_frame,
            text="Consulta marcaciones por corte, busca empleados y revisa jornadas registradas.",
            font=ctk.CTkFont(size=12),
            text_color="gray75",
            anchor="w",
        ).place(relx=0.00, rely=0.52)

        corte_card = ctk.CTkFrame(self.frame_vista, corner_radius=16, border_width=1, border_color="#3f3f46")
        corte_card.place(relx=0.03, rely=0.17, relwidth=0.94, relheight=0.15)
        corte_card.grid_columnconfigure(0, weight=2)
        corte_card.grid_columnconfigure(1, weight=3)
        corte_card.grid_columnconfigure(2, weight=1)
        corte_card.grid_columnconfigure(3, weight=1)
        corte_card.grid_columnconfigure(4, weight=1)
        corte_card.grid_columnconfigure(5, weight=4)
        corte_card.grid_rowconfigure(1, weight=1)

        screen_w = sipp.winfo_screenwidth()

        etiqueta_corte = ctk.CTkLabel(corte_card, text="", font=ctk.CTkFont(size=12, weight="bold"))
        etiqueta_corte.grid(row=0, column=0, columnspan=5, sticky="w", padx=14, pady=(10, 2))

        entrada_busqueda = ctk.CTkEntry(
            corte_card,
            width=max(220, min(360, int(screen_w * 0.22))),
            placeholder_text="Nombre o Dui",
            justify="center",
            font=ctk.CTkFont(size=16),
        )
        entrada_busqueda.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(14, 8), pady=(2, 12))

        boton_buscar = boton_buscar_profesional(corte_card, width=96)
        boton_buscar.grid(row=1, column=2, sticky="w", padx=(0, 12), pady=(2, 12))

        scrolframe = ctk.CTkScrollableFrame(self.frame_vista)
        scrolframe.place(relx=0.03, rely=0.35, relwidth=0.94, relheight=0.60)

        servicio_marcaciones = db.marcaciones()
        servicio_cortes = db.cortes_planilla()
        servicio_empleados = db.gestion_empleado()

        etiqueta_depto = ctk.CTkLabel(corte_card, text="Depto:", font=ctk.CTkFont(size=12, weight="bold"))
        etiqueta_depto.grid(row=1, column=3, sticky="e", padx=(0, 8), pady=(2, 12))
        
        combo_departamento = ctk.CTkComboBox(
            corte_card,
            width=max(140, min(220, int(screen_w * 0.15))),
            values=["TODOS"],
            command=lambda valor: cargar_marcaciones(),
        )
        combo_departamento.grid(row=1, column=4, sticky="ew", padx=(0, 8), pady=(2, 12))
        combo_departamento.set("TODOS")

        columnas = ["Nombre Completo", "Numero de Identificacion", "Entrada", "Salida", "Fecha"]

        padx_val = max(8, min(100, int(screen_w * 0.02)))

        def convertir_a_fecha(fecha_valor):
            if hasattr(fecha_valor, "year") and hasattr(fecha_valor, "month") and hasattr(fecha_valor, "day"):
                return fecha_valor
            texto = str(fecha_valor).strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto, formato).date()
                except ValueError:
                    continue
            return None

        def formatear_fecha_marcacion(fecha_valor):
            fecha_convertida = convertir_a_fecha(fecha_valor)
            if fecha_convertida:
                return fecha_convertida.strftime("%d/%m/%Y")
            return str(fecha_valor)

        def clave_orden_fecha(valor):
            try:
                return datetime.strptime(valor, "%d/%m/%Y")
            except ValueError:
                return datetime.min

        def dibujar_fila(parent, valores, es_encabezado=False):
            row_frame = ctk.CTkFrame(parent, fg_color="transparent", height=32)
            row_frame.pack(fill='x', padx=padx_val, pady=(0, 2))
            row_frame.grid_propagate(False)

            cantidad = len(columnas)
            for indice_col in range(cantidad * 2 - 1):
                if indice_col % 2 == 0:
                    row_frame.grid_columnconfigure(indice_col, weight=1, uniform="cols_marcaciones")
                else:
                    row_frame.grid_columnconfigure(indice_col, weight=0)

            color_linea = "red" if es_encabezado else "gray40"
            fuente = ctk.CTkFont(size=14, weight="bold") if es_encabezado else ctk.CTkFont(size=14)

            for indice, valor in enumerate(valores):
                columna = indice * 2
                celda = ctk.CTkFrame(row_frame, fg_color="transparent", height=32)
                celda.grid(row=0, column=columna, sticky="nsew")
                celda.grid_propagate(False)

                label_valor = ctk.CTkLabel(celda, text=str(valor), anchor="w", font=fuente)
                label_valor.pack(fill='both', padx=(10, 6), pady=4)

                if indice < cantidad - 1:
                    linea = ctk.CTkFrame(row_frame, width=2, fg_color=color_linea)
                    linea.grid(row=0, column=columna + 1, sticky="ns", pady=4)

        cortes_guardados = servicio_cortes.obtener_cortes_planilla()
        corte_mapa = {}
        valores_cortes = []

        if cortes_guardados:
            cortes_ordenados = sorted(cortes_guardados, key=lambda corte: corte[0], reverse=True)
            corte_activo = servicio_cortes.obtener_corte_activo()
            corte_activo_id = corte_activo[0] if corte_activo else None

            for corte in cortes_ordenados:
                corte_id, fecha_inicio, fecha_fin, estado, *_ = corte
                texto_corte = f"Corte {corte_id}: {formatear_fecha_marcacion(fecha_inicio)} a {formatear_fecha_marcacion(fecha_fin)} ({estado})"
                valores_cortes.append(texto_corte)
                corte_mapa[texto_corte] = (fecha_inicio, fecha_fin, estado, corte_id)

            valor_inicial = None
            if corte_activo_id is not None:
                for texto_corte, datos_corte in corte_mapa.items():
                    if datos_corte[3] == corte_activo_id:
                        valor_inicial = texto_corte
                        break
            if valor_inicial is None and valores_cortes:
                valor_inicial = valores_cortes[0]
        else:
            valor_inicial = ""

        combobox_corte = ctk.CTkComboBox(
            corte_card,
            width=max(240, min(460, int(screen_w * 0.28))),
            values=valores_cortes,
            command=lambda valor: cargar_marcaciones(valor),
        )
        combobox_corte.grid(row=1, column=5, sticky="ew", padx=(0, 14), pady=(0, 10))

        if valor_inicial:
            combobox_corte.set(valor_inicial)

        def cargar_marcaciones(valor_corte=None):
            for widget in scrolframe.winfo_children():
                widget.destroy()

            criterio = entrada_busqueda.get().strip().lower()
            departamento_seleccionado = combo_departamento.get().strip().upper() or "TODOS"
            corte_seleccionado = valor_corte or combobox_corte.get().strip()

            # Obtener departamentos de empleados para actualizar el combobox
            try:
                empleados_db = servicio_empleados.obtener_empleados()
                departamentos = set()
                dui_a_departamento = {}
                
                for empleado in empleados_db:
                    if len(empleado) < 8:
                        continue
                    dui_emp = str(empleado[2] or "").strip()
                    departamento_emp = str(empleado[7] or "SIN DEPARTAMENTO").strip().upper()
                    if not departamento_emp:
                        departamento_emp = "SIN DEPARTAMENTO"
                    
                    if dui_emp:
                        dui_a_departamento[dui_emp] = departamento_emp
                    departamentos.add(departamento_emp)
                
                valores_depto = ["TODOS"] + sorted(departamentos)
                combo_departamento.configure(values=valores_depto)
                
                if departamento_seleccionado not in valores_depto:
                    departamento_seleccionado = "TODOS"
                    combo_departamento.set("TODOS")
            except Exception:
                dui_a_departamento = {}

            if not corte_seleccionado or corte_seleccionado not in corte_mapa:
                etiqueta_corte.configure(text="No hay cortes disponibles", text_color="gray50")
                label_vacio = ctk.CTkLabel(scrolframe, text="No hay cortes para mostrar marcaciones.", anchor="w")
                label_vacio.pack(fill='x', padx=padx_val, pady=10)
                return

            fecha_inicio_corte, fecha_fin_corte, estado_corte, _ = corte_mapa[corte_seleccionado]
            fecha_inicio_corte = convertir_a_fecha(fecha_inicio_corte)
            fecha_fin_corte = convertir_a_fecha(fecha_fin_corte)

            etiqueta_corte.configure(
                text=corte_seleccionado,
                text_color="#2aa745" if str(estado_corte).upper() == "ACTIVO" else "#d08a00",
            )

            if not fecha_inicio_corte or not fecha_fin_corte:
                label_vacio = ctk.CTkLabel(scrolframe, text="No se pudo leer el rango del corte seleccionado.", anchor="w")
                label_vacio.pack(fill='x', padx=padx_val, pady=10)
                return

            marcaciones_db = servicio_marcaciones.obtener_marcaciones_por_rango(fecha_inicio_corte, fecha_fin_corte)

            marcaciones_por_fecha = {}
            for registro in marcaciones_db:
                if len(registro) < 6:
                    continue

                _, nombre, dui, hora_entrada, hora_salida, fecha_registro = registro[:6]
                nombre_txt = str(nombre or "")
                dui_txt = str(dui or "")
                departamento_txt = dui_a_departamento.get(dui_txt, "SIN DEPARTAMENTO")

                if criterio and criterio not in nombre_txt.lower() and criterio not in dui_txt.lower():
                    continue
                
                if departamento_seleccionado != "TODOS" and departamento_txt != departamento_seleccionado:
                    continue

                fecha_texto = formatear_fecha_marcacion(fecha_registro)
                marcaciones_por_fecha.setdefault(fecha_texto, []).append((
                    nombre_txt,
                    dui_txt,
                    str(hora_entrada or ""),
                    str(hora_salida or ""),
                    fecha_texto,
                ))

            if not marcaciones_por_fecha:
                mensaje = "No hay marcaciones dentro del corte seleccionado." if not criterio else "No se encontraron marcaciones para la busqueda dentro del corte seleccionado."
                label_vacio = ctk.CTkLabel(scrolframe, text=mensaje, anchor="w")
                label_vacio.pack(fill='x', padx=padx_val, pady=10)
                return

            fechas_ordenadas = sorted(marcaciones_por_fecha.keys(), key=clave_orden_fecha, reverse=True)

            for fecha_texto in fechas_ordenadas:
                label_fecha = ctk.CTkLabel(
                    scrolframe,
                    text=fecha_texto,
                    anchor="w",
                    font=ctk.CTkFont(size=16, weight="bold"),
                )
                label_fecha.pack(fill='x', padx=padx_val, pady=(8, 4))

                dibujar_fila(scrolframe, columnas, es_encabezado=True)

                for fila in marcaciones_por_fecha[fecha_texto]:
                    dibujar_fila(scrolframe, fila, es_encabezado=False)

                separador = ctk.CTkFrame(scrolframe, fg_color="gray35", height=2)
                separador.pack(fill='x', padx=padx_val, pady=(4, 8))

        boton_buscar.configure(command=cargar_marcaciones)
        entrada_busqueda.bind("<Return>", lambda event: cargar_marcaciones())
        cargar_marcaciones()
    


    def ver_registro(self, sipp):
        self._crear_frame_vista(sipp)
        self.frame_vista.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)

        cabecera = ctk.CTkFrame(self.frame_vista, fg_color="transparent")
        cabecera.place(relx=0.03, rely=0.03, relwidth=0.94, relheight=0.13)

        ctk.CTkLabel(
            cabecera,
            text="Buscar Registro",
            font=ctk.CTkFont(size=26, weight="bold"),
            anchor="w",
        ).place(relx=0.00, rely=0.02)
        ctk.CTkLabel(
            cabecera,
            text="Supervisa horas extras, filtra empleados y gestiona aprobaciones por corte activo.",
            font=ctk.CTkFont(size=12),
            text_color="gray75",
            anchor="w",
        ).place(relx=0.00, rely=0.52)

        etiqueta_corte_activo = ctk.CTkLabel(cabecera, text="", font=ctk.CTkFont(size=12, weight="bold"), anchor="w")
        etiqueta_corte_activo.place(relx=0.00, rely=0.76)

        screen_w = sipp.winfo_screenwidth()

        frame_filtros = ctk.CTkFrame(self.frame_vista, corner_radius=16, border_width=1, border_color="#3f3f46")
        frame_filtros.place(relx=0.03, rely=0.18, relwidth=0.94, relheight=0.13)
        frame_filtros.grid_columnconfigure(0, weight=0)
        frame_filtros.grid_columnconfigure(1, weight=0)
        frame_filtros.grid_columnconfigure(2, weight=0)
        frame_filtros.grid_columnconfigure(3, weight=1)
        frame_filtros.grid_rowconfigure(0, weight=1)

        departamento_filtro = ctk.StringVar(value="TODOS")
        combo_departamentos = ctk.CTkComboBox(
            frame_filtros,
            values=["TODOS"],
            width=180,
            variable=departamento_filtro,
            state="readonly",
        )
        combo_departamentos.grid(row=0, column=0, sticky="w", padx=(14, 8), pady=10)

        entrada_busqueda = ctk.CTkEntry(
            frame_filtros,
            width=250,
            placeholder_text="Buscar por nombre o DUI",
            justify="left",
            font=ctk.CTkFont(size=16),
        )
        entrada_busqueda.grid(row=0, column=1, sticky="w", padx=(0, 8), pady=10)
        boton_buscar = boton_buscar_profesional(frame_filtros, width=96)
        boton_buscar.grid(row=0, column=2, sticky="w", padx=(0, 14), pady=10)

        scrolframe = ctk.CTkScrollableFrame(self.frame_vista)
        scrolframe.place(relx=0.03, rely=0.35, relwidth=0.94, relheight=0.60)

        columnas_detalle = ["Inicio", "Fin", "H. Trabajadas", "H. Extras", "Fecha", "Justificar", "Acción"]
        servicio_marcaciones = db.marcaciones()
        servicio_cortes = db.cortes_planilla()
        servicio_empleados = db.gestion_empleado()
        usuario_desplegado = {"clave": None}

        padx_val = max(8, min(100, int(screen_w * 0.02)))

        def convertir_a_fecha(fecha_valor):
            if hasattr(fecha_valor, "year") and hasattr(fecha_valor, "month") and hasattr(fecha_valor, "day"):
                return fecha_valor
            texto = str(fecha_valor).strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto, formato).date()
                except ValueError:
                    continue
            return None

        def parsear_hora(valor):
            texto = str(valor or "").strip().upper()
            if not texto:
                return None

            for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                try:
                    return datetime.strptime(texto, formato)
                except ValueError:
                    continue
            return None

        def calcular_horas(entrada_txt, salida_txt):
            inicio = parsear_hora(entrada_txt)
            fin = parsear_hora(salida_txt)
            if not inicio or not fin:
                return 0.0, 0.0

            # Si la salida es menor que la entrada, se asume turno nocturno.
            if fin < inicio:
                fin = fin + timedelta(days=1)

            horas = max(0.0, (fin - inicio).total_seconds() / 3600)
            extras = max(0.0, horas - 8.0)
            return horas, extras

        def calcular_horas_para_pago(horas_trabajadas, horas_extras, estado_revision):
            """Devuelve las horas que realmente se pagan.

            Si las horas extra no están aprobadas, solo se paga la jornada normal
            de 8 horas y se excluyen las extras. Si está aprobada, se pueden pagar
            todas las horas registradas, incluyendo extras.
            """
            horas_trabajadas = max(0.0, float(horas_trabajadas or 0.0))
            horas_extras = max(0.0, float(horas_extras or 0.0))
            estado_revision = str(estado_revision or "PENDIENTE").upper()

            if horas_extras <= 0:
                return min(horas_trabajadas, 8.0), 0.0

            if estado_revision == "APROBADA":
                return horas_trabajadas, horas_extras

            return min(horas_trabajadas, 8.0), 0.0

        def dibujar_fila(parent, columnas, valores, es_encabezado=False, accion=None, meta_justificar=None):
            row_frame = ctk.CTkFrame(parent, fg_color="transparent", height=34)
            row_frame.pack(fill='x', padx=8, pady=(0, 2))
            row_frame.grid_propagate(False)

            total = len(columnas)
            for indice_col in range(total * 2 - 1):
                if indice_col % 2 == 0:
                    row_frame.grid_columnconfigure(indice_col, weight=1, uniform="cols_registro")
                else:
                    row_frame.grid_columnconfigure(indice_col, weight=0)

            color_linea = "red" if es_encabezado else "gray40"
            fuente = ctk.CTkFont(size=14, weight="bold") if es_encabezado else ctk.CTkFont(size=13)

            for indice, valor in enumerate(valores):
                columna = indice * 2
                celda = ctk.CTkFrame(row_frame, fg_color="transparent", height=34)
                celda.grid(row=0, column=columna, sticky="nsew")
                celda.grid_propagate(False)

                if indice == total - 2 and not es_encabezado and meta_justificar is not None:
                    texto_btn = meta_justificar.get("texto", "Revisar")
                    color_btn = meta_justificar.get("color", "#d08a00")
                    estado_btn = meta_justificar.get("estado", "normal")
                    comando_btn = meta_justificar.get("command")
                    boton_justificar = ctk.CTkButton(
                        celda,
                        text=texto_btn,
                        width=104,
                        height=26,
                        fg_color=color_btn,
                        state=estado_btn,
                        command=comando_btn,
                    )
                    boton_justificar.pack(padx=6, pady=4)
                elif accion is not None and indice == total - 1 and not es_encabezado:
                    boton_accion = ctk.CTkButton(celda, text="Editar", width=86, height=26, command=accion)
                    boton_accion.pack(padx=6, pady=4)
                else:
                    label_valor = ctk.CTkLabel(celda, text=str(valor), anchor="w", font=fuente)
                    label_valor.pack(fill='both', padx=(10, 6), pady=4)

                if indice < total - 1:
                    linea = ctk.CTkFrame(row_frame, width=2, fg_color=color_linea)
                    linea.grid(row=0, column=columna + 1, sticky="ns", pady=4)

        def render_lista_usuarios(registros_por_usuario, permitir_edicion_corte, aprobaciones_por_usuario, corte_id, callback_recarga):
            for (dui_usuario, nombre_usuario) in sorted(registros_por_usuario.keys(), key=lambda item: item[1].lower()):
                filas_usuario = registros_por_usuario[(dui_usuario, nombre_usuario)]
                filas_usuario.sort(key=lambda fila: fila["fecha"] if fila["fecha"] else datetime.min.date(), reverse=True)
                departamento_usuario = str(filas_usuario[0].get("departamento") or "SIN DEPARTAMENTO")
                aprobado_usuario, fecha_aprobacion_usuario = aprobaciones_por_usuario.get(str(dui_usuario), (False, None))
                pendientes_revision = sum(
                    1
                    for f in filas_usuario
                    if float(f.get("horas_extras_valor") or 0) > 0 and str(f.get("justificar") or "").upper() == "PENDIENTE"
                )
                tiene_pendientes_revision = any(
                    float(f.get("horas_extras_valor") or 0) > 0 and str(f.get("justificar") or "").upper() == "PENDIENTE"
                    for f in filas_usuario
                )
                permitir_edicion_usuario = permitir_edicion_corte and not aprobado_usuario

                tarjeta_usuario = ctk.CTkFrame(scrolframe, corner_radius=10, border_width=1, border_color="#3f3f3f")
                tarjeta_usuario.pack(fill='x', padx=padx_val, pady=(6, 4))

                encabezado_usuario = ctk.CTkFrame(tarjeta_usuario, fg_color="transparent")
                encabezado_usuario.pack(fill='x', padx=10, pady=(8, 6))

                fila_superior = ctk.CTkFrame(encabezado_usuario, fg_color="transparent")
                fila_superior.pack(fill='x')

                bloque_nombre = ctk.CTkFrame(fila_superior, fg_color="transparent")
                bloque_nombre.pack(side='left', fill='x', expand=True, padx=(0, 8))

                ctk.CTkLabel(
                    bloque_nombre,
                    text=nombre_usuario,
                    font=ctk.CTkFont(size=20, weight="bold"),
                    text_color="#f8fafc",
                    anchor="w",
                ).pack(fill='x')

                ctk.CTkLabel(
                    bloque_nombre,
                    text=f"DUI: {dui_usuario}  |  Departamento: {departamento_usuario}",
                    font=ctk.CTkFont(size=12),
                    text_color="gray75",
                    anchor="w",
                ).pack(fill='x', pady=(2, 0))

                bloque_acciones = ctk.CTkFrame(fila_superior, fg_color="transparent")
                bloque_acciones.pack(side='right', anchor='ne')

                fila_detalle = ctk.CTkFrame(encabezado_usuario, fg_color="transparent")
                fila_detalle.pack(fill='x', pady=(6, 0))

                clave_usuario = f"{dui_usuario}|{nombre_usuario}"

                def alternar_usuario(clave=clave_usuario):
                    if usuario_desplegado["clave"] == clave:
                        usuario_desplegado["clave"] = None
                    else:
                        usuario_desplegado["clave"] = clave
                    callback_recarga()

                texto_boton = "Aprobado" if aprobado_usuario else ("Ocultar" if usuario_desplegado["clave"] == clave_usuario else "Ver")
                ctk.CTkButton(
                    bloque_acciones,
                    text=texto_boton,
                    width=84,
                    height=30,
                    corner_radius=10,
                    font=("Segoe UI", 11, "bold"),
                    text_color="#ffffff",
                    text_color_disabled="#f8fafc",
                    fg_color="#64748b" if aprobado_usuario else "#2563eb",
                    hover_color="#1d4ed8",
                    state="disabled" if aprobado_usuario else "normal",
                    command=alternar_usuario,
                ).pack(side='right', padx=(0, 6))

                def aprobar_usuario_actual(dui=dui_usuario, nombre=nombre_usuario):
                    confirmar = messagebox.askyesno(
                        "Aprobar marcaciones",
                        f"Se aprobaran las marcaciones de {nombre} en este corte y quedaran en solo lectura.\n\n¿Desea continuar?"
                    )
                    if not confirmar:
                        return

                    guardado = servicio_marcaciones.aprobar_usuario_en_corte(corte_id, dui)
                    if guardado:
                        usuario_desplegado["clave"] = None
                        self._mostrar_animacion_aprobacion(sipp, nombre, dui)
                        callback_recarga()
                    else:
                        messagebox.showerror("Error", "No se pudo aprobar las marcaciones de este usuario.")

                texto_aprobar = "Aprobado" if aprobado_usuario else ("Pendiente" if tiene_pendientes_revision else "Aprobar")
                color_aprobar = "#475569" if aprobado_usuario else ("#b45309" if tiene_pendientes_revision else "#15803d")
                hover_aprobar = "#334155" if aprobado_usuario else ("#92400e" if tiene_pendientes_revision else "#166534")

                boton_aprobar_usuario = ctk.CTkButton(
                    bloque_acciones,
                    text=texto_aprobar,
                    width=96,
                    height=30,
                    corner_radius=10,
                    border_width=1,
                    border_color="#e2e8f0",
                    font=("Segoe UI", 11, "bold"),
                    text_color="#ffffff",
                    text_color_disabled="#f8fafc",
                    fg_color=color_aprobar,
                    hover_color=hover_aprobar,
                    state="disabled" if aprobado_usuario or not permitir_edicion_corte or tiene_pendientes_revision else "normal",
                    command=aprobar_usuario_actual,
                )
                boton_aprobar_usuario.pack(side='right', padx=(0, 6))

                ctk.CTkLabel(
                    fila_detalle,
                    text=f"Marcaciones: {len(filas_usuario)}",
                    font=ctk.CTkFont(size=12),
                    text_color="gray72",
                    anchor="w",
                ).pack(side='left', padx=(0, 10))

                ctk.CTkLabel(
                    fila_detalle,
                    text=f"Pendientes: {pendientes_revision}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#d08a00" if pendientes_revision > 0 else "#1A7F37",
                    anchor="w",
                ).pack(side='left', padx=(0, 10))

                estado_aprobacion_texto = "Aprobado" if aprobado_usuario else "Pendiente"
                color_aprobacion = "#1A7F37" if aprobado_usuario else "#d08a00"
                ctk.CTkLabel(
                    fila_detalle,
                    text=f"Estado: {estado_aprobacion_texto}",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=color_aprobacion,
                    anchor="w",
                ).pack(side='left', padx=(0, 10))

                if aprobado_usuario and fecha_aprobacion_usuario:
                    ctk.CTkLabel(
                        fila_detalle,
                        text=f"({fecha_aprobacion_usuario})",
                        font=ctk.CTkFont(size=11),
                        text_color="gray72",
                        anchor="w",
                    ).pack(side='left')

                if usuario_desplegado["clave"] != clave_usuario:
                    continue

                cuerpo_usuario = ctk.CTkFrame(tarjeta_usuario, corner_radius=8)
                cuerpo_usuario.pack(fill='x', padx=8, pady=(0, 8))

                dibujar_fila(cuerpo_usuario, columnas_detalle, columnas_detalle, es_encabezado=True)

                for fila in filas_usuario:
                    estado_revision = str(fila.get("justificar") or "PENDIENTE").upper()
                    horas_extras_valor = float(fila.get("horas_extras_valor") or 0)

                    if horas_extras_valor <= 0:
                        meta_justificar = {
                            "texto": "Sin extras",
                            "color": "#5b5b5b",
                            "estado": "disabled",
                            "command": None,
                        }
                    elif estado_revision == "APROBADA":
                        meta_justificar = {
                            "texto": "Aprobado",
                            "color": "#1A7F37",
                            "estado": "normal" if permitir_edicion_usuario else "disabled",
                            "command": (lambda f=fila: self._abrir_justificacion_registro(
                                sipp,
                                f["nombre"],
                                f["dui"],
                                f["fecha_texto"],
                                f["motivo"],
                                f["justificar"],
                                f["horas_extras_valor"],
                                permitir_edicion_usuario,
                                callback_recarga,
                            )) if permitir_edicion_usuario else None,
                        }
                    elif estado_revision == "NO APROBADA":
                        meta_justificar = {
                            "texto": "No aprobado",
                            "color": "#B42318",
                            "estado": "normal" if permitir_edicion_usuario else "disabled",
                            "command": (lambda f=fila: self._abrir_justificacion_registro(
                                sipp,
                                f["nombre"],
                                f["dui"],
                                f["fecha_texto"],
                                f["motivo"],
                                f["justificar"],
                                f["horas_extras_valor"],
                                permitir_edicion_usuario,
                                callback_recarga,
                            )) if permitir_edicion_usuario else None,
                        }
                    else:
                        meta_justificar = {
                            "texto": "Revisar",
                            "color": "#d08a00",
                            "estado": "normal" if permitir_edicion_usuario else "disabled",
                            "command": (lambda f=fila: self._abrir_justificacion_registro(
                                sipp,
                                f["nombre"],
                                f["dui"],
                                f["fecha_texto"],
                                f["motivo"],
                                f["justificar"],
                                f["horas_extras_valor"],
                                permitir_edicion_usuario,
                                callback_recarga,
                            )) if permitir_edicion_usuario else None,
                        }

                    horas_pagadas, extras_pagadas = calcular_horas_para_pago(
                        fila.get("horas_valor", 0.0),
                        fila.get("horas_extras_valor", 0.0),
                        fila.get("justificar", "PENDIENTE"),
                    )

                    valores_fila = [
                        fila["entrada"],
                        fila["salida"],
                        f"{horas_pagadas:.2f} h",
                        f"{extras_pagadas:.2f} h",
                        fila["fecha_texto"],
                        fila["justificar"],
                        "Editar" if permitir_edicion_usuario else "Bloqueado",
                    ]
                    dibujar_fila(
                        cuerpo_usuario,
                        columnas_detalle,
                        valores_fila,
                        es_encabezado=False,
                        accion=lambda f=fila: self._editar_marcacion_desde_registro(
                            sipp,
                            f["nombre"],
                            f["dui"],
                            f["fecha_texto"],
                            f["entrada"],
                            f["salida"],
                            permitir_edicion_usuario,
                            callback_recarga,
                        ) if permitir_edicion_usuario else None,
                        meta_justificar=meta_justificar,
                    )

        def obtener_mapa_departamentos():
            try:
                empleados_db = servicio_empleados.obtener_empleados()
            except Exception:
                return {}, ["TODOS"]

            dui_a_departamento = {}
            departamentos = set()
            for empleado in empleados_db:
                if len(empleado) < 8:
                    continue

                dui_emp = str(empleado[2] or "").strip()
                departamento_emp = str(empleado[7] or "SIN DEPARTAMENTO").strip().upper()
                if not departamento_emp:
                    departamento_emp = "SIN DEPARTAMENTO"

                if dui_emp:
                    dui_a_departamento[dui_emp] = departamento_emp
                departamentos.add(departamento_emp)

            valores_combo = ["TODOS"] + sorted(departamentos)
            return dui_a_departamento, valores_combo

        def actualizar_combo_departamentos(valores_combo):
            combo_departamentos.configure(values=valores_combo)
            valor_actual = departamento_filtro.get().strip().upper() or "TODOS"
            if valor_actual not in valores_combo:
                departamento_filtro.set("TODOS")

        def cargar_registros():
            for widget in scrolframe.winfo_children():
                widget.destroy()

            criterio = entrada_busqueda.get().strip().lower()
            departamento_seleccionado = departamento_filtro.get().strip().upper() or "TODOS"

            dui_a_departamento, valores_combo = obtener_mapa_departamentos()
            actualizar_combo_departamentos(valores_combo)
            departamento_seleccionado = departamento_filtro.get().strip().upper() or "TODOS"

            corte_activo = servicio_cortes.obtener_corte_activo()
            if not corte_activo:
                etiqueta_corte_activo.configure(text="No hay corte activo", text_color="#d08a00")
                ctk.CTkLabel(
                    scrolframe,
                    text="No hay corte activo. Active un corte para ver marcaciones.",
                    anchor="w",
                ).pack(fill='x', padx=padx_val, pady=10)
                return

            corte_id, fecha_inicio_corte, fecha_fin_corte, estado_corte = corte_activo
            estado_corte_texto = str(estado_corte or "").strip().upper()
            permitir_edicion = estado_corte_texto == "ACTIVO"
            fecha_inicio_corte = convertir_a_fecha(fecha_inicio_corte)
            fecha_fin_corte = convertir_a_fecha(fecha_fin_corte)
            if not fecha_inicio_corte or not fecha_fin_corte:
                etiqueta_corte_activo.configure(text="Corte activo con fechas invalidas", text_color="#c43c3c")
                ctk.CTkLabel(scrolframe, text="No se pudo interpretar el rango del corte activo.", anchor="w").pack(fill='x', padx=padx_val, pady=10)
                return

            etiqueta_corte_activo.configure(
                text=f"Corte {corte_id}: {fecha_inicio_corte.strftime('%d/%m/%Y')} a {fecha_fin_corte.strftime('%d/%m/%Y')} ({estado_corte})",
                text_color="#2aa745" if permitir_edicion else "#d08a00",
            )

            if not permitir_edicion:
                ctk.CTkLabel(
                    scrolframe,
                    text="Corte en solo lectura: sin edicion ni aprobaciones.",
                    anchor="w",
                    text_color="#d08a00",
                ).pack(fill='x', padx=padx_val, pady=(0, 8))

            try:
                marcaciones_db = servicio_marcaciones.obtener_marcaciones_por_rango(fecha_inicio_corte, fecha_fin_corte)
                revisiones_rango = servicio_marcaciones.obtener_revisiones_marcaciones_rango(fecha_inicio_corte, fecha_fin_corte)
            except Exception as exc:
                messagebox.showerror("Error", f"No se pudieron cargar las marcaciones: {exc}")
                return

            registros_filtrados = []
            for registro in marcaciones_db:
                if len(registro) < 6:
                    continue

                _, nombre, dui, hora_entrada, hora_salida, fecha_registro = registro[:6]
                nombre_txt = str(nombre or "")
                dui_txt = str(dui or "")
                departamento_txt = dui_a_departamento.get(dui_txt, "SIN DEPARTAMENTO")

                if departamento_seleccionado != "TODOS" and departamento_txt != departamento_seleccionado:
                    continue

                if criterio and criterio not in nombre_txt.lower() and criterio not in dui_txt.lower():
                    continue

                horas_trabajadas, horas_extras = calcular_horas(hora_entrada, hora_salida)
                fecha_valor = convertir_a_fecha(fecha_registro)
                fecha_texto = fecha_valor.strftime("%d/%m/%Y") if fecha_valor else str(fecha_registro or "")
                estado_aprobacion, motivo = revisiones_rango.get((dui_txt, fecha_valor), ("PENDIENTE", ""))
                if horas_extras <= 0:
                    estado_aprobacion = "SIN EXTRAS"

                horas_pagadas, extras_pagadas = calcular_horas_para_pago(horas_trabajadas, horas_extras, estado_aprobacion)

                registros_filtrados.append({
                    "nombre": nombre_txt,
                    "dui": dui_txt,
                    "entrada": str(hora_entrada or ""),
                    "salida": str(hora_salida or ""),
                    "horas": f"{horas_pagadas:.2f} h",
                    "extras": f"{extras_pagadas:.2f} h",
                    "horas_valor": horas_pagadas,
                    "horas_extras_valor": extras_pagadas,
                    "justificar": estado_aprobacion,
                    "fecha": fecha_valor,
                    "fecha_texto": fecha_texto,
                    "motivo": str(motivo or "").strip(),
                    "departamento": departamento_txt,
                })

            registros_filtrados.sort(
                key=lambda fila: (
                    fila["nombre"],
                    fila["dui"],
                    fila["fecha"] if fila["fecha"] else datetime.min.date(),
                ),
            )

            if not registros_filtrados:
                mensaje = "No hay marcaciones en el corte activo." if not criterio else "No se encontraron registros en el corte activo para la busqueda."
                ctk.CTkLabel(scrolframe, text=mensaje, anchor="w").pack(fill='x', padx=padx_val, pady=10)
                return

            registros_por_usuario = {}
            for fila in registros_filtrados:
                clave = (fila["dui"], fila["nombre"])
                registros_por_usuario.setdefault(clave, []).append(fila)

            aprobaciones_por_usuario = servicio_marcaciones.obtener_aprobaciones_por_corte(corte_id)
            render_lista_usuarios(registros_por_usuario, permitir_edicion, aprobaciones_por_usuario, corte_id, cargar_registros)

        boton_buscar.configure(command=cargar_registros)
        entrada_busqueda.bind("<Return>", lambda event: cargar_registros())
        combo_departamentos.configure(command=lambda valor: cargar_registros())
        departamento_filtro.set("TODOS")
        entrada_busqueda.focus_set()
        cargar_registros()

    def _abrir_edicion_registro(self, sipp, dui):
        self.añadir_marcacion_manual(sipp)
        self.entrada_busqueda_personal.delete(0, "end")
        self.entrada_busqueda_personal.insert(0, str(dui or ""))
        self.añadir_manual(self.frame_vista)

    def _editar_marcacion_desde_registro(self, sipp, nombre, dui, fecha_texto, entrada_actual, salida_actual, permitir_edicion=True, callback_recarga=None):
        if not permitir_edicion:
            messagebox.showinfo("Solo lectura", "Este corte no esta activo, no se puede editar.")
            return

        top_level = ctk.CTkToplevel(sipp)
        ancho_modal = 620
        alto_modal = 360
        x_modal = (top_level.winfo_screenwidth() - ancho_modal) // 2
        y_modal = (top_level.winfo_screenheight() - alto_modal) // 2
        top_level.geometry(f"{ancho_modal}x{alto_modal}+{x_modal}+{y_modal}")
        top_level.title("Editar Marcacion del Dia")
        top_level.resizable(False, False)
        top_level.grab_set()

        contenedor = ctk.CTkFrame(top_level, corner_radius=14, border_width=1, border_color="#3f3f3f")
        contenedor.place(relx=0.03, rely=0.05, relwidth=0.94, relheight=0.90)

        ctk.CTkLabel(
            contenedor,
            text="Editar Marcacion",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).place(relx=0.03, rely=0.04)

        ctk.CTkLabel(
            contenedor,
            text=f"{nombre} | DUI: {dui}",
            font=ctk.CTkFont(size=14),
            text_color="gray70",
            anchor="w",
        ).place(relx=0.03, rely=0.15)

        ctk.CTkLabel(
            contenedor,
            text=f"Fecha del registro: {fecha_texto}",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w",
        ).place(relx=0.03, rely=0.23)

        tarjeta_horas = ctk.CTkFrame(contenedor, corner_radius=10)
        tarjeta_horas.place(relx=0.03, rely=0.34, relwidth=0.64, relheight=0.40)

        def descomponer_hora(valor):
            texto = str(valor or "").strip().upper()
            for fmt in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
                try:
                    marca = datetime.strptime(texto, fmt)
                    return f"{marca.strftime('%I')}", f"{marca.strftime('%M')}", marca.strftime('%p')
                except ValueError:
                    continue
            return "00", "00", "AM"

        def validar_hora(text):
            if text == "":
                return True
            return text.isdigit() and 0 <= int(text) <= 12 and len(text) <= 2

        def validar_minuto(text):
            if text == "":
                return True
            return text.isdigit() and 0 <= int(text) <= 59 and len(text) <= 2

        e_h, e_m, e_z = descomponer_hora(entrada_actual)
        s_h, s_m, s_z = descomponer_hora(salida_actual)

        ctk.CTkLabel(tarjeta_horas, text="Hora de entrada", font=ctk.CTkFont(size=15, weight="bold")).place(relx=0.04, rely=0.18)
        fila_entrada = ctk.CTkFrame(tarjeta_horas, fg_color="transparent")
        fila_entrada.place(relx=0.34, rely=0.15)
        entrada_hora = ctk.CTkEntry(fila_entrada, width=45, justify="center", validate="key", validatecommand=(top_level.register(validar_hora), "%P"))
        entrada_hora.pack(side='left')
        ctk.CTkLabel(fila_entrada, text=":", width=14, anchor="center").pack(side='left', padx=3)
        entrada_minuto = ctk.CTkEntry(fila_entrada, width=45, justify="center", validate="key", validatecommand=(top_level.register(validar_minuto), "%P"))
        entrada_minuto.pack(side='left')
        entrada_zona = ctk.CTkComboBox(fila_entrada, values=["AM", "PM"], width=90, state="readonly")
        entrada_zona.pack(side='left', padx=(8, 0))

        ctk.CTkLabel(tarjeta_horas, text="Hora de salida", font=ctk.CTkFont(size=15, weight="bold")).place(relx=0.04, rely=0.58)
        fila_salida = ctk.CTkFrame(tarjeta_horas, fg_color="transparent")
        fila_salida.place(relx=0.34, rely=0.55)
        salida_hora = ctk.CTkEntry(fila_salida, width=45, justify="center", validate="key", validatecommand=(top_level.register(validar_hora), "%P"))
        salida_hora.pack(side='left')
        ctk.CTkLabel(fila_salida, text=":", width=14, anchor="center").pack(side='left', padx=3)
        salida_minuto = ctk.CTkEntry(fila_salida, width=45, justify="center", validate="key", validatecommand=(top_level.register(validar_minuto), "%P"))
        salida_minuto.pack(side='left')
        salida_zona = ctk.CTkComboBox(fila_salida, values=["AM", "PM"], width=90, state="readonly")
        salida_zona.pack(side='left', padx=(8, 0))

        entrada_hora.insert(0, e_h)
        entrada_minuto.insert(0, e_m)
        entrada_zona.set(e_z)
        salida_hora.insert(0, s_h)
        salida_minuto.insert(0, s_m)
        salida_zona.set(s_z)

        def guardar_edicion():
            try:
                fecha_db = datetime.strptime(str(fecha_texto), "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                fecha_db = str(fecha_texto)

            h_entrada = str(entrada_hora.get() or "0").zfill(2)
            m_entrada = str(entrada_minuto.get() or "0").zfill(2)
            h_salida = str(salida_hora.get() or "0").zfill(2)
            m_salida = str(salida_minuto.get() or "0").zfill(2)

            try:
                int_h_entrada = int(h_entrada)
                int_m_entrada = int(m_entrada)
                int_h_salida = int(h_salida)
                int_m_salida = int(m_salida)
            except ValueError:
                messagebox.showwarning("Datos invalidos", "Debe ingresar horas y minutos validos.")
                return

            if int_h_entrada < 0 or int_h_entrada > 12 or int_h_salida < 0 or int_h_salida > 12:
                messagebox.showwarning("Datos invalidos", "La hora debe estar entre 0 y 12.")
                return

            if int_m_entrada < 0 or int_m_entrada > 59 or int_m_salida < 0 or int_m_salida > 59:
                messagebox.showwarning("Datos invalidos", "El minuto debe estar entre 0 y 59.")
                return

            hora_entrada_txt = f"{h_entrada}:{m_entrada} {entrada_zona.get().upper()}"
            hora_salida_txt = f"{h_salida}:{m_salida} {salida_zona.get().upper()}"

            try:
                actualizado = db.marcaciones().actualizar_marcacion(
                    str(dui or "").strip(),
                    fecha_db,
                    hora_entrada_txt,
                    hora_salida_txt,
                    str(nombre or "").strip(),
                )
            except Exception as exc:
                messagebox.showerror("Error", f"No se pudo actualizar la marcacion: {exc}")
                return

            if actualizado:
                messagebox.showinfo("Actualizado", "Marcacion actualizada correctamente.")
                top_level.destroy()
                if callback_recarga:
                    callback_recarga()
            else:
                messagebox.showwarning("Sin cambios", "No se encontro una marcacion para actualizar con esos datos.")

        panel_info = ctk.CTkFrame(contenedor, corner_radius=10)
        panel_info.place(relx=0.70, rely=0.34, relwidth=0.27, relheight=0.40)
        ctk.CTkLabel(panel_info, text="Consejo", font=ctk.CTkFont(size=15, weight="bold")).place(relx=0.08, rely=0.10)
        ctk.CTkLabel(
            panel_info,
            text="Use formato 00:00 y seleccione AM/PM para evitar errores de captura.",
            wraplength=130,
            justify="left",
            text_color="gray75",
            font=ctk.CTkFont(size=12),
        ).place(relx=0.08, rely=0.28)

        barra_botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        barra_botones.place(relx=0.03, rely=0.79, relwidth=0.94, relheight=0.16)
        ctk.CTkButton(barra_botones, text="Cancelar", width=140, height=36, fg_color="#7a2830", command=top_level.destroy).pack(side='right', padx=(8, 0), pady=10)
        ctk.CTkButton(barra_botones, text="Guardar Cambios", width=160, height=36, fg_color="#0A9D58", command=guardar_edicion).pack(side='right', padx=(0, 8), pady=10)
        entrada_hora.focus_set()

    def _mostrar_animacion_aprobacion(self, sipp, nombre, dui):
        top_level = ctk.CTkToplevel(sipp)
        ancho_modal = 420
        alto_modal = 280
        x_modal = (top_level.winfo_screenwidth() - ancho_modal) // 2
        y_modal = (top_level.winfo_screenheight() - alto_modal) // 2
        top_level.geometry(f"{ancho_modal}x{alto_modal}+{x_modal}+{y_modal}")
        top_level.title("Aprobacion Exitosa")
        top_level.resizable(False, False)
        top_level.grab_set()

        panel = ctk.CTkFrame(top_level, corner_radius=14, border_width=1, border_color="#2b8a3e")
        panel.place(relx=0.04, rely=0.06, relwidth=0.92, relheight=0.88)

        ctk.CTkLabel(
            panel,
            text="Marcaciones Aprobadas",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="#1A7F37",
        ).place(relx=0.20, rely=0.08)

        ctk.CTkLabel(
            panel,
            text=f"{nombre} | DUI: {dui}",
            font=ctk.CTkFont(size=13),
            text_color="gray70",
        ).place(relx=0.17, rely=0.22)

        lienzo = tk.Canvas(panel, width=120, height=120, bg="#2b2b2b", highlightthickness=0)
        lienzo.place(relx=0.37, rely=0.34)
        circulo = lienzo.create_oval(10, 10, 110, 110, outline="#1A7F37", width=5)

        puntos = [(36, 62), (54, 80), (86, 44)]
        linea_check = None

        def animar_trazo(segmento=0, paso=0):
            nonlocal linea_check
            if segmento >= 2:
                return

            x1, y1 = puntos[segmento]
            x2, y2 = puntos[segmento + 1]
            total = 12
            t = paso / total
            xa = x1 + (x2 - x1) * t
            ya = y1 + (y2 - y1) * t

            if linea_check is not None:
                lienzo.delete(linea_check)

            if segmento == 0:
                linea_check = lienzo.create_line(puntos[0][0], puntos[0][1], xa, ya, fill="#19c37d", width=6, capstyle=tk.ROUND, joinstyle=tk.ROUND)
            else:
                linea_check = lienzo.create_line(puntos[0][0], puntos[0][1], puntos[1][0], puntos[1][1], xa, ya, fill="#19c37d", width=6, capstyle=tk.ROUND, joinstyle=tk.ROUND)

            if paso < total:
                top_level.after(18, lambda: animar_trazo(segmento, paso + 1))
            else:
                top_level.after(40, lambda: animar_trazo(segmento + 1, 0))

        animar_trazo()

        ctk.CTkButton(panel, text="Aceptar", width=120, fg_color="#1A7F37", command=top_level.destroy).place(relx=0.36, rely=0.86)

    def _abrir_justificacion_registro(self, sipp, nombre, dui, fecha_texto, motivo_actual, estado_actual, horas_extras_valor, permitir_edicion=True, callback_recarga=None):
        if not permitir_edicion:
            messagebox.showinfo("Solo lectura", "Este corte no esta activo, no se puede aprobar o rechazar horas extras.")
            return

        if float(horas_extras_valor or 0) <= 0:
            messagebox.showinfo("Sin horas extras", "Este registro no tiene horas extras para aprobar o no aprobar.")
            return

        top_level = ctk.CTkToplevel(sipp)
        ancho_modal = 620
        alto_modal = 430
        x_modal = (top_level.winfo_screenwidth() - ancho_modal) // 2
        y_modal = (top_level.winfo_screenheight() - alto_modal) // 2
        top_level.geometry(f"{ancho_modal}x{alto_modal}+{x_modal}+{y_modal}")
        top_level.title("Revision de Horas Extras")
        top_level.resizable(False, False)
        top_level.grab_set()

        contenedor = ctk.CTkFrame(top_level, corner_radius=14, border_width=1, border_color="#3f3f3f")
        contenedor.place(relx=0.03, rely=0.05, relwidth=0.94, relheight=0.90)

        ctk.CTkLabel(
            contenedor,
            text="Revision de Horas Extras",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).place(relx=0.03, rely=0.04)

        ctk.CTkLabel(
            contenedor,
            text=f"{nombre} | DUI: {dui}",
            font=ctk.CTkFont(size=14),
            text_color="gray70",
            anchor="w",
        ).place(relx=0.03, rely=0.15)

        ctk.CTkLabel(
            contenedor,
            text=f"Fecha: {fecha_texto}",
            font=ctk.CTkFont(size=14),
            anchor="w",
        ).place(relx=0.03, rely=0.24)

        ctk.CTkLabel(
            contenedor,
            text=f"Horas extras registradas: {float(horas_extras_valor):.2f} h",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#2aa745",
            anchor="w",
        ).place(relx=0.03, rely=0.31)

        tarjeta_revision = ctk.CTkFrame(contenedor, corner_radius=10)
        tarjeta_revision.place(relx=0.03, rely=0.40, relwidth=0.94, relheight=0.42)

        ctk.CTkLabel(
            tarjeta_revision,
            text="Decision de revision",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).place(relx=0.03, rely=0.08)

        estado_horas = ctk.StringVar(value=str(estado_actual or "PENDIENTE").upper())
        if estado_horas.get() not in ("PENDIENTE", "APROBADA", "NO APROBADA"):
            estado_horas.set("PENDIENTE")

        selector_estado = ctk.CTkSegmentedButton(
            tarjeta_revision,
            values=["PENDIENTE", "APROBADA", "NO APROBADA"],
            variable=estado_horas,
            width=380,
            height=34,
        )
        selector_estado.place(relx=0.03, rely=0.25)
        selector_estado.set(estado_horas.get())

        ctk.CTkLabel(
            tarjeta_revision,
            text="Motivo (opcional):",
            font=ctk.CTkFont(size=14),
        ).place(relx=0.03, rely=0.50)

        entrada_motivo = ctk.CTkTextbox(tarjeta_revision, width=540, height=90)
        entrada_motivo.place(relx=0.03, rely=0.63)
        if motivo_actual:
            entrada_motivo.insert("1.0", motivo_actual)

        def guardar_justificacion():
            motivo = entrada_motivo.get("1.0", "end").strip()
            estado = estado_horas.get().strip().upper() or "PENDIENTE"

            try:
                fecha_db = datetime.strptime(str(fecha_texto), "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                fecha_db = str(fecha_texto)

            try:
                guardado = db.marcaciones().guardar_justificacion_marcacion(
                    str(dui or "").strip(),
                    fecha_db,
                    motivo,
                    estado,
                )
            except Exception as exc:
                messagebox.showerror("Error", f"No se pudo guardar la aprobacion: {exc}")
                return

            if guardado:
                messagebox.showinfo("Guardado", "Decision guardada correctamente.")
                top_level.destroy()
                if callback_recarga:
                    callback_recarga()
            else:
                messagebox.showwarning("Aviso", "No se pudo guardar la decision.")

        barra_botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        barra_botones.place(relx=0.03, rely=0.84, relwidth=0.94, relheight=0.13)
        boton_cancelar = boton_buscar_profesional(
            barra_botones,
            text="Cancelar",
            command=top_level.destroy,
            width=140,
            height=34,
            fg_color="#4b5563",
            hover_color="#374151",
        )
        boton_cancelar.pack(side='right', padx=(10, 0), pady=8)
        boton_guardar = boton_buscar_profesional(
            barra_botones,
            text="Guardar decisión",
            command=guardar_justificacion,
            width=170,
            height=34,
            fg_color="#0F766E",
            hover_color="#115E59",
        )
        boton_guardar.pack(side='right', padx=(0, 8), pady=8)
        entrada_motivo.focus_set()

        

        

        

    def buscar_cortes_fechas(self, sipp):
        self._crear_frame_vista(sipp)
        self.frame_vista.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=1.0)

        cabecera = ctk.CTkFrame(self.frame_vista, fg_color="transparent")
        cabecera.place(relx=0.00, rely=0.00, relwidth=1.0, relheight=0.10)

        ctk.CTkLabel(
            cabecera,
            text="Cortes de Fechas",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).place(relx=0.00, rely=0.10)

        ctk.CTkLabel(
            cabecera,
            text="Consulta y revisa cortes guardados por rango",
            font=ctk.CTkFont(size=12),
            text_color="gray75",
        ).place(relx=0.00, rely=0.62)

        barra_busqueda = ctk.CTkFrame(self.frame_vista, corner_radius=14, border_width=1, border_color="#4a4a4a")
        barra_busqueda.place(relx=0.00, rely=0.11, relwidth=1.0, relheight=0.10)
        barra_busqueda.grid_columnconfigure(0, weight=0)
        barra_busqueda.grid_columnconfigure(1, weight=0)

        servicio_cortes = db.cortes_planilla()
        cortes_guardados = servicio_cortes.obtener_cortes_planilla()
        cortes_mapa = {}
        fechas = ["Todos"]

        def convertir_fecha_corte(valor):
            if hasattr(valor, "strftime"):
                return valor.strftime("%d/%m/%Y")
            texto = str(valor or "").strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto, formato).strftime("%d/%m/%Y")
                except ValueError:
                    continue
            return texto

        if cortes_guardados:
            cortes_ordenados = sorted(cortes_guardados, key=lambda corte: corte[0], reverse=True)
            for corte in cortes_ordenados:
                corte_id, fecha_inicio, fecha_fin, estado, *_ = corte
                texto_corte = f"Corte {corte_id}: {convertir_fecha_corte(fecha_inicio)} a {convertir_fecha_corte(fecha_fin)} ({estado})"
                fechas.append(texto_corte)
                cortes_mapa[texto_corte] = corte

        entrada_busqueda = ctk.CTkComboBox(
            barra_busqueda,
            width=340,
            values=fechas,
            font=ctk.CTkFont(size=14),
        )
        entrada_busqueda.grid(row=0, column=0, padx=(16, 10), pady=14, sticky="w")
        entrada_busqueda.set("Todos")

        boton_buscar = boton_buscar_profesional(barra_busqueda, width=96)
        boton_buscar.grid(row=0, column=1, padx=(0, 16), pady=14, sticky="w")

        scrolframe = ctk.CTkScrollableFrame(self.frame_vista)
        scrolframe.place(relx=0.00, rely=0.23, relwidth=1.0, relheight=0.69)

        datos = ["Fecha", "Periodo", "Estado", "Archivo"]

        screen_w = sipp.winfo_screenwidth()
        padx_val = max(8, min(100, int(screen_w * 0.02)))

        def dibujar_encabezado(parent):
            fila = ctk.CTkFrame(parent, fg_color="transparent")
            fila.pack(fill='x', padx=padx_val, pady=(0, 4))
            for indice, titulo in enumerate(datos):
                fila.grid_columnconfigure(indice, weight=1, uniform="cortes")
                label = ctk.CTkLabel(fila, text=titulo, font=ctk.CTkFont(size=13, weight="bold"), text_color="gray80", anchor="w")
                label.grid(row=0, column=indice, sticky="ew", padx=(10, 6), pady=4)

        def renderizar_cortes(filtro_texto=""):
            for widget in scrolframe.winfo_children():
                widget.destroy()

            dibujar_encabezado(scrolframe)

            filtro = str(filtro_texto or "").strip().lower()
            coincidencias = []
            for texto_corte, corte in cortes_mapa.items():
                if filtro in ("", "todos") or filtro in texto_corte.lower():
                    coincidencias.append((texto_corte, corte))

            if not coincidencias:
                mensaje = "No hay cortes guardados." if not cortes_guardados else "No se encontraron cortes con ese criterio."
                ctk.CTkLabel(scrolframe, text=mensaje, anchor="w", text_color="gray75").pack(fill='x', padx=padx_val, pady=10)
                return

            for texto_corte, corte in coincidencias:
                corte_id, fecha_inicio, fecha_fin, estado, *_ = corte
                fila = ctk.CTkFrame(scrolframe, corner_radius=10, border_width=1, border_color="#3f3f3f")
                fila.pack(fill='x', padx=padx_val, pady=(0, 8))

                cont = ctk.CTkFrame(fila, fg_color="transparent")
                cont.pack(fill='x', padx=10, pady=8)
                for indice in range(4):
                    cont.grid_columnconfigure(indice, weight=1, uniform="cortes")

                valores = [
                    convertir_fecha_corte(fecha_inicio),
                    f"{convertir_fecha_corte(fecha_inicio)} a {convertir_fecha_corte(fecha_fin)}",
                    estado,
                    f"Corte {corte_id}",
                ]
                colores = ["#e2e8f0", "#e2e8f0", "#d08a00" if str(estado).upper() != "ACTIVO" else "#1A7F37", "#cbd5e1"]

                for indice, valor in enumerate(valores):
                    label = ctk.CTkLabel(
                        cont,
                        text=valor,
                        font=ctk.CTkFont(size=13, weight="bold" if indice == 2 else "normal"),
                        text_color=colores[indice],
                        anchor="w",
                    )
                    label.grid(row=0, column=indice, sticky="ew", padx=(10, 6))

                boton_descargar = boton_buscar_profesional(
                    fila,
                    text="Descargar Excel",
                    command=lambda c=corte: descarga.descargar_corte_planilla_excel(c),
                    width=150,
                    height=30,
                    fg_color="#0F766E" if str(estado).upper() != "ACTIVO" else "#4b5563",
                    hover_color="#115E59" if str(estado).upper() != "ACTIVO" else "#374151",
                )
                if str(estado).upper() == "ACTIVO":
                    boton_descargar.configure(state="disabled")
                boton_descargar.pack(anchor="e", padx=10, pady=(0, 10))

        def ejecutar_busqueda():
            renderizar_cortes(entrada_busqueda.get())

        boton_buscar.configure(command=ejecutar_busqueda)
        entrada_busqueda.bind("<Return>", lambda event: ejecutar_busqueda())
        renderizar_cortes(entrada_busqueda.get())


# ==================== DESCUENTOS COMPARTIDOS ====================
# Lista centralizada de descuentos - sincronizada entre Descuentos y Deducciones
DESCUENTOS_SISTEMA = [
    "Préstamo personal",
    "Préstamo hipotecario",
    "Préstamo de auto",
    "Descuento por adelanto de salario",
    "Descuento por ausencia (faltas)",
    "Descuento por tardanzas",
    "Descuento por suspensión",
    "Descuento por incapacidad (si aplica)",
    "Descuento por pensión alimenticia",
    "Descuento por embargo judicial",
    "Descuento por cooperativa",
    "Descuento por sindicato",
    "Descuento por ahorro obligatorio",
    "Descuento por seguro privado",
    "Descuento por uniformes",
    "Descuento por herramientas o equipo",
    "Descuento por cafetería/comida",
    "Descuento por daños o pérdidas",
    "Descuento por servicios (transporte)",
    "Descuento por compras a crédito",
    "Otros descuentos autorizados"
]
# ================================================================


class Gestion_Descuentos:
    """Gestión completa de descuentos y préstamos con todas las funcionalidades profesionales"""

    def __init__(self):
        self.servicio_descuentos = None
        self.scrollframe = None
        self.sipp_ref = None
        self.filtros_activos = {
            "estado": "ACTIVO",
            "tipo": "",
            "empleado": "",
            "departamento": "Todos"
        }

    def lista_descuentos(self):
        """Tipos de descuentos disponibles en el sistema"""
        return DESCUENTOS_SISTEMA.copy()

    def ventana_gestion_prestamo(self, frame, sipp):
        """Interfaz principal de gestión de descuentos con tabla, búsqueda y acciones"""
        self.sipp_ref = sipp
        
        for widget in frame.winfo_children():
            widget.destroy()

        frame.configure(fg_color="transparent")

        # ===== HEADER =====
        header = ctk.CTkFrame(frame, corner_radius=20, fg_color=("#f8fafc", "#0f172a"), 
                             border_width=1, border_color=("#cbd5e1", "#334155"))
        header.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.18)

        ctk.CTkLabel(header, text="💰 Gestión de Descuentos y Préstamos",
                    font=ctk.CTkFont(size=28, weight="bold"),
                    text_color=("#0f172a", "#e2e8f0")).place(relx=0.04, rely=0.15)

        ctk.CTkLabel(header, text="Administre descuentos, préstamos y deducciones del personal",
                    font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8")).place(relx=0.04, rely=0.60)

        # ===== MÉTRICAS =====
        metrics = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#ffffff", "#111827"),
                              border_width=1, border_color=("#e2e8f0", "#334155"))
        metrics.place(relx=0.02, rely=0.22, relwidth=0.96, relheight=0.085)

        # Métrica 1: Total registrados
        ctk.CTkLabel(metrics, text="Activos", font=ctk.CTkFont(size=12, weight="bold")).place(relx=0.05, rely=0.20)
        self.label_total_registrados = ctk.CTkLabel(metrics, text="0", font=ctk.CTkFont(size=16, weight="bold"),
                                                     text_color="#2563eb")
        self.label_total_registrados.place(relx=0.05, rely=0.50)

        # Métrica 2: Pagados
        ctk.CTkLabel(metrics, text="Pagados", font=ctk.CTkFont(size=12, weight="bold")).place(relx=0.25, rely=0.20)
        self.label_pagados = ctk.CTkLabel(metrics, text="0", font=ctk.CTkFont(size=16, weight="bold"),
                                         text_color="#10b981")
        self.label_pagados.place(relx=0.25, rely=0.50)

        # Métrica 3: Anulados
        ctk.CTkLabel(metrics, text="Anulados", font=ctk.CTkFont(size=12, weight="bold")).place(relx=0.45, rely=0.20)
        self.label_anulados = ctk.CTkLabel(metrics, text="0", font=ctk.CTkFont(size=16, weight="bold"),
                                          text_color="#ef4444")
        self.label_anulados.place(relx=0.45, rely=0.50)

        # Métrica 4: Monto total
        ctk.CTkLabel(metrics, text="Monto Total", font=ctk.CTkFont(size=12, weight="bold")).place(relx=0.65, rely=0.20)
        self.label_total_monto = ctk.CTkLabel(metrics, text="B/. 0.00", font=ctk.CTkFont(size=16, weight="bold"),
                                             text_color="#f59e0b")
        self.label_total_monto.place(relx=0.65, rely=0.50)

        # ===== BÚSQUEDA Y FILTROS =====
        search_bar = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#f8fafc", "#111827"),
                                 border_width=1, border_color=("#e2e8f0", "#334155"))
        search_bar.place(relx=0.02, rely=0.34, relwidth=0.96, relheight=0.10)
        search_bar.pack_propagate(False)

        for i in range(8):
            search_bar.grid_columnconfigure(i, weight=1)

        self.entrada_buscar_prestamo = ctk.CTkEntry(search_bar, placeholder_text="Buscar por nombre o DUI",
                                                     font=ctk.CTkFont(size=11), justify="center")
        self.entrada_buscar_prestamo.grid(row=0, column=0, padx=(14, 8), pady=10, sticky="ew")

        btn_buscar = ctk.CTkButton(search_bar, text="Buscar", width=80, font=ctk.CTkFont(size=9, weight="bold"),
                                   fg_color="#2563eb", hover_color="#1d4ed8",
                                   command=self._ejecutar_busqueda_avanzada)
        btn_buscar.grid(row=0, column=1, padx=(0, 8), pady=10, sticky="ew")

        label_departamento = ctk.CTkLabel(search_bar, text="Departamento:", font=ctk.CTkFont(size=9, weight="bold"))
        label_departamento.grid(row=0, column=2, padx=(0, 4), pady=0, sticky="e")
        self.combo_departamento_filtro = ctk.CTkComboBox(search_bar, values=["Todos"],
                                                          font=ctk.CTkFont(size=9), width=150,
                                                          command=lambda _: self._ejecutar_busqueda_avanzada())
        self.combo_departamento_filtro.set("Todos")
        self.combo_departamento_filtro.grid(row=0, column=3, padx=(0, 8), pady=10, sticky="ew")
        self._actualizar_combo_departamentos()

        label_estado = ctk.CTkLabel(search_bar, text="Estado:", font=ctk.CTkFont(size=9, weight="bold"))
        label_estado.grid(row=0, column=4, padx=(0, 4), pady=0, sticky="e")
        self.combo_estado_filtro = ctk.CTkComboBox(search_bar, values=["ACTIVO", "PAGADO", "ANULADO", "TODOS"],
                                                    font=ctk.CTkFont(size=9), width=100,
                                                    command=lambda _: self._ejecutar_busqueda_avanzada())
        self.combo_estado_filtro.set("ACTIVO")
        self.combo_estado_filtro.grid(row=0, column=5, padx=(0, 8), pady=10, sticky="ew")

        btn_nuevo = ctk.CTkButton(search_bar, text="Nuevo", width=80, font=ctk.CTkFont(size=9, weight="bold"),
                                  fg_color="#10b981", hover_color="#059669",
                                  command=lambda: self.formulario_descuento(sipp))
        btn_nuevo.grid(row=0, column=6, padx=(6, 6), pady=10, sticky="ew")

        btn_reportes = ctk.CTkButton(search_bar, text="Reportes", width=80, font=ctk.CTkFont(size=9, weight="bold"),
                                     fg_color="#f59e0b", hover_color="#d97706",
                                     command=self._mostrar_reportes)
        btn_reportes.grid(row=0, column=7, padx=(0, 10), pady=10, sticky="ew")

        self.entrada_buscar_prestamo.bind("<KeyRelease>", lambda event: self._ejecutar_busqueda_avanzada())

        # ===== TABLA =====
        tabla_frame = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#ffffff", "#111827"),
                                  border_width=1, border_color=("#e2e8f0", "#334155"))
        tabla_frame.place(relx=0.02, rely=0.47, relwidth=0.96, relheight=0.51)

        # ScrollableFrame para la tabla
        self.scrollframe = ctk.CTkScrollableFrame(tabla_frame, corner_radius=12, fg_color="transparent")
        self.scrollframe.pack(fill="both", expand=True, padx=10, pady=10)

        # Encabezados
        encabezados = ["ID", "Nombre", "DUI", "Tipo", "Acreedor", "Monto", "Estado", "Acciones"]
        for i, encabezado in enumerate(encabezados):
            label = ctk.CTkLabel(self.scrollframe, text=encabezado, font=ctk.CTkFont(size=12, weight="bold"),
                                text_color=("#0f172a", "#e2e8f0"), anchor="w")
            label.grid(row=0, column=i, padx=8, pady=10, sticky="ew")
            self.scrollframe.grid_columnconfigure(i, weight=1)

        self.servicio_descuentos = db.descuentos() if db is not None else None
        self.cargar_info_descuentos()
        self.entrada_buscar_prestamo.bind("<Return>", lambda e: self._ejecutar_busqueda_avanzada())

    def _obtener_datos_departamentos(self):
        """Devuelve la lista de departamentos y el mapa DUI -> departamento para registros activos."""
        try:
            if db is None:
                return [], {}

            servicio_descuentos = getattr(self, "servicio_descuentos", None)
            if servicio_descuentos is None:
                servicio_descuentos = db.descuentos()
                self.servicio_descuentos = servicio_descuentos

            descuentos = servicio_descuentos.obtener_descuentos(estado=None)
            if not descuentos:
                return [], {}

            servicio_empleados = db.gestion_empleado()
            empleados = servicio_empleados.obtener_empleados()
            mapa_dui_departamento = {}
            for empleado in empleados:
                if len(empleado) <= 7:
                    continue
                dui = str(empleado[2] or "").strip()
                departamento = str(empleado[7] or "").strip()
                if dui:
                    mapa_dui_departamento[dui] = departamento

            departamentos = []
            mapa = {}
            for registro in descuentos:
                dui = str(registro[2] or "").strip()
                if dui:
                    mapa[dui] = mapa_dui_departamento.get(dui, "")
                    departamento = mapa_dui_departamento.get(dui, "").strip()
                    if departamento and departamento not in departamentos:
                        departamentos.append(departamento)

            return sorted(departamentos), mapa
        except Exception:
            return [], {}

    def _actualizar_combo_departamentos(self):
        """Actualiza el combobox de departamento con los empleados existentes."""
        if not hasattr(self, "combo_departamento_filtro") or self.combo_departamento_filtro is None:
            return
        departamentos, _ = self._obtener_datos_departamentos()
        valores = ["Todos"] + departamentos
        valor_actual = self.combo_departamento_filtro.get() if self.combo_departamento_filtro.winfo_exists() else "Todos"
        if valor_actual not in valores:
            valor_actual = "Todos"
        self.combo_departamento_filtro.configure(values=valores)
        self.combo_departamento_filtro.set(valor_actual)

    def _ejecutar_busqueda_avanzada(self):
        """Ejecuta búsqueda con filtros avanzados"""
        self.filtros_activos["empleado"] = self.entrada_buscar_prestamo.get().strip().lower()
        self.filtros_activos["tipo"] = ""
        departamento = self.combo_departamento_filtro.get() if hasattr(self, "combo_departamento_filtro") else "Todos"
        self.filtros_activos["departamento"] = "Todos" if departamento == "Todos" else departamento
        estado = self.combo_estado_filtro.get()
        self.filtros_activos["estado"] = None if estado == "TODOS" else estado
        self.cargar_info_descuentos()

    def _limpiar_tabla(self):
        """Limpia todas las filas de la tabla excepto el encabezado"""
        for widget in self.scrollframe.winfo_children():
            info = widget.grid_info()
            if info.get("row", 0) > 0:
                widget.destroy()

    def _render_fila_descuento(self, fila, descuento):
        """Renderiza una fila de descuento con botones de acción"""
        id_desc, nombre, dui, cargo, salario, tipo, acreedor, monto, estado, obs, fecha_c, fecha_a = descuento

        # ID
        ctk.CTkLabel(self.scrollframe, text=str(id_desc), font=ctk.CTkFont(size=11),
                    fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=0, padx=8, pady=6, sticky="ew")

        # Nombre
        ctk.CTkLabel(self.scrollframe, text=nombre or "-", font=ctk.CTkFont(size=11),
                    fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=1, padx=8, pady=6, sticky="ew")

        # DUI
        ctk.CTkLabel(self.scrollframe, text=dui or "-", font=ctk.CTkFont(size=11),
                    fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=2, padx=8, pady=6, sticky="ew")

        # Tipo
        ctk.CTkLabel(self.scrollframe, text=tipo or "-", font=ctk.CTkFont(size=11),
                    fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=3, padx=8, pady=6, sticky="ew")

        # Acreedor
        ctk.CTkLabel(self.scrollframe, text=acreedor or "-", font=ctk.CTkFont(size=11),
                    fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=4, padx=8, pady=6, sticky="ew")

        # Monto
        try:
            monto_fmt = f"B/. {float(monto):,.2f}"
        except:
            monto_fmt = str(monto or "")
        ctk.CTkLabel(self.scrollframe, text=monto_fmt, font=ctk.CTkFont(size=11, weight="bold"),
                    text_color="#2563eb", fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=5, padx=8, pady=6, sticky="ew")

        # Estado
        color_estado = {"ACTIVO": "#2563eb", "PAGADO": "#10b981", "ANULADO": "#ef4444"}.get(estado, "#64748b")
        ctk.CTkLabel(self.scrollframe, text=estado, font=ctk.CTkFont(size=11, weight="bold"),
                    text_color=color_estado, fg_color=("#f8fafc", "#101b2a") if fila % 2 == 0 else ("#ffffff", "#1e293b")).grid(
            row=fila + 1, column=6, padx=8, pady=6, sticky="ew")

        # Botones de acción
        frame_acciones = ctk.CTkFrame(self.scrollframe, fg_color="transparent")
        frame_acciones.grid(row=fila + 1, column=7, padx=8, pady=6, sticky="ew")

        btn_editar = ctk.CTkButton(frame_acciones, text="✏️ Editar", width=65, height=24,
                                   font=ctk.CTkFont(size=10), fg_color="#3b82f6", hover_color="#2563eb",
                                   command=lambda id_d=id_desc: self.formulario_descuento(self.sipp_ref, id_d))
        btn_editar.pack(side="left", padx=2)

        btn_cambiar_estado = ctk.CTkButton(frame_acciones, text="↻ Estado", width=65, height=24,
                                          font=ctk.CTkFont(size=10), fg_color="#f59e0b", hover_color="#d97706",
                                          command=lambda id_d=id_desc, est=estado: self._cambiar_estado_descuento(id_d, est))
        btn_cambiar_estado.pack(side="left", padx=2)

        btn_eliminar = ctk.CTkButton(frame_acciones, text="🗑️ Eliminar", width=65, height=24,
                                    font=ctk.CTkFont(size=10), fg_color="#ef4444", hover_color="#dc2626",
                                    command=lambda id_d=id_desc, nom=nombre: self._eliminar_descuento(id_d, nom))
        btn_eliminar.pack(side="left", padx=2)

    def cargar_info_descuentos(self):
        """Carga y filtra los descuentos en la tabla"""
        if not hasattr(self, "scrollframe") or self.servicio_descuentos is None:
            return

        self._limpiar_tabla()

        try:
            # Obtener descuentos con el estado del filtro
            estado_filtro = self.filtros_activos.get("estado", "ACTIVO")
            registros = self.servicio_descuentos.obtener_descuentos(estado=estado_filtro)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar descuentos: {exc}")
            registros = []

        # Aplicar filtros
        filtro_empleado = self.filtros_activos.get("empleado", "").lower()
        filtro_departamento = self.filtros_activos.get("departamento", "Todos")
        _, departamentos_por_dui = self._obtener_datos_departamentos()

        registros_filtrados = []
        for r in registros:
            nombre = (r[1] or "").lower()
            dui = (r[2] or "").strip().upper()
            departamento_registro = departamentos_por_dui.get(dui, "")

            coincide_empleado = (
                not filtro_empleado
                or filtro_empleado in nombre
                or filtro_empleado in dui.lower()
                or filtro_empleado in (departamento_registro or "").lower()
            )
            coincide_departamento = (
                filtro_departamento == "Todos"
                or departamento_registro == filtro_departamento
            )

            if coincide_empleado and coincide_departamento:
                registros_filtrados.append(r)

        # Renderizar filas
        for idx, registro in enumerate(registros_filtrados):
            self._render_fila_descuento(idx, registro)

        # Actualizar métricas
        self._actualizar_metricas()

    def _actualizar_metricas(self):
        """Actualiza las tarjetas de métricas"""
        if self.servicio_descuentos is None:
            return

        try:
            # Obtener resumen
            resumen = self.servicio_descuentos.obtener_resumen_descuentos()
            activos = pagados = anulados = 0
            monto_total = 0.0

            for estado, cantidad, monto in resumen:
                if estado == "ACTIVO":
                    activos = cantidad
                elif estado == "PAGADO":
                    pagados = cantidad
                elif estado == "ANULADO":
                    anulados = cantidad
                monto_total += float(monto or 0)

            if hasattr(self, 'label_total_registrados'):
                self.label_total_registrados.configure(text=str(activos))
            if hasattr(self, 'label_pagados'):
                self.label_pagados.configure(text=str(pagados))
            if hasattr(self, 'label_anulados'):
                self.label_anulados.configure(text=str(anulados))
            if hasattr(self, 'label_total_monto'):
                self.label_total_monto.configure(text=f"B/. {monto_total:,.2f}")
        except Exception as exc:
            logging.error(f"Error actualizando métricas: {exc}")

    def _cambiar_estado_descuento(self, id_desc, estado_actual):
        """Abre diálogo para cambiar estado de descuento"""
        if self.servicio_descuentos is None:
            messagebox.showerror("Error", "No hay conexión con la base de datos")
            return
            
        estados_disponibles = ["ACTIVO", "PAGADO", "ANULADO"]

        top = ctk.CTkToplevel(self.sipp_ref)
        ajustar_toplevel_a_pantalla(top, 420, 250, margen_x=120, margen_y=100, min_ancho=340, min_alto=210)
        top.title("Cambiar Estado")
        top.grab_set()
        top.resizable(False, False)

        frame = ctk.CTkFrame(top, corner_radius=15, fg_color=("#f8fafc", "#111827"), border_width=1)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        ctk.CTkLabel(frame, text="Cambiar Estado del Descuento", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=15)

        ctk.CTkLabel(frame, text=f"Estado actual: {estado_actual}", font=ctk.CTkFont(size=12)).pack(pady=5)

        combo = ctk.CTkComboBox(frame, values=estados_disponibles, font=ctk.CTkFont(size=12), state="readonly")
        combo.pack(padx=20, pady=10, fill="x")
        combo.set(estado_actual)

        def guardar_cambio():
            nuevo_estado = combo.get()
            try:
                actualizado, msg = self.servicio_descuentos.cambiar_estado_descuento(id_desc, nuevo_estado)
                if actualizado:
                    messagebox.showinfo("Éxito", f"✓ Estado cambiado a {nuevo_estado}")
                    self.cargar_info_descuentos()
                    top.destroy()
                else:
                    messagebox.showerror("Error", f"No se pudo cambiar estado:\n{msg}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al cambiar estado:\n{str(e)}")

        frame_botones = ctk.CTkFrame(frame, fg_color="transparent")
        frame_botones.pack(pady=20)

        ctk.CTkButton(frame_botones, text="Guardar", width=120, command=guardar_cambio,
                     fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=10)
        ctk.CTkButton(frame_botones, text="Cancelar", width=120, command=top.destroy,
                     fg_color="#64748b", hover_color="#4b5563", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=10)

    def _eliminar_descuento(self, id_desc, nombre):
        """Elimina (anula) un descuento con confirmación"""
        if messagebox.askyesno("Confirmar", f"¿Anular descuento de {nombre}?\n\nEsta acción no se puede deshacer."):
            eliminado = self.servicio_descuentos.eliminar_descuento(id_desc)
            if eliminado:
                messagebox.showinfo("Éxito", "Descuento anulado correctamente")
                self.cargar_info_descuentos()
            else:
                messagebox.showerror("Error", "No se pudo anular el descuento")

    def _mostrar_reportes(self):
        """Muestra y exporta descuentos aplicados dentro de un corte cerrado."""
        top = ctk.CTkToplevel(self.sipp_ref)
        ajustar_toplevel_a_pantalla(top, 900, 620, margen_x=120, margen_y=110, min_ancho=640, min_alto=430)
        top.title("Reportes de Descuentos")
        top.grab_set()
        top.resizable(False, False)

        frame = ctk.CTkFrame(top, corner_radius=15, fg_color=("#f8fafc", "#111827"), border_width=1)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        ctk.CTkLabel(
            frame,
            text="Reporte de descuentos por corte cerrado",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(15, 8))

        barra = ctk.CTkFrame(frame, fg_color="transparent")
        barra.pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkLabel(barra, text="Corte realizado:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 8))

        def formatear_fecha(fecha):
            if hasattr(fecha, "strftime"):
                return fecha.strftime("%d/%m/%Y")
            try:
                return datetime.strptime(str(fecha)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                return str(fecha or "")

        def convertir_fecha(fecha):
            if hasattr(fecha, "date"):
                return fecha.date()
            if hasattr(fecha, "year") and hasattr(fecha, "month") and hasattr(fecha, "day"):
                return fecha
            texto = str(fecha or "").strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto[:10], formato).date()
                except ValueError:
                    continue
            return None

        cortes_cerrados = []
        for corte in db.cortes_planilla().obtener_cortes_planilla() or []:
            if len(corte) >= 4 and str(corte[3] or "").strip().upper() == "CERRADO":
                texto_corte = f"Corte {corte[0]}: {formatear_fecha(corte[1])} a {formatear_fecha(corte[2])}"
                cortes_cerrados.append((texto_corte, corte))

        combo_corte = ctk.CTkComboBox(
            barra,
            values=[texto for texto, _ in cortes_cerrados] or ["Sin cortes realizados"],
            state="readonly",
            width=360,
        )
        combo_corte.pack(side="left", fill="x", expand=True, padx=(0, 10))
        if cortes_cerrados:
            combo_corte.set(cortes_cerrados[-1][0])
        else:
            combo_corte.set("Sin cortes realizados")

        tabla = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        tabla.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        encabezados = [_tr("Nombre"), _tr("DUI"), _tr("Tipo"), _tr("Acreedor"), _tr("Monto"), _tr("Estado"), _tr("Fecha")]
        filas_reporte = []

        def obtener_registros():
            if not cortes_cerrados:
                return []
            indice = next((i for i, (texto, _) in enumerate(cortes_cerrados) if texto == combo_corte.get()), 0)
            corte = cortes_cerrados[indice][1]
            inicio = convertir_fecha(corte[1])
            fin = convertir_fecha(corte[2])
            registros = []
            for registro in self.servicio_descuentos.obtener_descuentos(estado=None):
                fecha_descuento = convertir_fecha(registro[10]) if len(registro) > 10 else None
                estado = str(registro[8] or "").strip().upper()
                if estado != "ANULADO" and fecha_descuento and inicio <= fecha_descuento <= fin:
                    registros.append(registro)
            return registros

        def cargar_reporte():
            nonlocal filas_reporte
            for widget in tabla.winfo_children():
                widget.destroy()
            filas_reporte = obtener_registros()
            valores = [encabezados] + [
                [r[1] or "", r[2] or "", r[5] or "", r[6] or "", f"B/. {float(r[7] or 0):,.2f}", r[8] or "", formatear_fecha(r[10])]
                for r in filas_reporte
            ]
            for fila, valores_fila in enumerate(valores):
                for columna, valor in enumerate(valores_fila):
                    ctk.CTkLabel(
                        tabla,
                        text=str(valor),
                        font=ctk.CTkFont(size=11, weight="bold" if fila == 0 else "normal"),
                        anchor="w",
                    ).grid(row=fila, column=columna, padx=7, pady=5, sticky="ew")
            for columna in range(len(encabezados)):
                tabla.grid_columnconfigure(columna, weight=1)

        def exportar_excel():
            registros = obtener_registros()
            if not registros:
                messagebox.showinfo("Reporte", "El corte seleccionado no tiene descuentos aplicados.", parent=top)
                return
            ruta = filedialog.asksaveasfilename(
                parent=top, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                initialfile="reporte_descuentos_corte.xlsx",
            )
            if not ruta:
                return
            libro = Workbook()
            hoja = libro.active
            hoja.title = "Descuentos"
            hoja.append(encabezados)
            for registro in registros:
                hoja.append([registro[1], registro[2], registro[5], registro[6], float(registro[7] or 0), registro[8], convertir_fecha(registro[10])])
            for celda in hoja[1]:
                celda.font = celda.font.copy(bold=True)
            hoja.column_dimensions["A"].width = 28
            hoja.column_dimensions["D"].width = 22
            hoja.column_dimensions["E"].width = 14
            libro.save(ruta)
            messagebox.showinfo("Reporte", "Reporte Excel generado correctamente.", parent=top)

        def exportar_pdf():
            registros = obtener_registros()
            if not registros:
                messagebox.showinfo("Reporte", "El corte seleccionado no tiene descuentos aplicados.", parent=top)
                return
            ruta = filedialog.asksaveasfilename(
                parent=top, defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                initialfile="reporte_descuentos_corte.pdf",
            )
            if not ruta:
                return
            estilos = getSampleStyleSheet()
            documento = SimpleDocTemplate(ruta, pagesize=landscape(letter), rightMargin=0.35 * inch, leftMargin=0.35 * inch)
            datos_pdf = [[Paragraph(str(valor), estilos["Normal"]) for valor in encabezados]]
            for registro in registros:
                datos_pdf.append([
                    str(registro[1] or ""), str(registro[2] or ""), str(registro[5] or ""),
                    str(registro[6] or ""), f"B/. {float(registro[7] or 0):,.2f}",
                    str(registro[8] or ""), formatear_fecha(registro[10]),
                ])
            tabla_pdf = Table(datos_pdf, repeatRows=1)
            tabla_pdf.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            documento.build([Paragraph(_tr("Reporte de descuentos aplicados"), estilos["Title"]), Spacer(1, 10), tabla_pdf])
            messagebox.showinfo("Reporte", "Reporte PDF generado correctamente.", parent=top)

        combo_corte.configure(command=lambda _: cargar_reporte())
        ctk.CTkButton(barra, text="Excel", width=85, command=exportar_excel, fg_color="#10b981").pack(side="left", padx=3)
        ctk.CTkButton(barra, text="PDF", width=85, command=exportar_pdf, fg_color="#ef4444").pack(side="left", padx=3)
        ctk.CTkButton(barra, text="Cerrar", width=85, command=top.destroy, fg_color="#64748b").pack(side="left", padx=(3, 0))
        cargar_reporte()

    def formulario_descuento(self, sipp, id_editar=None):
        """Formulario completo para crear o editar descuentos"""
        es_edicion = id_editar is not None

        top_level = ctk.CTkToplevel(sipp)
        ajustar_toplevel_a_pantalla(top_level, 700, 750, margen_x=120, margen_y=120, min_ancho=520, min_alto=450)
        top_level.title("Editar Descuento" if es_edicion else "Nuevo Descuento")
        top_level.grab_set()
        top_level.resizable(False, False)

        frame_form = ctk.CTkFrame(top_level, corner_radius=16, fg_color=("#f8fafc", "#111827"), border_width=1)
        frame_form.pack(fill="both", expand=True, padx=16, pady=16)

        # TÍTULO
        title_frame = ctk.CTkFrame(frame_form, fg_color="transparent")
        title_frame.pack(padx=16, pady=(16, 8), fill="x")

        ctk.CTkLabel(title_frame, text="📝 " + ("Editar Descuento" if es_edicion else "Crear Nuevo Descuento"),
                    font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Complete los datos del descuento o préstamo",
                    font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8")).pack(anchor="w", pady=(2, 0))

        # BÚSQUEDA DE EMPLEADO
        search_frame = ctk.CTkFrame(frame_form, corner_radius=12, fg_color=("#eef6fb", "#1e293b"), border_width=1)
        search_frame.pack(padx=16, pady=(8, 12), fill="x")

        ctk.CTkLabel(search_frame, text="👤 Empleado", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=12, pady=(10, 0))

        buscar_frame = ctk.CTkFrame(search_frame, fg_color="transparent")
        buscar_frame.pack(padx=12, pady=(0, 10), fill="x")

        entrada_buscar = ctk.CTkEntry(buscar_frame, placeholder_text="Buscar por DUI o nombre",
                                     font=ctk.CTkFont(size=12))
        entrada_buscar.pack(side="left", fill="x", expand=True, padx=(0, 8))

        empleado_data = {"id": None, "nombre": None, "dui": None, "cargo": None, "salario": None}

        # INFO EMPLEADO
        empleado_frame = ctk.CTkFrame(frame_form, corner_radius=12, fg_color=("#ffffff", "#1e293b"), border_width=1)
        empleado_frame.pack(padx=16, pady=12, fill="x")

        info_labels = [
            ("Nombre:", "lbl_nombre"),
            ("DUI:", "lbl_dui"),
            ("Cargo:", "lbl_cargo"),
            ("Salario:", "lbl_salario")
        ]

        labels_dict = {}
        for i, (label_text, var_name) in enumerate(info_labels):
            ctk.CTkLabel(empleado_frame, text=label_text, font=ctk.CTkFont(size=11, weight="bold")).grid(
                row=i, column=0, padx=(12, 8), pady=8, sticky="w")
            lbl = ctk.CTkLabel(empleado_frame, text="-", font=ctk.CTkFont(size=11))
            lbl.grid(row=i, column=1, padx=8, pady=8, sticky="w")
            labels_dict[var_name] = lbl

        def llenar_empleado(empleado):
            if not empleado:
                empleado_data.update({"id": None, "nombre": None, "dui": None, "cargo": None, "salario": None})
                labels_dict["lbl_nombre"].configure(text="-")
                labels_dict["lbl_dui"].configure(text="-")
                labels_dict["lbl_cargo"].configure(text="-")
                labels_dict["lbl_salario"].configure(text="-")
                return

            id_e, nombre, dui, correo, direccion, telefono, puesto, departamento, salario, ingreso, nacimiento = empleado
            empleado_data.update({"id": id_e, "nombre": nombre, "dui": dui, "cargo": puesto, "salario": salario})
            labels_dict["lbl_nombre"].configure(text=nombre or "-")
            labels_dict["lbl_dui"].configure(text=dui or "-")
            labels_dict["lbl_cargo"].configure(text=puesto or "-")
            
            # Convertir salario de forma robusta
            try:
                salario_val = float(salario or 0)
                salario_text = f"B/. {salario_val:,.2f}"
            except (ValueError, TypeError):
                salario_text = f"B/. {salario or 0}"
            labels_dict["lbl_salario"].configure(text=salario_text)

        def buscar_empleado_fn(auto=False):
            texto = entrada_buscar.get().strip()
            if not texto:
                llenar_empleado(None)
                if not auto:
                    messagebox.showwarning("Aviso", "Ingrese el DUI o nombre del empleado")
                return
            try:
                service = db.gestion_empleado()
                texto_norm = texto.strip().lower()

                empleados = service.obtener_empleados()
                mejor_coincidencia = None
                mejor_puntaje = -1

                for emp in empleados:
                    if len(emp) < 11:
                        continue
                    id_e, nombre, dui, correo, direccion, telefono, puesto, departamento, salario, ingreso, nacimiento = emp
                    nombre_norm = (nombre or "").strip().lower()
                    dui_norm = (dui or "").strip().lower()

                    puntaje = 0
                    
                    # Búsqueda por DUI (prioridad máxima)
                    if dui_norm == texto_norm:
                        puntaje = 1000
                    elif dui_norm.startswith(texto_norm):
                        puntaje = 800
                    elif texto_norm in dui_norm:
                        puntaje = 600
                    
                    # Búsqueda por nombre
                    if nombre_norm == texto_norm:
                        puntaje = max(puntaje, 950)
                    elif nombre_norm.startswith(texto_norm):
                        puntaje = max(puntaje, 700)
                    elif texto_norm in nombre_norm:
                        puntaje = max(puntaje, 400)
                    
                    if puntaje > mejor_puntaje:
                        mejor_puntaje = puntaje
                        mejor_coincidencia = emp

                empleado = mejor_coincidencia

                if empleado and mejor_puntaje > 0:
                    llenar_empleado(empleado)
                    if not auto:
                        messagebox.showinfo("Éxito", f"Empleado encontrado: {empleado[1]} - DUI: {empleado[2]}")
                else:
                    llenar_empleado(None)
                    if not auto:
                        messagebox.showwarning("Aviso", f"No se encontró un empleado con DUI o nombre: {texto_norm}")
            except Exception as e:
                llenar_empleado(None)
                print(f"Error en búsqueda de empleado: {e}")
                if not auto:
                    messagebox.showerror("Error", f"Error al buscar empleado: {str(e)}")

        entrada_buscar.bind("<Return>", lambda event: buscar_empleado_fn())
        entrada_buscar.bind("<KeyRelease>", lambda event: buscar_empleado_fn(auto=True))

        def limpiar_busqueda():
            entrada_buscar.delete(0, "end")
            llenar_empleado(None)

        btn_limpiar = ctk.CTkButton(buscar_frame, text="🗑️ Limpiar", width=90, font=ctk.CTkFont(size=11, weight="bold"),
                                    fg_color="#ef4444", hover_color="#dc2626", command=limpiar_busqueda)
        btn_limpiar.pack(side="left")

        # Si es edición, cargar datos
        if es_edicion:
            try:
                descuento = self.servicio_descuentos.obtener_descuento_por_id(id_editar)
                if descuento:
                    _, nombre, dui, cargo, salario, _, _, _, _, _, _, _ = descuento
                    empleado_data.update({"nombre": nombre, "dui": dui, "cargo": cargo, "salario": salario})
                    labels_dict["lbl_nombre"].configure(text=nombre or "-")
                    labels_dict["lbl_dui"].configure(text=dui or "-")
                    labels_dict["lbl_cargo"].configure(text=cargo or "-")
                    
                    # Convertir salario de forma robusta
                    try:
                        salario_val = float(salario or 0)
                        salario_text = f"B/. {salario_val:,.2f}"
                    except (ValueError, TypeError):
                        salario_text = f"B/. {salario or 0}"
                    labels_dict["lbl_salario"].configure(text=salario_text)
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar descuento: {e}")

        # FORM SCROLLABLE
        scroll_form = ctk.CTkScrollableFrame(frame_form, corner_radius=12, fg_color="transparent")
        scroll_form.pack(padx=16, pady=12, fill="both", expand=False)

        # Tipo de retención
        ctk.CTkLabel(scroll_form, text="Tipo de Retención", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(10, 4))
        combo_tipo = ctk.CTkComboBox(scroll_form, values=self.lista_descuentos(), font=ctk.CTkFont(size=11))
        combo_tipo.pack(fill="x", padx=0, pady=(0, 12))
        combo_tipo.set(self.lista_descuentos()[0])
        # Registrar este combobox en la instancia de Deducciones para actualizarlo cuando se agreguen nuevos descuentos
        if hasattr(Deducciones, '_instance') and Deducciones._instance is not None:
            Deducciones._instance.registrar_combo_descuentos(combo_tipo)

        # Entidad/Acreedor
        ctk.CTkLabel(scroll_form, text="Entidad / Acreedor", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(10, 4))
        entrada_acreedor = ctk.CTkEntry(scroll_form, placeholder_text="Nombre de la entidad", font=ctk.CTkFont(size=11))
        entrada_acreedor.pack(fill="x", pady=(0, 12))

        # Monto
        ctk.CTkLabel(scroll_form, text="Monto (B/.)", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(10, 4))
        entrada_monto = ctk.CTkEntry(scroll_form, placeholder_text="Ej: 500 o 500.50", font=ctk.CTkFont(size=11))
        entrada_monto.pack(fill="x", pady=(0, 12))
        
        # Función para limpiar y validar entrada de monto en tiempo real
        def validar_monto_entrada(event=None):
            """Permite solo números y separadores decimales (punto y coma)"""
            valor = entrada_monto.get()
            # Permitir solo dígitos, punto y coma
            valor_limpio = ''.join(c for c in valor if c.isdigit() or c in '.,')
            # Limitar un solo punto o coma
            if valor_limpio.count('.') > 1 or valor_limpio.count(',') > 1:
                # Si hay más de un separador, remover los excedentes
                valor_limpio = valor_limpio.replace('.', '', valor_limpio.count('.') - 1)
                valor_limpio = valor_limpio.replace(',', '', valor_limpio.count(',') - 1)
            
            if valor != valor_limpio:
                entrada_monto.delete(0, tk.END)
                entrada_monto.insert(0, valor_limpio)
        
        entrada_monto.bind("<KeyRelease>", validar_monto_entrada)

        # Observaciones
        ctk.CTkLabel(scroll_form, text="Observaciones (opcional)", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(10, 4))
        entrada_obs = ctk.CTkTextbox(scroll_form, height=80, font=ctk.CTkFont(size=11), corner_radius=8)
        entrada_obs.pack(fill="x", pady=(0, 12))

        # Si es edición, cargar valores
        if es_edicion:
            try:
                descuento = self.servicio_descuentos.obtener_descuento_por_id(id_editar)
                if descuento:
                    _, _, _, _, _, tipo, acreedor, monto, _, obs, _, _ = descuento
                    combo_tipo.set(tipo or self.lista_descuentos()[0])
                    entrada_acreedor.insert(0, acreedor or "")
                    entrada_monto.insert(0, str(float(monto or 0)))
                    entrada_obs.insert("1.0", obs or "")
            except Exception as e:
                logging.error(f"Error cargando edición: {e}")

        # BOTONES
        frame_botones = ctk.CTkFrame(frame_form, fg_color="transparent")
        frame_botones.pack(padx=16, pady=(0, 16), fill="x")

        def guardar_descuento_fn():
            nombre = empleado_data.get("nombre")
            dui = empleado_data.get("dui")
            cargo = empleado_data.get("cargo")
            salario = empleado_data.get("salario")
            tipo = combo_tipo.get()
            acreedor = entrada_acreedor.get().strip()
            monto_str = entrada_monto.get().strip()
            obs = entrada_obs.get("1.0", "end").strip()

            # Validaciones
            if not nombre or not dui:
                messagebox.showwarning("Error", "Debe buscar y seleccionar un empleado")
                return
            if not acreedor:
                messagebox.showwarning("Error", "Ingrese la entidad/acreedor")
                return
            if not monto_str:
                messagebox.showwarning("Error", "Ingrese el monto")
                return

            # Función para convertir monto de diferentes formatos
            def convertir_monto(monto_text):
                """Convierte monto aceptando formatos: '500', '500.50', '500,50'"""
                try:
                    # Limpiar espacios
                    monto_text = monto_text.strip()
                    
                    if not monto_text:
                        raise ValueError("El monto no puede estar vacío")
                    
                    # Reemplazar coma por punto (formato europeo a inglés)
                    monto_text = monto_text.replace(',', '.')
                    
                    # Validar que solo tenga dígitos y un punto
                    if not all(c.isdigit() or c == '.' for c in monto_text):
                        raise ValueError("El monto contiene caracteres inválidos")
                    
                    # Convertir a float
                    monto = float(monto_text)
                    
                    # Validar que sea positivo
                    if monto <= 0:
                        raise ValueError("El monto debe ser mayor a 0")
                    
                    # Validar rango razonable (máximo 999,999.99)
                    if monto > 999999.99:
                        raise ValueError("El monto es demasiado grande (máximo 999,999.99)")
                    
                    # Redondear a 2 decimales
                    return round(monto, 2)
                except ValueError as e:
                    raise ValueError(f"Formato de monto inválido: {str(e)}")


            try:
                monto = convertir_monto(monto_str)
            except ValueError as e:
                messagebox.showwarning("Error de Monto", str(e))
                return

            try:
                if es_edicion:
                    actualizado, msg = self.servicio_descuentos.actualizar_descuento(
                        id_editar, nombre, dui, cargo, salario, tipo, acreedor, monto, obs
                    )
                    if actualizado:
                        messagebox.showinfo("Éxito", "Descuento actualizado correctamente")
                        self.cargar_info_descuentos()
                        top_level.destroy()
                    else:
                        if msg == "DUPLICADO":
                            messagebox.showwarning("Aviso", "Ya existe otro descuento activo igual para este empleado.")
                        else:
                            messagebox.showerror("Error", f"No se pudo actualizar: {msg}")
                else:
                    insertor, resultado = self.servicio_descuentos.insertar_descuento(
                        nombre, dui, cargo, salario, tipo, acreedor, monto, obs
                    )
                    if insertor:
                        messagebox.showinfo("Éxito", f"Descuento creado correctamente (ID: {resultado})")
                        self.cargar_info_descuentos()
                        top_level.destroy()
                    else:
                        if resultado == "DUPLICADO":
                            messagebox.showwarning("Aviso", "Este descuento ya existe para este empleado")
                        elif resultado == "MONTO_INVALIDO":
                            messagebox.showerror("Error de Monto", f"El monto no es válido. Ingrese un número (ej: 500 o 500.50). Valor ingresado: {monto_str}")
                        elif resultado == "MONTO_NEGATIVO":
                            messagebox.showerror("Error de Monto", "El monto debe ser mayor a 0")
                        elif resultado == "ERROR_DB":
                            messagebox.showerror("Error de Base de Datos", "No se pudo guardar el descuento en la base de datos. Intente nuevamente.")
                        else:
                            messagebox.showerror("Error", f"Error al guardar: {resultado}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar: {e}")

        ctk.CTkButton(frame_botones, text="✓ Guardar", width=140, font=ctk.CTkFont(size=12, weight="bold"),
                     fg_color="#10b981", hover_color="#059669", command=guardar_descuento_fn).pack(side="left", padx=(0, 8))
        ctk.CTkButton(frame_botones, text="✕ Cancelar", width=140, font=ctk.CTkFont(size=12, weight="bold"),
                     fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="left")


class Deducciones:
    # Instancia singleton para compartir referencias de combobox
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(Deducciones, cls).__new__(cls)
            cls._instance.combos_descuentos = []
        return cls._instance

    valores_por_defecto = {
        "CSS": "9.75%",
        "Seguro Educativo": "1.25%",
        "Impuesto Sobre la Renta": "ISR progresivo (Panamá)"
    }

    @staticmethod
    def parse_porcentaje(valor):
        """Convierte un valor de porcentaje a decimal o float, aceptando 9.75, '9.75%' y '9,75'."""
        if valor is None:
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)
        texto = str(valor).strip().replace('%', '').replace(',', '.')
        if not texto:
            return 0.0
        try:
            return float(texto)
        except ValueError:
            return 0.0

    @classmethod
    def porcentaje_decimal(cls, valor):
        """Convierte un porcentaje a decimal para el cálculo: 9.75 -> 0.0975."""
        return cls.parse_porcentaje(valor) / 100.0

    @staticmethod
    def redondear(valor):
        return round(float(valor or 0.0), 2)

    def formula_isr_panama(self, salario_mensual):
        """
        Fórmula reutilizable de ISR para Panamá.
        La base calculada es anual y la regla es progresiva por tramos.
        Retorna el impuesto mensual estimado según la base salarial mensual.
        """
        try:
            salario_mensual = float(salario_mensual)
        except (TypeError, ValueError):
            return 0.0

        if salario_mensual <= 0:
            return 0.0

        base_anual = salario_mensual * 12

        if base_anual <= 11000:
            isr_anual = 0.0
        elif base_anual <= 50000:
            isr_anual = (base_anual - 11000) * 0.15
        else:
            isr_anual = 5850 + (base_anual - 50000) * 0.25

        return round(isr_anual / 12, 2)

    def get_isr_formula_texto(self):
        return (
            "ISR (reglas actuales usadas): "
            "0% hasta B/.11,000 anual; 15% sobre el excedente entre B/.11,000 y B/.50,000; "
            "más de B/.50,000: B/.5,850 más 25% sobre lo que exceda B/.50,000."
        )

    def tasa_css(self):
        try:
            css_txt = str(self.valores_por_defecto.get("CSS", "0%"))
            return self.parse_porcentaje(css_txt)
        except Exception:
            return 0.0

    def tasa_seguro_educativo(self):
        try:
            edu_txt = str(self.valores_por_defecto.get("Seguro Educativo", "0%"))
            return self.parse_porcentaje(edu_txt)
        except Exception:
            return 0.0

    def calcular_css(self, salario_base, porcentaje=None):
        porcentaje_usado = self.tasa_css() if porcentaje is None else self.parse_porcentaje(porcentaje)
        return max(0.0, float(salario_base or 0.0)) * (porcentaje_usado / 100.0)

    def calcular_seguro_educativo(self, salario_base, porcentaje=None):
        porcentaje_usado = self.tasa_seguro_educativo() if porcentaje is None else self.parse_porcentaje(porcentaje)
        return max(0.0, float(salario_base or 0.0)) * (porcentaje_usado / 100.0)

    def calcular_salario_periodo(self, salario_mensual, dias_pagados):
        """Calcula el salario proporcional de un período de hasta quince días."""
        salario_mensual = max(0.0, float(salario_mensual or 0.0))
        dias_pagados = min(15.0, max(0.0, float(dias_pagados or 0.0)))
        return self.redondear(salario_mensual * dias_pagados / 30.0)

    def calcular_deducciones_nomina(self, salario_base, porcentaje_css=None, porcentaje_seguro_educativo=None, incluir_isr=True, otros_descuentos=0, salario_mensual=None, dias_pagados=None):
        """Calcula deducciones legales, otros descuentos y salario neto."""
        salario_base = float(salario_base or 0.0)
        otros_descuentos = max(0.0, float(otros_descuentos or 0.0))
        css = self.calcular_css(salario_base, porcentaje_css)
        seguro = self.calcular_seguro_educativo(salario_base, porcentaje_seguro_educativo)
        isr = self.formula_isr_panama(salario_base) if incluir_isr else 0.0
        if incluir_isr and salario_mensual is not None and dias_pagados is not None:
            salario_mensual = max(0.0, float(salario_mensual or 0.0))
            dias_pagados = min(30.0, max(0.0, float(dias_pagados or 0.0)))
            isr = self.formula_isr_panama(salario_mensual) * (dias_pagados / 30.0)

        total_deducciones = css + seguro + isr + otros_descuentos
        salario_neto = max(0.0, salario_base - total_deducciones)

        return {
            "salario_base": self.redondear(salario_base),
            "css": self.redondear(css),
            "seguro_educativo": self.redondear(seguro),
            "isr": self.redondear(isr),
            "otros_descuentos": self.redondear(otros_descuentos),
            "total_deducciones": self.redondear(total_deducciones),
            "salario_neto": self.redondear(salario_neto),
        }

    def info_deducciones(self, frame):
        self.valores_deducciones = dict(self.__class__.valores_por_defecto)
        self.labels_valor_deducciones = {}
        self.labels_metricos = {}

        frame_principal = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#eef6fb", "#101b2a"), border_width=1)
        frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        header = ctk.CTkFrame(frame_principal, corner_radius=12, fg_color=("#d9edf7", "#18263a"), border_width=1, height=92)
        header.place(relx=0.02, rely=0.02, relwidth=0.96)
        ctk.CTkLabel(header, text="Deducciones y Retenciones", font=ctk.CTkFont(size=25, weight="bold"), anchor="w").place(x=16, y=10)
        ctk.CTkLabel(header, text="Retenciones legales y descuentos autorizados de nómina",
                      font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w").place(x=18, y=50)

        metrics = ctk.CTkFrame(frame_principal, corner_radius=10, fg_color=("#ffffff", "#182334"), height=50)
        metrics.place(relx=0.02, rely=0.15, relwidth=0.96)
        ctk.CTkLabel(metrics, text="CSS", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=14, y=12)
        self.labels_metricos["CSS"] = ctk.CTkLabel(metrics, text=self.valores_deducciones["CSS"], font=ctk.CTkFont(size=13, weight="bold"), text_color="#18794e")
        self.labels_metricos["CSS"].place(x=90, y=10)
        ctk.CTkLabel(metrics, text="ISR", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=250, y=12)
        self.labels_metricos["Impuesto Sobre la Renta"] = ctk.CTkLabel(
            metrics,
            text="Progresivo",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#b45309",
            wraplength=280,
            justify="left"
        )
        self.labels_metricos["Impuesto Sobre la Renta"].place(x=325, y=10)
        ctk.CTkLabel(metrics, text="Seguro Educativo", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").place(x=455, y=12)
        self.labels_metricos["Seguro Educativo"] = ctk.CTkLabel(metrics, text=self.valores_deducciones["Seguro Educativo"], font=ctk.CTkFont(size=13, weight="bold"), text_color="#1d4ed8")
        self.labels_metricos["Seguro Educativo"].place(x=605, y=10)

        formula_isr = ctk.CTkLabel(
            frame_principal,
            text="Fórmula",
            font=ctk.CTkFont(size=10),
            text_color=("#526071", "#a1a9b8"),
            anchor="w",
            justify="left",
            wraplength=640
        )
        formula_isr.place(relx=0.02, rely=0.905, relwidth=0.96)

        toolbar = ctk.CTkFrame(frame_principal, corner_radius=10, fg_color=("#f8fafc", "#111b27"), height=52)
        toolbar.place(relx=0.02, rely=0.23, relwidth=0.96)

        self.entrada_buscar_descuento = ctk.CTkEntry(toolbar, width=280, placeholder_text="Buscar deducción", font=ctk.CTkFont(size=14), justify="center")
        self.entrada_buscar_descuento.place(x=14, y=10)
        ctk.CTkButton(toolbar, text="Buscar", width=110, font=ctk.CTkFont(size=13), fg_color="#4f46e5").place(x=340, y=8)
        ctk.CTkButton(toolbar, text="Nuevo ajuste", width=140, font=ctk.CTkFont(size=13), fg_color="#2563eb",
                       command=lambda: self.actualizar_otros_descuentos(toolbar)).place(x=470, y=8)

        content = ctk.CTkScrollableFrame(frame_principal, corner_radius=10)
        content.place(relx=0.02, rely=0.34, relwidth=0.96, relheight=0.58)

        self._render_deduccion_row(content, 0, "CSS", self.valores_deducciones["CSS"], "Seguro Social", "Editar", self.actualizar_css)
        self._render_deduccion_row(content, 1, "Seguro Educativo", self.valores_deducciones["Seguro Educativo"], "Aporte educativo", "Editar", self.actualizar_educativo)
        self._render_deduccion_row(content, 2, "Impuesto Sobre la Renta", "Fórmula", "ISR", "Editar fórmula", self.actualizar_i_renta)

        row_otras = ctk.CTkFrame(content, corner_radius=10, fg_color=("#ffffff", "#182334"), height=128)
        row_otras.grid(row=3, column=0, padx=12, pady=10, sticky="ew")
        row_otras.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(row_otras, text="Otros descuentos autorizados", font=ctk.CTkFont(size=16, weight="bold"), anchor="w").place(x=12, y=10)
        ctk.CTkLabel(row_otras, text="Rangos, aportes y conceptos complementarios", font=ctk.CTkFont(size=11), anchor="w", text_color=("#526071", "#a1a9b8")).place(x=12, y=40)
        # Guardar referencia al combobox para actualizarlo después
        self.combo_descuentos_deducciones = ctk.CTkComboBox(row_otras, values=self.lista_descuentos(), width=300)
        self.combo_descuentos_deducciones.place(x=12, y=72)
        self.combos_descuentos.append(self.combo_descuentos_deducciones)
        # botón eliminado para evitar duplicidad con el toolbar ('Nueva deducción')

        aviso = ctk.CTkLabel(
            content,
            text="*Aviso legal: las deducciones de nómina se aplican conforme a la normativa vigente y a la política interna de retenciones, incluyendo CSS, ISSS e ISR.",
            font=ctk.CTkFont(size=11, slant="italic"),
            fg_color="transparent",
            wraplength=700,
            justify="left"
        )
        aviso.grid(row=4, column=0, padx=12, pady=(10, 4), sticky="w")

        aviso2 = ctk.CTkLabel(
            content,
            text="Además de políticas de deducción empresarial, La empresa garantiza transparencia en los cálculos según normativa laboral vigente.",
            font=ctk.CTkFont(size=11, slant="italic"),
            fg_color="transparent",
            wraplength=700,
            justify="left"
        )
        aviso2.grid(row=5, column=0, padx=12, pady=(4, 12), sticky="w")

    def _render_deduccion_row(self, parent, row, titulo, valor, descripcion, boton_texto, command):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=("#ffffff", "#182334"), height=90)
        card.grid(row=row, column=0, padx=12, pady=8, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").place(x=12, y=10)
        valor_label = ctk.CTkLabel(card, text=valor, font=ctk.CTkFont(size=15, weight="bold"), anchor="w")
        valor_label.place(x=180, y=10)
        self.labels_valor_deducciones[titulo] = valor_label
        ctk.CTkLabel(card, text=descripcion, font=ctk.CTkFont(size=10), text_color=("#526071", "#a1a9b8"), anchor="w").place(x=12, y=44)
        ctk.CTkButton(card, text=boton_texto, fg_color="#999fdb", command=lambda: command(card), width=110).place(x=360, y=30)

    def _actualizar_deduccion(self, titulo, nuevo_valor):
        if titulo not in self.valores_deducciones:
            self.valores_deducciones[titulo] = nuevo_valor
        else:
            self.valores_deducciones[titulo] = nuevo_valor

        if titulo in self.labels_valor_deducciones:
            self.labels_valor_deducciones[titulo].configure(text=nuevo_valor)

        if titulo in self.labels_metricos:
            self.labels_metricos[titulo].configure(text=nuevo_valor)

        self.__class__.valores_por_defecto.update(self.valores_deducciones)

    def _abrir_modal_deduccion(self, master, titulo, descripcion, valor_actual):
        top_level = ctk.CTkToplevel(master)
        # aumentar tamaño para que el contenido se aprecie correctamente
        ajustar_toplevel_a_pantalla(top_level, 700, 460, margen_x=120, margen_y=120, min_ancho=500, min_alto=320)
        top_level.title(titulo)
        top_level.grab_set()
        top_level.resizable(False, False)

        frame_form = ctk.CTkFrame(top_level, corner_radius=14, fg_color=("#f8fbfd", "#111827"), border_width=1)
        frame_form.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(frame_form, text=titulo, font=ctk.CTkFont(size=20, weight="bold"), anchor="w").grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 4), sticky="w")

        descripcion_modal = descripcion
        if titulo == "Impuesto Sobre la Renta":
            descripcion_modal = self.get_isr_formula_texto()

        ctk.CTkLabel(frame_form, text=descripcion_modal, font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8"), anchor="w", justify="left", wraplength=620).grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 14), sticky="w")

        ctk.CTkLabel(frame_form, text="Valor actual:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=12, pady=8, sticky="w")
        valor_actual_label = ctk.CTkLabel(frame_form, text="Fórmula", font=ctk.CTkFont(size=13), anchor="w", justify="left", wraplength=280)
        valor_actual_label.grid(row=2, column=1, padx=12, pady=8, sticky="w")

        ctk.CTkLabel(frame_form, text="Nuevo valor:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=3, column=0, padx=12, pady=8, sticky="w")
        entrada = ctk.CTkEntry(frame_form, width=220, placeholder_text="00.00%", justify="center")
        entrada.grid(row=3, column=1, padx=12, pady=8, sticky="ew")

        if titulo == "Impuesto Sobre la Renta":
            ctk.CTkLabel(frame_form, text="Regla aplicable:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=4, column=0, padx=12, pady=8, sticky="w")
            regla = ctk.CTkLabel(
                frame_form,
                text="Fórmula",
                font=ctk.CTkFont(size=10),
                text_color=("#526071", "#a1a9b8"),
                justify="left",
                wraplength=350,
                anchor="w"
            )
            regla.grid(row=4, column=1, padx=12, pady=8, sticky="w")

        botones = ctk.CTkFrame(frame_form, fg_color="transparent")
        botones.grid(row=5 if titulo == "Impuesto Sobre la Renta" else 4, column=0, columnspan=2, padx=10, pady=(20, 4), sticky="e")

        def guardar_valor():
            texto = entrada.get().strip()
            if not texto:
                messagebox.showwarning("Aviso", "Ingrese un valor para continuar.")
                return
            try:
                valor = float(texto.replace('%', ''))
            except (TypeError, ValueError):
                messagebox.showwarning("Aviso", "Debe ingresar un valor numérico válido.")
                return
            if valor < 0:
                messagebox.showwarning("Aviso", "El porcentaje no puede ser negativo.")
                return

            nuevo_valor = f"{valor:.2f}%"
            self._actualizar_deduccion(titulo, nuevo_valor)
            messagebox.showinfo("Actualizado", f"Se registró el valor {nuevo_valor} para {titulo}.")
            top_level.destroy()

        ctk.CTkButton(botones, text="Guardar", width=120, command=guardar_valor, fg_color="#2f855a").pack(side="left", padx=(0, 12))
        ctk.CTkButton(botones, text="Cancelar", width=120, command=top_level.destroy, fg_color="#64748b").pack(side="left")

    def actualizar_css(self, frame_2):
        self._abrir_modal_deduccion(
            frame_2,
            "CSS",
            "Ajuste del aporte del seguro social",
            "9.75%"
        )

    def actualizar_educativo(self, frame_3):
        self._abrir_modal_deduccion(
            frame_3,
            "Seguro Educativo",
            "Ajuste del aporte educativo",
            "1.25%"
        )

    def actualizar_i_renta(self, frame_4):
        self._abrir_modal_deduccion(
            frame_4,
            "Impuesto Sobre la Renta",
            "Ajuste del impuesto sobre la renta",
            self.get_isr_formula_texto()
        )

    def lista_descuentos(self):
        return DESCUENTOS_SISTEMA.copy()
    
    def _actualizar_todos_los_combos(self):
        """Actualiza todos los combobox registrados con la lista de descuentos actual"""
        try:
            valores_nuevos = self.lista_descuentos()
            for combo in self.combos_descuentos:
                if combo.winfo_exists():
                    valor_actual = combo.get()
                    combo.configure(values=valores_nuevos)
                    # Mantener el valor actual si sigue disponible
                    if valor_actual in valores_nuevos:
                        combo.set(valor_actual)
                    elif valores_nuevos:
                        combo.set(valores_nuevos[0])
        except Exception as e:
            print(f"Error actualizando combobox de descuentos: {e}")
    
    def registrar_combo_descuentos(self, combo):
        """Registra un combobox para que se actualice cuando se agreguen nuevos descuentos"""
        if combo not in self.combos_descuentos:
            self.combos_descuentos.append(combo)

    def actualizar_otros_descuentos(self, frame_5):
        top_level = ctk.CTkToplevel(frame_5)
        # ancho aumentado para mejor visibilidad del contenido
        ajustar_toplevel_a_pantalla(top_level, 640, 360, margen_x=120, margen_y=120, min_ancho=460, min_alto=260)
        top_level.title("Nuevo descuento autorizado")
        top_level.grab_set()
        top_level.resizable(False, False)

        frame_form = ctk.CTkFrame(top_level, corner_radius=14, fg_color=("#f8fbfd", "#111827"), border_width=1)
        frame_form.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(frame_form, text="Nuevo descuento autorizado", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").grid(row=0, column=0, padx=12, pady=(12, 10), sticky="w")
        ctk.CTkLabel(frame_form, text="Ingrese un nuevo concepto de descuento o ajuste salarial", font=ctk.CTkFont(size=11), text_color=("#526071", "#a1a9b8")).grid(row=1, column=0, padx=12, pady=(0, 12), sticky="w")

        ctk.CTkLabel(frame_form, text="Concepto:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=12, pady=8, sticky="w")
        entrada = ctk.CTkEntry(frame_form, width=280, placeholder_text="Ej: préstamo extraordinario")
        entrada.grid(row=2, column=1, padx=12, pady=8, sticky="ew")

        ctk.CTkLabel(frame_form, text="Tipo:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=3, column=0, padx=12, pady=8, sticky="w")
        combo = ctk.CTkComboBox(frame_form, values=self.lista_descuentos(), width=280)
        combo.grid(row=3, column=1, padx=12, pady=8, sticky="ew")
        combo.set(self.lista_descuentos()[0])

        botones = ctk.CTkFrame(frame_form, fg_color="transparent")
        botones.grid(row=4, column=0, columnspan=2, padx=12, pady=(20, 8), sticky="e")

        def guardar_otros():
            concepto = entrada.get().strip()
            tipo = combo.get()
            if not concepto:
                messagebox.showwarning("Aviso", "Escriba el nombre del nuevo descuento.")
                return
            
            # Agregar a la lista compartida si no existe
            global DESCUENTOS_SISTEMA
            if concepto not in DESCUENTOS_SISTEMA:
                DESCUENTOS_SISTEMA.append(concepto)
            
            # Actualizar todos los combobox registrados
            self._actualizar_todos_los_combos()
            
            messagebox.showinfo("Guardado", f"Se registró '{concepto}' en el sistema.\n\nAhora estará disponible en Descuentos y Deducciones.")
            top_level.destroy()

        ctk.CTkButton(botones, text="Guardar", width=120, command=guardar_otros, fg_color="#2f855a").pack(side="left", padx=(0, 12))
        ctk.CTkButton(botones, text="Cancelar", width=120, command=top_level.destroy, fg_color="#64748b").pack(side="left")



class Remuneraciones:

    
 
    def info_remuneraciones(self, frame):
        self.sipp_ref = frame

        frame_principal = ctk.CTkFrame(frame)
        frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        label_Remuneraciones = ctk.CTkLabel(frame_principal, text="Remuneraciones", font=ctk.CTkFont(size=28, weight="bold"))
        label_Remuneraciones.place(relx=0.5, y=28, anchor="n")

        ctk.CTkLabel(
            frame_principal,
            text="Gestiona estados de pago y registra los pagos recientes de personal.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8")
        ).place(relx=0.5, y=58, anchor="n")

        self.estado_db = db.estado_remuneracion()
        self.rem_db = db.registro_remuneracion()
        self.primer_frame(frame_principal)

        marco = ctk.CTkFrame(frame_principal, width=1, fg_color=("#cbd5e1", "#334155"))
        marco.place(relx=0.50, y=84, relheight=0.84, anchor="n")
        self.segundo_frame(frame_principal)
        self.actualizar_resumen_detalles()
        self.actualizar_tabla_remuneraciones()

    def _crear_estado_row(self, parent, nombre, y_rel, estado_actual, comando_editar, boton_color="#999fdb"):
        frame = ctk.CTkFrame(parent, height=50)
        frame.place(rely=y_rel, relwidth=1)

        label = ctk.CTkLabel(frame, text=f"{nombre}: ", font=ctk.CTkFont(size=20))
        label.place(relx=0.04, rely=0.2)

        combo = ctk.CTkComboBox(frame, values=["Activado", "Desactivado"], width=180)
        combo.place(relx=0.30, rely=0.2)
        if estado_actual is not None:
            combo.set(estado_actual[1])
        combo.configure(state="disable")

        boton = ctk.CTkButton(frame, text="Editar", font=ctk.CTkFont(size=20), width=100,
                               fg_color=boton_color, command=comando_editar)
        boton.place(relx=0.72, rely=0.2)

        return frame, combo

    def _crear_estado_editor_row(self, parent, nombre, y_rel, estado_actual, comando_guardar, boton_color="#2563eb"):
        frame = ctk.CTkFrame(parent, height=50)
        frame.place(rely=y_rel, relwidth=1)

        label = ctk.CTkLabel(frame, text=f"{nombre}: ", font=ctk.CTkFont(size=20))
        label.place(relx=0.04, rely=0.2)

        combo = ctk.CTkComboBox(frame, values=["Activado", "Desactivado"], width=180)
        combo.place(relx=0.30, rely=0.2)
        if estado_actual is not None:
            combo.set(estado_actual[1])

        boton = ctk.CTkButton(frame, text="Guardar", font=ctk.CTkFont(size=20), width=100,
                               fg_color=boton_color, command=comando_guardar)
        boton.place(relx=0.72, rely=0.2)

        return frame, combo


    def primer_frame(self, frame1):
        frame1 = ctk.CTkFrame(frame1, corner_radius=20,
                              fg_color=("#f8fafc", "#111827"), border_width=1,
                              border_color=("#cbd5e1", "#334155"))
        frame1.place(relx=0.00, y=84, relwidth=0.48, relheight=0.84)

        frame_salario = ctk.CTkFrame(frame1, height=56, fg_color="transparent")
        frame_salario.place(rely=0.08, relwidth=1)
        salario_activado = ["Activado"]
        salario_label = ctk.CTkLabel(frame_salario, text="Salario: ", font=ctk.CTkFont(size=20))
        salario_label.place(relx=0.04, rely=0.2)
        combo_salario = ctk.CTkComboBox(frame_salario, values=salario_activado)
        combo_salario.place(relx=0.30, rely=0.2)
        combo_salario.configure(state="disable")

        frame_vacaciones = ctk.CTkFrame(frame1, height=50)
        frame_vacaciones.place(rely=0.20, relwidth=1)
        vacaciones_activado = ["Activado"]
        vacaciones_label = ctk.CTkLabel(frame_vacaciones, text="Vacaciones: ", font=ctk.CTkFont(size=20))
        vacaciones_label.place(relx=0.04, rely=0.2)
        combo_vacaciones = ctk.CTkComboBox(frame_vacaciones, values=vacaciones_activado)
        combo_vacaciones.place(relx=0.30, rely=0.2)
        combo_vacaciones.configure(state="disable")

        estado_base = db.estado_remuneracion()
        estado_horas = estado_base.datos_nombre("Horas Extras")
        self.frame_horas, _ = self._crear_estado_row(
            frame1,
            "Horas Extras",
            0.32,
            estado_horas,
            lambda: self.actualizar_estado_horas(frame1)
        )

        estado_comisiones = estado_base.datos_nombre("Comisiones")
        self.frame_comisiones, _ = self._crear_estado_row(
            frame1,
            "Comisiones",
            0.44,
            estado_comisiones,
            lambda: self.actualizar_estado_comisiones(frame1)
        )

        estado_bonos = estado_base.datos_nombre("Bonificaciones")
        self.frame_bonos, _ = self._crear_estado_row(
            frame1,
            "Bonificaciones",
            0.56,
            estado_bonos,
            lambda: self.actualizar_estado_bonos(frame1)
        )

        estado_otros = estado_base.datos_nombre("Otros")
        self.frame_otros, _ = self._crear_estado_row(
            frame1,
            "Otros",
            0.68,
            estado_otros,
            lambda: self.actualizar_estado_otros(frame1),
            boton_color="#2563eb"
        )
        

    def actualizar_estado_horas(self, frame1):
        self.frame_horas.destroy()
        nombre = "Horas Extras"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_horas, combo = self._crear_estado_editor_row(
            frame1,
            nombre,
            0.32,
            estado_actual,
            lambda: self.guardar_estado_horas_extras(nombre, combo.get().title(), frame1)
        )

    def guardar_estado_horas_extras(self, nombre, estado, frames):
        datos = db.estado_remuneracion()
        if datos:
            datos.actualizar_estados_remuneracion_nombre(nombre, estado)
            messagebox.showinfo("Actualizado", f"{nombre}, Actualizado correctamente")
            self.reiniciar_hora(frames)
        else:
            messagebox.showerror("Error", f"{nombre}, No fue encontrado... ")

    def reiniciar_hora(self, frame):
        self.frame_horas.destroy()
        nombre = "Horas Extras"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_horas, _ = self._crear_estado_row(
            frame,
            nombre,
            0.32,
            estado_actual,
            lambda: self.actualizar_estado_horas(frame)
        )

    def actualizar_estado_comisiones(self, frame1):
        self.frame_comisiones.destroy()
        nombre = "Comisiones"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_comisiones, combo = self._crear_estado_editor_row(
            frame1,
            nombre,
            0.44,
            estado_actual,
            lambda: self.guardar_estado_comisiones(nombre, combo.get().title(), frame1)
        )

    def guardar_estado_comisiones(self, nombre, estado, frames):
        datos = db.estado_remuneracion()
        if datos:
            datos.actualizar_estados_remuneracion_nombre(nombre, estado)
            messagebox.showinfo("Actualizado", f"{nombre}, Actualizado correctamente")
            self.reiniciar_comisiones(frames)
        else:
            messagebox.showerror("Error", f"{nombre}, No fue encontrado... ")

    def reiniciar_comisiones(self, frame):
        self.frame_comisiones.destroy()
        nombre = "Comisiones"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_comisiones, _ = self._crear_estado_row(
            frame,
            nombre,
            0.44,
            estado_actual,
            lambda: self.actualizar_estado_comisiones(frame)
        )

    def actualizar_estado_bonos(self, frame1):
        self.frame_bonos.destroy()
        nombre = "Bonificaciones"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_bonos, combo = self._crear_estado_editor_row(
            frame1,
            nombre,
            0.56,
            estado_actual,
            lambda: self.guardar_estado_bonos(nombre, combo.get().title(), frame1)
        )

    def guardar_estado_bonos(self, nombre, estado, frames):
        datos = db.estado_remuneracion()
        if datos:
            datos.actualizar_estados_remuneracion_nombre(nombre, estado)
            messagebox.showinfo("Actualizado", f"{nombre}, Actualizado correctamente")
            self.reiniciar_bonos(frames)
        else:
            messagebox.showerror("Error", f"{nombre}, No fue encontrado... ")

    def reiniciar_bonos(self, frame):
        self.frame_bonos.destroy()
        nombre = "Bonificaciones"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_bonos, _ = self._crear_estado_row(
            frame,
            nombre,
            0.56,
            estado_actual,
            lambda: self.actualizar_estado_bonos(frame)
        )

    def actualizar_estado_otros(self, frame1):
        self.frame_otros.destroy()
        nombre = "Otros"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_otros, combo = self._crear_estado_editor_row(
            frame1,
            nombre,
            0.68,
            estado_actual,
            lambda: self.guardar_estado_otros(nombre, combo.get().title(), frame1)
        )

    def guardar_estado_otros(self, nombre, estado, frames):
        datos = db.estado_remuneracion()
        if datos:
            datos.actualizar_estados_remuneracion_nombre(nombre, estado)
            messagebox.showinfo("Actualizado", f"{nombre}, Actualizado correctamente")
            self.reiniciar_otros(frames)
        else:
            messagebox.showerror("Error", f"{nombre}, No fue encontrado... ")

    def reiniciar_otros(self, frame):
        self.frame_otros.destroy()
        nombre = "Otros"
        estado_actual = db.estado_remuneracion().datos_nombre(nombre)
        self.frame_otros, _ = self._crear_estado_row(
            frame,
            nombre,
            0.68,
            estado_actual,
            lambda: self.actualizar_estado_otros(frame)
        )

    def segundo_frame(self, frame2):
        frame2 = ctk.CTkFrame(frame2, corner_radius=20,
                              fg_color=("#f8fafc", "#0f172a"), border_width=1,
                              border_color=("#cbd5e1", "#334155"))
        frame2.place(relx=0.50, y=84, relwidth=0.48, relheight=0.84)

        self.entada_personal = ctk.CTkEntry(frame2, font=ctk.CTkFont(size=16), width=220, placeholder_text="Nombre o DUI", justify="center")
        self.entada_personal.place(relx=0.12, rely=0.02)
        boton_buscar = boton_buscar_profesional(frame2, command=lambda : self.personal_remunerar(frame2), width=96)
        boton_buscar.place(relx=0.60, rely=0.02)

        self.scrol_remunerar(frame2)

    def personal_remunerar(self, frame2):
        busqueda = self.entada_personal.get().strip()
        if not busqueda:
            messagebox.showwarning("Atención", "Ingrese un Nombre o DUI antes de buscar.")
            return

        # Buscar empleado en la BD
        empleado = db.gestion_empleado().buscar_empleado(busqueda)
        if empleado is None:
            messagebox.showerror("Error", f"No se encontró empleado con: {busqueda}")
            self.entada_personal.delete(0, "end")
            return

        # Extraer datos del empleado
        id_emp, nombre_emp, dui_emp = empleado[0], empleado[1], empleado[2]
        
        # Guardar datos del empleado para usar luego
        self.empleado_seleccionado = {
            'id': id_emp,
            'nombre': nombre_emp,
            'dui': dui_emp
        }

        # Destruir frame anterior si existe
        if hasattr(self, 'frame_remunerar') and self.frame_remunerar and self.frame_remunerar.winfo_exists():
            self.frame_remunerar.destroy()

        # Crear nuevo frame para el formulario
        self.frame_remunerar = ctk.CTkFrame(frame2, fg_color=("#ffffff", "#111827"), corner_radius=20,
                                           border_width=1, border_color=("#cbd5e1", "#334155"))
        self.frame_remunerar.place(relx=0.02, rely=0.08, relwidth=0.96, relheight=0.40)

        # Encabezado
        header = ctk.CTkFrame(self.frame_remunerar, fg_color="transparent")
        header.place(relx=0.03, rely=0.02, relwidth=0.94, relheight=0.20)
        
        ctk.CTkLabel(header, text="Registrar Remuneración", font=ctk.CTkFont(size=20, weight="bold"), text_color=("#0f172a", "#e2e8f0")).pack(anchor="w")
        ctk.CTkLabel(header, text=f"Empleado: {nombre_emp} | DUI: {dui_emp}", font=ctk.CTkFont(size=12), text_color=("#475569", "#94a3b8")).pack(anchor="w", pady=(4, 0))

        # Línea divisoria
        divider = ctk.CTkFrame(self.frame_remunerar, height=1, fg_color=("#cbd5e1", "#475569"))
        divider.place(relx=0.03, rely=0.22, relwidth=0.94)

        # Fila 1: Tipo de remuneración y Monto
        ctk.CTkLabel(self.frame_remunerar, text="Tipo de remuneración", font=ctk.CTkFont(size=13, weight="bold")).place(relx=0.03, rely=0.27)
        tipos_disponibles = self.tipos_remuneracion_disponibles()
        self.combo_tipos = ctk.CTkComboBox(self.frame_remunerar, values=tipos_disponibles, font=ctk.CTkFont(size=12), corner_radius=10)
        self.combo_tipos.place(relx=0.03, rely=0.36, relwidth=0.42)

        ctk.CTkLabel(self.frame_remunerar, text="Monto (B/.)", font=ctk.CTkFont(size=13, weight="bold")).place(relx=0.53, rely=0.27)
        self.entrada_monto = ctk.CTkEntry(self.frame_remunerar, placeholder_text="0.00", font=ctk.CTkFont(size=12), justify="center", corner_radius=10)
        self.entrada_monto.place(relx=0.53, rely=0.36, relwidth=0.42)

        ctk.CTkLabel(self.frame_remunerar, text="Fecha de aplicación", font=ctk.CTkFont(size=13, weight="bold")).place(relx=0.03, rely=0.46)
        self.fecha_remuneracion = DateEntry(self.frame_remunerar, width=18, date_pattern="dd/MM/yyyy", justify="center")
        self.fecha_remuneracion.place(relx=0.03, rely=0.55, relwidth=0.42)

        # Fila 2: Motivo
        ctk.CTkLabel(self.frame_remunerar, text="Motivo (opcional)", font=ctk.CTkFont(size=13, weight="bold")).place(relx=0.53, rely=0.46)
        self.entrada_motivo = ctk.CTkEntry(self.frame_remunerar, placeholder_text="Descripción breve", font=ctk.CTkFont(size=12), justify="center", corner_radius=10)
        self.entrada_motivo.place(relx=0.53, rely=0.55, relwidth=0.42)

        # Botones al final
        boton_guardar = ctk.CTkButton(self.frame_remunerar, text="Guardar", width=130, height=36,
                                      font=ctk.CTkFont(size=13, weight="bold"),
                                      fg_color="#2563eb", hover_color="#1d4ed8",
                                      command=lambda: self.guardar_remuneracion(frame2))
        boton_guardar.place(relx=0.55, rely=0.78)

        boton_cancelar = ctk.CTkButton(self.frame_remunerar, text="Cancelar", width=130, height=36,
                                       font=ctk.CTkFont(size=13, weight="bold"),
                                       fg_color="#64748b", hover_color="#475569",
                                       command=lambda: self.cancelar_remuneracion(frame2))
        boton_cancelar.place(relx=0.03, rely=0.78)

    def cancelar_remuneracion(self, frame2):
        """Cancela el formulario y vuelve al frame de resumen."""
        self.entada_personal.delete(0, "end")
        if hasattr(self, 'frame_remunerar') and self.frame_remunerar and self.frame_remunerar.winfo_exists():
            self.frame_remunerar.destroy()
        self.empleado_seleccionado = None
        self.remuneracion_editando_id = None
        self.actualizar_tabla_remuneraciones()

    def guardar_remuneracion(self, frame2):
        """Guarda la remuneración registrada."""
        if not hasattr(self, 'empleado_seleccionado') or not self.empleado_seleccionado:
            messagebox.showwarning("Atención", "Seleccione un empleado primero.")
            return

        nombre = self.empleado_seleccionado['nombre']
        dui = self.empleado_seleccionado['dui']
        tipo = self.combo_tipos.get().strip() if hasattr(self, 'combo_tipos') else ""
        monto_text = self.entrada_monto.get().strip() if hasattr(self, 'entrada_monto') else ""
        motivo = self.entrada_motivo.get().strip() if hasattr(self, 'entrada_motivo') else ""
        fecha_texto = self.fecha_remuneracion.get().strip() if hasattr(self, 'fecha_remuneracion') else ""

        # Validaciones
        if not tipo:
            messagebox.showwarning("Atención", "Seleccione el tipo de remuneración.")
            return

        if not self.tipo_remuneracion_activa(tipo):
            messagebox.showwarning("Atención", f"La remuneración '{tipo}' está desactivada y no puede registrarse.")
            return

        if not monto_text:
            messagebox.showwarning("Atención", "Ingrese el monto de la remuneración.")
            return

        try:
            monto = float(monto_text.replace(',', '.'))
            if monto <= 0:
                raise ValueError("Monto debe ser mayor a 0")
        except ValueError:
            messagebox.showwarning("Atención", "Ingrese un monto válido mayor que 0.")
            return

        try:
            fecha = datetime.strptime(fecha_texto, "%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("Atención", "Seleccione una fecha válida.")
            return
        id_edicion = getattr(self, "remuneracion_editando_id", None)
        if self.rem_db.existe_remuneracion(dui, tipo, monto, fecha, excluir_id=id_edicion):
            messagebox.showwarning("Duplicado", "Ya existe una remuneración activa igual para este empleado y fecha.")
            return

        # Guardar en BD
        if id_edicion is not None:
            guardado = self.rem_db.actualizar_remuneracion(id_edicion, nombre, dui, tipo, monto, fecha, motivo)
            if not guardado:
                messagebox.showerror("Error", "No se pudo actualizar la remuneración.")
                return
        else:
            self.rem_db.insertar_remuneracion(nombre, tipo, monto, fecha, motivo, dui)
        self.remuneracion_editando_id = None
        
        # Limpiar formulario
        self.combo_tipos.set("")
        self.entrada_monto.delete(0, "end")
        self.entrada_motivo.delete(0, "end")
        self.entada_personal.delete(0, "end")
        
        # Mensaje de éxito
        messagebox.showinfo("Guardado", f"✓ Remuneración registrada\n\nEmpleado: {nombre}\nDUI: {dui}\nTipo: {tipo}\nMonto: B/. {monto:,.2f}")
        
        # Actualizar interfaz
        self.actualizar_resumen_detalles()
        self.actualizar_tabla_remuneraciones()
        
        # Ocultar formulario
        if hasattr(self, 'frame_remunerar') and self.frame_remunerar and self.frame_remunerar.winfo_exists():
            self.frame_remunerar.destroy()


    def scrol_remunerar(self, frame2):
        self.resumen_frame = ctk.CTkFrame(frame2, fg_color=("#ffffff", "#111827"), corner_radius=20,
                                          border_width=1, border_color=("#cbd5e1", "#334155"))
        self.resumen_frame.place(relx=0.02, rely=0.08, relwidth=0.96, relheight=0.18)
        self.crear_panel_resumen()

        self.frame_remunerar = None

        self.lista_remuneraciones = ctk.CTkScrollableFrame(frame2, corner_radius=20, fg_color=("#f8fafc", "#0f172a"))
        self.lista_remuneraciones.place(relx=0.02, rely=0.28, relwidth=0.96, relheight=0.70)

        header_frame = ctk.CTkFrame(self.lista_remuneraciones, height=26, corner_radius=16, fg_color=("#ffffff", "#0f172a"), border_width=1,
                                    border_color=("#cbd5e1", "#475569"))
        header_frame.pack(fill='x', padx=0, pady=(0, 0))
        header_frame.pack_propagate(False)

        columnas = ["Colaborador", "Concepto", "Importe", "Fecha", "Acciones"]
        for index, encabezado in enumerate(columnas):
            label = ctk.CTkLabel(header_frame, text=encabezado, font=ctk.CTkFont(size=11, weight="bold"), text_color=("#0f172a", "#e2e8f0"))
            label.pack(side='left', fill='x', expand=True, padx=(1, 1), pady=0)
            if index < len(columnas) - 1:
                linea = ctk.CTkFrame(header_frame, width=1, fg_color=("#cbd5e1", "#475569"))
                linea.pack(side="left", pady=0)

        self.tabla_contenido = ctk.CTkFrame(self.lista_remuneraciones, fg_color="transparent")
        self.tabla_contenido.pack(fill='both', expand=True, padx=0, pady=(0, 0))

    def crear_panel_resumen(self):
        self.resumen_frame.grid_rowconfigure(0, weight=1)
        self.resumen_frame.grid_columnconfigure(0, weight=1)
        self.resumen_frame.grid_columnconfigure(1, weight=1)

        tarjeta1 = ctk.CTkFrame(self.resumen_frame, fg_color=("#ffffff", "#0f172a"), corner_radius=18,
                                 border_width=1, border_color=("#cbd5e1", "#475569"))
        tarjeta1.grid(row=0, column=0, padx=(12, 6), pady=12, sticky='nsew')
        ctk.CTkLabel(tarjeta1, text="Total movimientos", font=ctk.CTkFont(size=13), text_color=("#475569", "#94a3b8")).pack(anchor='w', padx=14, pady=(14, 4))
        self.valor_total_registros = ctk.CTkLabel(tarjeta1, text="0 registros", font=ctk.CTkFont(size=22, weight="bold"), text_color=("#0f172a", "#e2e8f0"))
        self.valor_total_registros.pack(anchor='w', padx=14, pady=(0, 14))

        tarjeta2 = ctk.CTkFrame(self.resumen_frame, fg_color=("#ffffff", "#0f172a"), corner_radius=18,
                                 border_width=1, border_color=("#cbd5e1", "#475569"))
        tarjeta2.grid(row=0, column=1, padx=(6, 12), pady=12, sticky='nsew')
        ctk.CTkLabel(tarjeta2, text="Importe total", font=ctk.CTkFont(size=13), text_color=("#475569", "#94a3b8")).pack(anchor='w', padx=14, pady=(14, 4))
        self.valor_total_importe = ctk.CTkLabel(tarjeta2, text="B/. 0.00", font=ctk.CTkFont(size=22, weight="bold"), text_color=("#0f172a", "#e2e8f0"))
        self.valor_total_importe.pack(anchor='w', padx=14, pady=(0, 14))

    def actualizar_resumen_detalles(self):
        cantidad = self.rem_db.cantidad_remuneraciones()
        total = self.rem_db.total_importe()
        self.valor_total_registros.configure(text=f"{cantidad} registro{'s' if cantidad != 1 else ''}")
        self.valor_total_importe.configure(text=f"B/. {total:,.2f}")

    def tipo_remuneracion_activa(self, tipo):
        estado = self.estado_db.datos_nombre(tipo)
        if estado is None:
            return True
        return estado[1].strip().lower() != "desactivado"

    def tipos_remuneracion_disponibles(self):
        opciones = ["Salario", "Vacaciones", "Horas Extras", "Comisiones", "Bonificaciones", "Otros"]
        return [tipo for tipo in opciones if tipo in ("Salario", "Vacaciones") or self.tipo_remuneracion_activa(tipo)]

    def actualizar_tabla_remuneraciones(self):
        for widget in self.tabla_contenido.winfo_children():
            widget.destroy()

        remuneraciones = self.rem_db.obtener_remuneraciones(limit=10)
        if not remuneraciones:
            sin_registros = ctk.CTkLabel(self.tabla_contenido, text="Aún no hay remuneraciones registradas.", font=ctk.CTkFont(size=14), text_color=("#475569", "#94a3b8"))
            sin_registros.pack(pady=20)
            return

        for idx, registro in enumerate(remuneraciones):
            id_remuneracion, nombre, tipo, monto, fecha, motivo = registro
            background = ("#ffffff", "#111827") if idx % 2 == 0 else ("#f3f4f6", "#1f2937")
            row_frame = ctk.CTkFrame(self.tabla_contenido, height=26, fg_color=background, corner_radius=14)
            row_frame.pack(fill='x', padx=0, pady=0)
            row_frame.pack_propagate(False)

            valores = [nombre, tipo, f"B/. {monto:,.2f}", fecha.strftime('%d/%m/%Y') if hasattr(fecha, 'strftime') else str(fecha)]
            for index, valor in enumerate(valores):
                label = ctk.CTkLabel(row_frame, text=valor, font=ctk.CTkFont(size=11), text_color=("#0f172a", "#e2e8f0"))
                label.pack(side='left', fill='x', expand=True, padx=(1, 1), pady=0)
                if index < len(valores) - 1:
                    divider = ctk.CTkFrame(row_frame, width=1, fg_color=("#cbd5e1", "#475569"))
                    divider.pack(side="left", pady=0)
            acciones = ctk.CTkFrame(row_frame, fg_color="transparent")
            acciones.pack(side="left", fill="x", expand=True, padx=2, pady=0)
            ctk.CTkButton(acciones, text="Editar", width=58, height=22, font=ctk.CTkFont(size=10), fg_color="#2563eb", hover_color="#1d4ed8", command=lambda id_r=id_remuneracion: self.editar_remuneracion(id_r)).pack(side="left", padx=2)
            ctk.CTkButton(acciones, text="Anular", width=58, height=22, font=ctk.CTkFont(size=10), fg_color="#dc2626", hover_color="#b91c1c", command=lambda id_r=id_remuneracion: self.anular_remuneracion(id_r)).pack(side="left", padx=2)

    def anular_remuneracion(self, id_remuneracion):
        if not messagebox.askyesno("Confirmar anulación", "¿Desea anular esta remuneración?", parent=self.sipp_ref if hasattr(self, "sipp_ref") else None):
            return
        if self.rem_db.anular_remuneracion(id_remuneracion):
            messagebox.showinfo("Remuneración", "La remuneración fue anulada.")
            self.actualizar_resumen_detalles()
            self.actualizar_tabla_remuneraciones()
        else:
            messagebox.showerror("Remuneración", "No se pudo anular la remuneración.")

    def editar_remuneracion(self, id_remuneracion):
        registro = next((fila for fila in self.rem_db.obtener_remuneraciones() if fila[0] == id_remuneracion), None)
        if not registro:
            messagebox.showwarning("Remuneración", "El registro ya no está disponible.")
            return
        _, nombre, tipo, monto, fecha, motivo = registro
        empleado = db.gestion_empleado().buscar_empleado(nombre)
        if not empleado:
            messagebox.showwarning("Remuneración", "No se encontró el empleado asociado.")
            return
        self.empleado_seleccionado = {"id": empleado[0], "nombre": empleado[1], "dui": empleado[2]}
        self.personal_remunerar(self.sipp_ref)
        self.combo_tipos.set(tipo)
        self.entrada_monto.delete(0, "end")
        self.entrada_monto.insert(0, str(monto))
        self.fecha_remuneracion.delete(0, "end")
        self.fecha_remuneracion.insert(0, fecha.strftime("%d/%m/%Y") if hasattr(fecha, "strftime") else str(fecha))
        self.entrada_motivo.delete(0, "end")
        self.entrada_motivo.insert(0, motivo or "")
        self.remuneracion_editando_id = id_remuneracion

    def limpiar(self):
        """Limpia todos los widgets y referencias de la clase Remuneraciones."""
        try:
            # Destruir widgets principales
            widgets_a_limpiar = [
                'frame_remunerar', 'frame_horas', 'frame_comisiones', 
                'frame_bonos', 'frame_otros', 'lista_remuneraciones',
                'resumen_frame', 'tabla_contenido', 'entada_personal'
            ]
            
            for widget_name in widgets_a_limpiar:
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    if widget and hasattr(widget, 'winfo_exists'):
                        try:
                            if widget.winfo_exists():
                                widget.destroy()
                        except Exception:
                            pass
                    delattr(self, widget_name)
            
            # Limpiar referencias a objetos de base de datos
            if hasattr(self, 'estado_db'):
                delattr(self, 'estado_db')
            if hasattr(self, 'rem_db'):
                delattr(self, 'rem_db')
            if hasattr(self, 'sipp_ref'):
                delattr(self, 'sipp_ref')
                
        except Exception:
            pass


class Permisos_Ausencias:
    def __init__(self):
        self.frame_principal = None
        self.left_panel = None
        self.right_panel = None
        self.form_frame = None
        self.report_scroll = None
        self.report_rows_container = None
        self.search_entry = None
        self.report_search = None
        self.selected_employee = None
        self.nombre_label = None
        self.identificacion_label = None
        self.texto = None
        self.combobox_permisos = None
        self.combobox_pagado = None
        self.boton_limpiar_P_A = None
        self.boton_guardar_P_A = None
        self.boton_archivo = None
        self.adjunto_ruta = None
        self.right_search_entry = None
        self._empleados_derecha = []
        self.right_selected_record = None
        self.right_record_status = None
        self.right_records_scroll = None
        self.right_details_text = None
        self.fecha_inicio_entry = None
        self.fecha_fin_entry = None

    def info_Per_Aus(self, frame):
        if self.frame_principal is not None:
            self.frame_principal.destroy()

        self.frame_principal = ctk.CTkFrame(frame)
        self.frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        title = ctk.CTkLabel(
            self.frame_principal,
            text="Permisos y Ausencias",
            font=ctk.CTkFont(size=25, weight="bold")
        )
        title.place(relx=0.5, y=20, anchor="n")

        self.left_panel = ctk.CTkFrame(self.frame_principal, corner_radius=18, fg_color=("#f8fafc", "#111827"))
        self.left_panel.place(relx=0.02, y=60, relwidth=0.46, relheight=0.92)

        ctk.CTkFrame(self.frame_principal, width=2, fg_color="#d62828").place(relx=0.50, y=60, relheight=0.92, anchor="n")

        self.right_panel = ctk.CTkFrame(self.frame_principal, corner_radius=18, fg_color=("#ffffff", "#0f172a"))
        self.right_panel.place(relx=0.52, y=60, relwidth=0.46, relheight=0.92)

        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self):
        top_bar = ctk.CTkFrame(self.left_panel, fg_color="transparent", height=110)
        top_bar.place(relx=0.0, rely=0.0, relwidth=1)

        etiqueta_buscar = ctk.CTkLabel(
            top_bar,
            text="Buscar empleado",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        etiqueta_buscar.place(relx=0.02, y=16, anchor="w")

        etiqueta_subtitulo = ctk.CTkLabel(
            top_bar,
            text="Encuentra personal por nombre, documento o reporte asociado.",
            font=ctk.CTkFont(size=12),
            text_color=("#475569", "#cbd5e1")
        )
        etiqueta_subtitulo.place(relx=0.02, y=44, anchor="w")

        search_container = ctk.CTkFrame(
            top_bar,
            corner_radius=18,
            fg_color=("#f8fafc", "#111827"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            height=50
        )
        search_container.place(relx=0.02, y=80, relwidth=0.96, anchor="w")

        icono_buscar = ctk.CTkLabel(
            search_container,
            text="🔎",
            font=ctk.CTkFont(size=16)
        )
        icono_buscar.place(relx=0.03, rely=0.5, anchor="w")

        self.search_entry = ctk.CTkEntry(
            search_container,
            width=260,
            placeholder_text="Buscar empleado...",
            border_width=0,
            fg_color=("#ffffff", "#0f172a"),
            text_color=("#0f172a", "#e2e8f0"),
            font=ctk.CTkFont(size=14),
            justify="left"
        )
        self.search_entry.place(relx=0.10, rely=0.5, anchor="w")

        boton_buscar = ctk.CTkButton(
            search_container,
            text="Buscar",
            width=90,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            corner_radius=16,
            command=self._buscar_empleado
        )
        boton_buscar.place(relx=0.97, rely=0.5, anchor="e")

        self.form_frame = ctk.CTkFrame(self.left_panel, corner_radius=18, fg_color=("#ffffff", "#0f172a"))
        self.form_frame.place(relx=0.02, rely=0.18, relwidth=0.96, relheight=0.78)

        label_nombre = ctk.CTkLabel(self.form_frame, text="Nombre Personal", font=ctk.CTkFont(size=14, weight="bold"))
        label_nombre.place(relx=0.02, rely=0.03)
        self.nombre_label = ctk.CTkLabel(self.form_frame, text="--", anchor="w", font=ctk.CTkFont(size=14))
        self.nombre_label.place(relx=0.02, rely=0.12, relwidth=0.44)

        label_identificacion = ctk.CTkLabel(self.form_frame, text="Número de identificación", font=ctk.CTkFont(size=14, weight="bold"))
        label_identificacion.place(relx=0.52, rely=0.03)
        self.identificacion_label = ctk.CTkLabel(self.form_frame, text="--", anchor="w", font=ctk.CTkFont(size=14))
        self.identificacion_label.place(relx=0.52, rely=0.12, relwidth=0.44)

        label_tipo = ctk.CTkLabel(self.form_frame, text="Tipo", font=ctk.CTkFont(size=14, weight="bold"))
        label_tipo.place(relx=0.02, rely=0.24)
        self.combobox_permisos = ctk.CTkComboBox(
            self.form_frame,
            values=["Permiso", "Ausencia"],
            state="disabled"
        )
        self.combobox_permisos.place(relx=0.02, rely=0.31, relwidth=0.30)

        label_pagado = ctk.CTkLabel(self.form_frame, text="Pagado", font=ctk.CTkFont(size=14, weight="bold"))
        label_pagado.place(relx=0.34, rely=0.24)
        self.combobox_pagado = ctk.CTkComboBox(
            self.form_frame,
            values=["No Pagado", "Pagado"],
            state="disabled"
        )
        self.combobox_pagado.place(relx=0.34, rely=0.31, relwidth=0.30)

        self.boton_archivo = ctk.CTkButton(
            self.form_frame,
            text="Adjuntar archivo",
            fg_color="#ff8c00",
            hover_color="#d97706",
            state="disabled",
            command=self._seleccionar_archivo
        )
        self.boton_archivo.place(relx=0.68, rely=0.31, relwidth=0.30)
        
        label_fecha_inicio = ctk.CTkLabel(self.form_frame, text="Fecha Inicio", font=ctk.CTkFont(size=14, weight="bold"))
        label_fecha_inicio.place(relx=0.02, rely=0.41)
        self.fecha_inicio_entry = DateEntry(self.form_frame)
        self.fecha_inicio_entry.place(relx=0.02, rely=0.48, relwidth=0.30)
        
        label_fecha_fin = ctk.CTkLabel(self.form_frame, text="Fecha Fin", font=ctk.CTkFont(size=14, weight="bold"))
        label_fecha_fin.place(relx=0.34, rely=0.41)
        self.fecha_fin_entry = DateEntry(self.form_frame)
        self.fecha_fin_entry.place(relx=0.34, rely=0.48, relwidth=0.30)
        
        self.texto = ctk.CTkTextbox(self.form_frame, font=ctk.CTkFont(size=14), state="disabled")
        self.texto.insert("0.0", "Describa el motivo o reporte aquí...")
        self.texto.place(relx=0.02, rely=0.60, relwidth=0.96, relheight=0.25)

        footer = ctk.CTkFrame(self.form_frame, fg_color="transparent", height=70)
        footer.place(relx=0.0, rely=0.86, relwidth=1)

        self.boton_limpiar_P_A = ctk.CTkButton(
            footer,
            text="Limpiar",
            width=120,
            fg_color="#d62828",
            hover_color="#b91c1c",
            state="disabled",
            command=self._limpiar_form
        )
        self.boton_limpiar_P_A.place(relx=0.02, rely=0.5, anchor="w")

        self.boton_guardar_P_A = ctk.CTkButton(
            footer,
            text="Grabar",
            width=120,
            fg_color="#2fe81e",
            hover_color="#22c55e",
            state="disabled",
            command=self._guardar_permiso
        )
        self.boton_guardar_P_A.place(relx=0.98, rely=0.5, anchor="e")


    def _build_right_panel(self):
        top_bar = ctk.CTkFrame(self.right_panel, fg_color="transparent", height=110)
        top_bar.place(relx=0.0, rely=0.0, relwidth=1)

        etiqueta_titulo = ctk.CTkLabel(
            top_bar,
            text="Registro profesional",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        etiqueta_titulo.place(relx=0.02, rely=0.18, anchor="w")

        etiqueta_subtitulo = ctk.CTkLabel(
            top_bar,
            text="Visualiza ausencias y permisos guardados en un formato profesional.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#cbd5e1")
        )
        etiqueta_subtitulo.place(relx=0.02, rely=0.55, anchor="w")

        search_container = ctk.CTkFrame(
            self.right_panel,
            corner_radius=18,
            fg_color=("#f8fafc", "#111827"),
            border_width=1,
            border_color=("#cbd5e1", "#334155"),
            height=48
        )
        search_container.place(relx=0.02, rely=0.18, relwidth=0.96)

        icono_buscar_derecha = ctk.CTkLabel(
            search_container,
            text="🔎",
            font=ctk.CTkFont(size=16)
        )
        icono_buscar_derecha.place(relx=0.03, rely=0.5, anchor="w")

        self.right_search_entry = ctk.CTkEntry(
            search_container,
            placeholder_text="Buscar por nombre o DUI...",
            border_width=0,
            fg_color=("#ffffff", "#0f172a"),
            text_color=("#0f172a", "#e2e8f0"),
            font=ctk.CTkFont(size=14),
            justify="left"
        )
        # Entry más ancho para permitir que el botón quede sobrepuesto dentro del campo
        self.right_search_entry.place(relx=0.12, rely=0.5, relwidth=0.72, anchor="w")

        boton_buscar_derecha = ctk.CTkButton(
            search_container,
            text="Buscar",
            width=88,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            corner_radius=12,
            command=self._buscar_empleado_derecha
        )
        # Posicionar el botón ligeramente dentro del área del Entry para apariencia integrada
        boton_buscar_derecha.place(relx=0.80, rely=0.5, anchor="w")

        list_card = ctk.CTkFrame(
            self.right_panel,
            fg_color=("#ffffff", "#0f172a"),
            corner_radius=18,
            border_width=1,
            border_color=("#cbd5e1", "#475569")
        )
        list_card.place(relx=0.02, rely=0.28, relwidth=0.96, relheight=0.52)

        label_registros = ctk.CTkLabel(
            list_card,
            text="Registros guardados",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label_registros.place(relx=0.03, rely=0.05, anchor="w")

        self.right_record_status = ctk.CTkLabel(
            list_card,
            text="Cargando registros...",
            font=ctk.CTkFont(size=12),
            text_color=("#475569", "#cbd5e1")
        )
        self.right_record_status.place(relx=0.03, rely=0.12, anchor="w")

        header_row = ctk.CTkFrame(list_card, fg_color="transparent", height=32)
        header_row.place(relx=0.0, rely=0.18, relwidth=1)

        headers = [("Nombre", 0.40), ("Fecha", 0.30), ("Ver Detalles", 0.30)]
        current_x = 0.0
        for texto, relwidth in headers:
            etiqueta = ctk.CTkLabel(
                header_row,
                text=texto,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#0f172a", "#e2e8f0")
            )
            etiqueta.place(relx=current_x, rely=0.5, relheight=1, relwidth=relwidth, anchor="w")
            if texto != "Ver Detalles":
                separator = ctk.CTkFrame(header_row, width=1, fg_color=("#cbd5e1", "#475569"))
                separator.place(relx=current_x + relwidth - 0.005, rely=0.1, relheight=0.8)
            current_x += relwidth

        self.right_records_scroll = ctk.CTkScrollableFrame(
            list_card,
            fg_color="transparent"
        )
        self.right_records_scroll.place(relx=0.0, rely=0.28, relwidth=1, relheight=0.70)

        details_card = ctk.CTkFrame(
            self.right_panel,
            fg_color=("#f8fafc", "#111827"),
            corner_radius=18,
            border_width=1,
            border_color=("#cbd5e1", "#475569")
        )
        details_card.place(relx=0.02, rely=0.82, relwidth=0.96, relheight=0.16)

        details_title = ctk.CTkLabel(
            details_card,
            text="Detalles del registro",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        details_title.place(relx=0.03, rely=0.08, anchor="w")

        self.right_details_text = ctk.CTkLabel(
            details_card,
            text="Selecciona un registro para ver información detallada.",
            font=ctk.CTkFont(size=12),
            text_color=("#475569", "#cbd5e1"),
            justify="left",
            anchor="nw"
        )
        self.right_details_text.place(relx=0.03, rely=0.28, relwidth=0.94, relheight=0.64)

        self._cargar_empleados_derecha()

    def _mostrar_registros_derecha(self):
        if self.right_records_scroll is None:
            return

        for child in self.right_records_scroll.winfo_children():
            child.destroy()

        if not self._empleados_derecha:
            mensaje = ctk.CTkLabel(
                self.right_records_scroll,
                text="No hay registros guardados. Usa la búsqueda para filtrar por nombre o DUI.",
                font=ctk.CTkFont(size=13),
                text_color=("#475569", "#cbd5e1"),
                wraplength=420,
                justify="left"
            )
            mensaje.pack(anchor="w", padx=16, pady=14)
            return

        for registro in self._empleados_derecha:
            self._crear_fila_registro_derecha(self.right_records_scroll, registro)

    def _crear_fila_registro_derecha(self, parent, registro):
        item_frame = ctk.CTkFrame(parent, fg_color=("#ffffff", "#0f172a"), corner_radius=14, border_width=1, border_color=("#cbd5e1", "#475569"))
        item_frame.pack(fill='x', padx=10, pady=4, ipady=10)

        nombre = registro[1] or "--"
        fecha = self._formatear_fecha(registro[5]) or "--"

        label_nombre = ctk.CTkLabel(
            item_frame,
            text=nombre,
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left"
        )
        label_nombre.place(relx=0.02, rely=0.5, relwidth=0.40, anchor="w")

        label_fecha = ctk.CTkLabel(
            item_frame,
            text=fecha,
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left"
        )
        label_fecha.place(relx=0.42, rely=0.5, relwidth=0.30, anchor="w")

        boton_detalles = ctk.CTkButton(
            item_frame,
            text="Ver Detalles",
            width=100,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            corner_radius=14,
            command=lambda registro=registro: self._mostrar_detalles_empleado(registro)
        )
        boton_detalles.place(relx=0.95, rely=0.5, anchor="e")

    def _buscar_empleado_derecha(self):
        texto_busqueda = self.right_search_entry.get().strip()
        if not texto_busqueda:
            messagebox.showwarning("Validación", "Ingrese un nombre o DUI para buscar.")
            return

        registros = db.Permisos_y_ausencias().obtener_permisos_ausencias()
        texto_normalizado = texto_busqueda.lower()
        filtrados = [registro for registro in registros if texto_normalizado in (registro[1] or "").lower() or texto_normalizado in (registro[2] or "").lower()]
        self._empleados_derecha = filtrados
        self.right_record_status.configure(text=f"{len(filtrados)} registros encontrados")
        self._mostrar_registros_derecha()

        if not filtrados:
            messagebox.showinfo("Registro no encontrado", "No se encontraron ausencias o permisos que coincidan con la búsqueda.")

    def _cargar_empleados_derecha(self):
        registros = db.Permisos_y_ausencias().obtener_permisos_ausencias()
        self._empleados_derecha = registros or []
        self.right_record_status.configure(text=f"{len(self._empleados_derecha)} registros guardados")
        self._mostrar_registros_derecha()

    def _seleccionar_registro_derecha(self, registro):
        self.right_selected_record = registro
        self._actualizar_panel_derecho()

    def _actualizar_panel_derecho(self):
        if not self.right_selected_record:
            self.right_details_text.configure(text="Selecciona un registro para ver información detallada.")
            self.right_record_status.configure(text=f"{len(self._empleados_derecha)} registros guardados")
            return

        registro = self.right_selected_record
        nombre = registro[1] or "--"
        dui = registro[2] or "--"
        tipo = registro[3] or "--"
        pagado = registro[4] or "--"
        fecha_inicio = self._formatear_fecha(registro[5])
        fecha_fin = self._formatear_fecha(registro[6])
        motivo = registro[7] or "--"

        self.right_details_text.configure(
            text=(
                f"Nombre: {nombre}\n"
                f"DUI: {dui}\n"
                f"Tipo: {tipo}\n"
                f"Pagado: {pagado}\n"
                f"Período: {fecha_inicio} - {fecha_fin}\n"
                f"Motivo: {motivo}"
            )
        )
        self.right_record_status.configure(text=f"Registro seleccionado: {nombre} ({tipo})")

    def _mostrar_detalles_empleado(self, registro=None):
        if registro is None:
            registro = self.right_selected_record

        if not registro:
            messagebox.showinfo("Detalles", "Selecciona primero un registro en el panel derecho.")
            return

        self.right_selected_record = registro
        self._actualizar_panel_derecho()
        registro_id = registro[0]

        detalles_top = ctk.CTkToplevel(self.right_panel)
        detalles_top.title("Detalles del Permiso/Ausencia")
        ajustar_toplevel_a_pantalla(detalles_top, 600, 500, margen_x=120, margen_y=120, min_ancho=430, min_alto=340)
        detalles_top.resizable(False, False)
        detalles_top.grab_set()

        # Header con título
        header = ctk.CTkLabel(
            detalles_top,
            text="Detalles del Permiso/Ausencia",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        header.pack(pady=(15, 10))

        # Frame para pestañas
        tabview = ctk.CTkTabview(detalles_top)
        tabview.pack(fill="both", expand=True, padx=15, pady=10)

        # Pestaña 1: Detalles
        tabview.add("Detalles")
        panel_detalles = tabview.tab("Detalles")

        campos = [
            ("Nombre", registro[1] or "--"),
            ("DUI", registro[2] or "--"),
            ("Tipo", registro[3] or "--"),
            ("Pagado", registro[4] or "--"),
            ("Fecha inicio", self._formatear_fecha(registro[5])),
            ("Fecha fin", self._formatear_fecha(registro[6])),
            ("Motivo", registro[7] or "--"),
            ("Creado en", registro[10] or "--")
        ]

        for index, (etiqueta_texto, valor) in enumerate(campos):
            label_titulo = ctk.CTkLabel(
                panel_detalles,
                text=etiqueta_texto,
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w"
            )
            label_titulo.place(relx=0.05, rely=0.08 + index * 0.10, relwidth=0.35)

            label_valor = ctk.CTkLabel(
                panel_detalles,
                text=valor,
                font=ctk.CTkFont(size=12),
                anchor="w",
                wraplength=300,
                justify="left"
            )
            label_valor.place(relx=0.40, rely=0.08 + index * 0.10, relwidth=0.55)

        # Pestaña 2: Archivos
        tabview.add("Archivos")
        panel_archivos = tabview.tab("Archivos")

        if registro[8]:  # imagen_nombre
            archivo_nombre = registro[8]
            archivo_path = os.path.join(os.path.abspath('.'), 'uploads', archivo_nombre)

            archivo_info = ctk.CTkLabel(
                panel_archivos,
                text=f"📎 Archivo: {archivo_nombre}",
                font=ctk.CTkFont(size=13, weight="bold"),
                anchor="w"
            )
            archivo_info.pack(pady=(20, 10), padx=20, anchor="w")

            botones_archivo = ctk.CTkFrame(panel_archivos, fg_color="transparent")
            botones_archivo.pack(fill="x", padx=20, pady=(0, 10))

            boton_ver = ctk.CTkButton(
                botones_archivo,
                text="Ver archivo",
                width=140,
                fg_color="#2563eb",
                hover_color="#1d4ed8",
                command=lambda: self._abrir_archivo_adjuntado(archivo_path, archivo_nombre)
            )
            boton_ver.pack(side="left", padx=(0, 10))

            boton_descargar = ctk.CTkButton(
                botones_archivo,
                text="📥 Descargar",
                width=140,
                fg_color="#0ea5e9",
                hover_color="#0284c7",
                command=lambda: self._descargar_archivo(registro)
            )
            boton_descargar.pack(side="left")

            preview_frame = ctk.CTkFrame(
                panel_archivos,
                fg_color=("#f8fafc", "#0f172a"),
                corner_radius=16,
                border_width=1,
                border_color=("#cbd5e1", "#475569")
            )
            preview_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

            extension = os.path.splitext(archivo_nombre)[1].lower()
            if extension in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}:
                try:
                    with Image.open(archivo_path) as imagen:
                        imagen.thumbnail((420, 260))
                        imagen_ctk = ctk.CTkImage(
                            light_image=imagen,
                            dark_image=imagen,
                            size=(imagen.width, imagen.height)
                        )
                        label_imagen = ctk.CTkLabel(preview_frame, image=imagen_ctk, text="")
                        label_imagen.image = imagen_ctk
                        label_imagen.pack(pady=15)
                except Exception:
                    fallback = ctk.CTkLabel(
                        preview_frame,
                        text="No se pudo previsualizar la imagen.\nUse 'Ver archivo' para abrirla con el visor del sistema.",
                        font=ctk.CTkFont(size=12),
                        justify="center",
                        wraplength=300
                    )
                    fallback.pack(pady=30)
            elif extension == ".pdf":
                icon = ctk.CTkLabel(preview_frame, text="📄", font=ctk.CTkFont(size=54))
                icon.pack(pady=(25, 10))
                texto_pdf = ctk.CTkLabel(
                    preview_frame,
                    text="Archivo PDF adjunto\nEl documento se abrirá con el visor predeterminado del sistema.",
                    font=ctk.CTkFont(size=12),
                    justify="center",
                    wraplength=300
                )
                texto_pdf.pack(pady=(0, 15))
            else:
                texto_generico = ctk.CTkLabel(
                    preview_frame,
                    text="Archivo adjunto disponible.\nUse 'Ver archivo' para abrirlo.",
                    font=ctk.CTkFont(size=12),
                    justify="center",
                    wraplength=300
                )
                texto_generico.pack(pady=30)
        else:
            sin_archivo = ctk.CTkLabel(
                panel_archivos,
                text="📁 No hay archivos adjuntos",
                font=ctk.CTkFont(size=13),
                text_color="#64748b"
            )
            sin_archivo.pack(pady=50)

        # Frame de botones inferiores
        botones_frame = ctk.CTkFrame(detalles_top, fg_color="transparent")
        botones_frame.pack(fill="x", padx=15, pady=15)

        boton_editar = ctk.CTkButton(
            botones_frame,
            text="✏️ Editar",
            width=120,
            fg_color="#16a34a",
            hover_color="#15803d",
            corner_radius=12,
            command=lambda: self._editar_registro(registro, detalles_top)
        )
        boton_editar.pack(side="left", padx=5)

        boton_eliminar = ctk.CTkButton(
            botones_frame,
            text="🗑️ Eliminar",
            width=120,
            fg_color="#ef4444",
            hover_color="#dc2626",
            corner_radius=12,
            command=lambda: self._eliminar_y_cerrar(registro_id, detalles_top)
        )
        boton_eliminar.pack(side="left", padx=5)

        boton_cerrar = ctk.CTkButton(
            botones_frame,
            text="Cerrar",
            width=120,
            fg_color="#64748b",
            hover_color="#475569",
            corner_radius=12,
            command=detalles_top.destroy
        )
        boton_cerrar.pack(side="right", padx=5)

    def _limpiar_panel_derecho(self):
        self.right_selected_record = None
        self.right_search_entry.delete(0, "end")
        self._actualizar_panel_derecho()

    def _descargar_archivo(self, registro):
        """Descarga el archivo adjunto del registro"""
        imagen_nombre = registro[8]
        if not imagen_nombre:
            messagebox.showwarning("Error", "No hay archivo para descargar.")
            return
        
        uploads_dir = os.path.join(os.path.abspath('.'), 'uploads')
        archivo_path = os.path.join(uploads_dir, imagen_nombre)
        
        if not os.path.exists(archivo_path):
            messagebox.showerror("Error", f"El archivo no se encuentra: {archivo_path}")
            return
        
        try:
            destino = filedialog.asksaveasfilename(
                defaultextension=os.path.splitext(imagen_nombre)[1],
                filetypes=[("Todos", "*.*")],
                initialfile=imagen_nombre
            )
            if destino:
                shutil.copy2(archivo_path, destino)
                messagebox.showinfo("Éxito", f"Archivo descargado en:\n{destino}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo descargar: {e}")

    def _abrir_archivo_adjuntado(self, archivo_path, nombre_archivo=None):
        """Abre el archivo adjunto con el visor predeterminado del sistema."""
        if not archivo_path or not os.path.exists(archivo_path):
            messagebox.showerror("Error", "El archivo no existe o fue movido.")
            return

        try:
            if sys.platform.startswith('win'):
                os.startfile(archivo_path)
            elif sys.platform.startswith('darwin'):
                subprocess.run(['open', archivo_path], check=False)
            else:
                subprocess.run(['xdg-open', archivo_path], check=False)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    def _editar_registro(self, registro, ventana_detalles):
        """Abre el formulario para editar un registro (función pendiente)"""
        messagebox.showinfo("Editar", "La funcionalidad de edición está en desarrollo.")

    def _eliminar_y_cerrar(self, registro_id, ventana_detalles):
        """Elimina el registro y cierra la ventana de detalles"""
        respuesta = messagebox.askyesno(
            "Confirmar eliminación",
            "¿Estás seguro de que deseas eliminar este registro?\nEsta acción no se puede deshacer."
        )
        
        if not respuesta:
            return
        
        try:
            servicio = db.Permisos_y_ausencias()
            eliminado = servicio.eliminar_permiso_ausencia(registro_id)
            
            if eliminado:
                messagebox.showinfo("Éxito", "Registro eliminado correctamente.")
                ventana_detalles.destroy()
                self._cargar_empleados_derecha()
            else:
                messagebox.showwarning("Error", "No se pudo eliminar el registro.")
        except Exception as e:
            messagebox.showerror("Error al eliminar", f"Error: {e}")
            import traceback
            logging.error(f"Error eliminando registro: {traceback.format_exc()}")

    def _crear_fila_reporte(self, parent, valores, alterno=False):
        row_color = ("#ffffff", "#1f2937") if alterno else ("#f8fafc", "#111827")
        item_frame = ctk.CTkFrame(parent, fg_color=row_color, corner_radius=14, border_width=1, border_color=("#e2e8f0", "#334155"))
        item_frame.pack(fill='x', padx=5, pady=(0, 8), ipady=12)

        for index, valor in enumerate(valores):
            label = ctk.CTkLabel(
                item_frame,
                text=valor,
                font=ctk.CTkFont(size=13),
                text_color=("#0f172a", "#e2e8f0")
            )
            label.pack(side="left", fill="x", expand=True, padx=(0, 5))
            if index < len(valores) - 1:
                separator = ctk.CTkFrame(item_frame, width=1, fg_color=("#cbd5e1", "#475569"))
                separator.pack(side="left", pady=8)

    def _mostrar_reportes(self, reportes):
        if self.report_rows_container is None:
            return

        for child in self.report_rows_container.winfo_children():
            child.destroy()

        if not reportes:
            plantilla = ctk.CTkFrame(self.report_rows_container, fg_color=("#eef2ff", "#111827"), corner_radius=14)
            plantilla.pack(fill='x', padx=5, pady=10, ipady=20)
            mensaje = ctk.CTkLabel(
                plantilla,
                text="No hay reportes disponibles todavía. Usa el buscador para filtrar por nombre, tipo o estado.",
                font=ctk.CTkFont(size=14),
                text_color=("#475569", "#cbd5e1"),
                wraplength=720,
                justify="left"
            )
            mensaje.pack(anchor="w", padx=20)
            return

        for indice, reporte in enumerate(reportes):
            self._crear_fila_reporte(self.report_rows_container, reporte, alterno=(indice % 2 == 0))

    def _activar_formulario(self):
        self.combobox_permisos.configure(state="normal")
        self.combobox_pagado.configure(state="normal")
        self.boton_archivo.configure(state="normal")
        self.texto.configure(state="normal")
        if self.fecha_inicio_entry:
            self.fecha_inicio_entry.configure(state="normal")
        if self.fecha_fin_entry:
            self.fecha_fin_entry.configure(state="normal")
        self.texto.delete("0.0", "end")
        self.boton_limpiar_P_A.configure(state="normal")
        self.boton_guardar_P_A.configure(state="normal")

    def _limpiar_form(self):
        self.selected_employee = None
        self.nombre_label.configure(text="--")
        self.identificacion_label.configure(text="--")
        self.adjunto_ruta = None
        if self.boton_archivo is not None:
            self.boton_archivo.configure(text="Adjuntar archivo")
        self.combobox_permisos.set("")
        self.combobox_pagado.set("")
        if self.fecha_inicio_entry:
            self.fecha_inicio_entry.set_date(datetime.now())
        if self.fecha_fin_entry:
            self.fecha_fin_entry.set_date(datetime.now())
        self.texto.delete("0.0", "end")
        self.texto.insert("0.0", "Describa el motivo o reporte aquí...")

    def _guardar_permiso(self):
        nombre = self.nombre_label.cget("text").strip()
        identificacion = self.identificacion_label.cget("text").strip()
        tipo = self.combobox_permisos.get().strip()
        pagado = self.combobox_pagado.get().strip()
        motivo = self.texto.get("0.0", "end").strip()
        
        try:
            fecha_inicio = self.fecha_inicio_entry.get_date()
            fecha_fin = self.fecha_fin_entry.get_date()
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar las fechas: {e}")
            return

        if not nombre or nombre == "--":
            messagebox.showwarning("Validación", "Nombre es obligatorio.")
            return
            
        if not identificacion or identificacion == "--":
            messagebox.showwarning("Validación", "Número de identificación es obligatorio.")
            return

        if tipo not in {"Permiso", "Ausencia"}:
            messagebox.showwarning("Validación", "Seleccione Permiso o Ausencia.")
            return

        if pagado not in {"Pagado", "No Pagado"}:
            messagebox.showwarning("Validación", "Seleccione si está Pagado o No Pagado.")
            return

        if not motivo or motivo == "Describa el motivo o reporte aquí...":
            messagebox.showwarning("Validación", "Ingrese un motivo o descripción.")
            return

        if fecha_inicio > fecha_fin:
            messagebox.showwarning("Validación", "La fecha de inicio debe ser anterior a la fecha de fin.")
            return

        try:
            # Guardar en la base de datos
            servicio = db.Permisos_y_ausencias()
            imagen_nombre = os.path.basename(self.adjunto_ruta) if self.adjunto_ruta else None
            imagen_bytes = None
            if self.adjunto_ruta and os.path.exists(self.adjunto_ruta):
                try:
                    with open(self.adjunto_ruta, 'rb') as f:
                        imagen_bytes = f.read()
                except Exception:
                    imagen_nombre = None
                    imagen_bytes = None
            
            servicio.insertar_permiso_ausencia(
                nombre=nombre,
                dui=identificacion,
                fecha_inicio=fecha_inicio,
                fecha_fin=fecha_fin,
                tipo=tipo,
                motivo=motivo,
                pagado=pagado,
                imagen_nombre=imagen_nombre,
                imagen_bytes=imagen_bytes
            )
            
            messagebox.showinfo("Éxito", "Permiso/Ausencia guardado correctamente.")
            self._limpiar_form()
            self._cargar_empleados_derecha()
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el registro:\n{e}")
            import traceback
            logging.error(f"Error guardando permiso/ausencia: {traceback.format_exc()}")

    def _buscar_empleado(self):
        texto_busqueda = self.search_entry.get().strip()
        if not texto_busqueda:
            messagebox.showwarning("Validación", "Ingrese un nombre o DUI para buscar.")
            return

        empleado = db.gestion_empleado().buscar_empleado(texto_busqueda)
        if not empleado:
            messagebox.showinfo("Empleado no encontrado", "No se encontró ningún empleado con ese nombre o DUI.")
            self._limpiar_form()
            return

        self.selected_employee = empleado
        nombre = empleado[1] if len(empleado) > 1 else "--"
        identificacion = empleado[2] if len(empleado) > 2 else "--"
        self.nombre_label.configure(text=nombre or "--")
        self.identificacion_label.configure(text=identificacion or "--")
        self._activar_formulario()

    def _seleccionar_archivo(self):
        extensiones_permitidas = {".pdf", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[
                ("PDF e Imágenes", "*.pdf;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp"),
                ("Todos los archivos", "*.*")
            ]
        )

        if not ruta:
            return

        extension = os.path.splitext(ruta)[1].lower()
        if extension not in extensiones_permitidas:
            messagebox.showwarning(
                "Archivo no permitido",
                "Solo se permiten archivos PDF o imágenes: .png, .jpg, .jpeg, .gif, .bmp, .webp"
            )
            return

        uploads_dir = os.path.join(os.path.abspath('.'), 'uploads')
        os.makedirs(uploads_dir, exist_ok=True)
        nombre_archivo = os.path.basename(ruta)
        destino = os.path.join(uploads_dir, nombre_archivo)
        base, extension = os.path.splitext(nombre_archivo)
        contador = 1
        while os.path.exists(destino):
            destino = os.path.join(uploads_dir, f"{base}_{contador}{extension}")
            contador += 1
        try:
            shutil.copy2(ruta, destino)
            self.adjunto_ruta = destino
            self.boton_archivo.configure(text=os.path.basename(destino))
            messagebox.showinfo("Archivo subido", f"Archivo cargado en:\n{destino}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo subir el archivo:\n{e}")

    def _formatear_fecha(self, valor):
        """Formatea una fecha a formato dd/mm/yyyy"""
        if not valor:
            return "--"
        if hasattr(valor, "strftime"):
            return valor.strftime("%d/%m/%Y")
        
        texto = str(valor or "").strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(texto, formato).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return texto

    def _crear_encabezado_reportes(self):
        """Crea el encabezado para la sección de reportes"""
        if self.report_scroll is None:
            return
        
        header = ctk.CTkFrame(self.report_scroll, fg_color="transparent", height=40)
        header.pack(fill="x", padx=10, pady=(0, 10))
        
        etiquetas = ["Nombre", "DUI", "Tipo", "Pagado"]
        for etiqueta in etiquetas:
            label = ctk.CTkLabel(header, text=etiqueta, font=ctk.CTkFont(size=12, weight="bold"))
            label.pack(side="left", fill="x", expand=True)

    def scrol_permisos_ausencias(self, frame):
        if self.report_scroll is not None:
            self.report_scroll.destroy()

        self.report_scroll = ctk.CTkScrollableFrame(frame)
        self.report_scroll.place(relx=0.00, rely=0.2, relwidth=1.0, relheight=1)
        self._crear_encabezado_reportes()



    
class Vacaciones:

    def info_vacaciones(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

        frame.configure(fg_color="transparent")
        frame_principal = ctk.CTkFrame(frame, fg_color="transparent")
        frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        empleados = []
        vacaciones_db = None
        try:
            empleados = db.gestion_empleado().obtener_empleados() if db else []
            vacaciones_db = db.vacaciones_pagos() if db else None
        except Exception as exc:
            logging.exception("No se pudo cargar el personal para vacaciones")
            messagebox.showerror("Vacaciones", f"No se pudo cargar el personal: {exc}")

        def convertir_fecha(valor):
            if hasattr(valor, "date"):
                return valor.date()
            if hasattr(valor, "year") and hasattr(valor, "month") and hasattr(valor, "day"):
                return valor
            texto = str(valor or "").strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto[:10], formato).date()
                except ValueError:
                    continue
            return None

        def antiguedad(fecha_ingreso):
            fecha = convertir_fecha(fecha_ingreso)
            if not fecha:
                return "Sin fecha"
            dias = max(0, (datetime.now().date() - fecha).days)
            años, dias_restantes = divmod(dias, 365)
            meses = dias_restantes // 30
            return f"{años} año{'s' if años != 1 else ''}, {meses} mes{'es' if meses != 1 else ''}"

        ctk.CTkLabel(
            frame_principal,
            text="Gestión de Vacaciones",
            font=ctk.CTkFont(size=28, weight="bold"),
            anchor="w",
        ).place(relx=0.03, rely=0.025)
        ctk.CTkLabel(
            frame_principal,
            text="Consulta saldos, revisa períodos y registra nuevas solicitudes del personal.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8"),
            anchor="w",
        ).place(relx=0.03, rely=0.085)

        resumen = ctk.CTkFrame(frame_principal, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        resumen.place(relx=0.03, rely=0.14, relwidth=0.94, relheight=0.14)
        pendientes, dias_aprobados = vacaciones_db.contar_resumen() if vacaciones_db else (0, 0)
        metricas = [("Personal registrado", len(empleados), "#2563eb"), ("Solicitudes pendientes", pendientes, "#d97706"), ("Días aprobados", dias_aprobados, "#0f766e"), ("Política", "7 / 15 / 22 / 30", "#7c3aed")]
        for indice, (titulo, valor, color) in enumerate(metricas):
            bloque = ctk.CTkFrame(resumen, fg_color="transparent")
            bloque.grid(row=0, column=indice, padx=18, pady=12, sticky="nsew")
            resumen.grid_columnconfigure(indice, weight=1)
            ctk.CTkLabel(bloque, text=titulo, font=ctk.CTkFont(size=11, weight="bold"), text_color=("#64748b", "#94a3b8"), anchor="w").pack(anchor="w")
            ctk.CTkLabel(bloque, text=str(valor), font=ctk.CTkFont(size=23, weight="bold"), text_color=color, anchor="w").pack(anchor="w", pady=(3, 0))

        # Definir todas las funciones anidadas ANTES de crear los widgets
        def revisar_solicitudes():
            """Abre ventana para revisar y aprobar/rechazar solicitudes de vacaciones."""
            modal_revisar = ctk.CTkToplevel(frame_principal)
            modal_revisar.title("Revisión de solicitudes de vacaciones")
            ajustar_toplevel_a_pantalla(modal_revisar, 1200, 650, margen_x=120, margen_y=120, min_ancho=740, min_alto=420)
            modal_revisar.resizable(False, False)
            modal_revisar.transient(frame_principal.winfo_toplevel())
            modal_revisar.grab_set()
            
            contenedor_revisar = ctk.CTkFrame(modal_revisar, corner_radius=16, border_width=1, border_color=("#dbe4ee", "#334155"))
            contenedor_revisar.pack(fill="both", expand=True, padx=16, pady=16)
            
            ctk.CTkLabel(contenedor_revisar, text="Revisión de solicitudes", font=ctk.CTkFont(size=21, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 4))
            ctk.CTkLabel(contenedor_revisar, text="Aprueba o rechaza solicitudes de vacaciones pendientes.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=20, pady=(0, 16))
            
            tabla_revisar = ctk.CTkFrame(contenedor_revisar, fg_color="transparent")
            tabla_revisar.pack(fill="both", expand=True, padx=20, pady=(0, 12))
            tabla_revisar.grid_rowconfigure(0, weight=1)
            tabla_revisar.grid_columnconfigure(0, weight=1)
            
            columnas_solicitudes = (
                "ID", "Empleado", "Inicio", "Fin", "Días", "Valor diario",
                "Total bruto", "CSS", "Seguro educativo", "ISR", "Deducciones",
                "Pago neto", "Estado", "Observaciones",
            )
            solicitudes_tree = ttk.Treeview(
                tabla_revisar,
                columns=columnas_solicitudes,
                show="headings",
                height=15,
            )
            
            # Configurar columnas
            solicitudes_tree.heading("ID", text="ID")
            solicitudes_tree.heading("Empleado", text="Empleado")
            solicitudes_tree.heading("Inicio", text="Fecha Inicio")
            solicitudes_tree.heading("Fin", text="Fecha Fin")
            solicitudes_tree.heading("Días", text="Días")
            solicitudes_tree.heading("Valor diario", text="Valor diario")
            solicitudes_tree.heading("Total bruto", text="Total bruto")
            solicitudes_tree.heading("CSS", text="CSS")
            solicitudes_tree.heading("Seguro educativo", text="Seguro educativo")
            solicitudes_tree.heading("ISR", text="ISR")
            solicitudes_tree.heading("Deducciones", text="Deducciones")
            solicitudes_tree.heading("Pago neto", text="Pago neto")
            solicitudes_tree.heading("Estado", text="Estado")
            solicitudes_tree.heading("Observaciones", text="Observaciones")
            
            anchos_solicitudes = {
                "ID": 55,
                "Empleado": 180,
                "Inicio": 95,
                "Fin": 95,
                "Días": 55,
                "Valor diario": 100,
                "Total bruto": 105,
                "CSS": 90,
                "Seguro educativo": 125,
                "ISR": 90,
                "Deducciones": 110,
                "Pago neto": 105,
                "Estado": 95,
                "Observaciones": 220,
            }
            for col in columnas_solicitudes:
                solicitudes_tree.column(col, width=anchos_solicitudes[col], minwidth=55, anchor="center")
            
            solicitudes_tree.grid(row=0, column=0, sticky="nsew")
            scroll_tree = ttk.Scrollbar(tabla_revisar, orient="vertical", command=solicitudes_tree.yview)
            scroll_tree.grid(row=0, column=1, sticky="ns")
            scroll_horizontal = ttk.Scrollbar(tabla_revisar, orient="horizontal", command=solicitudes_tree.xview)
            scroll_horizontal.grid(row=1, column=0, sticky="ew")
            solicitudes_tree.configure(
                yscrollcommand=scroll_tree.set,
                xscrollcommand=scroll_horizontal.set,
            )
            
            def cargar_solicitudes_pendientes():
                for item in solicitudes_tree.get_children():
                    solicitudes_tree.delete(item)
                try:
                    solicitudes = vacaciones_db.obtener_solicitudes() if vacaciones_db else []
                    for solicitud in solicitudes:
                        sol_id, empleado, dui, fecha_inicio, fecha_fin, dias, salario, total_bruto, estado, observaciones = solicitud[:10]
                        fecha_inicio_fmt = fecha_inicio.strftime("%d/%m/%Y") if hasattr(fecha_inicio, "strftime") else str(fecha_inicio)
                        fecha_fin_fmt = fecha_fin.strftime("%d/%m/%Y") if hasattr(fecha_fin, "strftime") else str(fecha_fin)
                        salario_mensual = float(salario or 0)
                        pago_bruto = float(total_bruto or 0)
                        calculo_deducciones = Deducciones().calcular_deducciones_nomina(
                            pago_bruto,
                            salario_mensual=salario_mensual,
                            dias_pagados=float(dias or 0),
                        )
                        valor_diario = pago_bruto / float(dias) if float(dias or 0) else 0.0
                        obs_corta = (observaciones or "")[:30] + ("..." if len(observaciones or "") > 30 else "")
                        
                        solicitudes_tree.insert("", "end", iid=str(sol_id), values=(
                            sol_id,
                            empleado,
                            fecha_inicio_fmt,
                            fecha_fin_fmt,
                            dias,
                            f"B/. {valor_diario:,.2f}",
                            f"B/. {pago_bruto:,.2f}",
                            f"B/. {calculo_deducciones['css']:,.2f}",
                            f"B/. {calculo_deducciones['seguro_educativo']:,.2f}",
                            f"B/. {calculo_deducciones['isr']:,.2f}",
                            f"B/. {calculo_deducciones['total_deducciones']:,.2f}",
                            f"B/. {calculo_deducciones['salario_neto']:,.2f}",
                            estado,
                            obs_corta,
                        ))
                except Exception as exc:
                    messagebox.showerror("Error", f"No se pudieron cargar las solicitudes: {exc}")

            def aprobar_solicitud():
                item = solicitudes_tree.selection()
                if not item:
                    messagebox.showwarning("Selección", "Selecciona una solicitud para aprobar.")
                    return
                solicitud_id = item[0]
                try:
                    if vacaciones_db.actualizar_estado_solicitud(int(solicitud_id), "Aprobada"):
                        messagebox.showinfo("Aprobado", "La solicitud ha sido aprobada.")
                        cargar_solicitudes_pendientes()
                    else:
                        messagebox.showwarning("Error", "No se pudo actualizar la solicitud.")
                except Exception as exc:
                    messagebox.showerror("Error", str(exc))
            
            def rechazar_solicitud():
                item = solicitudes_tree.selection()
                if not item:
                    messagebox.showwarning("Selección", "Selecciona una solicitud para rechazar.")
                    return
                solicitud_id = item[0]
                try:
                    if vacaciones_db.actualizar_estado_solicitud(int(solicitud_id), "Rechazada"):
                        messagebox.showinfo("Rechazado", "La solicitud ha sido rechazada.")
                        cargar_solicitudes_pendientes()
                    else:
                        messagebox.showwarning("Error", "No se pudo actualizar la solicitud.")
                except Exception as exc:
                    messagebox.showerror("Error", str(exc))

            def calcular_pago_vacaciones(solicitud):
                _, _, _, _, _, dias, salario, total_bruto, _, _ = solicitud[:10]
                pago_bruto = float(total_bruto or 0)
                dias_pagados = float(dias or 0)
                deducciones = Deducciones().calcular_deducciones_nomina(
                    pago_bruto,
                    salario_mensual=float(salario or 0),
                    dias_pagados=dias_pagados,
                )
                return pago_bruto, deducciones

            def descargar_comprobante_vacaciones(solicitud, ventana):
                sol_id, empleado, dui, fecha_inicio, fecha_fin, dias, salario, _, _, observaciones = solicitud[:10]
                ruta = filedialog.asksaveasfilename(
                    parent=ventana,
                    title="Guardar comprobante de vacaciones",
                    defaultextension=".pdf",
                    filetypes=[("Documento PDF", "*.pdf")],
                    initialfile=f"comprobante_vacaciones_{sol_id}_{str(empleado or 'empleado').replace(' ', '_')}.pdf",
                )
                if not ruta:
                    return

                try:
                    pago_bruto, deducciones = calcular_pago_vacaciones(solicitud)
                    valor_diario = pago_bruto / float(dias) if float(dias or 0) else 0.0
                    fecha_inicio_txt = fecha_inicio.strftime("%d/%m/%Y") if hasattr(fecha_inicio, "strftime") else str(fecha_inicio)
                    fecha_fin_txt = fecha_fin.strftime("%d/%m/%Y") if hasattr(fecha_fin, "strftime") else str(fecha_fin)
                    empresa = datos_empresa.copy()
                    if db and hasattr(db, "empresa"):
                        empresa_guardada = db.empresa().obtener_empresa()
                        if empresa_guardada:
                            empresa.update({
                                "nombre": empresa_guardada[1] or "",
                                "ruc": empresa_guardada[2] or "",
                                "direccion": empresa_guardada[3] or "",
                                "telefono": empresa_guardada[4] or "",
                                "logo": empresa_guardada[5] or "",
                            })

                    estilos = getSampleStyleSheet()
                    estilo_titulo = estilos["Title"].clone("TituloVacaciones")
                    estilo_titulo.alignment = 1
                    estilo_titulo.fontSize = 18
                    estilo_detalle = estilos["Normal"].clone("DetalleVacaciones")
                    estilo_detalle.fontSize = 9
                    estilo_detalle.leading = 12
                    logo = empresa.get("logo") or ""
                    logo_pdf = PdfImage(logo, width=0.70 * inch, height=0.70 * inch) if logo and os.path.isfile(logo) else ""
                    datos_empresa_pdf = [
                        [Paragraph(f"<b>{escape(empresa.get('nombre') or _tr('Nombre de la empresa'))}</b>", estilo_detalle)],
                        [Paragraph(escape(f"{_tr('RUC')}: {empresa.get('ruc') or _tr('No especificado')}"), estilo_detalle)],
                        [Paragraph(escape(f"{_tr('Dirección')}: {empresa.get('direccion') or _tr('No especificada')}"), estilo_detalle)],
                        [Paragraph(escape(f"{_tr('Teléfono')}: {empresa.get('telefono') or _tr('No especificado')}"), estilo_detalle)],
                    ]
                    encabezado = Table([[logo_pdf, Table(datos_empresa_pdf, colWidths=[5.7 * inch])]], colWidths=[0.85 * inch, 5.7 * inch])
                    encabezado.setStyle(TableStyle([
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ]))
                    conceptos = [
                        ["Concepto", "Importe"],
                        ["Salario mensual de referencia", f"B/. {float(salario or 0):,.2f}"],
                        ["Valor diario", f"B/. {valor_diario:,.2f}"],
                        [f"Pago bruto por {dias} día(s) de vacaciones", f"B/. {pago_bruto:,.2f}"],
                        ["CSS", f"B/. {deducciones['css']:,.2f}"],
                        ["Seguro Educativo", f"B/. {deducciones['seguro_educativo']:,.2f}"],
                        ["ISR proporcional", f"B/. {deducciones['isr']:,.2f}"],
                        ["Total deducciones", f"B/. {deducciones['total_deducciones']:,.2f}"],
                        ["PAGO NETO DE VACACIONES", f"B/. {deducciones['salario_neto']:,.2f}"],
                    ]
                    tabla_conceptos = Table(conceptos, colWidths=[4.5 * inch, 2.0 * inch], repeatRows=1)
                    tabla_conceptos.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#dcfce7")),
                        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
                        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                        ("PADDING", (0, 0), (-1, -1), 6),
                    ]))
                    documento = SimpleDocTemplate(ruta, pagesize=letter, rightMargin=45, leftMargin=45, topMargin=40, bottomMargin=40)
                    contenido = [
                        encabezado,
                        Spacer(1, 18),
                        Paragraph("COMPROBANTE DE VACACIONES", estilo_titulo),
                        Spacer(1, 12),
                        Paragraph(escape(f"Comprobante No. {sol_id} | Fecha de emisión: {datetime.now():%d/%m/%Y}"), estilo_detalle),
                        Spacer(1, 8),
                        Paragraph(f"<b>Empleado:</b> {escape(str(empleado or '-'))}<br/><b>DUI:</b> {escape(str(dui or '-'))}<br/><b>Período de vacaciones:</b> {escape(fecha_inicio_txt)} al {escape(fecha_fin_txt)}<br/><b>Días aprobados:</b> {escape(str(dias))}", estilo_detalle),
                        Spacer(1, 14),
                        tabla_conceptos,
                        Spacer(1, 14),
                        Paragraph(f"<b>Observaciones:</b> {escape(str(observaciones or 'Sin observaciones.'))}", estilo_detalle),
                        Spacer(1, 44),
                        Table([["____________________________", "____________________________"], ["Firma del empleado", "Firma autorizada"]], colWidths=[3.2 * inch, 3.2 * inch], style=TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTNAME", (0, 1), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, -1), 9)])),
                    ]
                    documento.build(contenido)
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.abspath(ruta))
                    messagebox.showinfo("Vacaciones", "Comprobante generado correctamente.", parent=ventana)
                except Exception as exc:
                    logging.exception("No se pudo generar el comprobante de vacaciones")
                    messagebox.showerror("Vacaciones", f"No se pudo generar el comprobante: {exc}", parent=ventana)

            def mostrar_aprobadas():
                aprobadas = vacaciones_db.obtener_solicitudes("Aprobada") if vacaciones_db else []
                ventana_aprobadas = ctk.CTkToplevel(modal_revisar)
                ventana_aprobadas.title("Vacaciones aprobadas")
                ajustar_toplevel_a_pantalla(ventana_aprobadas, 760, 480, margen_x=120, margen_y=120, min_ancho=560, min_alto=360)
                ventana_aprobadas.transient(modal_revisar)
                ventana_aprobadas.grab_set()
                contenedor = ctk.CTkFrame(ventana_aprobadas, corner_radius=14, border_width=1, border_color=("#dbe4ee", "#334155"))
                contenedor.pack(fill="both", expand=True, padx=14, pady=14)
                ctk.CTkLabel(contenedor, text="Vacaciones aprobadas", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").pack(fill="x", padx=16, pady=(16, 3))
                ctk.CTkLabel(contenedor, text="Descarga el comprobante individual con el detalle del pago y las deducciones.", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=16, pady=(0, 10))
                lista_aprobadas = ctk.CTkScrollableFrame(contenedor, fg_color="transparent")
                lista_aprobadas.pack(fill="both", expand=True, padx=16, pady=(0, 10))
                if not aprobadas:
                    ctk.CTkLabel(lista_aprobadas, text="No hay solicitudes de vacaciones aprobadas.", font=ctk.CTkFont(size=13), text_color=("#64748b", "#94a3b8")).pack(pady=30)
                for solicitud in aprobadas:
                    _, empleado, dui, fecha_inicio, fecha_fin, dias, _, _, _, _ = solicitud[:10]
                    pago_bruto, deducciones = calcular_pago_vacaciones(solicitud)
                    inicio = fecha_inicio.strftime("%d/%m/%Y") if hasattr(fecha_inicio, "strftime") else str(fecha_inicio)
                    fin = fecha_fin.strftime("%d/%m/%Y") if hasattr(fecha_fin, "strftime") else str(fecha_fin)
                    fila = ctk.CTkFrame(lista_aprobadas, corner_radius=8, fg_color=("#f8fafc", "#1f2937"))
                    fila.pack(fill="x", pady=4)
                    ctk.CTkLabel(fila, text=f"{empleado}\nDUI: {dui} | {inicio} al {fin} | {dias} día(s)\nBruto: B/. {pago_bruto:,.2f} | Deducciones: B/. {deducciones['total_deducciones']:,.2f} | Neto: B/. {deducciones['salario_neto']:,.2f}", font=ctk.CTkFont(size=11), justify="left", anchor="w").pack(side="left", fill="x", expand=True, padx=12, pady=10)
                    ctk.CTkButton(fila, text="Descargar PDF", width=125, height=32, fg_color="#0f766e", hover_color="#115e59", command=lambda solicitud=solicitud: descargar_comprobante_vacaciones(solicitud, ventana_aprobadas)).pack(side="right", padx=10)
                ctk.CTkButton(contenedor, text="Cerrar", width=100, fg_color="#64748b", hover_color="#475569", command=ventana_aprobadas.destroy).pack(anchor="e", padx=16, pady=(0, 14))
            
            botones_revisar = ctk.CTkFrame(contenedor_revisar, fg_color="transparent")
            botones_revisar.pack(fill="x", padx=20, pady=16)
            ctk.CTkButton(botones_revisar, text="Aprobar", width=120, fg_color="#0f766e", hover_color="#115e59", command=aprobar_solicitud).pack(side="left", padx=(0, 8))
            ctk.CTkButton(botones_revisar, text="Rechazar", width=120, fg_color="#dc2626", hover_color="#b91c1c", command=rechazar_solicitud).pack(side="left", padx=(0, 8))
            ctk.CTkButton(botones_revisar, text="Aprobadas", width=120, fg_color="#2563eb", hover_color="#1d4ed8", command=mostrar_aprobadas).pack(side="left", padx=(0, 8))
            ctk.CTkButton(botones_revisar, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=modal_revisar.destroy).pack(side="right")
            
            cargar_solicitudes_pendientes()

        barra = ctk.CTkFrame(frame_principal, corner_radius=14, fg_color=("#f8fafc", "#1f2937"), border_width=1, border_color=("#dbe4ee", "#334155"))
        barra.place(relx=0.03, rely=0.30, relwidth=0.94, relheight=0.09)
        
        # ComboBox para Departamento
        ctk.CTkLabel(barra, text="Depto:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(14, 8), pady=10)
        departamentos_lista = ["Todos"] + vacaciones_db.obtener_departamentos_unicos() if vacaciones_db else ["Todos"]
        combo_departamento = ctk.CTkComboBox(barra, values=departamentos_lista, width=150, state="readonly")
        combo_departamento.set("Todos")
        combo_departamento.pack(side="left", padx=(0, 12), pady=10)
        
        entrada_busqueda = ctk.CTkEntry(barra, width=310, height=34, placeholder_text="Buscar por nombre o DUI", font=ctk.CTkFont(size=13))
        entrada_busqueda.pack(side="left", padx=(0, 8), pady=10)
        ctk.CTkButton(barra, text="Revisar solicitudes", width=145, height=34, fg_color="#d97706", hover_color="#b45309", command=revisar_solicitudes).pack(side="right", padx=(0, 8), pady=10)
        ctk.CTkButton(barra, text="Actualizar", width=105, height=34, fg_color="#2563eb", hover_color="#1d4ed8", command=lambda: cargar_empleados()).pack(side="right", padx=(0, 8), pady=10)

        tabla = ctk.CTkFrame(frame_principal, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        tabla.place(relx=0.03, rely=0.41, relwidth=0.94, relheight=0.54)
        ctk.CTkLabel(tabla, text="Resumen por empleado", font=ctk.CTkFont(size=17, weight="bold"), anchor="w").pack(fill="x", padx=16, pady=(14, 8))
        lista = ctk.CTkScrollableFrame(tabla, fg_color="transparent")
        lista.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columnas = ["Empleado", "DUI", "Antigüedad", "Aprobadas", "Pendientes", "Estado", "Acción"]
        # Pesos de columnas para alineación consistente
        pesos_columnas = [4, 2, 2, 2, 2, 2, 1]

        def dibujar_encabezado():
            encabezado = ctk.CTkFrame(lista, fg_color=("#e8f1f5", "#1e293b"), height=40, corner_radius=8)
            encabezado.pack(fill="x", pady=(0, 6), padx=2)
            # Configurar pesos de columnas
            for indice, peso in enumerate(pesos_columnas):
                encabezado.grid_columnconfigure(indice, weight=peso)
            # Dibujar encabezados
            for indice, texto in enumerate(columnas):
                ctk.CTkLabel(encabezado, text=texto, font=ctk.CTkFont(size=12, weight="bold"), anchor="w", text_color=("#0f172a", "#e2e8f0")).grid(row=0, column=indice, padx=10, pady=10, sticky="ew")

        def abrir_solicitud(empleado=None):
            modal = ctk.CTkToplevel(frame_principal)
            modal.title("Nueva solicitud de vacaciones")
            ajustar_toplevel_a_pantalla(modal, 560, 460, margen_x=120, margen_y=120, min_ancho=420, min_alto=340)
            modal.resizable(False, False)
            modal.transient(frame_principal.winfo_toplevel())
            modal.grab_set()
            contenedor = ctk.CTkFrame(modal, corner_radius=16, border_width=1, border_color=("#dbe4ee", "#334155"))
            contenedor.pack(fill="both", expand=True, padx=16, pady=16)
            ctk.CTkLabel(contenedor, text="Nueva solicitud de vacaciones", font=ctk.CTkFont(size=21, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 4))
            ctk.CTkLabel(contenedor, text="Complete el período que desea enviar a aprobación.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=20, pady=(0, 16))
            formulario = ctk.CTkFrame(contenedor, fg_color="transparent")
            formulario.pack(fill="x", padx=20)
            formulario.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(formulario, text="Empleado:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(0, 12), pady=8, sticky="w")
            nombres = [str(empleado[1]) for empleado in empleados if len(empleado) > 1]
            selector = ctk.CTkComboBox(formulario, values=nombres or ["Sin empleados"], state="readonly")
            selector.grid(row=0, column=1, pady=8, sticky="ew")
            if empleado and len(empleado) > 1:
                selector.set(str(empleado[1]))
            ctk.CTkLabel(formulario, text="Desde:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=1, column=0, padx=(0, 12), pady=8, sticky="w")
            fecha_inicio = DateEntry(formulario)
            fecha_inicio.grid(row=1, column=1, pady=8, sticky="ew")
            ctk.CTkLabel(formulario, text="Hasta:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=2, column=0, padx=(0, 12), pady=8, sticky="w")
            fecha_fin = DateEntry(formulario)
            fecha_fin.grid(row=2, column=1, pady=8, sticky="ew")
            ctk.CTkLabel(formulario, text="Observaciones:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=3, column=0, padx=(0, 12), pady=8, sticky="nw")
            observaciones = ctk.CTkTextbox(formulario, height=80)
            observaciones.grid(row=3, column=1, pady=8, sticky="ew")

            def guardar_solicitud():
                inicio = fecha_inicio.get_date()
                fin = fecha_fin.get_date()
                
                # Validar que la fecha inicial no sea posterior a la final
                if inicio > fin:
                    messagebox.showwarning("Validación", "La fecha inicial no puede ser posterior a la fecha final.", parent=modal)
                    return
                
                # Validar que no sea fecha pasada
                if inicio < datetime.now().date():
                    messagebox.showwarning("Validación", "No se pueden crear solicitudes para fechas pasadas.", parent=modal)
                    return
                
                seleccionado = next((item for item in empleados if str(item[1]) == selector.get()), None)
                if not seleccionado or not vacaciones_db:
                    messagebox.showwarning("Validación", "Seleccione un empleado válido.", parent=modal)
                    return
                
                dias = (fin - inicio).days + 1
                try:
                    # Pasar el salario sin limpiar, la función lo hará internamente
                    vacaciones_db.insertar_solicitud(
                        seleccionado[0], inicio, fin, dias, seleccionado[8],
                        observaciones.get("1.0", tk.END).strip()
                    )
                except Exception as exc:
                    try:
                        vacaciones_db.conexion.rollback()
                    except Exception:
                        pass
                    messagebox.showerror("Solicitud", str(exc), parent=modal)
                    return
                messagebox.showinfo("Solicitud registrada", f"Se registraron {dias} días para revisión.", parent=modal)
                modal.destroy()
                cargar_empleados()

            botones = ctk.CTkFrame(contenedor, fg_color="transparent")
            botones.pack(fill="x", padx=20, pady=18)
            ctk.CTkButton(botones, text="Cancelar", width=120, fg_color="#64748b", hover_color="#475569", command=modal.destroy).pack(side="right")
            ctk.CTkButton(botones, text="Enviar solicitud", width=145, fg_color="#0f766e", hover_color="#115e59", command=guardar_solicitud).pack(side="right", padx=(0, 10))

        def cargar_empleados():
            for widget in lista.winfo_children():
                widget.destroy()
            
            filtro_nombre_dui = entrada_busqueda.get().strip().lower()
            filtro_departamento = combo_departamento.get()
            
            dibujar_encabezado()
            
            # Filtrar por departamento y búsqueda de texto
            visibles = []
            for empleado in empleados:
                # Filtro por departamento
                if filtro_departamento != "Todos":
                    depto = str(empleado[7] if len(empleado) > 7 else "")
                    if depto != filtro_departamento:
                        continue
                
                # Filtro por nombre o DUI
                if filtro_nombre_dui:
                    nombre = str(empleado[1] if len(empleado) > 1 else "").lower()
                    dui = str(empleado[2] if len(empleado) > 2 else "").lower()
                    if filtro_nombre_dui not in nombre and filtro_nombre_dui not in dui:
                        continue
                
                visibles.append(empleado)
            
            # Dibujar empleados filtrados
            for indice, empleado in enumerate(visibles):
                nombre = str(empleado[1] if len(empleado) > 1 else "Sin nombre")
                dui = str(empleado[2] if len(empleado) > 2 else "-")
                resumen_empleado = vacaciones_db.obtener_resumen_empleado(empleado) if vacaciones_db else {}
                dias_disponibles = resumen_empleado.get("dias_disponibles", 0)
                dias_usados = resumen_empleado.get("dias_usados", 0)
                dias_pendientes = resumen_empleado.get("dias_pendientes", 0)
                estado = "Disponible" if dias_disponibles > 0 else ("Pendiente" if dias_pendientes else "Sin saldo")
                fila = ctk.CTkFrame(lista, fg_color=("#ffffff", "#172033") if indice % 2 == 0 else ("#f8fafc", "#1f2937"), corner_radius=8, height=44)
                fila.pack(fill="x", pady=2, padx=2)
                # Configurar pesos de columnas igual al encabezado
                for col_idx, peso in enumerate(pesos_columnas):
                    fila.grid_columnconfigure(col_idx, weight=peso)
                valores = [nombre, dui, antiguedad(empleado[9] if len(empleado) > 9 else None), f"{dias_usados} días", f"{dias_pendientes} días", f"{estado} ({dias_disponibles})"]
                # Dibujar valores en las primeras 6 columnas
                for columna, valor in enumerate(valores):
                    ctk.CTkLabel(fila, text=valor, font=ctk.CTkFont(size=11), anchor="w").grid(row=0, column=columna, padx=10, pady=10, sticky="ew")
                # Botón en la columna 6
                ctk.CTkButton(fila, text="Solicitar", width=70, height=26, fg_color="#0f766e", hover_color="#115e59", command=lambda empleado=empleado: abrir_solicitud(empleado)).grid(row=0, column=6, padx=10, pady=10, sticky="ew")
            
            if not visibles:
                ctk.CTkLabel(lista, text="No se encontraron empleados con ese criterio.", text_color=("#64748b", "#94a3b8")).pack(pady=24)

        # Conectar eventos de filtro
        entrada_busqueda.bind("<KeyRelease>", lambda event: cargar_empleados())
        combo_departamento.configure(command=lambda value: cargar_empleados())
        cargar_empleados()




class Contrato_trabajo:
    def _info_contrato_trabajo_legacy(self, frame):
        self._info_contrato_trabajo_anterior(frame)

    def _info_contrato_trabajo_anterior(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

        frame.configure(fg_color="transparent")
        servicio_empleados = db.gestion_empleado() if db else None
        servicio_contratos = db.expedientes_laborales() if db else None
        empleados = servicio_empleados.obtener_empleados_todos() if servicio_empleados else []
        seleccionado = {"empleado": None, "contrato": None}

        def fecha_a_texto(valor):
            if not valor:
                return "Sin definir"
            if hasattr(valor, "strftime"):
                return valor.strftime("%d/%m/%Y")
            return str(valor)

        def fecha_desde_control(control):
            try:
                return control.get_date()
            except Exception:
                texto = control.get().strip()
                for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        return datetime.strptime(texto, formato).date()
                    except ValueError:
                        continue
                return None

        ctk.CTkLabel(
            frame,
            text="Expedientes laborales",
            font=ctk.CTkFont(size=28, weight="bold"),
            anchor="w",
        ).place(relx=0.03, rely=0.025)
        ctk.CTkLabel(
            frame,
            text="Administra la vigencia, jornada y documentación de los expedientes laborales del personal.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8"),
            anchor="w",
        ).place(relx=0.03, rely=0.085)

        panel_empleados = ctk.CTkFrame(
            frame, corner_radius=16, fg_color=("#ffffff", "#111827"),
            border_width=1, border_color=("#dbe4ee", "#334155"),
        )
        panel_empleados.place(relx=0.03, rely=0.14, relwidth=0.37, relheight=0.80)
        panel_detalle = ctk.CTkFrame(
            frame, corner_radius=16, fg_color=("#ffffff", "#111827"),
            border_width=1, border_color=("#dbe4ee", "#334155"),
        )
        panel_detalle.place(relx=0.42, rely=0.14, relwidth=0.55, relheight=0.80)

        ctk.CTkLabel(panel_empleados, text="Personal", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(fill="x", padx=16, pady=(16, 4))
        ctk.CTkLabel(panel_empleados, text="Selecciona un empleado para consultar su contrato.", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w", wraplength=280).pack(fill="x", padx=16, pady=(0, 12))
        entrada_busqueda = ctk.CTkEntry(panel_empleados, placeholder_text="Buscar por nombre o DUI", height=34)
        entrada_busqueda.pack(fill="x", padx=16, pady=(0, 10))

        # Crear frame scrollable para la lista de empleados
        lista_frame = ctk.CTkScrollableFrame(panel_empleados, fg_color="transparent", corner_radius=0)
        lista_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        lista_frame.grid_columnconfigure(0, weight=1)
        
        # Diccionario para guardar referencias a los botones y frames de empleados
        empleados_widgets = {}

        titulo_detalle = ctk.CTkLabel(panel_detalle, text="Selecciona un empleado", font=ctk.CTkFont(size=20, weight="bold"), anchor="w")
        titulo_detalle.pack(fill="x", padx=20, pady=(20, 2))
        subtitulo_detalle = ctk.CTkLabel(panel_detalle, text="", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w")
        subtitulo_detalle.pack(fill="x", padx=20, pady=(0, 16))

        resumen = ctk.CTkFrame(panel_detalle, corner_radius=12, fg_color=("#f8fafc", "#1f2937"))
        resumen.pack(fill="x", padx=20, pady=(0, 14))
        resumen.grid_columnconfigure(1, weight=1)
        campos_resumen = {}
        for fila, etiqueta in enumerate(("Tipo de contrato", "Vigencia", "Jornada", "Horario", "Estado", "Documento")):
            ctk.CTkLabel(resumen, text=f"{etiqueta}:", font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=fila, column=0, padx=(14, 10), pady=7, sticky="w")
            campos_resumen[etiqueta] = ctk.CTkLabel(resumen, text="Sin contrato", font=ctk.CTkFont(size=12), anchor="w", wraplength=300)
            campos_resumen[etiqueta].grid(row=fila, column=1, padx=(0, 14), pady=7, sticky="w")

        estado_vacio = ctk.CTkLabel(
            panel_detalle,
            text="Selecciona un empleado para consultar o registrar su contrato laboral.",
            font=ctk.CTkFont(size=13), text_color=("#64748b", "#94a3b8"), justify="left", anchor="nw",
        )
        estado_vacio.pack(fill="both", expand=True, padx=20, pady=10)

        acciones = ctk.CTkFrame(panel_detalle, fg_color="transparent")
        acciones.pack(fill="x", padx=20, pady=(0, 18))

        def mostrar_detalle(empleado=None):
            if empleado is None:
                empleado = seleccionado["empleado"]
            if empleado is None:
                return
            
            seleccionado["empleado"] = empleado
            contrato = servicio_contratos.obtener_contrato(empleado[0]) if servicio_contratos else None
            seleccionado["contrato"] = contrato
            
            # Actualizar la selección visual (cambiar color del botón)
            for idx, (emp, widgets) in enumerate(empleados_widgets.items()):
                if emp[0] == empleado[0]:  # Comparar por ID
                    widgets["boton"].configure(fg_color="#0f766e")
                else:
                    widgets["boton"].configure(fg_color="#1f2937")
            
            titulo_detalle.configure(text=empleado[1] or "Empleado")
            subtitulo_detalle.configure(text=f"DUI: {empleado[2] or '-'}  |  {empleado[6] or 'Sin cargo'}")
            if contrato:
                campos_resumen["Tipo de contrato"].configure(text=contrato[1])
                campos_resumen["Vigencia"].configure(text=f"{fecha_a_texto(contrato[2])} al {fecha_a_texto(contrato[3])}")
                campos_resumen["Jornada"].configure(text=contrato[4] or "Sin definir")
                campos_resumen["Horario"].configure(text=f"{contrato[5] or '-'} a {contrato[6] or '-'}")
                campos_resumen["Estado"].configure(text=contrato[7] or "Activo", text_color="#15803d" if str(contrato[7]).lower() == "activo" else "#b45309")
                campos_resumen["Documento"].configure(text=contrato[8] or "Sin adjunto")
                estado_vacio.configure(text="Contrato cargado. Puedes actualizar sus datos o consultar el expediente laboral.")
            else:
                for campo in campos_resumen.values():
                    campo.configure(text="Sin contrato", text_color=("#64748b", "#94a3b8"))
                estado_vacio.configure(text="Este empleado aún no tiene un contrato registrado.")

        def cargar_empleados(event=None):
            criterio = entrada_busqueda.get().strip().lower()
            
            # Limpiar la lista anterior
            for widget in lista_frame.winfo_children():
                widget.destroy()
            empleados_widgets.clear()
            
            # Crear un botón para cada empleado que coincida con la búsqueda
            for indice, empleado in enumerate(empleados):
                nombre = str(empleado[1] or "")
                dui = str(empleado[2] or "")
                if criterio and criterio not in nombre.lower() and criterio not in dui.lower():
                    continue
                
                contrato = servicio_contratos.obtener_contrato(empleado[0]) if servicio_contratos else None
                estado = contrato[7] if contrato else "Sin contrato"
                
                # Crear un frame para cada empleado
                empleado_frame = ctk.CTkFrame(lista_frame, corner_radius=8, fg_color=("#f1f5f9", "#1f2937"), border_width=1, border_color=("#cbd5e1", "#334155"))
                empleado_frame.grid(row=indice, column=0, sticky="ew", padx=0, pady=6)
                empleado_frame.grid_columnconfigure(0, weight=1)
                
                # Contenido del frame
                info_frame = ctk.CTkFrame(empleado_frame, fg_color="transparent")
                info_frame.pack(fill="x", padx=12, pady=10)
                info_frame.grid_columnconfigure(0, weight=1)
                
                # Nombre del empleado
                ctk.CTkLabel(info_frame, text=nombre, font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(row=0, column=0, sticky="w")
                
                # DUI y Estado
                ctk.CTkLabel(info_frame, text=f"DUI: {dui} | Estado: {estado}", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 0))
                
                # Botón de acceso
                boton_revisar = ctk.CTkButton(
                    empleado_frame, 
                    text="Acceder", 
                    width=80, 
                    height=32,
                    fg_color="#1f2937",
                    hover_color="#334155",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    command=lambda emp=empleado: mostrar_detalle(emp)
                )
                boton_revisar.pack(side="right", padx=12, pady=10)
                
                # Guardar referencias
                empleados_widgets[empleado] = {
                    "frame": empleado_frame,
                    "boton": boton_revisar
                }

        def abrir_formulario():
            empleado = seleccionado["empleado"]
            if not empleado:
                messagebox.showwarning("Contratos", "Selecciona un empleado antes de continuar.")
                return

            contrato_actual = seleccionado["contrato"]
            modal = ctk.CTkToplevel(panel_detalle)
            ajustar_toplevel_a_pantalla(modal, 640, 570, margen_x=120, margen_y=120, min_ancho=500, min_alto=420)
            modal.title("Actualizar contrato" if contrato_actual else "Nuevo contrato")
            modal.transient(panel_detalle.winfo_toplevel())
            modal.grab_set()
            contenedor = ctk.CTkFrame(modal, corner_radius=16, border_width=1, border_color=("#dbe4ee", "#334155"))
            contenedor.pack(fill="both", expand=True, padx=16, pady=16)
            ctk.CTkLabel(contenedor, text=modal.title(), font=ctk.CTkFont(size=21, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 4))
            ctk.CTkLabel(contenedor, text=f"{empleado[1]}  |  DUI: {empleado[2]}", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=20, pady=(0, 14))

            formulario = ctk.CTkFrame(contenedor, fg_color="transparent")
            formulario.pack(fill="x", padx=20)
            formulario.grid_columnconfigure(1, weight=1)
            controles = {}
            definiciones = (("Tipo de contrato", "tipo", ["Indefinido", "Definido", "Por obra o servicio"]), ("Jornada", "jornada", ["Tiempo completo", "Medio tiempo", "Por horas"]), ("Estado", "estado", ["Activo", "Finalizado", "Suspendido"]))
            for fila, (etiqueta, clave, valores) in enumerate(definiciones):
                ctk.CTkLabel(formulario, text=f"{etiqueta}:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=fila, column=0, padx=(0, 14), pady=8, sticky="w")
                controles[clave] = ctk.CTkComboBox(formulario, values=valores, state="readonly")
                controles[clave].grid(row=fila, column=1, pady=8, sticky="ew")
            for fila, (etiqueta, clave, valor) in enumerate((("Fecha de inicio", "inicio", datetime.now()), ("Fecha de fin", "fin", None)), start=3):
                ctk.CTkLabel(formulario, text=f"{etiqueta}:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=fila, column=0, padx=(0, 14), pady=8, sticky="w")
                controles[clave] = DateEntry(formulario, year=valor.year if valor else datetime.now().year, month=valor.month if valor else datetime.now().month, day=valor.day if valor else datetime.now().day)
                controles[clave].grid(row=fila, column=1, pady=8, sticky="ew")
            for fila, (etiqueta, clave, placeholder) in enumerate((("Horario de entrada", "entrada", "08:00 AM"), ("Horario de salida", "salida", "04:00 PM")), start=5):
                ctk.CTkLabel(formulario, text=f"{etiqueta}:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=fila, column=0, padx=(0, 14), pady=8, sticky="w")
                controles[clave] = ctk.CTkEntry(formulario, placeholder_text=placeholder)
                controles[clave].grid(row=fila, column=1, pady=8, sticky="ew")

            archivo = {"ruta": None}
            archivo_label = ctk.CTkLabel(formulario, text="Sin archivo seleccionado", text_color=("#64748b", "#94a3b8"), anchor="w")
            archivo_label.grid(row=7, column=1, pady=8, sticky="w")

            def seleccionar_archivo():
                ruta = filedialog.askopenfilename(parent=modal, title="Seleccionar contrato", filetypes=[("Documentos", "*.pdf *.doc *.docx *.png *.jpg *.jpeg"), ("Todos los archivos", "*.*")])
                if ruta:
                    archivo["ruta"] = ruta
                    archivo_label.configure(text=os.path.basename(ruta))

            ctk.CTkLabel(formulario, text="Documento:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=7, column=0, padx=(0, 14), pady=8, sticky="w")
            ctk.CTkButton(formulario, text="Adjuntar", width=100, command=seleccionar_archivo).grid(row=7, column=2, padx=(8, 0), pady=8)
            if contrato_actual:
                controles["tipo"].set(contrato_actual[1] or "Indefinido")
                controles["jornada"].set(contrato_actual[4] or "Tiempo completo")
                controles["estado"].set(contrato_actual[7] or "Activo")
                controles["entrada"].insert(0, contrato_actual[5] or "")
                controles["salida"].insert(0, contrato_actual[6] or "")
                if contrato_actual[8]:
                    archivo_label.configure(text=contrato_actual[8])

            def guardar():
                inicio = fecha_desde_control(controles["inicio"])
                fin = fecha_desde_control(controles["fin"])
                if not inicio or not fin or inicio > fin:
                    messagebox.showwarning("Validación", "La fecha de inicio debe ser válida y anterior o igual a la fecha de fin.", parent=modal)
                    return
                if not controles["entrada"].get().strip() or not controles["salida"].get().strip():
                    messagebox.showwarning("Validación", "Completa los horarios del contrato.", parent=modal)
                    return
                nombre_archivo = os.path.basename(archivo["ruta"]) if archivo["ruta"] else (contrato_actual[8] if contrato_actual else None)
                bytes_archivo = None
                if archivo["ruta"]:
                    try:
                        with open(archivo["ruta"], "rb") as contenido:
                            bytes_archivo = contenido.read()
                    except OSError as exc:
                        messagebox.showerror("Documento", f"No se pudo leer el archivo: {exc}", parent=modal)
                        return
                try:
                    servicio_contratos.guardar_contrato(empleado[0], controles["tipo"].get(), inicio, fin, controles["jornada"].get(), controles["entrada"].get().strip(), controles["salida"].get().strip(), controles["estado"].get(), nombre_archivo, bytes_archivo)
                    modal.destroy()
                    cargar_empleados()
                    mostrar_detalle()
                    messagebox.showinfo("Contratos", "El contrato se guardó correctamente.")
                except Exception as exc:
                    messagebox.showerror("Contratos", f"No se pudo guardar el contrato: {exc}", parent=modal)

            botones = ctk.CTkFrame(contenedor, fg_color="transparent")
            botones.pack(fill="x", padx=20, pady=18)
            ctk.CTkButton(botones, text="Cancelar", width=120, fg_color="#64748b", hover_color="#475569", command=modal.destroy).pack(side="right")
            ctk.CTkButton(botones, text="Guardar contrato", width=150, fg_color="#0f766e", hover_color="#115e59", command=guardar).pack(side="right", padx=(0, 10))

        def limpiar_seleccion():
            seleccionado["empleado"] = None
            seleccionado["contrato"] = None
            titulo_detalle.configure(text="Selecciona un empleado")
            subtitulo_detalle.configure(text="")
            for campo in campos_resumen.values():
                campo.configure(text="Sin contrato", text_color=("#64748b", "#94a3b8"))
            estado_vacio.configure(text="Selecciona un empleado para consultar o registrar su contrato laboral.")

        ctk.CTkButton(acciones, text="Nuevo / actualizar contrato", width=190, command=abrir_formulario, fg_color="#0f766e", hover_color="#115e59").pack(side="left")
        ctk.CTkButton(acciones, text="Limpiar", width=100, command=limpiar_seleccion, fg_color="#64748b", hover_color="#475569").pack(side="left", padx=(10, 0))
        entrada_busqueda.bind("<KeyRelease>", cargar_empleados)
        cargar_empleados()

    def info_contrato_trabajo(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

        frame.configure(fg_color="transparent")
        servicio_empleados = db.gestion_empleado() if db else None
        servicio_expedientes = db.expedientes_laborales() if db else None
        empleados = servicio_empleados.obtener_empleados() if servicio_empleados else []
        seleccionado = {"empleado": None}

        ctk.CTkLabel(
            frame, text="Expedientes laborales", font=ctk.CTkFont(size=28, weight="bold"), anchor="w"
        ).place(relx=0.03, rely=0.025)
        ctk.CTkLabel(
            frame,
            text="Guarda el contrato emitido por la empresa y consulta el historial completo del personal.",
            font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w",
        ).place(relx=0.03, rely=0.085)

        panel_personal = ctk.CTkFrame(
            frame, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1,
            border_color=("#dbe4ee", "#334155"),
        )
        panel_personal.place(relx=0.03, rely=0.14, relwidth=0.39, relheight=0.80)
        panel_acciones = ctk.CTkFrame(
            frame, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1,
            border_color=("#dbe4ee", "#334155"),
        )
        panel_acciones.place(relx=0.44, rely=0.14, relwidth=0.53, relheight=0.80)

        ctk.CTkLabel(panel_personal, text="Personal", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(fill="x", padx=16, pady=(16, 4))
        ctk.CTkLabel(panel_personal, text="Selecciona a la persona cuyo expediente deseas gestionar.", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w", wraplength=300).pack(fill="x", padx=16, pady=(0, 12))
        departamentos = sorted({
            str(empleado[7]).strip()
            for empleado in empleados
            if len(empleado) > 7 and str(empleado[7] or "").strip()
        })
        combo_departamento = ctk.CTkComboBox(
            panel_personal,
            values=["Todos"] + departamentos,
            state="readonly",
            height=34,
            command=lambda valor: cargar_tabla(),
        )
        combo_departamento.set("Todos")
        combo_departamento.pack(fill="x", padx=16, pady=(0, 8))
        busqueda = ctk.CTkEntry(panel_personal, placeholder_text="Buscar por nombre o DUI", height=34)
        busqueda.pack(fill="x", padx=16, pady=(0, 10))

        marco_tabla = ctk.CTkFrame(
            panel_personal,
            corner_radius=12,
            fg_color=("#f8fafc", "#0f172a"),
            border_width=1,
            border_color=("#dbe4ee", "#334155"),
        )
        marco_tabla.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        estilo_contratos = ttk.Style()
        modo_oscuro = ctk.get_appearance_mode().lower() == "dark"
        color_fondo_tabla = "#0f172a" if modo_oscuro else "#ffffff"
        color_texto_tabla = "#f8fafc" if modo_oscuro else "#0f172a"
        color_fila_contrato = "#143b2a" if modo_oscuro else "#f0fdf4"
        color_texto_contrato = "#bbf7d0" if modo_oscuro else "#166534"
        color_fila_pendiente = "#422006" if modo_oscuro else "#fff7ed"
        color_texto_pendiente = "#fed7aa" if modo_oscuro else "#9a3412"
        estilo_contratos.configure(
            "Contratos.Treeview",
            background=color_fondo_tabla,
            fieldbackground=color_fondo_tabla,
            foreground=color_texto_tabla,
            rowheight=38,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        estilo_contratos.configure(
            "Contratos.Treeview.Heading",
            background="#0f766e",
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=(8, 9),
        )
        estilo_contratos.map(
            "Contratos.Treeview",
            background=[("selected", "#115e59" if modo_oscuro else "#ccfbf1")],
            foreground=[("selected", "#f0fdfa" if modo_oscuro else "#134e4a")],
        )
        estilo_contratos.map(
            "Contratos.Treeview.Heading",
            background=[
                ("active", "#115e59"),
                ("pressed", "#0d5f58"),
                ("!disabled", "#0f766e"),
            ],
            foreground=[
                ("active", "#ffffff"),
                ("pressed", "#ffffff"),
                ("!disabled", "#ffffff"),
            ],
        )
        # Crear un scrollable frame en lugar de tabla
        lista_empleados = ctk.CTkScrollableFrame(marco_tabla, fg_color="transparent")
        lista_empleados.pack(fill="both", expand=True)
        lista_empleados.grid_columnconfigure(0, weight=1)
        
        # Diccionario para almacenar los frames de empleados (para actualizarlos después)
        frames_empleados = {}

        titulo = ctk.CTkLabel(panel_acciones, text="Selecciona un empleado", font=ctk.CTkFont(size=21, weight="bold"), anchor="w")
        titulo.pack(fill="x", padx=20, pady=(20, 4))
        detalle = ctk.CTkLabel(panel_acciones, text="", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w")
        detalle.pack(fill="x", padx=20, pady=(0, 18))

        tarjeta = ctk.CTkFrame(panel_acciones, corner_radius=12, fg_color=("#f8fafc", "#1f2937"))
        tarjeta.pack(fill="x", padx=20, pady=(0, 16))
        etiqueta_archivo = ctk.CTkLabel(tarjeta, text="Ningún contrato guardado", font=ctk.CTkFont(size=14, weight="bold"), anchor="w", wraplength=430)
        etiqueta_archivo.pack(fill="x", padx=16, pady=(16, 4))
        estado_archivo = ctk.CTkLabel(tarjeta, text="Selecciona un empleado para comenzar.", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w", wraplength=430)
        estado_archivo.pack(fill="x", padx=16, pady=(0, 16))

        ctk.CTkLabel(panel_acciones, text="El contrato es elaborado y firmado por la empresa. La aplicación solo almacena su archivo y lo vincula al expediente.", font=ctk.CTkFont(size=12), text_color=("#475569", "#cbd5e1"), justify="left", wraplength=470, anchor="w").pack(fill="x", padx=20, pady=(0, 18))
        barra = ctk.CTkFrame(panel_acciones, fg_color="transparent")
        barra.pack(fill="x", padx=20, pady=(0, 18))

        def contrato_de(empleado):
            return servicio_expedientes.obtener_contrato(empleado[0]) if servicio_expedientes else None

        def cargar_tabla(event=None):
            criterio = busqueda.get().strip().lower()
            departamento_filtro = combo_departamento.get().strip()
            
            # Limpiar frames anteriores
            for frame in lista_empleados.winfo_children():
                frame.destroy()
            frames_empleados.clear()
            
            for indice, empleado in enumerate(empleados):
                nombre, dui = str(empleado[1] or ""), str(empleado[2] or "")
                if criterio and criterio not in nombre.lower() and criterio not in dui.lower():
                    continue
                departamento = str(empleado[7] or "").strip() if len(empleado) > 7 else ""
                if departamento_filtro != "Todos" and departamento != departamento_filtro:
                    continue
                contrato = contrato_de(empleado)
                tiene_contrato = bool(contrato and contrato[8])
                estado_text = "Guardado" if tiene_contrato else "Pendiente"
                
                # Crear frame para cada empleado
                frame_empleado = ctk.CTkFrame(
                    lista_empleados,
                    corner_radius=10,
                    fg_color=color_fila_contrato if tiene_contrato else color_fila_pendiente,
                    border_width=1,
                    border_color=("#cbd5e1", "#475569")
                )
                frame_empleado.pack(fill="x", padx=8, pady=6, ipady=8)
                frame_empleado.grid_columnconfigure(1, weight=1)
                
                # Información del empleado
                info_frame = ctk.CTkFrame(frame_empleado, fg_color="transparent")
                info_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(8, 4))
                
                nombre_label = ctk.CTkLabel(
                    info_frame,
                    text=nombre,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=color_texto_contrato if tiene_contrato else color_texto_pendiente,
                    anchor="w"
                )
                nombre_label.pack(anchor="w")
                
                detalle_label = ctk.CTkLabel(
                    info_frame,
                    text=f"DUI: {dui}  |  Estado: {estado_text}",
                    font=ctk.CTkFont(size=10),
                    text_color=color_texto_contrato if tiene_contrato else color_texto_pendiente,
                    anchor="w"
                )
                detalle_label.pack(anchor="w", pady=(2, 0))
                
                # Botón de revisar
                def crear_revisar(emp_selec):
                    def revisar():
                        seleccionado["empleado"] = emp_selec
                        mostrar_empleado()
                    return revisar
                
                boton_revisar = ctk.CTkButton(
                    frame_empleado,
                    text="Acceder",
                    width=100,
                    height=32,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    fg_color="#0f766e",
                    hover_color="#115e59",
                    text_color="#ffffff",
                    command=crear_revisar(empleado)
                )
                boton_revisar.grid(row=0, column=1, padx=(12, 12), pady=8, sticky="e")
                
                frames_empleados[indice] = frame_empleado

        def mostrar_empleado(event=None):
            empleado = seleccionado["empleado"]
            if not empleado:
                return
            contrato = contrato_de(empleado)
            titulo.configure(text=empleado[1] or "Empleado")
            detalle.configure(text=f"DUI: {empleado[2] or '-'}  |  {empleado[6] or 'Sin cargo'}")
            if contrato and contrato[8]:
                etiqueta_archivo.configure(text=contrato[8])
                estado_archivo.configure(text="Documento contractual guardado en la base de datos.", text_color="#15803d")
            else:
                etiqueta_archivo.configure(text="Ningún contrato guardado")
                estado_archivo.configure(text="Busca el archivo original para asociarlo a este empleado.", text_color=("#64748b", "#94a3b8"))

        def guardar_archivo():
            empleado = seleccionado["empleado"]
            if not empleado:
                messagebox.showwarning("Contratos", "Selecciona un empleado primero.")
                return
            ruta = filedialog.askopenfilename(
                title="Seleccionar contrato de trabajo",
                filetypes=[("Documentos", "*.pdf *.doc *.docx *.odt *.png *.jpg *.jpeg"), ("Todos los archivos", "*.*")],
            )
            if not ruta:
                return
            try:
                with open(ruta, "rb") as archivo:
                    contenido = archivo.read()
                servicio_expedientes.guardar_archivo_contrato(empleado[0], os.path.basename(ruta), contenido)
                cargar_tabla()
                mostrar_empleado()
                messagebox.showinfo("Contratos", "El archivo contractual se guardó correctamente.")
            except (OSError, Exception) as exc:
                messagebox.showerror("Contratos", f"No se pudo guardar el archivo: {exc}")

        def mostrar_informe():
            empleado = seleccionado["empleado"]
            if not empleado:
                messagebox.showwarning("Informe", "Selecciona un empleado para generar su informe.")
                return

            def convertir_fecha_informe(valor):
                if hasattr(valor, "date"):
                    return valor.date()
                if hasattr(valor, "year") and hasattr(valor, "month") and hasattr(valor, "day"):
                    return valor
                texto_fecha = str(valor or "").strip()
                for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                    try:
                        return datetime.strptime(texto_fecha[:10], formato).date()
                    except ValueError:
                        continue
                return None

            def fecha_con_dia(valor):
                fecha = convertir_fecha_informe(valor)
                if not fecha:
                    return "No registrada"
                dias_semana = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")
                return f"{dias_semana[fecha.weekday()]} {fecha.strftime('%d/%m/%Y')}"

            def duracion_contrato(fecha_inicio, fecha_fin):
                inicio = convertir_fecha_informe(fecha_inicio)
                fin = convertir_fecha_informe(fecha_fin)
                if not inicio:
                    return "No estipulada"
                if not fin:
                    return "Indefinido"
                dias = max(0, (fin - inicio).days + 1)
                años, resto = divmod(dias, 365)
                meses, dias_restantes = divmod(resto, 30)
                partes = []
                if años:
                    partes.append(f"{años} año{'s' if años != 1 else ''}")
                if meses:
                    partes.append(f"{meses} mes{'es' if meses != 1 else ''}")
                if dias_restantes or not partes:
                    partes.append(f"{dias_restantes} día{'s' if dias_restantes != 1 else ''}")
                return ", ".join(partes)

            modal = ctk.CTkToplevel(panel_acciones)
            ajustar_toplevel_a_pantalla(modal, 900, 650, margen_x=100, margen_y=100, min_ancho=650, min_alto=450)
            modal.title(f"Informe integral | {empleado[1]}")
            modal.transient(panel_acciones.winfo_toplevel())
            modal.grab_set()
            contenedor = ctk.CTkFrame(modal, corner_radius=16, border_width=1, border_color=("#dbe4ee", "#334155"))
            contenedor.pack(fill="both", expand=True, padx=16, pady=16)
            ctk.CTkLabel(contenedor, text="Informe integral del expediente", font=ctk.CTkFont(size=22, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(20, 3))
            ctk.CTkLabel(contenedor, text=f"{empleado[1]}  |  DUI: {empleado[2] or '-'}  |  Cargo: {empleado[6] or '-'}", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=20, pady=(0, 12))
            informe = tk.Text(contenedor, wrap="word", state="disabled", font=("Consolas", 10), bg="#0f172a", fg="#e2e8f0", insertbackground="#e2e8f0")
            informe.pack(fill="both", expand=True, padx=20, pady=(0, 12))
            try:
                datos = servicio_expedientes.obtener_expediente(empleado, datetime(1900, 1, 1).date(), datetime.now().date())
                contrato = contrato_de(empleado)
                fecha_ingreso = empleado[9] if len(empleado) > 9 else None
                fecha_inicio_contrato = contrato[2] if contrato and len(contrato) > 2 and contrato[2] else fecha_ingreso
                fecha_fin_contrato = contrato[3] if contrato and len(contrato) > 3 else None
                lineas = [
                    "INFORME INTEGRAL DEL PERSONAL", "=" * 42,
                    f"Empleado: {empleado[1]}", f"DUI: {empleado[2]}", f"Departamento: {empleado[7] or '-'}",
                    f"Inicio en la empresa: {fecha_con_dia(fecha_ingreso)}",
                    f"Inicio del contrato: {fecha_con_dia(fecha_inicio_contrato)}",
                    f"Tiempo estipulado: {duracion_contrato(fecha_inicio_contrato, fecha_fin_contrato)}",
                    f"Finalización estipulada: {fecha_con_dia(fecha_fin_contrato) if fecha_fin_contrato else 'Indefinido'}", "",
                    "CONTRATO Y DOCUMENTACIÓN", "-" * 28,
                    f"Archivo contractual: {contrato[8] if contrato and contrato[8] else 'No registrado'}", "",
                    "RESUMEN DE PROCESOS", "-" * 22,
                    f"Marcaciones: {len(datos['marcaciones'])}", f"Días trabajados: {datos['dias_trabajados']}", f"Tardanzas: {datos['tardanzas']}", f"Ausencias estimadas: {datos['ausencias']}", f"Horas extra: {datos['sobretiempo_horas']:.2f}",
                    f"Permisos y ausencias: {len(datos['permisos'])} registro(s)", f"Días de permiso: {datos['dias_permiso']}", f"Solicitudes de vacaciones: {len(datos['vacaciones'])}", f"Remuneraciones: {len(datos['remuneraciones'])} | Total: B/. {datos['total_remuneraciones']:.2f}", f"Planillas: {len(datos['planilla'])} | Total pagado: B/. {datos['total_planilla']:.2f}", f"Total deducciones registradas: B/. {datos['total_deducciones']:.2f}",
                ]
                texto = "\n".join(lineas)
            except Exception as exc:
                texto = f"No se pudo generar el informe: {exc}"
            informe.configure(state="normal")
            informe.insert("1.0", texto)
            informe.configure(state="disabled")

            def descargar_informe_pdf():
                ruta = filedialog.asksaveasfilename(
                    parent=modal,
                    title="Descargar informe integral",
                    defaultextension=".pdf",
                    filetypes=[("Documento PDF", "*.pdf")],
                    initialfile=f"informe_integral_{str(empleado[1] or 'empleado').replace(' ', '_')}.pdf",
                )
                if not ruta:
                    return

                def texto_pdf(valor):
                    return escape(str(valor if valor is not None else "-"))

                def fecha_pdf(valor):
                    if hasattr(valor, "strftime"):
                        return valor.strftime("%d/%m/%Y")
                    return str(valor or "-")

                def tabla_pdf(encabezados, filas, anchos):
                    datos_tabla = [[Paragraph(texto_pdf(valor), estilos["Normal"]) for valor in encabezados]]
                    datos_tabla.extend([
                        [Paragraph(texto_pdf(valor), estilos["Normal"]) for valor in fila]
                        for fila in filas
                    ])
                    tabla = Table(datos_tabla, repeatRows=1, colWidths=anchos)
                    tabla.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0fdf4")]),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("PADDING", (0, 0), (-1, -1), 5),
                    ]))
                    return tabla

                try:
                    estilos = getSampleStyleSheet()
                    documento = SimpleDocTemplate(
                        ruta,
                        pagesize=landscape(letter),
                        rightMargin=0.35 * inch,
                        leftMargin=0.35 * inch,
                        topMargin=0.35 * inch,
                        bottomMargin=0.35 * inch,
                    )
                    contrato_pdf = contrato[8] if contrato and contrato[8] else "No registrado"
                    fecha_ingreso_pdf = empleado[9] if len(empleado) > 9 else None
                    fecha_inicio_contrato_pdf = contrato[2] if contrato and len(contrato) > 2 and contrato[2] else fecha_ingreso_pdf
                    fecha_fin_contrato_pdf = contrato[3] if contrato and len(contrato) > 3 else None
                    historia = [
                        Paragraph("Informe integral del expediente laboral", estilos["Title"]),
                        Spacer(1, 8),
                        Paragraph(
                            f"<b>Empleado:</b> {texto_pdf(empleado[1])} &nbsp;&nbsp; "
                            f"<b>DUI:</b> {texto_pdf(empleado[2])} &nbsp;&nbsp; "
                            f"<b>Cargo:</b> {texto_pdf(empleado[6])} &nbsp;&nbsp; "
                            f"<b>Departamento:</b> {texto_pdf(empleado[7])}",
                            estilos["Normal"],
                        ),
                        Spacer(1, 8),
                        Paragraph("Contrato y documentación", estilos["Heading2"]),
                        tabla_pdf(
                            ["Inicio en la empresa", "Inicio del contrato", "Tiempo estipulado", "Finalización estipulada"],
                            [[fecha_con_dia(fecha_ingreso_pdf), fecha_con_dia(fecha_inicio_contrato_pdf), duracion_contrato(fecha_inicio_contrato_pdf, fecha_fin_contrato_pdf), fecha_con_dia(fecha_fin_contrato_pdf) if fecha_fin_contrato_pdf else "Indefinido"]],
                            [1.7 * inch, 1.7 * inch, 1.7 * inch, 1.8 * inch],
                        ),
                        Spacer(1, 6),
                        Paragraph(f"Archivo contractual: {texto_pdf(contrato_pdf)}", estilos["Normal"]),
                        Spacer(1, 8),
                        Paragraph("Resumen general", estilos["Heading2"]),
                        tabla_pdf(
                            ["Marcaciones", "Días trabajados", "Tardanzas", "Ausencias", "Permisos", "Vacaciones", "Remuneraciones", "Planillas"],
                            [[len(datos["marcaciones"]), datos["dias_trabajados"], datos["tardanzas"], datos["ausencias"], len(datos["permisos"]), len(datos["vacaciones"]), f"B/. {datos['total_remuneraciones']:.2f}", f"B/. {datos['total_planilla']:.2f}"]],
                            [0.9 * inch] * 8,
                        ),
                        Spacer(1, 12),
                    ]

                    filas_marcaciones = [
                        [fecha_pdf(fila[0]), texto_pdf(fila[1]), texto_pdf(fila[2]), texto_pdf(fila[3])]
                        for fila in datos["marcaciones"]
                    ]
                    historia.append(Paragraph("Marcaciones registradas", estilos["Heading2"]))
                    historia.append(tabla_pdf(["Fecha", "Entrada", "Salida", "Estado"], filas_marcaciones or [["Sin registros", "-", "-", "-"]], [1.1 * inch, 1.1 * inch, 1.1 * inch, 2.2 * inch]))
                    historia.append(Spacer(1, 10))

                    filas_permisos = [[fecha_pdf(fila[2]), fecha_pdf(fila[3]), texto_pdf(fila[0]), texto_pdf(fila[1]), texto_pdf(fila[4])] for fila in datos["permisos"]]
                    historia.append(Paragraph("Permisos y ausencias", estilos["Heading2"]))
                    historia.append(tabla_pdf(["Inicio", "Fin", "Tipo", "Pagado", "Motivo"], filas_permisos or [["Sin registros", "-", "-", "-", "-"]], [1.0 * inch, 1.0 * inch, 1.2 * inch, 1.0 * inch, 4.0 * inch]))
                    historia.append(Spacer(1, 10))

                    filas_vacaciones = [[fecha_pdf(fila[0]), fecha_pdf(fila[1]), texto_pdf(fila[2]), f"B/. {float(fila[3] or 0):.2f}", texto_pdf(fila[4])] for fila in datos["vacaciones"]]
                    historia.append(Paragraph("Solicitudes de vacaciones", estilos["Heading2"]))
                    historia.append(tabla_pdf(["Inicio", "Fin", "Días", "Total bruto", "Estado"], filas_vacaciones or [["Sin registros", "-", "-", "-", "-"]], [1.0 * inch, 1.0 * inch, 0.7 * inch, 1.2 * inch, 1.4 * inch]))
                    historia.append(Spacer(1, 10))

                    filas_remuneraciones = [[texto_pdf(fila[0]), f"B/. {float(fila[1] or 0):.2f}", fecha_pdf(fila[2]), texto_pdf(fila[3])] for fila in datos["remuneraciones"]]
                    historia.append(Paragraph("Remuneraciones", estilos["Heading2"]))
                    historia.append(tabla_pdf(["Tipo", "Monto", "Fecha", "Motivo"], filas_remuneraciones or [["Sin registros", "-", "-", "-"]], [1.5 * inch, 1.2 * inch, 1.4 * inch, 4.0 * inch]))
                    historia.append(Spacer(1, 10))

                    filas_planilla = [[f"B/. {float(fila[0] or 0):.2f}", f"B/. {float(fila[1] or 0):.2f}", f"B/. {float(fila[2] or 0):.2f}", f"B/. {float(fila[3] or 0):.2f}", f"B/. {float(fila[5] or 0):.2f}", fecha_pdf(fila[6])] for fila in datos["planilla"]]
                    historia.append(Paragraph("Historial de planilla", estilos["Heading2"]))
                    historia.append(tabla_pdf(["Salario", "CSS", "Seguro educativo", "ISR", "Total pagado", "Fecha"], filas_planilla or [["Sin registros", "-", "-", "-", "-", "-"]], [1.0 * inch] * 6))
                    historia.append(Spacer(1, 10))
                    historia.append(Paragraph("Informe generado por SiPP a partir de la información registrada en la base de datos.", estilos["Italic"]))
                    documento.build(historia)
                    messagebox.showinfo("Informe", "El informe integral PDF se descargó correctamente.", parent=modal)
                except Exception as exc:
                    logging.exception("No se pudo descargar el informe integral")
                    messagebox.showerror("Informe", f"No se pudo generar el PDF: {exc}", parent=modal)

            botones_informe = ctk.CTkFrame(contenedor, fg_color="transparent")
            botones_informe.pack(fill="x", padx=20, pady=(0, 18))
            ctk.CTkButton(botones_informe, text="Descargar PDF", width=145, fg_color="#0f766e", hover_color="#115e59", command=descargar_informe_pdf).pack(side="left")
            ctk.CTkButton(botones_informe, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=modal.destroy).pack(side="right")

        ctk.CTkButton(barra, text="Buscar archivo y guardar", width=190, fg_color="#0f766e", hover_color="#115e59", command=guardar_archivo).pack(side="left")
        ctk.CTkButton(barra, text="Ver informe integral", width=170, fg_color="#2563eb", hover_color="#1d4ed8", command=mostrar_informe).pack(side="left", padx=(10, 0))
        busqueda.bind("<KeyRelease>", cargar_tabla)
        cargar_tabla()


class Liquidaciones:
    def info_liquidaciones(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()
        frame.configure(fg_color="transparent")

        servicio = db.liquidaciones_db() if db else None
        empleados = db.gestion_empleado().obtener_empleados_todos() if db else []
        estado = {"empleado": None, "calculo": None}

        encabezado = ctk.CTkFrame(frame, fg_color="transparent")
        encabezado.pack(fill="x", padx=16, pady=(14, 8))
        ctk.CTkLabel(encabezado, text="Liquidaciones", font=ctk.CTkFont(size=30, weight="bold"), anchor="w").pack(anchor="w")
        ctk.CTkLabel(encabezado, text="Calcula, revisa y conserva el historial de finiquitos del personal.", font=ctk.CTkFont(size=13), text_color=("#64748b", "#94a3b8"), anchor="w").pack(anchor="w", pady=(4, 0))

        cuerpo = ctk.CTkFrame(frame, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        cuerpo.grid_columnconfigure(0, weight=3, uniform="liquidaciones")
        cuerpo.grid_columnconfigure(1, weight=5, uniform="liquidaciones")
        cuerpo.grid_rowconfigure(0, weight=1)

        panel_personal = ctk.CTkFrame(cuerpo, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        panel_personal.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        panel_formulario = ctk.CTkFrame(cuerpo, corner_radius=16, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
        panel_formulario.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(panel_personal, text="Personal", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(18, 3))
        ctk.CTkLabel(panel_personal, text="Selecciona un empleado para preparar su liquidación.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w", wraplength=360).pack(fill="x", padx=18, pady=(0, 12))
        busqueda = ctk.CTkEntry(panel_personal, placeholder_text="Buscar por nombre o DUI", height=36)
        busqueda.pack(fill="x", padx=18, pady=(0, 12))

        # Crear frame scrollable para la lista de empleados
        lista_frame = ctk.CTkScrollableFrame(panel_personal, fg_color="transparent", corner_radius=0)
        lista_frame.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        lista_frame.grid_columnconfigure(0, weight=1)
        
        # Diccionario para guardar referencias a los botones y frames de empleados
        empleados_widgets = {}

        ctk.CTkLabel(panel_formulario, text="Nueva liquidación", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(18, 3))
        seleccionado_label = ctk.CTkLabel(
            panel_formulario,
            text="Selecciona un empleado",
            width=1,
            wraplength=420,
            justify="left",
            font=ctk.CTkFont(size=12),
            text_color=("#64748b", "#94a3b8"),
            anchor="w",
        )
        seleccionado_label.pack(fill="x", padx=20, pady=(0, 12))

        formulario = ctk.CTkScrollableFrame(panel_formulario, fg_color="transparent")
        formulario.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        formulario.grid_columnconfigure(0, minsize=165, weight=0)
        formulario.grid_columnconfigure(1, weight=1)
        controles = {}
        definiciones = [
            ("Fecha de salida", "fecha", "date"), ("Motivo de salida", "motivo", "combo"),
            ("Días trabajados pendientes", "dias_trabajados", "entry"), ("Días de vacaciones", "dias_vacaciones", "entry"),
            ("Días de décimo proporcional", "dias_decimo", "entry"), ("Preaviso / otros ingresos", "preaviso", "entry"),
            ("Indemnización", "indemnizacion", "entry"), ("Otros ingresos", "otros_ingresos", "entry"),
            ("Otros descuentos", "otros_descuentos", "entry"),
        ]
        campos_automaticos = {"dias_trabajados", "dias_vacaciones", "dias_decimo", "preaviso", "indemnizacion", "otros_ingresos", "otros_descuentos"}
        for fila, (etiqueta, clave, tipo) in enumerate(definiciones):
            ctk.CTkLabel(
                formulario,
                text=f"{etiqueta}:",
                width=165,
                wraplength=165,
                justify="left",
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w",
            ).grid(row=fila, column=0, padx=(4, 8), pady=7, sticky="w")
            if tipo == "date":
                control = DateEntry(formulario, height=34)
            elif tipo == "combo":
                control = ctk.CTkComboBox(formulario, values=["Renuncia", "Despido", "Fin de contrato", "Mutuo acuerdo", "Jubilación"], state="readonly", height=34)
                control.set("Renuncia")
            else:
                control = ctk.CTkEntry(formulario, placeholder_text="Automático", height=34)
                control.insert(0, "0")
                if clave in campos_automaticos:
                    control.configure(state="disabled", text_color=("#475569", "#cbd5e1"))
            control.grid(row=fila, column=1, padx=(0, 4), pady=7, sticky="ew")
            controles[clave] = control

        fila_obs = len(definiciones)
        ctk.CTkLabel(formulario, text="Observaciones:", width=165, font=ctk.CTkFont(size=12, weight="bold"), anchor="nw").grid(row=fila_obs, column=0, padx=(4, 8), pady=7, sticky="nw")
        controles["observaciones"] = ctk.CTkTextbox(formulario, height=76)
        controles["observaciones"].grid(row=fila_obs, column=1, padx=(0, 4), pady=7, sticky="ew")

        resumen = ctk.CTkFrame(panel_formulario, corner_radius=12, fg_color=("#f0fdf4", "#143b2a"))
        resumen.pack(fill="x", padx=20, pady=(0, 10))
        etiqueta_resumen = ctk.CTkLabel(
            resumen,
            text="Selecciona un empleado y calcula la liquidación.",
            width=1,
            wraplength=460,
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="w",
        )
        etiqueta_resumen.pack(fill="x", padx=14, pady=12)

        botones = ctk.CTkFrame(panel_formulario, fg_color="transparent")
        botones.pack(fill="x", padx=20, pady=(0, 18))

        def numero(clave):
            try:
                return float(controles[clave].get().replace(",", ".") or 0)
            except (TypeError, ValueError):
                raise ValueError(f"El campo '{clave.replace('_', ' ')}' debe ser numérico.")

        def establecer_numero(clave, valor):
            control = controles[clave]
            control.configure(state="normal")
            control.delete(0, "end")
            control.insert(0, str(valor))
            control.configure(state="disabled")

        def cargar_datos_automaticos():
            if not estado["empleado"] or not servicio:
                return
            try:
                datos = servicio.obtener_datos_automaticos(
                    estado["empleado"], controles["fecha"].get_date()
                )
                for clave in ("dias_trabajados", "dias_vacaciones", "dias_decimo", "preaviso", "indemnizacion", "otros_ingresos", "otros_descuentos"):
                    establecer_numero(clave, datos[clave])
                etiqueta_resumen.configure(
                    text=(
                        f"Datos automáticos cargados | Pendientes desde: {datos['inicio_periodo_pendiente']}\n"
                        f"Trabajo: {datos['dias_trabajados']} días | Vacaciones: {datos['dias_vacaciones']} días | Décimo: {datos['dias_decimo']} días\n"
                        "Preaviso e indemnización requieren una regla contractual configurada."
                    )
                )
                estado["calculo"] = None
                calcular()
            except (ValueError, Exception) as exc:
                etiqueta_resumen.configure(text=f"No se pudieron cargar los datos automáticos: {exc}")

        def calcular():
            if not estado["empleado"]:
                messagebox.showwarning("Liquidaciones", "Selecciona un empleado primero.")
                return
            try:
                # Mapear motivo a causal legal
                motivo_texto = str(controles["motivo"].get() or "Renuncia").strip().upper()
                mapeo_causales = {
                    "RENUNCIA": "RENUNCIA",
                    "DESPIDO": "DESPIDO_INJUSTIFICADO",
                    "FIN DE CONTRATO": "CAUSAS_JUSTIFICADAS",
                    "MUTUO ACUERDO": "CAUSAS_JUSTIFICADAS",
                    "JUBILACIÓN": "JUBILACION",
                }
                causal = mapeo_causales.get(motivo_texto, "DESPIDO_INJUSTIFICADO")
                
                calculo = servicio.calcular_liquidacion(
                    estado["empleado"][8], numero("dias_trabajados"), numero("dias_vacaciones"), numero("dias_decimo"),
                    numero("preaviso"), numero("indemnizacion"), numero("otros_ingresos"), numero("otros_descuentos"),
                    causal_terminacion=causal
                )
                estado["calculo"] = calculo
                
                # Construir resumen con advertencias
                texto_resumen = (
                    f"Salario mensual: B/. {calculo['salario_mensual']:,.2f}\n"
                    f"Total bruto: B/. {calculo['total_bruto']:,.2f}    "
                    f"Deducciones: B/. {calculo['total_deducciones']:,.2f}\n"
                    f"TOTAL NETO: B/. {calculo['total_neto']:,.2f}"
                )
                
                # Agregar advertencias si existen
                if calculo.get('advertencias'):
                    texto_resumen += "\n\n⚠️ ADVERTENCIAS LEGALES:"
                    for adv in calculo['advertencias']:
                        texto_resumen += f"\n• {adv}"
                
                # Color diferente si hay advertencias
                if calculo.get('advertencias'):
                    resumen.configure(fg_color=("#fef08a", "#3f3003"))  # Amarillo claro
                else:
                    resumen.configure(fg_color=("#f0fdf4", "#143b2a"))  # Verde claro
                
                etiqueta_resumen.configure(text=texto_resumen)
            except ValueError as exc:
                messagebox.showwarning("Validación", str(exc))

        def guardar():
            if not estado["empleado"]:
                messagebox.showwarning("Liquidaciones", "Selecciona un empleado primero.")
                return
            try:
                liquidacion_id, calculo = servicio.guardar_liquidacion(
                    estado["empleado"], controles["fecha"].get_date(), controles["motivo"].get(),
                    numero("dias_trabajados"), numero("dias_vacaciones"), numero("dias_decimo"),
                    numero("preaviso"), numero("indemnizacion"), numero("otros_ingresos"), numero("otros_descuentos"),
                    controles["observaciones"].get("1.0", "end").strip(), "CALCULADA",
                )
                estado["calculo"] = calculo
                messagebox.showinfo("Liquidación guardada", f"Liquidación #{liquidacion_id} guardada correctamente.")
            except Exception as exc:
                try:
                    servicio.conexion.rollback()
                except Exception:
                    pass
                messagebox.showerror("Liquidaciones", str(exc))

        ctk.CTkButton(botones, text="Calcular", width=120, command=calcular, fg_color="#2563eb", hover_color="#1d4ed8").pack(side="left")
        ctk.CTkButton(botones, text="Guardar liquidación", width=170, command=guardar, fg_color="#0f766e", hover_color="#115e59").pack(side="left", padx=(10, 0))

        def cargar_empleados(event=None):
            criterio = busqueda.get().strip().lower()
            
            # Limpiar la lista anterior
            for widget in lista_frame.winfo_children():
                widget.destroy()
            empleados_widgets.clear()
            
            # Crear un botón para cada empleado que coincida con la búsqueda
            for indice, empleado in enumerate(empleados):
                nombre, dui, cargo = str(empleado[1] or ""), str(empleado[2] or ""), str(empleado[6] or "")
                if criterio and criterio not in nombre.lower() and criterio not in dui.lower():
                    continue
                
                # Crear un frame para cada empleado
                empleado_frame = ctk.CTkFrame(lista_frame, corner_radius=8, fg_color=("#f1f5f9", "#1f2937"), border_width=1, border_color=("#cbd5e1", "#334155"))
                empleado_frame.grid(row=indice, column=0, sticky="ew", padx=0, pady=6)
                empleado_frame.grid_columnconfigure(0, weight=1)
                
                # Contenido del frame
                info_frame = ctk.CTkFrame(empleado_frame, fg_color="transparent")
                info_frame.pack(fill="x", padx=12, pady=10)
                info_frame.grid_columnconfigure(0, weight=1)
                
                # Nombre del empleado
                ctk.CTkLabel(info_frame, text=nombre, font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(row=0, column=0, sticky="w")
                
                # DUI y Cargo
                ctk.CTkLabel(info_frame, text=f"DUI: {dui} | Cargo: {cargo}", font=ctk.CTkFont(size=11), text_color=("#64748b", "#94a3b8"), anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 0))
                
                # Botón de acceder
                boton_acceder = ctk.CTkButton(
                    empleado_frame, 
                    text="Acceder", 
                    width=80, 
                    height=32,
                    fg_color="#1f2937",
                    hover_color="#334155",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    command=lambda emp=empleado: seleccionar_empleado(emp)
                )
                boton_acceder.pack(side="right", padx=12, pady=10)
                
                # Guardar referencias
                empleados_widgets[empleado] = {
                    "frame": empleado_frame,
                    "boton": boton_acceder
                }

        def seleccionar_empleado(empleado=None):
            if empleado is None:
                return
            
            estado["empleado"] = empleado
            estado["calculo"] = None
            
            # Actualizar la selección visual (cambiar color del botón)
            for emp, widgets in empleados_widgets.items():
                if emp[0] == empleado[0]:  # Comparar por ID
                    widgets["boton"].configure(fg_color="#0f766e")
                else:
                    widgets["boton"].configure(fg_color="#1f2937")
            
            seleccionado_label.configure(
                text=f"Empleado: {empleado[1]}\nDUI: {empleado[2]}  |  Salario: B/. {servicio._numero(empleado[8]):,.2f}"
            )
            cargar_datos_automaticos()

        def abrir_liquidaciones_registradas():
            top_level = ctk.CTkToplevel(frame.winfo_toplevel())
            top_level.title("Liquidaciones registradas")
            ajustar_toplevel_a_pantalla(top_level, 1120, 700, margen_x=80, margen_y=80, min_ancho=760, min_alto=520)
            top_level.transient(frame.winfo_toplevel())
            top_level.grab_set()
            top_level.resizable(False, False)

            contenedor = ctk.CTkFrame(top_level, corner_radius=16, fg_color=("#f8fafc", "#0f172a"), border_width=1, border_color=("#dbe4ee", "#334155"))
            contenedor.pack(fill="both", expand=True, padx=14, pady=14)
            ctk.CTkLabel(contenedor, text="Liquidaciones registradas", font=ctk.CTkFont(size=24, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(16, 2))
            ctk.CTkLabel(contenedor, text="Busca, filtra y consulta el detalle completo de cada liquidación guardada.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=18, pady=(0, 12))

            filtros = ctk.CTkFrame(contenedor, corner_radius=12, fg_color=("#ffffff", "#111827"), border_width=1, border_color=("#dbe4ee", "#334155"))
            filtros.pack(fill="x", padx=18, pady=(0, 12))
            filtros.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(filtros, text="Buscar:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(12, 8), pady=10)
            entrada_filtro = ctk.CTkEntry(filtros, placeholder_text="Nombre o DUI", height=34)
            entrada_filtro.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")
            ctk.CTkLabel(filtros, text="Departamento:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(0, 8), pady=10)
            departamentos = sorted({str(empleado[7] or "").strip() for empleado in empleados if len(empleado) > 7 and str(empleado[7] or "").strip()})
            combo_departamento = ctk.CTkComboBox(filtros, values=["Todos"] + departamentos, state="readonly", width=190, height=34)
            combo_departamento.set("Todos")
            combo_departamento.grid(row=0, column=3, padx=(0, 12), pady=10)

            area_tabla = ctk.CTkFrame(contenedor, fg_color="transparent")
            area_tabla.pack(fill="both", expand=True, padx=18, pady=(0, 10))
            area_tabla.grid_rowconfigure(0, weight=1)
            area_tabla.grid_columnconfigure(0, weight=1)
            columnas_registros = ("ID", "Empleado", "DUI", "Departamento", "Salida", "Motivo", "Bruto", "Neto", "Estado")
            tabla_registros = ttk.Treeview(area_tabla, columns=columnas_registros, show="headings", selectmode="browse")
            anchos = (45, 150, 100, 145, 90, 120, 95, 95, 90)
            for columna, ancho in zip(columnas_registros, anchos):
                tabla_registros.heading(columna, text=columna)
                tabla_registros.column(columna, width=ancho, minwidth=45, anchor="w", stretch=True)
            tabla_registros.grid(row=0, column=0, sticky="nsew")
            scroll_registros = ttk.Scrollbar(area_tabla, orient="vertical", command=tabla_registros.yview)
            scroll_registros.grid(row=0, column=1, sticky="ns")
            tabla_registros.configure(yscrollcommand=scroll_registros.set)

            detalle = ctk.CTkLabel(contenedor, text="Selecciona una liquidación para ver todos sus conceptos.", font=ctk.CTkFont(size=12), justify="left", anchor="nw", wraplength=1000)
            detalle.pack(fill="x", padx=18, pady=(0, 10))
            acciones = ctk.CTkFrame(contenedor, fg_color="transparent")
            acciones.pack(fill="x", padx=18, pady=(0, 14))

            registros = []
            empleados_por_dui = {str(empleado[2] or "").strip(): empleado for empleado in empleados if len(empleado) > 2}

            def cargar_registros(event=None):
                criterio = entrada_filtro.get().strip().casefold()
                departamento = combo_departamento.get().strip()
                for item in tabla_registros.get_children():
                    tabla_registros.delete(item)
                for registro in registros:
                    empleado = empleados_por_dui.get(str(registro[3] or "").strip())
                    departamento_registro = str(empleado[7] or "").strip() if empleado and len(empleado) > 7 else "Sin departamento"
                    coincide = not criterio or criterio in str(registro[2] or "").casefold() or criterio in str(registro[3] or "").casefold()
                    coincide_departamento = departamento == "Todos" or departamento_registro == departamento
                    if coincide and coincide_departamento:
                        tabla_registros.insert("", "end", iid=str(registro[0]), values=(registro[0], registro[2], registro[3], departamento_registro, registro[4], registro[5], f"B/. {float(registro[6] or 0):,.2f}", f"B/. {float(registro[8] or 0):,.2f}", registro[9]))

            def mostrar_detalle(event=None):
                seleccion = tabla_registros.selection()
                if not seleccion:
                    return
                registro = servicio.obtener_liquidacion(int(seleccion[0]))
                if not registro:
                    return
                etiquetas = ("Salario mensual", "Salario pendiente", "Vacaciones pendientes", "Décimo proporcional", "Preaviso", "Indemnización", "Otros ingresos", "CSS", "Seguro educativo", "ISR", "Otros descuentos", "Total bruto", "Total deducciones", "Total neto")
                indices = (6, 8, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)
                conceptos = " | ".join(f"{etiqueta}: B/. {float(registro[indice] or 0):,.2f}" for etiqueta, indice in zip(etiquetas, indices))
                detalle.configure(text=f"Liquidación #{registro[0]} | {registro[2]} | DUI: {registro[3]} | Salida: {registro[4]} | Motivo: {registro[5]} | Estado: {registro[24]}\n{conceptos}\nObservaciones: {registro[23] or 'Sin observaciones'}")

            def eliminar_seleccionada():
                seleccion = tabla_registros.selection()
                if not seleccion:
                    messagebox.showwarning("Liquidaciones", "Selecciona una liquidación para eliminar.", parent=top_level)
                    return
                valores = tabla_registros.item(seleccion[0], "values")
                if not messagebox.askyesno("Eliminar liquidación", f"¿Deseas eliminar la liquidación #{valores[0]} de {valores[1]}?\n\nEsta acción no se puede deshacer.", parent=top_level):
                    return
                try:
                    if servicio.eliminar_liquidacion(int(valores[0])):
                        cargar_registros()
                        cargar_empleados()
                        detalle.configure(text="Liquidación eliminada. Selecciona otra para consultar su detalle.")
                        messagebox.showinfo("Liquidaciones", "Liquidación eliminada correctamente.", parent=top_level)
                    else:
                        messagebox.showwarning("Liquidaciones", "La liquidación ya no existe.", parent=top_level)
                except Exception as exc:
                    messagebox.showerror("Liquidaciones", f"No se pudo eliminar: {exc}", parent=top_level)

            ctk.CTkButton(acciones, text="Eliminar seleccionada", width=170, fg_color="#dc2626", hover_color="#b91c1c", command=eliminar_seleccionada).pack(side="left")
            ctk.CTkButton(acciones, text="Actualizar", width=110, fg_color="#2563eb", hover_color="#1d4ed8", command=cargar_registros).pack(side="left", padx=(10, 0))
            ctk.CTkButton(acciones, text="Cerrar", width=110, fg_color="#64748b", hover_color="#475569", command=top_level.destroy).pack(side="right")
            tabla_registros.bind("<<TreeviewSelect>>", mostrar_detalle)
            entrada_filtro.bind("<KeyRelease>", cargar_registros)
            combo_departamento.configure(command=cargar_registros)
            try:
                registros = servicio.obtener_liquidaciones() if servicio else []
                cargar_registros()
            except Exception as exc:
                detalle.configure(text=f"No se pudo cargar el registro: {exc}")

        acciones_principales = ctk.CTkFrame(panel_formulario, fg_color="transparent")
        acciones_principales.pack(fill="x", padx=20, pady=(0, 18))
        ctk.CTkButton(acciones_principales, text="Liquidaciones registradas", width=210, height=36, fg_color="#0f766e", hover_color="#115e59", command=abrir_liquidaciones_registradas).pack(side="left")

        busqueda.bind("<KeyRelease>", cargar_empleados)
        controles["fecha"].bind("<<DateEntrySelected>>", lambda event: cargar_datos_automaticos())
        cargar_empleados()

class Cartas_de_trabajo:
    def info_cartas(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()
        frame.configure(fg_color="transparent")

        empleados = db.gestion_empleado().obtener_empleados() if db else []
        empresa = datos_empresa.copy()
        if db and hasattr(db, "empresa"):
            try:
                empresa_guardada = db.empresa().obtener_empresa()
                if empresa_guardada:
                    empresa.update({
                        "nombre": empresa_guardada[1] or "",
                        "ruc": empresa_guardada[2] or "",
                        "direccion": empresa_guardada[3] or "",
                        "telefono": empresa_guardada[4] or "",
                        "logo": empresa_guardada[5] or "",
                    })
            except Exception:
                logging.exception("No se pudieron cargar los datos de empresa para cartas")

        ctk.CTkLabel(frame, text="Cartas de trabajo", font=ctk.CTkFont(size=28, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(16, 2))
        ctk.CTkLabel(frame, text="Genera una carta laboral profesional para cada empleado registrado.", font=ctk.CTkFont(size=12), text_color=("#64748b", "#94a3b8"), anchor="w").pack(fill="x", padx=18, pady=(0, 12))

        filtros = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#f8fafc", "#1f2937"))
        filtros.pack(fill="x", padx=18, pady=(0, 12))
        filtros.grid_columnconfigure(1, weight=1)
        filtros.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(filtros, text="Buscar:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, padx=(12, 8), pady=10)
        busqueda = ctk.CTkEntry(filtros, placeholder_text="Nombre o DUI", height=34)
        busqueda.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")
        ctk.CTkLabel(filtros, text="Departamento:", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(0, 8), pady=10)
        departamentos = sorted({str(empleado[7] or "").strip() for empleado in empleados if len(empleado) > 7 and str(empleado[7] or "").strip()})
        selector_departamento = ctk.CTkComboBox(filtros, values=["Todos"] + departamentos, state="readonly", width=190, height=34)
        selector_departamento.set("Todos")
        selector_departamento.grid(row=0, column=3, padx=(0, 12), pady=10, sticky="ew")

        area = ctk.CTkFrame(frame, corner_radius=12, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        area.grid_rowconfigure(0, weight=1)
        area.grid_columnconfigure(0, weight=1)
        tabla = ttk.Treeview(area, columns=("Nombre", "DUI", "Departamento", "Cargo", "Ingreso"), show="headings", selectmode="browse")
        for columna, ancho in (("Nombre", 220), ("DUI", 130), ("Departamento", 190), ("Cargo", 230), ("Ingreso", 110)):
            tabla.heading(columna, text=columna)
            tabla.column(columna, width=ancho, minwidth=80, anchor="w", stretch=True)
        tabla.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(area, orient="vertical", command=tabla.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        tabla.configure(yscrollcommand=scroll.set)

        def convertir_fecha(valor):
            if hasattr(valor, "date"):
                return valor.date()
            if hasattr(valor, "year") and hasattr(valor, "month") and hasattr(valor, "day"):
                return valor
            texto = str(valor or "").strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto[:10], formato).date()
                except ValueError:
                    continue
            return None

        def antiguedad(fecha_ingreso):
            inicio = convertir_fecha(fecha_ingreso)
            if not inicio:
                return "No especificada"
            hoy = datetime.now().date()
            años = hoy.year - inicio.year - ((hoy.month, hoy.day) < (inicio.month, inicio.day))
            meses = (hoy.month - inicio.month) % 12
            if hoy.day < inicio.day:
                meses = max(0, meses - 1)
            partes = []
            if años:
                partes.append(f"{años} año{'s' if años != 1 else ''}")
            if meses:
                partes.append(f"{meses} mes{'es' if meses != 1 else ''}")
            return ", ".join(partes) if partes else "Menos de un mes"

        def cargar_empleados(event=None):
            criterio = busqueda.get().strip().casefold()
            departamento = selector_departamento.get().strip()
            for item in tabla.get_children():
                tabla.delete(item)
            for indice, empleado in enumerate(empleados):
                nombre = str(empleado[1] or "")
                dui = str(empleado[2] or "")
                departamento_empleado = str(empleado[7] or "Sin departamento").strip()
                if criterio and criterio not in nombre.casefold() and criterio not in dui.casefold():
                    continue
                if departamento != "Todos" and departamento_empleado != departamento:
                    continue
                ingreso = empleado[9] if len(empleado) > 9 else ""
                ingreso_texto = convertir_fecha(ingreso).strftime("%d/%m/%Y") if convertir_fecha(ingreso) else "No especificada"
                tabla.insert("", "end", iid=str(indice), values=(nombre, dui, departamento_empleado, empleado[6] or "Sin cargo", ingreso_texto))

        def generar_carta():
            seleccion = tabla.selection()
            if not seleccion:
                messagebox.showwarning("Cartas de trabajo", "Seleccione un empleado.", parent=frame.winfo_toplevel())
                return
            empleado = empleados[int(seleccion[0])]
            ruta = filedialog.asksaveasfilename(
                parent=frame.winfo_toplevel(),
                title="Guardar carta de trabajo",
                defaultextension=".pdf",
                filetypes=[("Documento PDF", "*.pdf")],
                initialfile=f"carta_trabajo_{str(empleado[1] or 'empleado').replace(' ', '_')}.pdf",
            )
            if not ruta:
                return
            try:
                estilos = getSampleStyleSheet()
                estilo_cuerpo = estilos["Normal"].clone("CartaCuerpo")
                estilo_cuerpo.fontSize = 11
                estilo_cuerpo.leading = 17
                estilo_titulo = estilos["Title"].clone("CartaTitulo")
                estilo_titulo.alignment = 1
                nombre_empresa = empresa.get("nombre") or "La empresa"
                fecha_ingreso = convertir_fecha(empleado[9] if len(empleado) > 9 else None)
                fecha_ingreso_texto = fecha_ingreso.strftime("%d/%m/%Y") if fecha_ingreso else "una fecha no especificada"
                tiempo = antiguedad(empleado[9] if len(empleado) > 9 else None)
                fecha_actual = datetime.now().strftime("%d de %B de %Y")
                logo = empresa.get("logo") or ""
                logo_pdf = PdfImage(logo, width=0.8 * inch, height=0.8 * inch) if logo and os.path.isfile(logo) else ""
                encabezado_empresa = Table([[logo_pdf, Paragraph(
                    f"<b>{escape(nombre_empresa)}</b><br/>"
                    f"{_tr('RUC')}: {escape(empresa.get('ruc') or _tr('No especificado'))}<br/>"
                    f"{escape(empresa.get('direccion') or _tr('No especificada'))}<br/>"
                    f"{_tr('Teléfono')}: {escape(empresa.get('telefono') or _tr('No especificado'))}", estilo_cuerpo
                )]], colWidths=[0.9 * inch, 5.9 * inch])
                encabezado_empresa.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
                texto_carta = (
                    f"Por medio de la presente hacemos constar que el señor(a) <b>{escape(str(empleado[1] or ''))}</b>, "
                    f"con documento de identidad <b>{escape(str(empleado[2] or ''))}</b>, labora o laboró en nuestra organización "
                    f"desempeñándose como <b>{escape(str(empleado[6] or 'colaborador'))}</b> en el departamento de "
                    f"<b>{escape(str(empleado[7] or 'Sin departamento'))}</b>."
                )
                contenido = [
                    encabezado_empresa, Spacer(1, 22), Paragraph(_tr("CARTA DE TRABAJO"), estilo_titulo), Spacer(1, 22),
                    Paragraph(escape(_tr("A quien corresponda:")), estilo_cuerpo), Spacer(1, 12),
                    Paragraph(texto_carta, estilo_cuerpo), Spacer(1, 12),
                    Paragraph(escape(f"El vínculo laboral inició el {fecha_ingreso_texto} y registra un tiempo de servicio de {tiempo}. Durante este período ha demostrado responsabilidad, compromiso, disposición y un desempeño profesional satisfactorio en las funciones asignadas."), estilo_cuerpo), Spacer(1, 12),
                    Paragraph("Extendemos la presente carta a solicitud del interesado para los fines que estime convenientes. Certificamos que la información consignada es verdadera según nuestros registros laborales.", estilo_cuerpo), Spacer(1, 30),
                    Paragraph(escape(f"{_tr('Atentamente,')}\\n\\n{nombre_empresa}\\n{empresa.get('telefono') or ''}"), estilo_cuerpo), Spacer(1, 28),
                    Paragraph(escape(f"{_tr('Emitida el')} {fecha_actual}"), estilos["Normal"]),
                ]
                documento = SimpleDocTemplate(ruta, pagesize=letter, rightMargin=54, leftMargin=54, topMargin=45, bottomMargin=45)
                documento.build(contenido)
                ruta = os.path.abspath(ruta)
                if sys.platform.startswith("win"):
                    os.startfile(ruta)
                messagebox.showinfo("Cartas de trabajo", f"Carta generada correctamente en:\n{ruta}", parent=frame.winfo_toplevel())
            except Exception as exc:
                logging.exception("No se pudo generar la carta de trabajo")
                messagebox.showerror("Cartas de trabajo", f"No se pudo generar la carta: {exc}", parent=frame.winfo_toplevel())

        acciones = ctk.CTkFrame(frame, fg_color="transparent")
        acciones.pack(fill="x", padx=18, pady=(0, 14))
        ctk.CTkButton(acciones, text="Generar carta PDF", width=160, fg_color="#0f766e", hover_color="#115e59", command=generar_carta).pack(side="left")
        ctk.CTkButton(acciones, text="Actualizar", width=110, fg_color="#2563eb", hover_color="#1d4ed8", command=cargar_empleados).pack(side="left", padx=(10, 0))
        ctk.CTkButton(acciones, text="Limpiar", width=100, fg_color="#64748b", hover_color="#475569", command=lambda: (busqueda.delete(0, "end"), selector_departamento.set("Todos"), cargar_empleados())).pack(side="left", padx=(10, 0))

        busqueda.bind("<KeyRelease>", cargar_empleados)
        busqueda.bind("<Return>", cargar_empleados)
        selector_departamento.configure(command=cargar_empleados)
        cargar_empleados()

class Manuales:
    """Presentación tipo diapositivas que explica, módulo por módulo, para qué sirve SiPP."""

    DIAPOSITIVAS = [
        {
            "titulo": "Bienvenido a SiPP",
            "resumen": "Sistema integral de Planilla y Personal",
            "imagen": "bienvenida",
            "puntos": [
                "SiPP centraliza en una sola aplicación la administración de empleados, marcaciones de asistencia, planilla, préstamos, deducciones legales, remuneraciones, permisos, vacaciones, contratos, liquidaciones y cartas laborales.",
                "El menú lateral (izquierda) da acceso a cada módulo principal; el menú superior agrupa Configuración, Seguridad, Reportes y Ayuda.",
                "Todos los módulos comparten la misma base de empleados registrada en Gestión Personal, así que un dato correcto ahí se refleja automáticamente en planilla, contratos, cartas y liquidaciones.",
                "Esta presentación recorre, diapositiva por diapositiva, para qué sirve cada sección y qué acciones concretas puedes realizar en ella.",
            ],
        },
        {
            "titulo": "Gestión Personal",
            "resumen": "Registro y administración de empleados",
            "imagen": "gestion_personal",
            "puntos": [
                "Permite crear, editar y eliminar empleados con sus datos completos: nombre, DUI, correo, dirección, teléfono, puesto, departamento, salario, fecha de ingreso y fecha de nacimiento.",
                "Incluye una barra de búsqueda que filtra en tiempo real por nombre o DUI, y un filtro adicional por departamento para localizar rápidamente a un grupo de empleados.",
                "La tabla de resultados se puede reordenar y ajustar por columnas para revisar la información de forma más cómoda.",
                "Es la base de datos maestra del sistema: los módulos de planilla, marcaciones, descuentos, contratos, vacaciones, liquidaciones y cartas obtienen aquí los datos del empleado (nombre, DUI, salario, cargo).",
            ],
        },
        {
            "titulo": "Gestión Planilla",
            "resumen": "Cálculo y control de la nómina",
            "imagen": "gestion_planilla",
            "puntos": [
                "Calcula la planilla de pago de todo el personal a partir del salario base, las horas trabajadas registradas en Gestión de Tiempo, las deducciones legales y los descuentos activos.",
                "Permite filtrar la planilla por departamento y por corte de planilla (período de pago) para revisar solo lo que necesitas.",
                "Incluye un análisis de planilla con totales y comparativos, además del cálculo del décimo tercer mes con su propio comprobante descargable.",
                "Genera comprobantes de pago individuales por empleado, listos para imprimir o descargar, y permite exportar la planilla completa a un archivo CSV.",
                "El corte de planilla organiza el proceso en períodos cerrados, evitando que se dupliquen pagos ya procesados.",
            ],
        },
        {
            "titulo": "Gestión de Tiempo",
            "resumen": "Marcaciones, horarios y asistencia",
            "imagen": "gestion_tiempo",
            "puntos": [
                "Registra la entrada y salida de cada empleado de forma manual (hora, minutos y AM/PM) o de forma automática mediante un lector biométrico conectado por CSV, JSON, REST o SQLite (configurable en Configuración → API Lector).",
                "Calcula automáticamente las horas trabajadas y las horas extra de cada marcación, y detecta tardanzas o ausencias comparando contra los horarios laborales configurados.",
                "Permite editar una marcación ya registrada (corregir hora de entrada/salida) y justificar inasistencias o tardanzas con un motivo, quedando el registro disponible para aprobación.",
                "Las aprobaciones de corte quedan documentadas por usuario, con una confirmación visual, antes de que esas marcaciones pasen a formar parte de la planilla.",
                "Incluye búsqueda de cortes por rango de fechas para auditar la asistencia de un período específico.",
            ],
        },
        {
            "titulo": "Descuentos y Préstamos",
            "resumen": "Control de préstamos y descuentos al personal",
            "imagen": "descuentos_prestamos",
            "puntos": [
                "Registra préstamos o descuentos autorizados a un empleado, con un formulario que valida el monto y la información antes de guardarlo.",
                "Cada descuento se puede activar o desactivar sin necesidad de eliminarlo, para pausar temporalmente su aplicación en la planilla.",
                "Incluye búsqueda avanzada por nombre/DUI y por departamento, además de métricas generales (total de descuentos activos, montos, etc.).",
                "Genera reportes exportables en Excel y PDF con el detalle de todos los descuentos y préstamos registrados.",
            ],
        },
        {
            "titulo": "Deducciones y Retenciones",
            "resumen": "Retenciones legales de nómina",
            "imagen": "deducciones_retenciones",
            "puntos": [
                "Calcula automáticamente el CSS (Seguro Social) y el Seguro Educativo aplicando el porcentaje legal vigente sobre el salario base.",
                "Calcula el Impuesto Sobre la Renta (ISR) usando la fórmula progresiva panameña, mostrando la fórmula aplicada para efectos de transparencia y auditoría.",
                "Permite editar manualmente los porcentajes de CSS y Seguro Educativo si la ley cambia, sin necesidad de modificar código.",
                "Estas deducciones se aplican automáticamente al calcular la planilla de cada empleado.",
            ],
        },
        {
            "titulo": "Remuneraciones",
            "resumen": "Bonificaciones, horas extra y comisiones",
            "imagen": "remuneraciones",
            "puntos": [
                "Administra cuatro tipos de pago adicional al salario base: horas extra, comisiones, bonificaciones y otros conceptos; cada tipo se puede activar o desactivar de forma independiente.",
                "Cuenta con un formulario para asignar una remuneración a un empleado específico, indicando tipo, monto, fecha y motivo.",
                "Muestra una tabla con el historial de remuneraciones otorgadas, con la posibilidad de editar o anular un registro si hubo un error.",
                "Incluye un panel resumen con los totales por tipo de remuneración, útil para revisar el gasto adicional antes de correr la planilla.",
            ],
        },
        {
            "titulo": "Ausencias y Permisos",
            "resumen": "Solicitudes de permiso e inasistencias",
            "imagen": "ausencias_permisos",
            "puntos": [
                "Registra permisos, incapacidades y ausencias justificadas de cada empleado mediante un formulario con búsqueda de empleado y adjunto de archivo de respaldo (constancia médica, etc.).",
                "El panel izquierdo lista a los empleados y el panel derecho muestra el detalle de sus permisos/ausencias registrados, con opción de descargar o abrir el archivo adjunto.",
                "Permite editar o eliminar un registro existente directamente desde la ventana de detalles.",
                "Genera reportes de permisos y ausencias, sirviendo como respaldo legal y afectando el cálculo de asistencia en Gestión de Tiempo y Planilla.",
            ],
        },
        {
            "titulo": "Vacaciones",
            "resumen": "Períodos y pagos de vacaciones",
            "imagen": "vacaciones",
            "puntos": [
                "Calcula automáticamente la antigüedad de cada empleado a partir de su fecha de ingreso, para determinar los días de vacaciones que le corresponden.",
                "Permite abrir una solicitud de vacaciones por empleado, indicando el período a disfrutar.",
                "Incluye una vista para revisar las solicitudes pendientes y su estado (aprobada, pendiente, etc.).",
                "Controla los períodos de vacaciones acumulados y ya disfrutados, y calcula el pago correspondiente cuando aplica.",
            ],
        },
        {
            "titulo": "Expedientes laborales",
            "resumen": "Expedientes laborales",
            "imagen": "contratos_trabajo",
            "puntos": [
                "Presenta una lista de empleados con buscador; al seleccionar uno se muestra el resumen de su contrato: tipo, vigencia, jornada, horario, estado y documento adjunto.",
                "El formulario de contrato permite definir tipo (Indefinido, Definido, Por obra o servicio), jornada, fechas de inicio y fin, horario de entrada/salida y estado (Activo, Finalizado, Suspendido).",
                "Permite adjuntar el documento del contrato (PDF, Word o imagen), el cual queda guardado junto al registro del empleado como parte de su expediente laboral.",
                "Centraliza el expediente laboral requerido para auditorías o inspecciones del Ministerio de Trabajo.",
            ],
        },
        {
            "titulo": "Liquidaciones",
            "resumen": "Cálculo de finiquito al terminar la relación laboral",
            "imagen": "liquidaciones",
            "puntos": [
                "Presenta un panel con la lista de empleados (con buscador) y un formulario para preparar la liquidación del empleado seleccionado.",
                "El formulario solicita fecha de salida, motivo de salida, días trabajados pendientes de pago y días de vacaciones no gozados, entre otros conceptos de ley.",
                "Calcula el finiquito final conforme a la legislación laboral panameña, incluyendo prestaciones, vacaciones proporcionales y demás rubros aplicables.",
                "Conserva un historial de las liquidaciones ya generadas para consulta posterior.",
            ],
        },
        {
            "titulo": "Cartas",
            "resumen": "Generación de cartas y constancias laborales",
            "imagen": "cartas",
            "puntos": [
                "Lista a todos los empleados en una tabla con búsqueda por nombre/DUI y filtro por departamento.",
                "Genera una carta de trabajo profesional para el empleado seleccionado, usando automáticamente los datos ya registrados (nombre, cargo, fecha de ingreso, antigüedad) y los datos de la empresa (nombre, RUC, dirección, logo) configurados en Configuración.",
                "Calcula la antigüedad del empleado a partir de su fecha de ingreso para incluirla en el texto de la carta.",
                "Evita tener que redactar manualmente cada carta, reduciendo errores y ahorrando tiempo administrativo.",
            ],
        },
        {
            "titulo": "Configuración",
            "resumen": "Ajustes generales de la aplicación",
            "imagen": "configuracion",
            "puntos": [
                "Datos de la empresa: nombre, RUC, dirección, teléfono y logo, usados en reportes y cartas generadas.",
                "Idioma de la interfaz (Español/English), horarios laborales (jornadas y límites diarios) y notificaciones (planilla, marcaciones).",
                "Tamaño de texto y opciones visuales (alto contraste, reducir movimiento) para adaptar la interfaz a cada usuario.",
                "API Lector: configura la conexión con un lector biométrico externo (CSV, JSON, REST o SQLite), define las columnas de identificación (DUI, fecha, hora, tipo) y permite probar la conexión y sincronizar marcaciones.",
                "Gestión de usuarios del sistema: crear, identificar y eliminar usuarios con acceso a la aplicación.",
            ],
        },
        {
            "titulo": "Seguridad",
            "resumen": "Revisión de la aplicación y copias de seguridad",
            "imagen": "seguridad",
            "puntos": [
                "Revisión: ejecuta una verificación general del estado de la aplicación y sus componentes, útil para detectar problemas antes de que afecten la operación.",
                "Copia de seguridad: crea un respaldo comprimido (.zip) que incluye el volcado completo de la base de datos (vía pg_dump) y todos los documentos subidos en la carpeta uploads/.",
                "Permite ver la lista de copias existentes (ordenadas de la más reciente a la más antigua) y restaurar cualquiera de ellas, o directamente la última, con una confirmación previa por ser una acción irreversible.",
                "La restauración usa pg_restore para reponer la base de datos y copia de vuelta los archivos adjuntos, dejando el sistema como estaba al momento de esa copia.",
            ],
        },
        {
            "titulo": "Reportes y Ayuda",
            "resumen": "Información de apoyo",
            "imagen": "reportes_ayuda",
            "puntos": [
                "Reportes de tardanzas y ausencias: resume, por empleado y por período, quiénes llegaron tarde o faltaron, con base en las marcaciones y justificaciones registradas.",
                "Información de la aplicación: muestra datos generales de la versión instalada de SiPP.",
                "Buscar actualizaciones: verifica si existe una versión más reciente de la aplicación disponible para instalar.",
            ],
        },
    ]

    CARPETA_IMAGENES = os.path.join(os.path.abspath("."), "imagenes", "manuales")


    def __init__(self):
        self.indice = 0
        self.slide_frame = None
        self.botones_indice = []
        self.etiqueta_contador = None

    def info_manuales(self, frame):
        self.indice = 0
        self.botones_indice = []

        frame_principal = ctk.CTkFrame(frame, corner_radius=12, fg_color=("#eef6fb", "#101b2a"), border_width=1)
        frame_principal.place(x=0, y=0, relwidth=1, relheight=1)

        header = ctk.CTkFrame(frame_principal, corner_radius=12, fg_color=("#d9edf7", "#18263a"), border_width=1, height=92)
        header.place(relx=0.02, rely=0.02, relwidth=0.96)
        ctk.CTkLabel(header, text="Manuales", font=ctk.CTkFont(size=25, weight="bold"), anchor="w").place(x=16, y=10)
        ctk.CTkLabel(
            header,
            text="Presentación guiada: qué hace cada parte de SiPP",
            font=ctk.CTkFont(size=11),
            text_color=("#526071", "#a1a9b8"),
            anchor="w",
        ).place(x=18, y=50)

        cuerpo = ctk.CTkFrame(frame_principal, fg_color="transparent")
        cuerpo.place(relx=0.02, rely=0.18, relwidth=0.96, relheight=0.72)
        cuerpo.grid_columnconfigure(0, weight=0)
        cuerpo.grid_columnconfigure(1, weight=1)
        cuerpo.grid_rowconfigure(0, weight=1)

        indice_frame = ctk.CTkScrollableFrame(cuerpo, corner_radius=10, fg_color=("#ffffff", "#182334"), width=220)
        indice_frame.grid(row=0, column=0, sticky="ns", padx=(0, 14))

        for posicion, diapositiva in enumerate(self.DIAPOSITIVAS):
            boton = ctk.CTkButton(
                indice_frame,
                text=f"{posicion + 1}. {diapositiva['titulo']}",
                anchor="w",
                fg_color="transparent",
                text_color=("#1e293b", "#e2e8f0"),
                hover_color=("#e2e8f0", "#233042"),
                command=lambda posicion=posicion: self.ir_a(posicion),
            )
            boton.pack(fill="x", padx=6, pady=4)
            self.botones_indice.append(boton)

        self.slide_frame = ctk.CTkFrame(cuerpo, corner_radius=10, fg_color=("#ffffff", "#182334"))
        self.slide_frame.grid(row=0, column=1, sticky="nsew")

        navegacion = ctk.CTkFrame(frame_principal, fg_color="transparent")
        navegacion.place(relx=0.02, rely=0.91, relwidth=0.96)
        ctk.CTkButton(navegacion, text="◀ Anterior", width=120, fg_color="#64748b", hover_color="#475569", command=self.anterior).pack(side="left")
        self.etiqueta_contador = ctk.CTkLabel(navegacion, text="", font=ctk.CTkFont(size=12, weight="bold"))
        self.etiqueta_contador.pack(side="left", expand=True)
        ctk.CTkButton(navegacion, text="Siguiente ▶", width=120, fg_color="#2563eb", hover_color="#1d4ed8", command=self.siguiente).pack(side="right")

        self.ir_a(0)

    def ir_a(self, posicion):
        self.indice = max(0, min(posicion, len(self.DIAPOSITIVAS) - 1))
        self._renderizar_slide()

    def anterior(self):
        self.ir_a(self.indice - 1)

    def siguiente(self):
        self.ir_a(self.indice + 1)

    def _renderizar_slide(self):
        if self.slide_frame is None or not self.slide_frame.winfo_exists():
            return
        for widget in self.slide_frame.winfo_children():
            widget.destroy()

        diapositiva = self.DIAPOSITIVAS[self.indice]
        contenedor = ctk.CTkScrollableFrame(self.slide_frame, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)

        ctk.CTkLabel(
            contenedor, text=diapositiva["titulo"], font=ctk.CTkFont(size=22, weight="bold"), anchor="w",
        ).pack(fill="x", padx=24, pady=(24, 4))
        ctk.CTkLabel(
            contenedor, text=diapositiva["resumen"], font=ctk.CTkFont(size=13), anchor="w",
            text_color=("#526071", "#a1a9b8"),
        ).pack(fill="x", padx=24, pady=(0, 14))

        self._renderizar_captura(contenedor, diapositiva["imagen"])

        for punto in diapositiva["puntos"]:
            fila = ctk.CTkFrame(contenedor, fg_color="transparent")
            fila.pack(fill="x", padx=24, pady=6)
            ctk.CTkLabel(fila, text="•", font=ctk.CTkFont(size=14, weight="bold"), width=18).pack(side="left", anchor="n")
            ctk.CTkLabel(
                fila, text=punto, font=ctk.CTkFont(size=13), anchor="w", justify="left", wraplength=640,
            ).pack(side="left", fill="x", expand=True)

        for posicion, boton in enumerate(self.botones_indice):
            if posicion == self.indice:
                boton.configure(fg_color=("#2563eb", "#1d4ed8"), text_color="#ffffff")
            else:
                boton.configure(fg_color="transparent", text_color=("#1e293b", "#e2e8f0"))

        if self.etiqueta_contador is not None:
            self.etiqueta_contador.configure(text=f"{self.indice + 1} / {len(self.DIAPOSITIVAS)}")

    def _renderizar_captura(self, contenedor, nombre_imagen):
        """Muestra la captura de pantalla de la sección si existe en imagenes/manuales/."""
        ruta = os.path.join(self.CARPETA_IMAGENES, f"{nombre_imagen}.png")
        if not os.path.isfile(ruta):
            ctk.CTkLabel(
                contenedor,
                text=f"(Sin captura todavía. Guarda una imagen como imagenes/manuales/{nombre_imagen}.png para que aparezca aquí.)",
                font=ctk.CTkFont(size=10, slant="italic"),
                text_color=("#94a3b8", "#64748b"),
                anchor="w",
            ).pack(fill="x", padx=24, pady=(0, 14))
            return
        try:
            imagen = Image.open(ruta)
            imagen.thumbnail((680, 360), Image.Resampling.LANCZOS)
            imagen_ctk = ctk.CTkImage(light_image=imagen, dark_image=imagen, size=imagen.size)
            etiqueta_imagen = ctk.CTkLabel(contenedor, image=imagen_ctk, text="")
            etiqueta_imagen.image = imagen_ctk
            etiqueta_imagen.pack(padx=24, pady=(0, 14), anchor="w")
        except (OSError, ValueError):
            logging.exception("No se pudo cargar la captura del manual: %s", ruta)


if __name__ == "__main__":
    app = SiPP()
    if not getattr(app, "_inicio_cancelado", False):
        try:
            app.mainloop()
        except KeyboardInterrupt:
            try:
                app.destroy()
            except Exception:
                pass
            print("Aplicación cerrada por interrupción del usuario.")
    else:
        app.destroy()