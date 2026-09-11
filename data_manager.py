"""
Manejo de datos: lectura de maestras y BD de OPs desde red,
definición de campos por grupo de máquina, y escritura al Excel consolidado.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import threading
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s [DataManager] %(message)s')

from config import Config, INTERVALO_RECARGA_MIN

# ── Grupos de máquinas y sus campos ──────────────────────────────────────────
CAMPOS_POR_GRUPO = {

    # Máquinas de impresión
    "impresion": [
        {"name": "op",           "label": "OP",                  "tipo": "op_search",   "col_excel": "OP"},
        {"name": "tipo_trabajo",  "label": "Tipo de Trabajo",     "tipo": "readonly",    "col_excel": "TIPO_TRABAJO"},
        {"name": "cod_referencia","label": "Cod. Referencia",     "tipo": "readonly",    "col_excel": "COD_REFERENCIA"},
        {"name": "referencia",    "label": "Referencia",          "tipo": "readonly",    "col_excel": "REFERENCIA"},
        {"name": "nit_cliente",   "label": "NIT Cliente",         "tipo": "readonly",    "col_excel": "NIT_CLIENTE"},
        {"name": "cliente",       "label": "Cliente",             "tipo": "readonly",    "col_excel": "CLIENTE"},
        {"name": "evento",        "label": "Evento",              "tipo": "combo",       "col_excel": "EVENTO"},
        {"name": "operario",      "label": "Operario",            "tipo": "combo",       "col_excel": "OPERARIO"},
        {"name": "cantidad",      "label": "Cantidad (Tiros/Unds)","tipo": "entry",      "col_excel": "CANTIDAD"},
        {"name": "cavidad",       "label": "Cavidad",             "tipo": "entry",       "col_excel": "CAVIDAD"},
        {"name": "planchas",      "label": "Planchas",            "tipo": "entry",       "col_excel": "PLANCHAS"},
        {"name": "observacion",   "label": "Observación",         "tipo": "text",        "col_excel": "OBSERVACION"},
    ],

    # Grupo general: troqueladoras, pegadoras, formadoras y similares.
    # Es el grupo por defecto para cualquier máquina no listada en MAQUINA_GRUPO.
    "general": [
        {"name": "op",            "label": "OP",                  "tipo": "op_search",   "col_excel": "OP"},
        {"name": "tipo_trabajo",  "label": "Tipo de Trabajo",     "tipo": "readonly",    "col_excel": "TIPO_TRABAJO"},
        {"name": "cod_referencia","label": "Cod. Referencia",     "tipo": "readonly",    "col_excel": "COD_REFERENCIA"},
        {"name": "referencia",    "label": "Referencia",          "tipo": "readonly",    "col_excel": "REFERENCIA"},
        {"name": "nit_cliente",   "label": "NIT Cliente",         "tipo": "readonly",    "col_excel": "NIT_CLIENTE"},
        {"name": "cliente",       "label": "Cliente",             "tipo": "readonly",    "col_excel": "CLIENTE"},
        {"name": "evento",        "label": "Evento",              "tipo": "combo",       "col_excel": "EVENTO"},
        {"name": "operario",      "label": "Operario",            "tipo": "combo",       "col_excel": "OPERARIO"},
        {"name": "cantidad",      "label": "Cantidad (Tiros/Unds)","tipo": "entry",      "col_excel": "CANTIDAD"},
        {"name": "cavidad",       "label": "Cavidad",             "tipo": "entry",       "col_excel": "CAVIDAD"},
        {"name": "observacion",   "label": "Observación",         "tipo": "text",        "col_excel": "OBSERVACION"},
    ],

    # Conversión de bobinas (campos propios de material y corte)
    "conversion": [
        {"name": "op",            "label": "OP",                  "tipo": "op_search",   "col_excel": "OP"},
        {"name": "tipo_trabajo",  "label": "Tipo de Trabajo",     "tipo": "readonly",    "col_excel": "TIPO_TRABAJO"},
        {"name": "cod_referencia","label": "Cod. Referencia",     "tipo": "readonly",    "col_excel": "COD_REFERENCIA"},
        {"name": "referencia",    "label": "Referencia",          "tipo": "readonly",    "col_excel": "REFERENCIA"},
        {"name": "nit_cliente",   "label": "NIT Cliente",         "tipo": "readonly",    "col_excel": "NIT_CLIENTE"},
        {"name": "cliente",       "label": "Cliente",             "tipo": "readonly",    "col_excel": "CLIENTE"},
        {"name": "evento",        "label": "Evento",              "tipo": "combo",       "col_excel": "EVENTO"},
        {"name": "operario",      "label": "Operario",            "tipo": "combo",       "col_excel": "OPERARIO"},
        {"name": "cantidad",      "label": "Cantidad (Tiros/Unds)","tipo": "entry",      "col_excel": "CANTIDAD"},
        {"name": "cavidad",       "label": "Cavidad",             "tipo": "entry",       "col_excel": "CAVIDAD"},
        {"name": "material",      "label": "Material",            "tipo": "entry",       "col_excel": "MATERIAL"},
        {"name": "no_bobina",     "label": "No. Bobina",          "tipo": "entry",       "col_excel": "NO_BOBINA"},
        {"name": "calibre",       "label": "Calibre",             "tipo": "entry",       "col_excel": "CALIBRE"},
        {"name": "ancho",         "label": "Ancho",               "tipo": "entry",       "col_excel": "ANCHO"},
        {"name": "corte",         "label": "Corte",               "tipo": "entry",       "col_excel": "CORTE"},
        {"name": "observacion",   "label": "Observación",         "tipo": "text",        "col_excel": "OBSERVACION"},
    ],
}

# Mapeo COD MAQUINA → grupo de campos.
# El código de cada máquina sale de la maestra de máquinas; aquí solo se indica
# qué formulario le corresponde. Las que no estén listadas usan "general".
MAQUINA_GRUPO = {
    "1": "impresion",
    "2": "impresion",
    "3": "conversion",
    "4": "general",
    "5": "general",
    "6": "general",
}

# Todas las columnas del consolidado (orden final)
COLUMNAS_CONSOLIDADO = [
    "TIMESTAMP_REGISTRO",
    "FECHA", "TURNO", "FECHA_TURNO",
    "COD_MAQUINA", "MAQUINA", "UBICACION",
    "COD_AREA", "AREA",
    "OP",
    "TIPO_TRABAJO",
    "COD_REFERENCIA", "REFERENCIA",
    "NIT_CLIENTE", "CLIENTE",
    "COD_EVENTO", "EVENTO",
    "COD_OPERARIO", "OPERARIO",
    "FECHA_INICIAL", "HORA_INICIAL", "FECHA_FINAL", "HORA_FINAL", "MINUTOS",
    "CANTIDAD", "CAVIDAD", "CANTIDAD_FINALIZADA",
    "PLANCHAS",
    "MATERIAL", "NO_BOBINA", "CALIBRE", "ANCHO", "CORTE",
    "OBSERVACION",
]

_excel_lock = threading.Lock()


def calcular_turno(fecha_ini: str, hora_ini: str) -> tuple:
    try:
        from datetime import datetime, timedelta
        dt = datetime.strptime(f"{fecha_ini} {hora_ini}", "%Y-%m-%d %H:%M:%S")
        h = dt.hour
        if 6 <= h < 14:
            turno = "MAÑANA"
            fecha_turno = dt.strftime("%Y-%m-%d")
        elif 14 <= h < 22:
            turno = "TARDE"
            fecha_turno = dt.strftime("%Y-%m-%d")
        else:
            turno = "NOCHE"
            # Si es madrugada (00:00-05:59) la fecha del turno es el día anterior
            if h < 6:
                fecha_turno = (dt - timedelta(days=1)).strftime("%Y-%m-%d")
            else:  # 22:00-23:59
                fecha_turno = dt.strftime("%Y-%m-%d")
        return turno, fecha_turno
    except:
        return "", ""


class DataManager:

    def __init__(self, config: Config):
        self.config   = config
        self.maestras = {}
        self.ops      = {}
        self._lock    = threading.Lock()
        self._recargar()
        self._iniciar_recarga_automatica()

    # ── Carga de datos ────────────────────────────────────────────────────────

    def _recargar(self):
        logging.info("Recargando maestras y OPs desde red...")
        nuevas_maestras = self._cargar_maestras()   # lanza ConnectionError si falla
        nuevas_ops      = self._cargar_ops()
        with self._lock:
            self.maestras = nuevas_maestras
            self.ops      = nuevas_ops
        logging.info(f"  → {len(self.ops)} OPs | "
                     f"{len(self.maestras.get('eventos', []))} eventos | "
                     f"{len(self.maestras.get('operarios', []))} operarios cargados")

    def _cargar_maestras(self):
        ruta = self.config.get_ruta_maestras()
        result = {k: [] for k in ("maquinas", "areas", "eventos", "operarios")}

        if not os.path.exists(ruta):
            raise ConnectionError(
                f"No se puede acceder a la red.\n\n"
                f"Ruta no disponible:\n{ruta}\n\n"
                f"Verifica que el PC esté conectado a la red de la planta\n"
                f"y vuelve a intentarlo. Si el problema persiste, llama al encargado."
            )

        try:
            wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
            result["maquinas"]  = self._sheet_to_dicts(wb, "MAESTRA MAQUINAS")
            result["areas"]     = self._sheet_to_dicts(wb, "MAESTRA AREAS")
            result["eventos"]   = self._sheet_to_dicts(wb, "MAESTRA EVENTOS")
            result["operarios"] = self._sheet_to_dicts(wb, "MAESTRA OPERARIOS")
            wb.close()
        except ConnectionError:
            raise
        except Exception as e:
            raise ConnectionError(
                f"Error al leer las maestras desde la red.\n\n"
                f"Detalle técnico: {e}\n\n"
                f"Llama al encargado."
            )
        return result

    def _cargar_ops(self):
        data = {}
        ruta = self.config.get_ruta_op()

        if not os.path.exists(ruta):
            logging.warning(f"No se encontró BD de OPs en: {ruta}")
            return data

        try:
            wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
            ws = wb.active

            # Detectar fila de encabezados en las primeras 5 filas
            headers = []
            header_row = 1
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                row_vals = [str(v).strip() if v else "" for v in row]
                if "OP" in row_vals:
                    headers = row_vals
                    header_row = i
                    break

            if not headers:
                logging.warning("No se encontró la columna 'OP' en la base de órdenes de producción")
                wb.close()
                return data

            def idx(nombres):
                for nombre in (nombres if isinstance(nombres, list) else [nombres]):
                    for i, h in enumerate(headers):
                        if nombre.lower() in h.lower():
                            return i
                return None

            idx_op   = idx("OP")
            idx_tipo = idx(["Tipo de Trabajo", "TIPO DE TRABAJO"])
            idx_ref  = idx(["Cód. Producto", "Cod. Producto", "COD PRODUCTO", "COD_REFERENCIA"])
            idx_nom  = idx(["Referencia", "REFERENCIA"])
            idx_nit  = idx(["Doc. Ident", "NIT", "DOC IDENT"])

            # Cliente: buscar coincidencia exacta para evitar confusión con "O.C. Cliente"
            idx_cli = None
            for i, h in enumerate(headers):
                if h.strip().upper() == "CLIENTE":
                    idx_cli = i
                    break

            logging.info(f"Columnas OP → OP:{idx_op} | Tipo:{idx_tipo} | "
                         f"CodRef:{idx_ref} | Ref:{idx_nom} | Cliente:{idx_cli} | NIT:{idx_nit}")

            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                if idx_op is not None and row[idx_op]:
                    op_key = str(row[idx_op]).strip()
                    data[op_key] = {
                        "tipo_trabajo":  str(row[idx_tipo]).strip()  if idx_tipo is not None and row[idx_tipo]  else "",
                        "cod_referencia":str(row[idx_ref]).strip()   if idx_ref  is not None and row[idx_ref]   else "",
                        "referencia":    str(row[idx_nom]).strip()   if idx_nom  is not None and row[idx_nom]   else "",
                        "cliente":       str(row[idx_cli]).strip()   if idx_cli  is not None and row[idx_cli]   else "",
                        "nit_cliente":   str(row[idx_nit]).strip()   if idx_nit  is not None and row[idx_nit]   else "",
                    }
            wb.close()
        except Exception as e:
            logging.error(f"Error leyendo la base de órdenes de producción: {e}")
        return data

    def _iniciar_recarga_automatica(self):
        def loop():
            while True:
                threading.Event().wait(INTERVALO_RECARGA_MIN * 60)
                try:
                    self._recargar()
                except ConnectionError as e:
                    logging.warning(f"Recarga automática fallida (red no disponible): {e}")
                except Exception as e:
                    logging.error(f"Error en recarga automática: {e}")
        t = threading.Thread(target=loop, daemon=True)
        t.start()
        logging.info(f"Recarga automática configurada cada {INTERVALO_RECARGA_MIN} minutos")

    @staticmethod
    def _sheet_to_dicts(wb, sheet_name):
        if sheet_name not in wb.sheetnames:
            return []
        ws   = wb[sheet_name]
        rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
        if not rows:
            return []
        headers = [str(h).strip() if h else f"col{i}" for i, h in enumerate(rows[0])]
        result  = []
        for row in rows[1:]:
            if any(v is not None for v in row):
                result.append({headers[i]: (str(row[i]).strip() if row[i] is not None else "")
                                for i in range(len(headers))})
        return result

    def reload(self):
        """Recarga forzada (llamada antes de guardar cada registro)."""
        self._recargar()

    # ── Máquina ───────────────────────────────────────────────────────────────

    def get_maquina_info(self, cod_maquina: str) -> dict:
        for m in self.maestras.get("maquinas", []):
            if str(m.get("COD MAQUINA", "")).strip() == str(cod_maquina).strip():
                info = dict(m)
                info["grupo"] = MAQUINA_GRUPO.get(str(cod_maquina).strip(), "general")
                for a in self.maestras.get("areas", []):
                    if str(a.get("COD AREA", "")) == str(m.get("COD AREA", "")):
                        info["AREA"] = a.get("AREA", "")
                        break
                return info
        return {"grupo": "general", "MAQUINA": f"Máquina {cod_maquina}", "AREA": ""}

    def get_campos(self, grupo: str) -> list:
        return CAMPOS_POR_GRUPO.get(grupo, CAMPOS_POR_GRUPO["general"])

    # ── Búsqueda de OP ────────────────────────────────────────────────────────

    def buscar_op(self, op: str) -> dict:
        """Retorna los datos autocomplete de una OP."""
        return self.ops.get(str(op).strip(), {})

    def get_lista_ops(self) -> list:
        """Retorna lista de códigos de OP para autocomplete."""
        return sorted(self.ops.keys())

    # ── Guardar registro ──────────────────────────────────────────────────────

    def _preparar_registro(self, registro: dict) -> dict:
        """Completa el registro con todos los campos derivados (máquina, evento,
        operario, turno, cantidad finalizada). NO escribe nada: solo prepara."""
        registro["TIMESTAMP_REGISTRO"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        registro["FECHA"] = datetime.now().strftime("%Y-%m-%d")

        # ── Máquina ──────────────────────────────────────────────────────────
        maq_info = self.get_maquina_info(str(registro.get("MAQUINA", "")))
        registro["COD_MAQUINA"] = maq_info.get("COD MAQUINA", "")
        registro["MAQUINA"]     = maq_info.get("MAQUINA", "")
        registro["UBICACION"]   = maq_info.get("UBICACIÓN", "")
        registro["COD_AREA"]    = maq_info.get("COD AREA", "")
        registro["AREA"]        = maq_info.get("AREA", "")

        # ── Evento ────────────────────────────────────────────────────────────
        val = registro.get("EVENTO", "")
        if " - " in str(val):
            parts = val.split(" - ", 1)
            registro["COD_EVENTO"] = parts[0].strip()
            registro["EVENTO"]     = parts[1].strip()
        else:
            registro["COD_EVENTO"] = val
            for e in self.maestras.get("eventos", []):
                if e.get("COD EVENTO", "") == val:
                    registro["EVENTO"] = e.get("EVENTO", val); break

        # ── Operario ──────────────────────────────────────────────────────────
        val = registro.get("OPERARIO", "")
        if " - " in str(val):
            parts = val.split(" - ", 1)
            registro["COD_OPERARIO"] = parts[0].strip()
            registro["OPERARIO"]     = parts[1].strip()
        else:
            registro["COD_OPERARIO"] = val
            for o in self.maestras.get("operarios", []):
                if o.get("ID", "") == val:
                    registro["OPERARIO"] = o.get("OPERARIO", val); break

        # ── Turno ─────────────────────────────────────────────────────────────
        turno, fecha_turno = calcular_turno(
            registro.get("FECHA_INICIAL", ""),
            registro.get("HORA_INICIAL", "")
        )
        registro["TURNO"]       = turno
        registro["FECHA_TURNO"] = fecha_turno

        # ── Cantidad finalizada ───────────────────────────────────────────────
        try:
            cant = float(registro.get("CANTIDAD", 0) or 0)
            cav  = float(registro.get("CAVIDAD", 0) or 0)
            registro["CANTIDAD_FINALIZADA"] = int(cant * cav) if cant and cav else ""
        except (ValueError, TypeError):
            registro["CANTIDAD_FINALIZADA"] = ""

        return registro

    def guardar_registro(self, registro: dict):
        """OPCIÓN A: en vez de escribir el Excel, deja un archivito JSON único en la
        carpeta PENDIENTES de la red. Si la red no está disponible, lo guarda en una
        carpeta local del PC y se subirá en el próximo guardado con red. Así 16
        máquinas nunca chocan: cada una escribe SU propio archivo, nunca el Excel."""
        import json, uuid

        # Asegura datos frescos para los campos derivados (evento/operario/máquina)
        try:
            self._recargar()
        except ConnectionError:
            # Si las maestras no cargan, igual preparamos con lo que haya en memoria
            logging.warning("No se pudieron recargar maestras antes de guardar; uso cache.")

        registro = self._preparar_registro(registro)

        # Nombre del archivo: NOMBREMAQUINA_CODIGO_DDMMYYYY_HHMMSS_xxxxxx.json
        # Ej: PEGADORA1_04_28052026_143207_a8f3.json
        # - El nombre y código son legibles a simple vista (sabes de quién es).
        # - La fecha y hora ordenan el archivo cronológicamente.
        # - El sufijo aleatorio garantiza que JAMÁS dos archivos tengan el mismo
        #   nombre, ni aunque dos máquinas guarden en el mismo segundo.
        nombre_maq = str(registro.get("MAQUINA", "") or "MAQ").upper()
        nombre_maq = "".join(c for c in nombre_maq if c.isalnum()) or "MAQ"
        cod_maq    = str(registro.get("COD_MAQUINA", "") or "NA")
        cod_maq    = "".join(c for c in cod_maq if c.isalnum()) or "NA"
        fecha_str  = datetime.now().strftime("%d%m%Y_%H%M%S")
        sufijo     = uuid.uuid4().hex[:6]
        nombre     = f"{nombre_maq}_{cod_maq}_{fecha_str}_{sufijo}.json"

        contenido = json.dumps(registro, ensure_ascii=False, indent=2)

        # 1) Primero intentamos subir cualquier respaldo local pendiente
        self._subir_respaldos_locales()

        # 2) Escribimos el archivito de este registro
        carpeta_red = self.config.get_carpeta_pendientes()
        try:
            os.makedirs(carpeta_red, exist_ok=True)
            self._escribir_atomico(os.path.join(carpeta_red, nombre), contenido)
            logging.info(f"Registro guardado en PENDIENTES: {nombre}")
        except Exception as e:
            # Red caída → respaldo local, no se pierde el dato
            logging.warning(f"Red no disponible al guardar ({e}). Guardo respaldo local.")
            carpeta_local = self.config.get_carpeta_respaldo_local()
            try:
                os.makedirs(carpeta_local, exist_ok=True)
                self._escribir_atomico(os.path.join(carpeta_local, nombre), contenido)
                logging.info(f"Registro guardado en respaldo local: {nombre}")
            except Exception as e2:
                raise ConnectionError(
                    "No se pudo guardar el registro.\n"
                    "Revisa la conexión de red y avisa al encargado."
                ) from e2

    def _subir_respaldos_locales(self):
        """Si hay archivitos en el respaldo local (de cuando se cayó la red), los
        sube a PENDIENTES y los borra de local. Silencioso si no hay red todavía."""
        carpeta_local = self.config.get_carpeta_respaldo_local()
        if not os.path.isdir(carpeta_local):
            return
        carpeta_red = self.config.get_carpeta_pendientes()
        try:
            os.makedirs(carpeta_red, exist_ok=True)
        except Exception:
            return  # red sigue caída; lo intentamos la próxima vez
        for fn in os.listdir(carpeta_local):
            if not fn.lower().endswith(".json"):
                continue
            origen = os.path.join(carpeta_local, fn)
            try:
                with open(origen, "r", encoding="utf-8") as f:
                    contenido = f.read()
                self._escribir_atomico(os.path.join(carpeta_red, fn), contenido)
                os.remove(origen)
                logging.info(f"Respaldo local subido a PENDIENTES: {fn}")
            except Exception as e:
                logging.warning(f"No se pudo subir respaldo local {fn}: {e}")
                return  # si falla uno, paramos; reintentamos luego

    @staticmethod
    def _escribir_atomico(ruta_destino, contenido: str):
        """Escribe primero a un archivo temporal y luego renombra. El rename es
        atómico, así el consolidador nunca lee un archivo a medio escribir."""
        tmp = ruta_destino + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(contenido)
        os.replace(tmp, ruta_destino)

    def _write_header(self, ws):
        headers_display = {
            "TIMESTAMP_REGISTRO":  "Timestamp",
            "FECHA":               "Fecha",
            "TURNO":               "Turno",
            "FECHA_TURNO":         "Fecha Turno",
            "COD_MAQUINA":         "Cod Maquina",
            "MAQUINA":             "Maquina",
            "UBICACION":           "Ubicación",
            "COD_AREA":            "Cod Area",
            "AREA":                "Area",
            "OP":                  "OP",
            "TIPO_TRABAJO":        "Tipo Trabajo",
            "COD_REFERENCIA":      "Cod Referencia",
            "REFERENCIA":          "Referencia",
            "NIT_CLIENTE":         "NIT Cliente",
            "CLIENTE":             "Cliente",
            "COD_EVENTO":          "Cod Evento",
            "EVENTO":              "Evento",
            "COD_OPERARIO":        "Cod Operario",
            "OPERARIO":            "Operario",
            "FECHA_INICIAL":       "Fecha Ini",
            "HORA_INICIAL":        "Hora Ini",
            "FECHA_FINAL":         "Fecha Fin",
            "HORA_FINAL":          "Hora Fin",
            "MINUTOS":             "Minutos",
            "CANTIDAD":            "Cantidad (Tiros/Unds)",
            "CAVIDAD":             "Cavidad",
            "CANTIDAD_FINALIZADA": "Cant. Final.",
            "PLANCHAS":            "Planchas",
            "MATERIAL":            "Material",
            "NO_BOBINA":           "No. Bobina",
            "CALIBRE":             "Calibre",
            "ANCHO":               "Ancho",
            "CORTE":               "Corte",
            "OBSERVACION":         "Observación",
        }
        header_font = Font(name="Calibri", size=10, bold=True)
        header_fill = PatternFill("solid", fgColor="1A5C2A")
        header_font_color = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        for i, col in enumerate(COLUMNAS_CONSOLIDADO, 1):
            cell = ws.cell(row=1, column=i, value=headers_display.get(col, col))
            cell.font      = header_font_color
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        widths = {
            "TIMESTAMP_REGISTRO": 20, "FECHA": 12, "TURNO": 10, "FECHA_TURNO": 13,
            "COD_MAQUINA": 12, "MAQUINA": 14, "UBICACION": 10,
            "COD_AREA": 10, "AREA": 14,
            "OP": 10, "TIPO_TRABAJO": 18,
            "COD_REFERENCIA": 16, "REFERENCIA": 40,
            "NIT_CLIENTE": 16, "CLIENTE": 35,
            "COD_EVENTO": 12, "EVENTO": 28,
            "COD_OPERARIO": 14, "OPERARIO": 22,
            "FECHA_INICIAL": 13, "HORA_INICIAL": 11,
            "FECHA_FINAL": 13, "HORA_FINAL": 11,
            "MINUTOS": 10, "CANTIDAD": 22, "CAVIDAD": 9, "CANTIDAD_FINALIZADA": 14,
            "PLANCHAS": 10,
            "MATERIAL": 14, "NO_BOBINA": 12, "CALIBRE": 9, "ANCHO": 9, "CORTE": 9,
            "OBSERVACION": 35,
        }
        for i, col in enumerate(COLUMNAS_CONSOLIDADO, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths.get(col, 12)

        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "A2"

    # ── Historial ─────────────────────────────────────────────────────────────

    def get_history(self, limit=50) -> list:
        """Devuelve los últimos registros de ESTA máquina, combinando:
        1) lo que ya está volcado en el Excel consolidado, y
        2) lo que esta máquina tiene pendiente (PENDIENTES de red + respaldo local),
        para que el operario vea su registro al instante aunque el consolidador
        todavía no lo haya pasado al Excel."""
        cfg        = self.config.get_machine_config()
        cod_maq    = str(cfg.get("maquina", "")).strip()
        nombre_maq = str(self.get_maquina_info(cod_maq).get("MAQUINA", "")).strip()

        # --- 2) Pendientes de esta máquina (se muestran primero, son los más nuevos)
        pendientes = self._leer_pendientes_de_maquina(cod_maq, nombre_maq)

        # --- 1) Excel consolidado
        result = list(pendientes)
        ruta = self.config.get_ruta_excel()
        if os.path.exists(ruta):
            try:
                wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                wb.close()
                if len(rows) >= 2:
                    headers = [str(h) if h else "" for h in rows[0]]
                    for row in reversed(rows[1:]):
                        d = {headers[i]: (str(row[i]) if row[i] is not None else "")
                             for i in range(len(headers))}
                        fila_cod    = str(d.get("Cod Maquina", "")).strip()
                        fila_nombre = str(d.get("Maquina", "")).strip()
                        if (cod_maq and fila_cod == cod_maq) or (nombre_maq and fila_nombre == nombre_maq):
                            result.append(d)
                        if len(result) >= limit:
                            break
            except Exception as e:
                logging.warning(f"No se pudo leer el Excel para historial: {e}")
        return result[:limit]

    def _leer_pendientes_de_maquina(self, cod_maq: str, nombre_maq: str) -> list:
        """Lee los archivitos JSON aún no consolidados de esta máquina (red + local)
        y los devuelve en el mismo formato de columnas que el Excel ('Hora Ini', etc.)."""
        import json
        carpetas = [self.config.get_carpeta_pendientes(),
                    self.config.get_carpeta_respaldo_local()]
        registros = []
        for carpeta in carpetas:
            if not carpeta or not os.path.isdir(carpeta):
                continue
            try:
                archivos = [a for a in os.listdir(carpeta) if a.lower().endswith(".json")]
            except Exception:
                continue
            for fn in archivos:
                try:
                    with open(os.path.join(carpeta, fn), "r", encoding="utf-8") as f:
                        r = json.load(f)
                except Exception:
                    continue
                fila_cod    = str(r.get("COD_MAQUINA", "")).strip()
                fila_nombre = str(r.get("MAQUINA", "")).strip()
                if not ((cod_maq and fila_cod == cod_maq) or
                        (nombre_maq and fila_nombre == nombre_maq)):
                    continue
                # Mapear nombres internos → encabezados que usa el panel de historial
                registros.append({
                    "Hora Ini": r.get("HORA_INICIAL", ""),
                    "OP":       r.get("OP", ""),
                    "Evento":   r.get("EVENTO", ""),
                    "Minutos":  r.get("MINUTOS", ""),
                    "Maquina":  r.get("MAQUINA", ""),
                    "_ts":      r.get("TIMESTAMP_REGISTRO", ""),
                })
        # Más recientes primero
        registros.sort(key=lambda x: x.get("_ts", ""), reverse=True)
        return registros

