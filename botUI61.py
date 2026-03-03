import threading
import traceback
from datetime import date, datetime
from pathlib import Path
from typing import Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Keep these defaults hardcoded in the UI so it can start even if task deps are missing.
EXCEL_INGRESO_DEMANDAS_DEFAULT = (
    "C:\\Applications\\RPA 06 - INGRESO DE DEMANDAS Y DOCUMENTOS EN PODER JUDICIAL\\input\\Itau_ddas_pjud\\BOT_MATRIZ_DEMANDAS.xlsx"
)
CARATULAS_FOLDER_PATH_DEFAULT = (
    "C:\\Applications\\RPA 06 - INGRESO DE DEMANDAS Y DOCUMENTOS EN PODER JUDICIAL\\input\\Itau_ddas_pjud\\Caratulas"
)
FECHA_FILTRO_CARATULAS_DEFAULT = "24/02/2026"


class BotUI61:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("BOT RPA 06.1 - Descarga de Caratulas")
        self.root.minsize(980, 500)
        self.root.configure(bg="#eef3f9")
        self._configure_styles()

        default_date = self._parse_default_date(FECHA_FILTRO_CARATULAS_DEFAULT)

        self.excel_ingreso_var = tk.StringVar(value=EXCEL_INGRESO_DEMANDAS_DEFAULT)
        self.excel_informe_var = tk.StringVar(value="")
        self.caratulas_folder_var = tk.StringVar(value=CARATULAS_FOLDER_PATH_DEFAULT)
        self.day_var = tk.StringVar(value=f"{default_date.day:02d}")
        self.month_var = tk.StringVar(value=f"{default_date.month:02d}")
        self.year_var = tk.StringVar(value=f"{default_date.year:04d}")
        self.status_var = tk.StringVar(value="Listo para ejecutar.")
        self.run_button = None
        self.download_informe_button = None

        self._build_ui()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background="#eef3f9")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure(
            "Title.TLabel",
            background="#ffffff",
            foreground="#16243a",
            font=("Segoe UI", 14, "bold"),
        )
        style.configure(
            "Body.TLabel",
            background="#ffffff",
            foreground="#3a4b64",
            font=("Segoe UI", 10),
        )
        style.configure(
            "FieldLabel.TLabel",
            background="#ffffff",
            foreground="#1f2f4a",
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "Status.TLabel",
            background="#ffffff",
            foreground="#2f3f57",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            foreground="#ffffff",
            background="#1e63d5",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#174ea8"), ("disabled", "#97b3e6")],
            foreground=[("disabled", "#f5f7fb")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(12, 6),
        )
        style.configure(
            "AltPrimary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            foreground="#ffffff",
            background="#0f8a6c",
            borderwidth=0,
        )
        style.map(
            "AltPrimary.TButton",
            background=[("active", "#0c6b54"), ("disabled", "#8ec8ba")],
            foreground=[("disabled", "#f5f7fb")],
        )

    def _build_ui(self) -> None:
        app = ttk.Frame(self.root, style="App.TFrame", padding=16)
        app.pack(fill="both", expand=True)

        info_card = ttk.Frame(app, style="Card.TFrame", padding=18)
        info_card.pack(fill="x", pady=(0, 12))

        ttk.Label(
            info_card,
            text="BOT RPA 06.1 | Descarga de Caratulas PJUD",
            style="Title.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            info_card,
            style="Body.TLabel",
            wraplength=920,
            text=(
                "Este bot ingresa a PJUD, filtra 'Demandas Enviadas' por la fecha seleccionada, "
                "descarga las caratulas PDF y genera un archivo final 'CaratulasUnidas.pdf' "
                "en la carpeta de caratulas."
            ),
        ).pack(anchor="w", pady=(8, 0))

        form_card = ttk.Frame(app, style="Card.TFrame", padding=18)
        form_card.pack(fill="both", expand=True)

        ttk.Label(
            form_card, text="Excel Ingreso Demandas (hardcoded, solo lectura)", style="FieldLabel.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Entry(
            form_card, textvariable=self.excel_ingreso_var, state="readonly", width=110
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 12))

        ttk.Label(form_card, text="Excel Informe PJUD (obligatorio)", style="FieldLabel.TLabel").grid(
            row=2, column=0, sticky="w"
        )
        ttk.Entry(form_card, textvariable=self.excel_informe_var, width=90).grid(
            row=3, column=0, sticky="ew", pady=(4, 12)
        )
        ttk.Button(
            form_card,
            text="Seleccionar archivo .xlsx",
            style="Secondary.TButton",
            command=self._browse_excel_informe,
        ).grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(4, 12))

        ttk.Label(
            form_card, text="Carpeta Caratulas (hardcoded, solo lectura)", style="FieldLabel.TLabel"
        ).grid(row=4, column=0, sticky="w")
        ttk.Entry(
            form_card, textvariable=self.caratulas_folder_var, state="readonly", width=110
        ).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 12))

        ttk.Label(form_card, text="Fecha de filtro", style="FieldLabel.TLabel").grid(row=6, column=0, sticky="w")
        date_frame = ttk.Frame(form_card, style="Card.TFrame")
        date_frame.grid(row=7, column=0, sticky="w", pady=(4, 16))

        ttk.Combobox(
            date_frame,
            textvariable=self.day_var,
            values=[f"{i:02d}" for i in range(1, 32)],
            width=4,
            state="readonly",
        ).grid(row=0, column=0)
        ttk.Label(date_frame, text="/", style="Body.TLabel").grid(row=0, column=1, padx=4)
        ttk.Combobox(
            date_frame,
            textvariable=self.month_var,
            values=[f"{i:02d}" for i in range(1, 13)],
            width=4,
            state="readonly",
        ).grid(row=0, column=2)
        ttk.Label(date_frame, text="/", style="Body.TLabel").grid(row=0, column=3, padx=4)
        current_year = date.today().year
        ttk.Combobox(
            date_frame,
            textvariable=self.year_var,
            values=[str(y) for y in range(current_year - 5, current_year + 6)],
            width=6,
            state="readonly",
        ).grid(row=0, column=4)

        ttk.Label(
            form_card,
            style="Body.TLabel",
            text=(
                "Accion 1: descarga caratulas desde PJUD para la fecha elegida, renombra archivos "
                "y genera CaratulasUnidas.pdf."
            ),
            wraplength=920,
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(
            form_card,
            style="Body.TLabel",
            text=(
                "Accion 2: exporta el Excel 'Demandas Enviadas' desde PJUD. "
                "Ese archivo tambien se descarga en la carpeta de caratulas mostrada arriba."
            ),
            wraplength=920,
        ).grid(row=9, column=0, columnspan=2, sticky="w", pady=(0, 10))

        actions_frame = ttk.Frame(form_card, style="Card.TFrame")
        actions_frame.grid(row=10, column=0, columnspan=2, sticky="w")

        self.run_button = ttk.Button(
            actions_frame,
            text="Iniciar descarga de caratulas desde PJUD",
            style="Primary.TButton",
            command=self._on_run_clicked,
        )
        self.run_button.grid(row=0, column=0, sticky="w")

        self.download_informe_button = ttk.Button(
            actions_frame,
            text="Descargar Informe PJUD (Excel)",
            style="AltPrimary.TButton",
            command=self._on_download_informe_clicked,
        )
        self.download_informe_button.grid(row=0, column=1, sticky="w", padx=(10, 0))

        ttk.Label(form_card, textvariable=self.status_var, style="Status.TLabel").grid(
            row=11, column=0, columnspan=2, sticky="w", pady=(12, 0)
        )

        form_card.columnconfigure(0, weight=1)
        form_card.columnconfigure(1, weight=0)

    def _browse_excel_informe(self) -> None:
        selected = filedialog.askopenfilename(
            title="Seleccionar Excel Informe PJUD",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if selected:
            self.excel_informe_var.set(selected)

    def _parse_default_date(self, value: str) -> date:
        try:
            return datetime.strptime(value, "%d/%m/%Y").date()
        except ValueError:
            return date.today()

    def _get_selected_date(self) -> str:
        day = int(self.day_var.get())
        month = int(self.month_var.get())
        year = int(self.year_var.get())
        validated = date(year, month, day)
        return validated.strftime("%d/%m/%Y")

    def _validate_inputs(self) -> Tuple[bool, str]:
        excel_informe = self.excel_informe_var.get().strip()
        if not excel_informe:
            return False, "Debes seleccionar 'Excel Informe PJUD'."

        path_excel = Path(excel_informe)
        if path_excel.suffix.lower() != ".xlsx":
            return False, "El archivo de 'Excel Informe PJUD' debe ser .xlsx."
        if not path_excel.exists():
            return False, f"No existe el archivo: {path_excel}"

        path_ingreso = Path(self.excel_ingreso_var.get().strip())
        if not path_ingreso.exists():
            return (
                False,
                "No existe el archivo hardcoded 'Excel Ingreso Demandas'. "
                f"Ruta actual: {path_ingreso}",
            )

        path_caratulas = Path(self.caratulas_folder_var.get().strip())
        if not path_caratulas.exists() or not path_caratulas.is_dir():
            return (
                False,
                "No existe la carpeta hardcoded 'Carpeta Caratulas'. "
                f"Ruta actual: {path_caratulas}",
            )

        try:
            self._get_selected_date()
        except ValueError:
            return False, "La fecha seleccionada no es valida."

        return True, ""

    def _on_run_clicked(self) -> None:
        valid, error_message = self._validate_inputs()
        if not valid:
            messagebox.showerror("Datos invalidos", error_message)
            return

        self._set_running_state(True)
        self.status_var.set("Ejecutando bot... Esto puede tardar varios minutos.")

        thread = threading.Thread(target=self._run_bot, daemon=True)
        thread.start()

    def _on_download_informe_clicked(self) -> None:
        path_caratulas = Path(self.caratulas_folder_var.get().strip())
        if not path_caratulas.exists() or not path_caratulas.is_dir():
            messagebox.showerror(
                "Datos invalidos",
                "No existe la carpeta hardcoded 'Carpeta Caratulas'. "
                f"Ruta actual: {path_caratulas}",
            )
            return

        self._set_running_state(True)
        self.status_var.set("Descargando Informe PJUD... Esto puede tardar varios minutos.")

        thread = threading.Thread(target=self._run_download_informe, daemon=True)
        thread.start()

    def _run_bot(self) -> None:
        try:
            # Lazy import to avoid crashing the UI at startup when opening by double-click.
            from tasks import run_get_caratulas

            run_get_caratulas(
                excel_ingreso_demandas=self.excel_ingreso_var.get().strip(),
                excel_informe_pjud=self.excel_informe_var.get().strip(),
                caratulas_folder_path=self.caratulas_folder_var.get().strip(),
                fecha_filtro=self._get_selected_date(),
            )
        except Exception as exc:
            error_detail = f"{exc}\n\n{traceback.format_exc()}"
            self.root.after(
                0, lambda: self._handle_finish(False, f"Error al ejecutar el bot:\n{error_detail}")
            )
            return

        self.root.after(0, lambda: self._handle_finish(True, "Ejecucion finalizada."))

    def _run_download_informe(self) -> None:
        try:
            # Lazy import to avoid crashing the UI at startup when opening by double-click.
            from tasks import RPA_06_DescargarInformePjud

            RPA_06_DescargarInformePjud()
        except Exception as exc:
            error_detail = f"{exc}\n\n{traceback.format_exc()}"
            self.root.after(
                0,
                lambda: self._handle_finish(
                    False, f"Error al descargar Informe PJUD:\n{error_detail}"
                ),
            )
            return

        self.root.after(
            0,
            lambda: self._handle_finish(
                True,
                "Informe PJUD descargado. Revisa la carpeta de caratulas mostrada en esta UI.",
            ),
        )

    def _set_running_state(self, is_running: bool) -> None:
        state = "disabled" if is_running else "normal"
        if self.run_button is not None:
            self.run_button.configure(state=state)
        if self.download_informe_button is not None:
            self.download_informe_button.configure(state=state)

    def _handle_finish(self, success: bool, message: str) -> None:
        self._set_running_state(False)
        if success:
            self.status_var.set("Proceso terminado.")
            messagebox.showinfo("BOT RPA 06.1", message)
        else:
            self.status_var.set("Proceso finalizo con error.")
            messagebox.showerror("BOT RPA 06.1", message)


def main() -> None:
    root = tk.Tk()
    BotUI61(root)
    root.mainloop()


if __name__ == "__main__":
    main()
