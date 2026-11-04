from pathlib import Path
import argparse
import re
import subprocess
import sys
import tempfile
from datetime import datetime
import shutil

from openpyxl import load_workbook
from playwright.sync_api import (
    BrowserContext,
    Page,
    sync_playwright,
)


ROOT = Path(__file__).resolve().parent

REPORT_URL = (
    "https://app.powerbi.com/groups/me/reports/"
    "fc4a9202-ac78-405d-bc5b-5ed1d59e738a/"
    "e4a0d105e89f296ddf02"
    "?experience=power-bi"
)

PROFILE_DIR = ROOT / "perfil_powerbi"

ALLOWED_ORGANIZATIONS = {"MMC", "TFC", "MMB"}

def validate_download(path: Path) -> None:
    """Validar el RAW sin mantener abierto el archivo físico."""
    from io import BytesIO

    # El archivo físico se abre y cierra dentro de read_bytes().
    file_bytes = path.read_bytes()

    with BytesIO(file_bytes) as memory_file:
        workbook = load_workbook(
            memory_file,
            read_only=False,
            data_only=True,
        )

        try:
            if "Export" in workbook.sheetnames:
                sheet = workbook["Export"]
            else:
                sheet = workbook.worksheets[0]

            rows = sheet.iter_rows(values_only=True)
            first_row = next(rows, None)

            if first_row is None:
                raise ValueError(
                    "El archivo descargado está vacío."
                )

            headers = [
                str(value or "").strip()
                for value in first_row
            ]

            for required in ("ORG", "ITEM"):
                if required not in headers:
                    raise ValueError(
                        f"La descarga no contiene la columna {required}. "
                        "Comprueba que se exportó "
                        "RAW DATA DAILY EXCEPTION. "
                        f"Encabezados encontrados: {headers[:12]}"
                    )

            org_column = headers.index("ORG")
            item_column = headers.index("ITEM")

            organizations = set()
            item_count = 0

            for row in rows:
                item = row[item_column]

                if item is None or item == "":
                    continue

                item_count += 1

                organization = str(
                    row[org_column] or ""
                ).strip().upper()

                organizations.add(organization)

            if item_count == 0:
                raise ValueError(
                    "La descarga no contiene artículos."
                )

            unexpected = (
                organizations - ALLOWED_ORGANIZATIONS
            )

            if unexpected:
                raise ValueError(
                    "La exportación contiene organizaciones "
                    "distintas a MMC, MMB y TFC: "
                    f"{sorted(unexpected)}. "
                    "No se ejecutará Oracle ni se generará "
                    "el reporte con estos datos."
                )

            print(
                f"Descarga validada: {item_count} artículos."
            )
            print(
                "Organizaciones encontradas:",
                ", ".join(sorted(organizations)),
            )

            absent = (
                ALLOWED_ORGANIZATIONS - organizations
            )

            if absent:
                print(
                    "AVISO: no hay filas para:",
                    ", ".join(sorted(absent)),
                    ". Comprueba el filtro si esperabas "
                    "datos de esas organizaciones.",
                )

        finally:
            workbook.close()

def select_organizations(page: Page) -> None:
    """
    Reproduce la selección grabada por nombre.

    Requiere que ORG comience en All. No utiliza posiciones
    nth-child ni hace clic en gráficas.
    """
    org_filter = page.get_by_role(
        "combobox",
        name="ORG",
        exact=True,
    )

    org_filter.wait_for(
        state="visible",
        timeout=180_000,
    )

    org_filter.click()

    org_list = page.get_by_role(
        "listbox",
        name="ORG",
        exact=True,
    )

    org_list.wait_for(state="visible")

    org_list.get_by_text(
        "MMB",
        exact=True,
    ).click()
    
    org_list.get_by_text(
        "MMC",
        exact=True,
    ).click()

    org_list.get_by_text(
        "TFC",
        exact=True,
    ).click()

    org_filter.click()

    print("Selección de MMC, MMB y TFC realizada.")


def export_table(page: Page, destination: Path, table_name: str) -> None:
    """Exportar únicamente el visual identificado por su título."""
    title_pattern = re.compile(
        r"^RAW DATA\s*:?\s+" + re.escape(table_name) + r"\s*$",
        re.IGNORECASE,
    )
    raw_title = page.get_by_role("heading", name=title_pattern)
    raw_title.wait_for(state="visible", timeout=120_000)
    raw_title.scroll_into_view_if_needed()
    raw_title.hover()

    # Buscar el contenedor más cercano del título que tenga su propio menú.
    visual = raw_title.locator(
        'xpath=ancestor::*[.//*[@data-testid="visual-more-options-btn"]][1]'
    )
    more_options = visual.get_by_test_id("visual-more-options-btn")
    if more_options.count() != 1:
        raise RuntimeError(
            f"No se pudo identificar un único menú para RAW DATA: {table_name}. "
            "Es necesario revisar el selector del visual."
        )
    more_options.wait_for(state="visible", timeout=30_000)
    more_options.click()
    page.get_by_test_id("pbimenu-item.Export data").click()

    dialog = page.get_by_role(
        "dialog", name=re.compile(r"Which data do you want to", re.IGNORECASE)
    )
    dialog.wait_for(state="visible")
    dialog.get_by_text("Data with current layout", exact=True).click()
    with page.expect_download(timeout=180_000) as download_info:
        dialog.get_by_test_id("export-btn").click()
    download = download_info.value
    failure = download.failure()
    if failure:
        raise RuntimeError(f"Falló la descarga de {table_name}: {failure}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    download.save_as(str(destination))
    dialog.wait_for(state="hidden")
    print(f"Descargado RAW DATA: {table_name}")


def validate_po_download(path: Path) -> None:
    """Validar el XLSX sin suponer columnas de PO que aún no conocemos."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["Export"] if "Export" in workbook.sheetnames else workbook.worksheets[0]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers or not any(value is not None for value in headers):
            raise ValueError("OPEN PO EXCEPTION no contiene encabezados.")
        count = sum(1 for row in rows if any(value is not None for value in row))
        print(f"OPEN PO EXCEPTION: {count} filas no vacías después del encabezado.")
        if count == 0:
            print("AVISO: OPEN PO EXCEPTION no tiene filas de datos.")
    finally:
        workbook.close()


def download_powerbi(
    destination: Path,
    setup: bool,
    po_destination: Path,
) -> None:
    with sync_playwright() as playwright:
        context: BrowserContext = (
            playwright.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                channel="msedge",
                device_scale_factor=1,
                headless=False,
                accept_downloads=True,
                viewport={
                    "width": 1920,
                    "height": 1080,
                },
            )
        )

        try:
            page = (
                context.pages[0]
                if context.pages
                else context.new_page()
            )

            page.set_default_timeout(60_000)

            page.goto(
                REPORT_URL,
                wait_until="domcontentloaded",
                timeout=120_000,
            )

            if setup:
                print(
                    "\nPREPARACIÓN EN EL NAVEGADOR\n"
                    "1. Inicia sesión si es necesario.\n"
                    "2. Espera a que cargue el reporte.\n"
                    "3. Deja ORG en All.\n"
                    "4. Cierra el desplegable ORG.\n"
                    "5. Regresa a esta terminal.\n"
                )

                input(
                    "Presiona Enter para seleccionar "
                    "MMC, MMB y TFC y descargar ambas tablas... "
                )

            select_organizations(page)

            export_table(page, destination, "DAILY EXCEPTION")
            validate_download(destination)
            export_table(page, po_destination, "OPEN PO EXCEPTIONS")
            validate_po_download(po_destination)

        finally:
            context.close()



def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Descargar RAW desde Power BI, ejecutar Oracle "
            "y generar Revision Enhanced."
        )
    )

    parser.add_argument(
        "--setup",
        action="store_true",
        help=(
            "Pausa para iniciar sesión y dejar "
            "el filtro ORG en All."
        ),
    )

    parser.add_argument(
        "--config",
        type=Path,
        default=ROOT / "config.json",
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("C:/Users/E1568382/OneDrive - Emerson/TF Purchasing Analytics - Documents/General Reports/2. Inventory Management/7. Enhanced Daily Exception Messages"),
    )

    parser.add_argument(
        "--log-dir",
        type=Path,
        default=ROOT / "logs",
    )

    parser.add_argument(
        "--lookup",
        choices=["item", "org_item"],
        default="item",
    )

    parser.add_argument(
        "--missing",
        choices=["blank", "na"],
        default="blank",
    )

    args = parser.parse_args()

    main_script = ROOT / "main.py"

    if not main_script.is_file():
        raise FileNotFoundError(
            "Coloca exportar_powerbi.py "
            "en la misma carpeta que main.py."
        )

    config_path = (
        args.config.expanduser().resolve()
    )

    if not config_path.is_file():
        raise FileNotFoundError(
            f"No existe la configuración: {config_path}"
        )

    output_dir = (
        args.output_dir.expanduser().resolve()
    )
    log_dir = (
        args.log_dir.expanduser().resolve()
    )

    if output_dir == log_dir:
        raise ValueError(
            "La carpeta de logs debe ser diferente "
            "a la carpeta del Excel final."
        )

    # El RAW se elimina al finalizar, incluso si hay un error.
    with tempfile.TemporaryDirectory(
        prefix="daily_exception_"
    ) as temporary_directory:

        raw_file = (
            Path(temporary_directory)
            / "RAW DATA DAILY EXCEPTION.xlsx"
        )

        po_file = Path(temporary_directory) / "RAW DATA OPEN PO EXCEPTION.xlsx"
        print("1. Abriendo Power BI y descargando ambas tablas...")

        download_powerbi(
            destination=raw_file,
            setup=args.setup,
            po_destination=po_file,
        )

        # Al no pasar --enhance, main.py consulta Oracle.
        subprocess.run(
            [
                sys.executable,
                str(main_script),
                "--config",
                str(config_path),
                "--raw",
                str(raw_file),
                "--po",
                str(po_file),
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(log_dir),
                "--lookup",
                args.lookup,
                "--missing",
                args.missing,
            ],
            cwd=ROOT,
            check=True,
        )

    print("Proceso terminado.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit("Proceso cancelado.")
    except subprocess.CalledProcessError as error:
        raise SystemExit(
            "main.py terminó con un error. "
            "Revisa el mensaje anterior. "
            f"Código de salida: {error.returncode}"
        ) from None
    except Exception as error:
        raise SystemExit(
            f"ERROR: {error}"
        ) from None