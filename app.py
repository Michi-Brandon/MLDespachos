import asyncio
import os
import socket
import subprocess
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

try:
    from playwright.async_api import TimeoutError as PlaywrightTimeoutError
    from playwright.async_api import async_playwright
except ImportError:  # Playwright no instalado
    async_playwright = None  # type: ignore[assignment]
    PlaywrightTimeoutError = Exception  # type: ignore[assignment]


LISTING_URL = "https://www.mercadolibre.cl/ventas/omni/listado"
DETAIL_URL_TEMPLATE = "https://www.mercadolibre.cl/ventas/{code}/detalle"
LOGIN_URL = "https://www.mercadolibre.cl/ventas/omni/listado"

# Perfil dedicado para el login (con cookies)
AUTOMATION_PROFILE_DIR = Path(__file__).parent / "ml_profile"

# Para apertura manual del listado con tu Chrome normal
DEFAULT_PROFILE_NAME = "Default"
DEFAULT_USER_DATA_DIR = Path(os.path.expandvars(r"%LocalAppData%\Google\Chrome\User Data"))

REMOTE_DEBUG_PORT: int | None = None
CHROME_PROCESS: subprocess.Popen | None = None


def find_chrome_executable() -> str | None:
    """
    Busca chrome.exe en rutas habituales de Windows.
    """
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def build_chrome_command(chrome_exe: str, url: str) -> list[str]:
    """
    Prepara el comando para abrir Chrome con el perfil por defecto y la URL objetivo.
    """
    return [
        chrome_exe,
        f"--user-data-dir={DEFAULT_USER_DATA_DIR}",
        f"--profile-directory={DEFAULT_PROFILE_NAME}",
        "--new-window",
        url,
    ]


def open_with_url(url: str) -> None:
    chrome_exe = find_chrome_executable()
    if not chrome_exe:
        messagebox.showerror(
            "Chrome no encontrado",
            "No se encontro Google Chrome en las rutas tipicas. "
            "Instala Chrome o ajusta el codigo con la ruta correcta.",
        )
        return

    try:
        subprocess.Popen(
            build_chrome_command(chrome_exe, url),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:  # pragma: no cover
        messagebox.showerror("Error al abrir", f"No se pudo abrir Chrome:\n{exc}")


def open_listing() -> None:
    open_with_url(LISTING_URL)


def open_login() -> None:
    print("[login] Abriendo ventana para iniciar sesion con perfil ml_profile...")
    threading.Thread(target=start_login_browser, daemon=True).start()


def open_detail(code: str) -> None:
    clean_code = code.strip()
    if not clean_code:
        messagebox.showwarning("Codigo requerido", "Ingresa un codigo de venta.")
        return

    url = DETAIL_URL_TEMPLATE.format(code=clean_code)
    print(f"[{clean_code}] Abriendo detalle y extrayendo Envíos...")

    threading.Thread(
        target=lambda: asyncio.run(open_detail_and_extract(clean_code, url)),
        daemon=True,
    ).start()


def center_window(win: tk.Tk, width: int = 480, height: int = 320) -> None:
    win.update_idletasks()
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 3) - (height / 2))
    win.geometry(f"{width}x{height}+{x}+{y}")


def main() -> None:
    root = tk.Tk()
    root.title("Despachos ML")
    root.resizable(False, False)
    center_window(root)

    root.configure(bg="#f2f2f2")

    title = tk.Label(
        root,
        text="Abrir ventas de Mercado Libre",
        font=("Segoe UI", 12, "bold"),
        bg="#f2f2f2",
    )
    title.pack(pady=(20, 10))

    description = tk.Label(
        root,
        text="1) Inicia sesion (boton abajo). 2) Abre detalle y lee Envíos.",
        font=("Segoe UI", 9),
        bg="#f2f2f2",
    )
    description.pack(pady=(0, 16))

    login_button = tk.Button(
        root,
        text="Login y guardar cookies (perfil ML)",
        font=("Segoe UI", 10, "bold"),
        bg="#03a9f4",
        fg="#fff",
        activebackground="#039be5",
        activeforeground="#fff",
        relief=tk.FLAT,
        padx=16,
        pady=8,
        command=open_login,
        cursor="hand2",
    )
    login_button.pack(pady=(0, 14))

    open_button = tk.Button(
        root,
        text="Abrir listado (omni)",
        font=("Segoe UI", 10, "bold"),
        bg="#ffc107",
        fg="#000",
        activebackground="#ffb300",
        activeforeground="#000",
        relief=tk.FLAT,
        padx=16,
        pady=8,
        command=open_listing,
        cursor="hand2",
    )
    open_button.pack(pady=(0, 18))

    # Seccion para abrir detalle por codigo de venta
    detail_frame = tk.Frame(root, bg="#f2f2f2")
    detail_frame.pack(pady=(6, 12), fill="x")

    code_label = tk.Label(
        detail_frame,
        text="Codigo de venta",
        font=("Segoe UI", 10, "bold"),
        bg="#f2f2f2",
        anchor="w",
    )
    code_label.pack(anchor="w", padx=12, pady=(0, 6))

    sale_code_var = tk.StringVar()
    code_entry = tk.Entry(
        detail_frame,
        textvariable=sale_code_var,
        font=("Segoe UI", 11),
        width=36,
        relief=tk.SOLID,
        borderwidth=1,
    )
    code_entry.pack(padx=12, fill="x", pady=(0, 10))

    detail_button = tk.Button(
        detail_frame,
        text="Abrir detalle por codigo",
        font=("Segoe UI", 10, "bold"),
        bg="#4caf50",
        fg="#fff",
        activebackground="#43a047",
        activeforeground="#fff",
        relief=tk.FLAT,
        padx=14,
        pady=10,
        command=lambda: open_detail(sale_code_var.get()),
        cursor="hand2",
    )
    detail_button.pack(padx=12, pady=(0, 4), fill="x")

    root.mainloop()


async def open_detail_and_extract(code: str, url: str) -> None:
    """
    Se conecta al Chrome abierto con el boton de login (puerto CDP) y lee el valor de Envíos.
    Se crea una instancia nueva de Playwright por llamada para evitar conexiones rotas.
    """
    if async_playwright is None:
        print(
            f"[{code}] Playwright no esta instalado. Ejecuta: pip install playwright && python -m playwright install"
        )
        return

    if REMOTE_DEBUG_PORT is None:
        print(f"[{code}] No hay puerto de depuracion. Pulsa primero el boton de login.")
        return

    if not wait_for_port("localhost", REMOTE_DEBUG_PORT, attempts=10, delay=0.4):
        print(f"[{code}] No se pudo alcanzar el puerto {REMOTE_DEBUG_PORT}.")
        return

    endpoint = f"http://localhost:{REMOTE_DEBUG_PORT}"
    playwright = await async_playwright().start()
    try:
        browser = await playwright.chromium.connect_over_cdp(endpoint)
        if not browser.contexts:
            print(f"[{code}] No hay contextos en Chrome. ¿Cerraste la ventana de login?")
            return

        context = browser.contexts[0]
        try:
            page = await context.new_page()
        except Exception as exc:
            print(f"[{code}] No se pudo abrir una nueva pestaña: {exc}")
            return

        page.set_default_timeout(20000)
        await page.goto(url, wait_until="domcontentloaded")

        text = await extract_amount_text(page, "Envíos", timeout_ms=5000)
        source = "Envíos"

        if text is None:
            # Bonificaciones debería resolverse rápido si no hay Envíos; usa timeout corto
            text = await extract_amount_text(page, "Bonificaciones", timeout_ms=2000)
            source = "Bonificaciones" if text else None

        if text is None:
            print(f"[{code}] No se encontraron Envíos ni Bonificaciones. Valor: $ 0")
            return

        parsed = parse_amount(text)
        if parsed is None:
            print(f"[{code}] No se pudo interpretar el valor de {source}: {text}")
            return

        if parsed < 0:
            parsed = 0

        print(f"[{code}] Envíos ({source}): {format_amount(parsed)}")
    except Exception as exc:
        print(f"[{code}] Error conectando a Chrome: {exc}")
    finally:
        try:
            await playwright.stop()
        except Exception:
            pass


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("localhost", 0))
        return s.getsockname()[1]


def wait_for_port(host: str, port: int, attempts: int = 15, delay: float = 0.4) -> bool:
    for _ in range(attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(delay)
            if sock.connect_ex((host, port)) == 0:
                return True
        time.sleep(delay)
    return False


def start_login_browser(start_url: str = LOGIN_URL) -> None:
    """
    Lanza Chrome normal (sin banderas de automatizacion) con remote debugging para que el usuario
    haga login y guarde cookies en ./ml_profile. Mantiene el puerto en REMOTE_DEBUG_PORT.
    """
    global REMOTE_DEBUG_PORT, CHROME_PROCESS

    chrome_exe = find_chrome_executable()
    if not chrome_exe:
        messagebox.showerror(
            "Chrome no encontrado",
            "No se encontro Google Chrome en las rutas tipicas. Ajusta la ruta en el codigo.",
        )
        return

    REMOTE_DEBUG_PORT = find_free_port()
    AUTOMATION_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    args = [
        chrome_exe,
        f"--remote-debugging-port={REMOTE_DEBUG_PORT}",
        f"--user-data-dir={AUTOMATION_PROFILE_DIR}",
        "--profile-directory=Default",
        "--start-maximized",
        "--no-default-browser-check",
        "--no-first-run",
        start_url,
    ]

    try:
        CHROME_PROCESS = subprocess.Popen(
            args,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        print(f"[login] No se pudo lanzar Chrome: {exc}")
        return

    if wait_for_port("localhost", REMOTE_DEBUG_PORT):
        print(
            f"[login] Chrome abierto en {start_url}. Puerto CDP {REMOTE_DEBUG_PORT}. "
            "Inicia sesion; las cookies se guardan en ./ml_profile."
        )
    else:
        print("[login] No se pudo confirmar el puerto de depuracion. Reintenta.")


async def extract_amount_text(page, title: str, timeout_ms: int = 15000) -> str | None:
    """
    Intenta obtener el texto del subtotal para una fila con el titulo dado.
    """
    try:
        row = page.locator("div.sc-account-rows__row", has_text=title).first
        await row.wait_for(state="attached", timeout=timeout_ms)
        text = await row.locator("span.sc-account-rows__row__subTotal").first.text_content()
        return text.strip() if text else None
    except PlaywrightTimeoutError:
        return None
    except Exception:
        return None


def parse_amount(text: str) -> int | None:
    """
    Convierte un texto como "$ 3.090" o "-$ 2.276" a entero (pesos).
    """
    clean = text.replace("\xa0", " ").replace("$", "")
    sign = -1 if "-" in clean else 1
    digits = "".join(ch for ch in clean if ch.isdigit())
    if not digits:
        return None
    try:
        value = int(digits) * sign
    except ValueError:
        return None
    return value


def format_amount(value: int) -> str:
    """
    Devuelve una cadena tipo "$ 3.090" con separador de miles usando punto.
    """
    return "$ " + f"{value:,}".replace(",", ".")


if __name__ == "__main__":
    main()
