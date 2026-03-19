import tkinter as tk
from tkinter import ttk
import threading
import os
import sys


class LoadingScreen:
    """Janela de loading com barra de progresso durante a coleta"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ElysianOS - Coletando dados...")
        self.root.resizable(False, False)
        self.root.overrideredirect(True)

        w, h = 500, 320
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.configure(bg="#1a1a2e")

        self.root.attributes("-topmost", True)

        self._build_ui()

    def _build_ui(self):
        """Constrói a interface da janela de loading"""

        logo_frame = tk.Frame(self.root, bg="#1a1a2e")
        logo_frame.pack(pady=(30, 5))

        try:
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            logo_path = os.path.join(base, "assets", "logo.png")
            from PIL import Image, ImageTk
            img = Image.open(logo_path).resize((60, 60), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(logo_frame, image=self.logo_img, bg="#1a1a2e").pack()
        except:
            tk.Label(logo_frame, text="⬡", font=("Arial", 36),
                     fg="#00BCD4", bg="#1a1a2e").pack()

        tk.Label(
            self.root, text="ElysianOS",
            font=("Segoe UI", 18, "bold"),
            fg="#00BCD4", bg="#1a1a2e"
        ).pack()

        tk.Label(
            self.root, text="Sistema de Inventário Integrado",
            font=("Segoe UI", 9),
            fg="#7f8c8d", bg="#1a1a2e"
        ).pack()

        tk.Frame(self.root, bg="#00BCD4", height=1, width=400).pack(pady=15)

        self.status_var = tk.StringVar(value="Iniciando...")
        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            fg="#ecf0f1", bg="#1a1a2e"
        )
        self.status_label.pack(pady=(0, 10))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Elysia.Horizontal.TProgressbar",
            troughcolor="#16213e",
            background="#00BCD4",
            lightcolor="#00BCD4",
            darkcolor="#0097A7",
            bordercolor="#1a1a2e",
            thickness=12
        )

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.root,
            variable=self.progress_var,
            maximum=100,
            length=420,
            style="Elysia.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=5)

        self.percent_var = tk.StringVar(value="0%")
        tk.Label(
            self.root,
            textvariable=self.percent_var,
            font=("Segoe UI", 9),
            fg="#00BCD4", bg="#1a1a2e"
        ).pack(pady=(5, 0))

        tk.Label(
            self.root,
            text="Desenvolvido por Luis Mesquitela",
            font=("Segoe UI", 8),
            fg="#4a4a6a", bg="#1a1a2e"
        ).pack(side="bottom", pady=10)

    def update_progress(self, step, total, message):
        """Atualiza barra de progresso e mensagem"""
        percent = int((step / total) * 100)
        self.progress_var.set(percent)
        self.percent_var.set(f"{percent}%")
        self.status_var.set(message)
        self.root.update_idletasks()

    def close(self):
        """Fecha a janela de loading"""
        self.root.destroy()

    def start(self):
        """Inicia o loop da janela"""
        self.root.mainloop()


def collect_with_loading(collaborator_name=""):
    """Coleta dados mostrando tela de loading"""
    result = {}
    error = []

    screen = LoadingScreen()

    steps = [
        (1, 7, "Coletando informações do sistema..."),
        (2, 7, "Coletando informações de hardware..."),
        (3, 7, "Coletando informações de rede..."),
        (4, 7, "Coletando softwares instalados..."),
        (5, 7, "Coletando licenças Windows e Office..."),
        (6, 7, "Coletando impressoras..."),
        (7, 7, "Coletando mapeamentos de rede..."),
    ]

    def run_collection():
        try:
            from collectors.system_info import get_system_info
            screen.update_progress(1, 7, "Coletando informações do sistema...")
            result["system"] = get_system_info(collaborator_name)

            from collectors.hardware_info import get_hardware_info
            screen.update_progress(2, 7, "Coletando informações de hardware...")
            result["hardware"] = get_hardware_info()

            from collectors.network_info import get_network_info
            screen.update_progress(3, 7, "Coletando informações de rede...")
            result["network"] = get_network_info()

            from collectors.software_info import get_software_info
            screen.update_progress(4, 7, "Coletando softwares instalados...")
            result["software"] = get_software_info()

            from collectors.license_info import get_license_info
            screen.update_progress(5, 7, "Coletando licenças Windows e Office...")
            result["licenses"] = get_license_info()

            from collectors.printer_info import get_printer_info
            screen.update_progress(6, 7, "Coletando impressoras...")
            result["printers"] = get_printer_info()

            from collectors.mapping_info import get_mapping_info
            screen.update_progress(7, 7, "Finalizando coleta...")
            result["mappings"] = get_mapping_info()

        except Exception as e:
            error.append(str(e))
        finally:
            screen.root.after(600, screen.close)

    thread = threading.Thread(target=run_collection, daemon=True)
    thread.start()
    screen.start()

    if error:
        raise Exception(f"Erro na coleta: {error[0]}")

    return result


if __name__ == "__main__":
    data = collect_with_loading("Luis Mesquitela")
    print(f"\nColeta concluída! {len(data)} módulos carregados.")
    for key in data:
        print(f"  ✓ {key}")
