import os
import sys
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import webbrowser
import http.server
import socketserver
import json

# Garante que o diretório raiz está no path
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


# ══════════════════════════════════════════════════════
#  TELA DE BOAS VINDAS — Pede nome do colaborador
# ══════════════════════════════════════════════════════

class WelcomeScreen:
    def __init__(self):
        self.collaborator_name = None
        self.root = tk.Tk()
        self.root.title("ElysianOS")
        self.root.resizable(False, False)
        self.root.overrideredirect(True)

        w, h = 480, 380
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.configure(bg="#1a1a2e")
        self.root.attributes("-topmost", True)

        self._build_ui()

    def _build_ui(self):
        # Logo
        try:
            from PIL import Image, ImageTk
            logo_path = os.path.join(BASE_DIR, "assets", "logo.png")
            img = Image.open(logo_path).resize((70, 70), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(self.root, image=self.logo_img,
                     bg="#1a1a2e").pack(pady=(30, 5))
        except:
            tk.Label(self.root, text="⬡", font=("Arial", 40),
                     fg="#00BCD4", bg="#1a1a2e").pack(pady=(30, 5))

        # Títulos
        tk.Label(self.root, text="ElysianOS",
                 font=("Segoe UI", 20, "bold"),
                 fg="#00BCD4", bg="#1a1a2e").pack()
        tk.Label(self.root, text="Sistema de Inventário Integrado",
                 font=("Segoe UI", 9),
                 fg="#7f8c8d", bg="#1a1a2e").pack()

        # Separador
        tk.Frame(self.root, bg="#00BCD4",
                 height=1, width=400).pack(pady=15)

        # Label input
        tk.Label(self.root,
                 text="Informe o nome do colaborador:",
                 font=("Segoe UI", 10),
                 fg="#ecf0f1", bg="#1a1a2e").pack(pady=(0, 8))

        # Entry estilizado
        entry_frame = tk.Frame(self.root, bg="#00BCD4",
                                bd=0, padx=1, pady=1)
        entry_frame.pack(padx=40, fill="x")

        self.name_entry = tk.Entry(
            entry_frame,
            font=("Segoe UI", 12),
            bg="#16213e",
            fg="#ffffff",
            insertbackground="#00BCD4",
            relief="flat",
            bd=8
        )
        self.name_entry.pack(fill="x")
        self.name_entry.focus()
        self.name_entry.bind("<Return>", lambda e: self._start())

        # Botão iniciar
        tk.Button(
            self.root,
            text="▶  INICIAR INVENTÁRIO",
            font=("Segoe UI", 11, "bold"),
            bg="#00BCD4",
            fg="#0f0f1a",
            activebackground="#0097A7",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
            command=self._start
        ).pack(pady=20, padx=40, fill="x")

        # Rodapé
        tk.Label(self.root,
                 text="Desenvolvido por Luis Mesquitela",
                 font=("Segoe UI", 8),
                 fg="#4a4a6a", bg="#1a1a2e").pack(
                     side="bottom", pady=10)

    def _start(self):
        name = self.name_entry.get().strip()
        if not name:
            self.name_entry.config(bg="#2a1a1a")
            self.root.after(500,
                lambda: self.name_entry.config(bg="#16213e"))
            return
        self.collaborator_name = name
        self.root.destroy()

    def run(self):
        self.root.mainloop()
        return self.collaborator_name


# ══════════════════════════════════════════════════════
#  SERVIDOR HTTP LOCAL — para os botões do HTML
# ══════════════════════════════════════════════════════

class ReportServer:
    """Servidor HTTP local que recebe comandos dos botões do HTML"""

    def __init__(self, data, html_path):
        self.data      = data
        self.html_path = html_path
        self.port      = 19876
        self.server    = None

    def start(self):
        data      = self.data
        html_path = self.html_path

        class Handler(http.server.BaseHTTPRequestHandler):
            def log_message(self, format, *args):
                pass  # Silencia logs

            def do_GET(self):
                if self.path == "/export-xls":
                    self._export_xls()
                elif self.path == "/export-pdf":
                    self._export_pdf()
                elif self.path == "/consolidate":
                    self._consolidate()
                else:
                    self.send_response(404)
                    self.end_headers()

            def _send_ok(self, msg="OK"):
                self.send_response(200)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(msg.encode("utf-8"))

            def _export_xls(self):
                try:
                    from exporters.xls_exporter import generate_xls
                    from tkinter import filedialog
                    import tkinter as tk

                    machine  = data.get("system",{}).get("machine_name","PC")
                    date_str = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
                    default  = f"inventario_{machine}_{date_str}.xlsx"

                    root = tk.Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    filepath = filedialog.asksaveasfilename(
                        title="Salvar Relatório XLS",
                        defaultextension=".xlsx",
                        filetypes=[("Excel files","*.xlsx"),
                                   ("All files","*.*")],
                        initialfile=default
                    )
                    root.destroy()

                    if filepath:
                        path = generate_xls(data, output_dir=os.path.dirname(filepath))
                        # Renomeia para o nome escolhido
                        if path != filepath:
                            import shutil
                            shutil.move(path, filepath)
                        os.startfile(filepath)
                        self._send_ok(f"XLS salvo em: {filepath}")
                    else:
                        self._send_ok("Cancelado")
                except Exception as e:
                    self._send_ok(f"Erro: {e}")

            def _export_pdf(self):
                try:
                    from exporters.pdf_exporter import generate_pdf
                    path = generate_pdf(data, ask_location=True)
                    if path:
                        os.startfile(path)
                        self._send_ok(f"PDF salvo em: {path}")
                    else:
                        self._send_ok("Cancelado")
                except Exception as e:
                    self._send_ok(f"Erro: {e}")

            def _consolidate(self):
                try:
                    from consolidator.consolidator import run_consolidator
                    run_consolidator()
                    self._send_ok("Consolidador aberto")
                except Exception as e:
                    self._send_ok(f"Erro: {e}")

        socketserver.TCPServer.allow_reuse_address = True
        self.server = socketserver.TCPServer(
            ("localhost", self.port), Handler)
        thread = threading.Thread(
            target=self.server.serve_forever, daemon=True)
        thread.start()

    def stop(self):
        if self.server:
            self.server.shutdown()


# ══════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════

def main():
    # 1. Tela de boas vindas
    welcome = WelcomeScreen()
    collaborator = welcome.run()

    if not collaborator:
        sys.exit(0)

    # 2. Coleta com loading
    from loading_screen import collect_with_loading
    try:
        data = collect_with_loading(collaborator)
    except Exception as e:
        messagebox.showerror(
            "ElysianOS — Erro",
            f"Erro durante a coleta de dados:\n{e}"
        )
        sys.exit(1)

    # 3. Gera HTML
    from exporters.html_exporter import save_and_open_html
    try:
        html_path = save_and_open_html(data)
    except Exception as e:
        messagebox.showerror(
            "ElysianOS — Erro",
            f"Erro ao gerar relatório HTML:\n{e}"
        )
        sys.exit(1)

    # 4. Inicia servidor para botões do HTML
    server = ReportServer(data, html_path)
    server.start()

    # 5. Atualiza HTML com URLs do servidor local
    _patch_html_buttons(html_path, server.port)

    # 6. Reabre o HTML atualizado
    webbrowser.open(
        f"file:///{html_path.replace(os.sep, '/')}")

    # 7. Mantém processo vivo enquanto browser estiver aberto
    _wait_and_cleanup(server, html_path)


def _patch_html_buttons(html_path, port):
    """Substitui os protocolos dos botões pelo servidor local"""
    try:
        with open(html_path, "r", encoding="utf-8") as f:
            content = f.read()

        content = content.replace(
            "window.location.href = 'elysianos://export-xls'",
            f"fetch('http://localhost:{port}/export-xls')"
                ".then(r=>r.text())"
                ".then(m=>{ if(m&&!m.includes('Cancelado'))"
                " alert('✅ XLS salvo com sucesso!') })"
        )
        content = content.replace(
            "window.location.href = 'elysianos://export-pdf'",
            f"fetch('http://localhost:{port}/export-pdf')"
                ".then(r=>r.text())"
                ".then(m=>{ if(m&&!m.includes('Cancelado'))"
                " alert('✅ PDF salvo com sucesso!') })"
        )
        content = content.replace(
            "window.location.href = 'elysianos://consolidate'",
            f"fetch('http://localhost:{port}/consolidate')"
                ".then(r=>r.text())"
                ".then(m=>alert('🗂️ Consolidador aberto!'))"
        )

        with open(html_path, "w", encoding="utf-8") as f:
            f.write(content)
    except:
        pass


def _wait_and_cleanup(server, html_path):
    """Mantém o processo vivo e limpa ao fechar"""
    import time
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        server.stop()
        # Remove HTML temporário
        try:
            os.remove(html_path)
        except:
            pass


if __name__ == "__main__":
    main()