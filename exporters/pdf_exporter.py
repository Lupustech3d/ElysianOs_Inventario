import os
import sys
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import Image as RLImage


# ── Cores ElysianOS ───────────────────────────────────
C_DARK     = colors.HexColor("#1A1A2E")
C_MID      = colors.HexColor("#16213E")
C_LIGHT    = colors.HexColor("#0F3460")
C_TEAL     = colors.HexColor("#00BCD4")
C_WHITE    = colors.white
C_GRAY     = colors.HexColor("#7F8C8D")
C_GREEN    = colors.HexColor("#00BC64")
C_RED      = colors.HexColor("#E74C3C")
C_YELLOW   = colors.HexColor("#F39C12")
C_ROW_EVEN = colors.HexColor("#141D35")
C_ROW_ODD  = colors.HexColor("#16213E")


def s(text):
    """Sanitiza texto para o ReportLab — remove/substitui caracteres problemáticos"""
    if text is None:
        return "—"
    text = str(text)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    # Backslash duplo vira legível
    text = text.replace("\\\\", r"\\")
    return text if text.strip() else "—"


def get_logo_path():
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.normpath(os.path.join(base, "..", "assets", "logo.png"))


def get_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ElyTitle", fontName="Helvetica-Bold", fontSize=22,
        textColor=C_TEAL, alignment=TA_CENTER, spaceAfter=4))
    styles.add(ParagraphStyle(
        name="ElySubtitle", fontName="Helvetica", fontSize=11,
        textColor=C_GRAY, alignment=TA_CENTER, spaceAfter=2))
    styles.add(ParagraphStyle(
        name="ElySectionTitle", fontName="Helvetica-Bold", fontSize=11,
        textColor=C_TEAL, alignment=TA_LEFT, spaceBefore=12, spaceAfter=6))
    styles.add(ParagraphStyle(
        name="ElyNormal", fontName="Helvetica", fontSize=8,
        textColor=C_WHITE, alignment=TA_LEFT, spaceAfter=2))
    styles.add(ParagraphStyle(
        name="ElySmall", fontName="Helvetica", fontSize=7,
        textColor=C_GRAY, alignment=TA_CENTER, spaceAfter=2))
    return styles


def header_footer(canvas, doc, data):
    """Header e footer em todas as páginas"""
    canvas.saveState()
    w, h = A4

    # Header background
    canvas.setFillColor(C_DARK)
    canvas.rect(0, h - 45*mm, w, 45*mm, fill=1, stroke=0)

    # Logo
    try:
        canvas.drawImage(get_logo_path(), 15*mm, h - 38*mm,
                         width=22*mm, height=22*mm,
                         preserveAspectRatio=True, mask="auto")
    except:
        pass

    # Títulos header
    canvas.setFillColor(C_TEAL)
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawString(42*mm, h - 20*mm, "ElysianOS")
    canvas.setFillColor(C_GRAY)
    canvas.setFont("Helvetica", 8)
    canvas.drawString(42*mm, h - 27*mm, "Sistema de Inventário Integrado")

    # Info direita
    system  = data.get("system", {})
    network = data.get("network", {})
    canvas.setFillColor(C_WHITE)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawRightString(w - 15*mm, h - 18*mm,
                           str(system.get("machine_name", "—")))
    canvas.setFillColor(C_GRAY)
    canvas.setFont("Helvetica", 7)
    canvas.drawRightString(w - 15*mm, h - 24*mm,
                           str(system.get("inventory_datetime", "—")))
    canvas.drawRightString(w - 15*mm, h - 30*mm,
                           f"IP: {network.get('main_ip','—')}")

    # Linha teal
    canvas.setStrokeColor(C_TEAL)
    canvas.setLineWidth(1.5)
    canvas.line(0, h - 45*mm, w, h - 45*mm)

    # Footer
    canvas.setFillColor(C_MID)
    canvas.rect(0, 0, w, 12*mm, fill=1, stroke=0)
    canvas.setStrokeColor(C_TEAL)
    canvas.setLineWidth(0.5)
    canvas.line(0, 12*mm, w, 12*mm)
    canvas.setFillColor(C_GRAY)
    canvas.setFont("Helvetica", 7)
    canvas.drawString(15*mm, 4*mm,
                      f"ElysianOS {datetime.now().year} | Desenvolvido por Luis Mesquitela")
    canvas.drawRightString(w - 15*mm, 4*mm, f"Página {doc.page}")
    canvas.restoreState()


def make_table(headers, rows, col_widths=None):
    """Tabela estilizada padrão"""
    pw = A4[0] - 30*mm
    if col_widths is None:
        col_widths = [pw / len(headers)] * len(headers)

    # Sanitiza todos os valores
    clean_rows = []
    for row in rows:
        clean_rows.append([s(cell) for cell in row])

    table_data = [headers] + clean_rows

    style = TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_LIGHT),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  C_TEAL),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  8),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_ROW_EVEN, C_ROW_ODD]),
        ("TEXTCOLOR",     (0, 1), (-1, -1), C_WHITE),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 7.5),
        ("ALIGN",         (0, 1), (-1, -1), "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#1E3A5F")),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ])

    return Table(table_data, colWidths=col_widths,
                 style=style, repeatRows=1, hAlign="LEFT")


def make_kv_table(fields, col_widths=None):
    """Tabela chave-valor"""
    pw = A4[0] - 30*mm
    if col_widths is None:
        col_widths = [pw * 0.35, pw * 0.65]

    rows = [[s(k), s(v)] for k, v in fields]

    style = TableStyle([
        ("BACKGROUND",    (0, 0), (0, -1), C_LIGHT),
        ("BACKGROUND",    (1, 0), (1, -1), C_MID),
        ("TEXTCOLOR",     (0, 0), (0, -1), C_TEAL),
        ("TEXTCOLOR",     (1, 0), (1, -1), C_WHITE),
        ("FONTNAME",      (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",      (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#1E3A5F")),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
    ])

    return Table(rows, colWidths=col_widths, style=style, hAlign="LEFT")


def section_title(text, styles):
    return Paragraph(f"▌ {text}", styles["ElySectionTitle"])


def ask_save_location(default_filename):
    """Abre diálogo para escolher onde salvar"""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    filepath = filedialog.asksaveasfilename(
        title="Salvar Relatório PDF",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialfile=default_filename
    )
    root.destroy()
    return filepath


def generate_pdf(data, output_dir=None, ask_location=False):
    """Gera PDF completo do inventário"""

    machine  = data.get("system", {}).get("machine_name", "PC")
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_filename = f"inventario_{machine}_{date_str}.pdf"

    # Pergunta onde salvar
    if ask_location:
        filepath = ask_save_location(default_filename)
        if not filepath:
            return None  # Usuário cancelou
    else:
        if output_dir is None:
            base = getattr(sys, '_MEIPASS',
                           os.path.dirname(os.path.abspath(__file__)))
            output_dir = os.path.normpath(os.path.join(base, ".."))
        os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, default_filename)

    system   = data.get("system", {})
    hardware = data.get("hardware", {})
    network  = data.get("network", {})
    software = data.get("software", {})
    licenses = data.get("licenses", {})
    printers = data.get("printers", {})
    mappings = data.get("mappings", {})

    styles = get_styles()
    story  = []
    pw     = A4[0] - 30*mm

    # ── CAPA ─────────────────────────────────────────
    story.append(Spacer(1, 20*mm))
    try:
        logo = RLImage(get_logo_path(), width=50*mm, height=50*mm)
        logo.hAlign = "CENTER"
        story.append(logo)
    except:
        pass

    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("ElysianOS", styles["ElyTitle"]))
    story.append(Paragraph("Sistema de Inventário Integrado",
                            styles["ElySubtitle"]))
    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=C_TEAL, spaceAfter=8*mm))

    capa_data = [
        ["RELATÓRIO DE INVENTÁRIO", ""],
        ["Colaborador",       s(system.get("collaborator"))],
        ["Usuário Windows",   s(system.get("windows_user"))],
        ["Máquina",           s(system.get("machine_name"))],
        ["Fabricante/Modelo", s(f"{system.get('manufacturer','—')} {system.get('model','—')}")],
        ["Nº de Série",       s(system.get("serial_number"))],
        ["Sistema Operacional", s(licenses.get("windows_edition"))],
        ["IP Principal",      s(network.get("main_ip"))],
        ["Data/Hora",         s(system.get("inventory_datetime"))],
    ]

    capa_style = TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  C_LIGHT),
        ("SPAN",          (0, 0), (-1, 0)),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  C_WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  11),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("BACKGROUND",    (0, 1), (0, -1),  C_LIGHT),
        ("BACKGROUND",    (1, 1), (1, -1),  C_MID),
        ("TEXTCOLOR",     (0, 1), (0, -1),  C_TEAL),
        ("TEXTCOLOR",     (1, 1), (1, -1),  C_WHITE),
        ("FONTNAME",      (0, 1), (0, -1),  "Helvetica-Bold"),
        ("FONTNAME",      (1, 1), (1, -1),  "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#1E3A5F")),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
    ])

    story.append(Table(capa_data,
                       colWidths=[pw * 0.35, pw * 0.65],
                       style=capa_style, hAlign="CENTER"))
    story.append(Spacer(1, 10*mm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=C_TEAL, spaceAfter=4*mm))
    story.append(Paragraph("Desenvolvido por Luis Mesquitela",
                            styles["ElySmall"]))
    story.append(PageBreak())

    # ── SISTEMA ───────────────────────────────────────
    story.append(section_title("INFORMAÇÕES DO SISTEMA", styles))
    story.append(make_kv_table([
        ("Nome da Máquina",     system.get("machine_name", "—")),
        ("Colaborador",         system.get("collaborator", "—")),
        ("Usuário Windows",     system.get("windows_user", "—")),
        ("Sistema Operacional", system.get("os_full", "—")),
        ("Arquitetura",         system.get("architecture", "—")),
        ("Tipo de Equipamento", system.get("pc_type", "—")),
        ("Fabricante",          system.get("manufacturer", "—")),
        ("Modelo",              system.get("model", "—")),
        ("Número de Série",     system.get("serial_number", "—")),
        ("Data do Inventário",  system.get("inventory_datetime", "—")),
    ]))
    story.append(Spacer(1, 5*mm))

    # ── HARDWARE ──────────────────────────────────────
    story.append(section_title("HARDWARE — PROCESSADOR", styles))
    story.append(make_kv_table([
        ("Modelo CPU",      hardware.get("cpu_name", "—")),
        ("Núcleos Físicos", str(hardware.get("cpu_cores_physical", "—"))),
        ("Núcleos Lógicos", str(hardware.get("cpu_cores_logical", "—"))),
        ("Velocidade Máx.", hardware.get("cpu_speed", "—")),
        ("Uso Atual",       hardware.get("cpu_usage", "—")),
    ]))

    story.append(section_title("HARDWARE — MEMÓRIA RAM", styles))
    story.append(make_kv_table([
        ("Total",        hardware.get("ram_total", "—")),
        ("Em Uso",       hardware.get("ram_used", "—")),
        ("Disponível",   hardware.get("ram_available", "—")),
        ("Uso %",        hardware.get("ram_percent", "—")),
        ("Slots Usados", str(hardware.get("ram_slots_used", "—"))),
    ]))

    if hardware.get("ram_slots"):
        story.append(Spacer(1, 3*mm))
        story.append(make_table(
            ["Slot", "Capacidade", "Velocidade"],
            [[sl.get("locator","—"), sl.get("capacity","—"),
              sl.get("speed","—")]
             for sl in hardware["ram_slots"]],
            col_widths=[pw*0.4, pw*0.3, pw*0.3]
        ))

    story.append(section_title("HARDWARE — ARMAZENAMENTO", styles))
    if hardware.get("disks"):
        story.append(make_table(
            ["Modelo do Disco", "Tamanho"],
            [[d.get("model","—"), d.get("size","—")]
             for d in hardware["disks"]],
            col_widths=[pw*0.7, pw*0.3]
        ))
    if hardware.get("partitions"):
        story.append(Spacer(1, 3*mm))
        story.append(make_table(
            ["Drive", "Sistema", "Total", "Usado", "Livre", "Uso %"],
            [[p.get("drive","—"), p.get("filesystem","—"),
              p.get("total","—"), p.get("used","—"),
              p.get("free","—"), p.get("percent","—")]
             for p in hardware["partitions"]],
            col_widths=[pw*0.12, pw*0.12, pw*0.16,
                        pw*0.16, pw*0.16, pw*0.12]
        ))

    if hardware.get("gpus"):
        story.append(section_title("HARDWARE — GPU", styles))
        story.append(make_table(
            ["Nome da GPU", "VRAM"],
            [[g.get("name","—"), g.get("ram","—")]
             for g in hardware["gpus"]],
            col_widths=[pw*0.7, pw*0.3]
        ))

    story.append(section_title("HARDWARE — PLACA MÃE & BIOS", styles))
    story.append(make_kv_table([
        ("Placa Mãe", hardware.get("motherboard", "—")),
        ("BIOS",      hardware.get("bios", "—")),
    ]))
    story.append(Spacer(1, 5*mm))

    # ── REDE ──────────────────────────────────────────
    story.append(section_title("REDE", styles))
    story.append(make_kv_table([
        ("Hostname",      network.get("hostname", "—")),
        ("IP Principal",  network.get("main_ip", "—")),
        ("MAC Principal", network.get("main_mac", "—")),
        ("Gateway",       network.get("main_gateway", "—")),
        ("Internet",      network.get("internet", "—")),
        ("WiFi SSID",     network.get("wifi_ssid", "—")),
    ]))
    if network.get("adapters"):
        story.append(Spacer(1, 3*mm))
        story.append(make_table(
            ["Adaptador", "IPv4", "MAC", "Gateway", "DNS", "DHCP"],
            [[a.get("description","—"), a.get("ipv4","—"),
              a.get("mac","—"),         a.get("gateway","—"),
              a.get("dns","—"),         a.get("dhcp","—")]
             for a in network["adapters"]],
            col_widths=[pw*0.25, pw*0.15, pw*0.16,
                        pw*0.16, pw*0.15, pw*0.09]
        ))
    story.append(Spacer(1, 5*mm))

    # ── LICENÇAS ──────────────────────────────────────
    story.append(PageBreak())
    story.append(section_title("LICENÇA WINDOWS", styles))
    story.append(make_kv_table([
        ("Edição",          licenses.get("windows_edition", "—")),
        ("Build",           licenses.get("windows_build", "—")),
        ("Status",          licenses.get("windows_status", "—")),
        ("Chave (parcial)", licenses.get("windows_key_partial", "—")),
        ("Tipo de Licença", licenses.get("windows_license_type", "—")),
    ]))

    story.append(section_title("LICENÇA OFFICE", styles))
    story.append(make_kv_table([
        ("Versão",          licenses.get("office_version", "—")),
        ("Produto",         licenses.get("office_product", "—")),
        ("Status",          licenses.get("office_status", "—")),
        ("Chave (parcial)", licenses.get("office_key_partial", "—")),
        ("Conta Vinculada", licenses.get("office_account", "—")),
        ("Apps Instalados", ", ".join(licenses.get("office_apps", []))),
    ]))
    story.append(Spacer(1, 5*mm))

    # ── IMPRESSORAS ───────────────────────────────────
    story.append(section_title(
        f"IMPRESSORAS — {printers.get('printers_count', 0)} encontradas",
        styles))
    if printers.get("printers"):
        story.append(make_table(
            ["Nome", "Status", "IP", "Driver", "Compartilhada"],
            [[p.get("name","—"),   p.get("status","—"),
              p.get("ip","—"),     p.get("driver","—"),
              p.get("shared","—")]
             for p in printers["printers"]],
            col_widths=[pw*0.28, pw*0.10, pw*0.16,
                        pw*0.34, pw*0.12]
        ))
    else:
        story.append(Paragraph("Nenhuma impressora encontrada.",
                                styles["ElyNormal"]))
    story.append(Spacer(1, 5*mm))

    # ── MAPEAMENTOS ───────────────────────────────────
    story.append(section_title("MAPEAMENTOS DE REDE", styles))
    if mappings.get("mappings"):
        story.append(make_table(
            ["Drive", "Caminho UNC", "Servidor", "Tamanho", "Livre", "Status"],
            [[m.get("drive","—"),    m.get("unc_path","—"),
              m.get("server","—"),   m.get("size","—"),
              m.get("free","—"),     m.get("status","—")]
             for m in mappings["mappings"]],
            col_widths=[pw*0.08, pw*0.30, pw*0.18,
                        pw*0.12, pw*0.12, pw*0.12]
        ))
    else:
        story.append(Paragraph("Nenhum drive mapeado encontrado.",
                                styles["ElyNormal"]))

    if mappings.get("local_shares"):
        story.append(Spacer(1, 3*mm))
        story.append(section_title("COMPARTILHAMENTOS LOCAIS", styles))
        story.append(make_table(
            ["Nome", "Caminho", "Descrição"],
            [[sh.get("name","—"), sh.get("path","—"),
              sh.get("description","—")]
             for sh in mappings["local_shares"]],
            col_widths=[pw*0.2, pw*0.45, pw*0.35]
        ))
    story.append(Spacer(1, 5*mm))

    # ── SOFTWARE ──────────────────────────────────────
    story.append(PageBreak())
    story.append(section_title(
        f"SOFTWARES INSTALADOS — {software.get('software_count', 0)} programas",
        styles))
    story.append(make_table(
        ["Nome", "Versão", "Fabricante", "Instalado em", "Tamanho"],
        [[sw.get("name","—"),         sw.get("version","—"),
          sw.get("publisher","—"),    sw.get("install_date","—"),
          sw.get("size","—")]
         for sw in software.get("software_list", [])],
        col_widths=[pw*0.35, pw*0.14, pw*0.25, pw*0.13, pw*0.13]
    ))

    # ── BUILD ─────────────────────────────────────────
    doc = SimpleDocTemplate(
        filepath,
        pagesize=A4,
        topMargin=50*mm,
        bottomMargin=18*mm,
        leftMargin=15*mm,
        rightMargin=15*mm,
        title=f"ElysianOS - Inventário {machine}",
        author="Luis Mesquitela",
        subject="Relatório de Inventário de TI"
    )

    doc.build(
        story,
        onFirstPage=lambda c, d: header_footer(c, d, data),
        onLaterPages=lambda c, d: header_footer(c, d, data)
    )

    return filepath