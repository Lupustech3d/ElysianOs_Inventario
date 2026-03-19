import os
import sys
import base64
from datetime import datetime


def get_logo_base64():
    """Converte logo para base64 para embutir no HTML"""
    try:
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        logo_path = os.path.join(base, "..", "assets", "logo.png")
        logo_path = os.path.normpath(logo_path)
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except:
        return ""


def generate_html(data):
    """Gera relatório HTML completo com todos os dados coletados"""

    system    = data.get("system", {})
    hardware  = data.get("hardware", {})
    network   = data.get("network", {})
    software  = data.get("software", {})
    licenses  = data.get("licenses", {})
    printers  = data.get("printers", {})
    mappings  = data.get("mappings", {})

    logo_b64  = get_logo_base64()
    logo_tag  = f'<img src="data:image/png;base64,{logo_b64}" alt="Logo" class="logo-img">' if logo_b64 else '<span class="logo-fallback">⬡</span>'

    # ── Software rows ────────────────────────────────────
    sw_rows = ""
    for i, sw in enumerate(software.get("software_list", [])):
        row_class = "row-even" if i % 2 == 0 else "row-odd"
        sw_rows += f"""
        <tr class="{row_class}">
            <td>{sw.get('name','—')}</td>
            <td>{sw.get('version','—')}</td>
            <td>{sw.get('publisher','—')}</td>
            <td>{sw.get('install_date','—')}</td>
            <td>{sw.get('size','—')}</td>
        </tr>"""

    # ── Printer rows ─────────────────────────────────────
    printer_rows = ""
    for p in printers.get("printers", []):
        status_class = "badge-online" if p.get("status") == "Online" else "badge-offline"
        default_badge = '<span class="badge-default">Padrão</span>' if p.get("default") == "Sim" else ""
        printer_rows += f"""
        <tr>
            <td>{p.get('name','—')} {default_badge}</td>
            <td><span class="{status_class}">{p.get('status','—')}</span></td>
            <td>{p.get('ip','—')}</td>
            <td>{p.get('port','—')}</td>
            <td>{p.get('driver','—')}</td>
            <td>{p.get('shared','—')}</td>
        </tr>"""

    # ── Mapping rows ─────────────────────────────────────
    mapping_rows = ""
    for m in mappings.get("mappings", []):
        mapping_rows += f"""
        <tr>
            <td><span class="drive-badge">{m.get('drive','—')}</span></td>
            <td>{m.get('unc_path','—')}</td>
            <td>{m.get('server','—')}</td>
            <td>{m.get('size','—')}</td>
            <td>{m.get('free','—')}</td>
            <td><span class="badge-online">{m.get('status','—')}</span></td>
        </tr>"""
    if not mapping_rows:
        mapping_rows = '<tr><td colspan="6" class="empty-row">Nenhum drive mapeado encontrado</td></tr>'

    # ── Network adapter rows ──────────────────────────────
    adapter_rows = ""
    for a in network.get("adapters", []):
        adapter_rows += f"""
        <tr>
            <td>{a.get('description','—')}</td>
            <td><code>{a.get('ipv4','—')}</code></td>
            <td><code>{a.get('mac','—')}</code></td>
            <td>{a.get('gateway','—')}</td>
            <td>{a.get('dns','—')}</td>
            <td>{a.get('dhcp','—')}</td>
        </tr>"""

    # ── RAM slot rows ─────────────────────────────────────
    ram_rows = ""
    for slot in hardware.get("ram_slots", []):
        ram_rows += f"""
        <tr>
            <td>{slot.get('locator','—')}</td>
            <td>{slot.get('capacity','—')}</td>
            <td>{slot.get('speed','—')}</td>
        </tr>"""
    if not ram_rows:
        ram_rows = '<tr><td colspan="3" class="empty-row">Informação disponível com privilégios de administrador</td></tr>'

    # ── Disk rows ─────────────────────────────────────────
    disk_rows = ""
    for d in hardware.get("disks", []):
        disk_rows += f"""
        <tr>
            <td>{d.get('model','—')}</td>
            <td>{d.get('size','—')}</td>
        </tr>"""

    # ── Partition rows ────────────────────────────────────
    partition_rows = ""
    for p in hardware.get("partitions", []):
        pct = float(p.get("percent","0%").replace("%","")) if p.get("percent") else 0
        bar_color = "#e74c3c" if pct > 85 else "#f39c12" if pct > 70 else "#00BCD4"
        partition_rows += f"""
        <tr>
            <td><strong>{p.get('drive','—')}</strong></td>
            <td>{p.get('filesystem','—')}</td>
            <td>{p.get('total','—')}</td>
            <td>{p.get('used','—')}</td>
            <td>{p.get('free','—')}</td>
            <td>
                <div class="progress-wrap">
                    <div class="progress-bar-fill" style="width:{pct}%;background:{bar_color}"></div>
                </div>
                <small>{p.get('percent','—')}</small>
            </td>
        </tr>"""

    # ── Office apps ───────────────────────────────────────
    office_apps_html = ""
    for app in licenses.get("office_apps", []):
        office_apps_html += f'<span class="app-badge">{app}</span> '
    if not office_apps_html:
        office_apps_html = "—"

    # ── Windows status badge ──────────────────────────────
    win_status = licenses.get("windows_status", "—")
    win_badge_class = "badge-online" if win_status == "Ativado" else "badge-offline"

    # ── Office status badge ───────────────────────────────
    off_status = licenses.get("office_status", "—")
    off_badge_class = "badge-online" if off_status == "Ativado" else "badge-warning"

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ElysianOS - Relatório de Inventário</title>
    <style>
        * {{ margin:0; padding:0; box-sizing:border-box; }}
        body {{ font-family:'Segoe UI',Arial,sans-serif; background:#0f0f1a; color:#e0e0e0; }}

        /* ── Header ── */
        .header {{
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            padding: 20px 40px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 2px solid #00BCD4;
            position: sticky; top: 0; z-index: 100;
            box-shadow: 0 2px 20px rgba(0,188,212,0.3);
        }}
        .header-left {{ display:flex; align-items:center; gap:15px; }}
        .logo-img {{ width:50px; height:50px; object-fit:contain; }}
        .logo-fallback {{ font-size:36px; color:#00BCD4; }}
        .header-title h1 {{ font-size:1.4rem; color:#00BCD4; letter-spacing:1px; }}
        .header-title p {{ font-size:0.75rem; color:#7f8c8d; }}
        .header-right {{ text-align:right; }}
        .header-right .machine-name {{ font-size:1rem; font-weight:bold; color:#fff; }}
        .header-right .date-info {{ font-size:0.75rem; color:#7f8c8d; }}

        /* ── Info bar ── */
        .info-bar {{
            background:#16213e;
            padding:12px 40px;
            display:flex;
            gap:30px;
            flex-wrap:wrap;
            border-bottom:1px solid #1e3a5f;
        }}
        .info-item {{ display:flex; flex-direction:column; }}
        .info-item .label {{ font-size:0.68rem; color:#7f8c8d; text-transform:uppercase; letter-spacing:1px; }}
        .info-item .value {{ font-size:0.9rem; color:#00BCD4; font-weight:600; }}

        /* ── Navigation tabs ── */
        .nav-tabs {{
            background:#16213e;
            padding:0 40px;
            display:flex;
            gap:5px;
            border-bottom:2px solid #1e3a5f;
            overflow-x:auto;
        }}
        .tab-btn {{
            padding:12px 20px;
            background:transparent;
            border:none;
            color:#7f8c8d;
            cursor:pointer;
            font-size:0.85rem;
            font-family:'Segoe UI',Arial,sans-serif;
            border-bottom:3px solid transparent;
            transition:all 0.2s;
            white-space:nowrap;
        }}
        .tab-btn:hover {{ color:#00BCD4; }}
        .tab-btn.active {{ color:#00BCD4; border-bottom:3px solid #00BCD4; font-weight:600; }}

        /* ── Content ── */
        .content {{ padding:30px 40px; max-width:1400px; margin:0 auto; }}
        .tab-content {{ display:none; }}
        .tab-content.active {{ display:block; }}

        /* ── Section title ── */
        .section-title {{
            font-size:1.1rem;
            color:#00BCD4;
            font-weight:600;
            margin-bottom:15px;
            padding-bottom:8px;
            border-bottom:1px solid #1e3a5f;
            display:flex;
            align-items:center;
            gap:8px;
        }}

        /* ── Cards grid ── */
        .cards-grid {{
            display:grid;
            grid-template-columns:repeat(auto-fit, minmax(250px,1fr));
            gap:15px;
            margin-bottom:25px;
        }}
        .card {{
            background:#16213e;
            border:1px solid #1e3a5f;
            border-radius:8px;
            padding:18px;
            transition:border-color 0.2s;
        }}
        .card:hover {{ border-color:#00BCD4; }}
        .card .card-label {{ font-size:0.72rem; color:#7f8c8d; text-transform:uppercase; letter-spacing:1px; margin-bottom:6px; }}
        .card .card-value {{ font-size:1rem; color:#fff; font-weight:500; word-break:break-word; }}
        .card .card-icon {{ font-size:1.5rem; margin-bottom:8px; }}

        /* ── Tables ── */
        .table-wrap {{ overflow-x:auto; border-radius:8px; border:1px solid #1e3a5f; margin-bottom:25px; }}
        table {{ width:100%; border-collapse:collapse; }}
        thead tr {{ background:#0f3460; }}
        thead th {{
            padding:12px 15px;
            text-align:left;
            font-size:0.78rem;
            color:#00BCD4;
            text-transform:uppercase;
            letter-spacing:0.5px;
            white-space:nowrap;
        }}
        tbody tr {{ border-bottom:1px solid #1a2a4a; transition:background 0.15s; }}
        tbody tr:hover {{ background:#1e3a5f; }}
        .row-even {{ background:#16213e; }}
        .row-odd {{ background:#141d35; }}
        td {{ padding:10px 15px; font-size:0.85rem; }}
        .empty-row {{ text-align:center; color:#7f8c8d; padding:20px; font-style:italic; }}

        /* ── Badges ── */
        .badge-online {{
            background:rgba(0,188,100,0.15);
            color:#00bc64;
            padding:3px 10px;
            border-radius:20px;
            font-size:0.75rem;
            font-weight:600;
            border:1px solid rgba(0,188,100,0.3);
        }}
        .badge-offline {{
            background:rgba(231,76,60,0.15);
            color:#e74c3c;
            padding:3px 10px;
            border-radius:20px;
            font-size:0.75rem;
            font-weight:600;
            border:1px solid rgba(231,76,60,0.3);
        }}
        .badge-warning {{
            background:rgba(243,156,18,0.15);
            color:#f39c12;
            padding:3px 10px;
            border-radius:20px;
            font-size:0.75rem;
            font-weight:600;
            border:1px solid rgba(243,156,18,0.3);
        }}
        .badge-default {{
            background:rgba(0,188,212,0.15);
            color:#00BCD4;
            padding:2px 8px;
            border-radius:20px;
            font-size:0.7rem;
            border:1px solid rgba(0,188,212,0.3);
            margin-left:5px;
        }}
        .drive-badge {{
            background:#0f3460;
            color:#00BCD4;
            padding:3px 10px;
            border-radius:4px;
            font-weight:bold;
            font-family:monospace;
        }}
        .app-badge {{
            background:#0f3460;
            color:#ecf0f1;
            padding:4px 12px;
            border-radius:20px;
            font-size:0.8rem;
            margin:2px;
            display:inline-block;
            border:1px solid #1e3a5f;
        }}

        /* ── Progress bar ── */
        .progress-wrap {{
            background:#1a2a4a;
            border-radius:10px;
            height:8px;
            width:120px;
            overflow:hidden;
            display:inline-block;
            vertical-align:middle;
            margin-right:6px;
        }}
        .progress-bar-fill {{
            height:100%;
            border-radius:10px;
            transition:width 0.3s;
        }}

        /* ── Export buttons ── */
        .export-bar {{
            position:fixed;
            bottom:0; left:0; right:0;
            background:#16213e;
            border-top:2px solid #00BCD4;
            padding:12px 40px;
            display:flex;
            gap:12px;
            align-items:center;
            z-index:100;
            box-shadow:0 -4px 20px rgba(0,0,0,0.4);
        }}
        .btn {{
            padding:10px 24px;
            border:none;
            border-radius:6px;
            cursor:pointer;
            font-size:0.88rem;
            font-family:'Segoe UI',Arial,sans-serif;
            font-weight:600;
            transition:all 0.2s;
            display:flex;
            align-items:center;
            gap:8px;
        }}
        .btn-xls {{ background:#1d6f42; color:#fff; }}
        .btn-xls:hover {{ background:#27ae60; transform:translateY(-1px); }}
        .btn-pdf {{ background:#c0392b; color:#fff; }}
        .btn-pdf:hover {{ background:#e74c3c; transform:translateY(-1px); }}
        .btn-consolidate {{ background:#0f3460; color:#00BCD4; border:1px solid #00BCD4; }}
        .btn-consolidate:hover {{ background:#00BCD4; color:#0f3460; transform:translateY(-1px); }}
        .export-label {{ color:#7f8c8d; font-size:0.8rem; margin-left:auto; }}

        /* ── Footer space ── */
        .footer-space {{ height:70px; }}

        code {{ background:#0f3460; padding:2px 6px; border-radius:4px; font-size:0.82rem; color:#00BCD4; }}

        /* ── Search box ── */
        .search-wrap {{ margin-bottom:15px; }}
        .search-input {{
            width:100%;
            padding:10px 15px;
            background:#16213e;
            border:1px solid #1e3a5f;
            border-radius:6px;
            color:#e0e0e0;
            font-size:0.9rem;
            outline:none;
            transition:border-color 0.2s;
        }}
        .search-input:focus {{ border-color:#00BCD4; }}
        .search-input::placeholder {{ color:#4a4a6a; }}

        /* ── Scrollbar ── */
        ::-webkit-scrollbar {{ width:6px; height:6px; }}
        ::-webkit-scrollbar-track {{ background:#0f0f1a; }}
        ::-webkit-scrollbar-thumb {{ background:#1e3a5f; border-radius:3px; }}
        ::-webkit-scrollbar-thumb:hover {{ background:#00BCD4; }}
    </style>
</head>
<body>

<!-- HEADER -->
<div class="header">
    <div class="header-left">
        {logo_tag}
        <div class="header-title">
            <h1>ElysianOS</h1>
            <p>Sistema de Inventário Integrado</p>
        </div>
    </div>
    <div class="header-right">
        <div class="machine-name">{system.get('machine_name','—')}</div>
        <div class="date-info">{system.get('inventory_datetime','—')}</div>
    </div>
</div>

<!-- INFO BAR -->
<div class="info-bar">
    <div class="info-item">
        <span class="label">Colaborador</span>
        <span class="value">{system.get('collaborator','—')}</span>
    </div>
    <div class="info-item">
        <span class="label">Usuário Windows</span>
        <span class="value">{system.get('windows_user','—')}</span>
    </div>
    <div class="info-item">
        <span class="label">IP Principal</span>
        <span class="value">{network.get('main_ip','—')}</span>
    </div>
    <div class="info-item">
        <span class="label">Sistema Operacional</span>
        <span class="value">{licenses.get('windows_edition','—')}</span>
    </div>
    <div class="info-item">
        <span class="label">Tipo</span>
        <span class="value">{system.get('pc_type','—')}</span>
    </div>
    <div class="info-item">
        <span class="label">Internet</span>
        <span class="value">{network.get('internet','—')}</span>
    </div>
</div>

<!-- NAVIGATION TABS -->
<div class="nav-tabs">
    <button class="tab-btn active" onclick="showTab('sistema')">💻 Sistema</button>
    <button class="tab-btn" onclick="showTab('hardware')">🔧 Hardware</button>
    <button class="tab-btn" onclick="showTab('rede')">🌐 Rede</button>
    <button class="tab-btn" onclick="showTab('software')">📦 Software ({software.get('software_count',0)})</button>
    <button class="tab-btn" onclick="showTab('licencas')">🔑 Licenças</button>
    <button class="tab-btn" onclick="showTab('impressoras')">🖨️ Impressoras ({printers.get('printers_count',0)})</button>
    <button class="tab-btn" onclick="showTab('mapeamentos')">📁 Mapeamentos</button>
</div>

<!-- CONTENT -->
<div class="content">

    <!-- ABA SISTEMA -->
    <div id="sistema" class="tab-content active">
        <div class="section-title">💻 Informações do Sistema</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Nome da Máquina</div><div class="card-value">{system.get('machine_name','—')}</div></div>
            <div class="card"><div class="card-label">Colaborador</div><div class="card-value">{system.get('collaborator','—')}</div></div>
            <div class="card"><div class="card-label">Usuário Windows</div><div class="card-value">{system.get('windows_user','—')}</div></div>
            <div class="card"><div class="card-label">Sistema Operacional</div><div class="card-value">{system.get('os_full','—')}</div></div>
            <div class="card"><div class="card-label">Arquitetura</div><div class="card-value">{system.get('architecture','—')}</div></div>
            <div class="card"><div class="card-label">Tipo de Equipamento</div><div class="card-value">{system.get('pc_type','—')}</div></div>
            <div class="card"><div class="card-label">Fabricante</div><div class="card-value">{system.get('manufacturer','—')}</div></div>
            <div class="card"><div class="card-label">Modelo</div><div class="card-value">{system.get('model','—')}</div></div>
            <div class="card"><div class="card-label">Número de Série</div><div class="card-value">{system.get('serial_number','—')}</div></div>
            <div class="card"><div class="card-label">Data do Inventário</div><div class="card-value">{system.get('inventory_datetime','—')}</div></div>
        </div>
    </div>

    <!-- ABA HARDWARE -->
    <div id="hardware" class="tab-content">
        <div class="section-title">⚙️ Processador</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Modelo CPU</div><div class="card-value">{hardware.get('cpu_name','—')}</div></div>
            <div class="card"><div class="card-label">Núcleos Físicos</div><div class="card-value">{hardware.get('cpu_cores_physical','—')}</div></div>
            <div class="card"><div class="card-label">Núcleos Lógicos</div><div class="card-value">{hardware.get('cpu_cores_logical','—')}</div></div>
            <div class="card"><div class="card-label">Velocidade Máx.</div><div class="card-value">{hardware.get('cpu_speed','—')}</div></div>
            <div class="card"><div class="card-label">Uso Atual</div><div class="card-value">{hardware.get('cpu_usage','—')}</div></div>
        </div>

        <div class="section-title">🧠 Memória RAM</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Total</div><div class="card-value">{hardware.get('ram_total','—')}</div></div>
            <div class="card"><div class="card-label">Em Uso</div><div class="card-value">{hardware.get('ram_used','—')}</div></div>
            <div class="card"><div class="card-label">Disponível</div><div class="card-value">{hardware.get('ram_available','—')}</div></div>
            <div class="card"><div class="card-label">Uso %</div><div class="card-value">{hardware.get('ram_percent','—')}</div></div>
            <div class="card"><div class="card-label">Slots Usados</div><div class="card-value">{hardware.get('ram_slots_used','—')}</div></div>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Slot</th><th>Capacidade</th><th>Velocidade</th></tr></thead>
                <tbody>{ram_rows}</tbody>
            </table>
        </div>

        <div class="section-title">💾 Armazenamento</div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Modelo</th><th>Tamanho</th></tr></thead>
                <tbody>{''.join(f"<tr><td>{d.get('model','—')}</td><td>{d.get('size','—')}</td></tr>" for d in hardware.get('disks',[]))}</tbody>
            </table>
        </div>

        <div class="section-title">📊 Partições</div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Drive</th><th>Sistema</th><th>Total</th><th>Usado</th><th>Livre</th><th>Uso</th></tr></thead>
                <tbody>{partition_rows}</tbody>
            </table>
        </div>

        <div class="section-title">🎮 GPU</div>
        <div class="cards-grid">
            {''.join(f'<div class="card"><div class="card-label">GPU</div><div class="card-value">{g.get("name","—")}</div><div class="card-label" style="margin-top:5px">VRAM</div><div class="card-value">{g.get("ram","—")}</div></div>' for g in hardware.get('gpus',[]))}
        </div>

        <div class="section-title">🖥️ Placa Mãe & BIOS</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Placa Mãe</div><div class="card-value">{hardware.get('motherboard','—')}</div></div>
            <div class="card"><div class="card-label">BIOS</div><div class="card-value">{hardware.get('bios','—')}</div></div>
        </div>
    </div>

    <!-- ABA REDE -->
    <div id="rede" class="tab-content">
        <div class="section-title">🌐 Informações de Rede</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Hostname</div><div class="card-value">{network.get('hostname','—')}</div></div>
            <div class="card"><div class="card-label">IP Principal</div><div class="card-value">{network.get('main_ip','—')}</div></div>
            <div class="card"><div class="card-label">MAC Principal</div><div class="card-value">{network.get('main_mac','—')}</div></div>
            <div class="card"><div class="card-label">Gateway</div><div class="card-value">{network.get('main_gateway','—')}</div></div>
            <div class="card"><div class="card-label">Internet</div><div class="card-value">{network.get('internet','—')}</div></div>
            <div class="card"><div class="card-label">WiFi SSID</div><div class="card-value">{network.get('wifi_ssid','—')}</div></div>
        </div>

        <div class="section-title">📡 Adaptadores de Rede</div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Adaptador</th><th>IPv4</th><th>MAC</th><th>Gateway</th><th>DNS</th><th>DHCP</th></tr></thead>
                <tbody>{adapter_rows}</tbody>
            </table>
        </div>
    </div>

    <!-- ABA SOFTWARE -->
    <div id="software" class="tab-content">
        <div class="section-title">📦 Softwares Instalados
            <span style="font-size:0.8rem;color:#7f8c8d;font-weight:normal">
                — {software.get('software_count',0)} programas encontrados
            </span>
        </div>
        <div class="search-wrap">
            <input type="text" class="search-input" id="sw-search"
                   placeholder="🔍 Pesquisar software..."
                   onkeyup="filterTable('sw-search','sw-table')">
        </div>
        <div class="table-wrap">
            <table id="sw-table">
                <thead><tr><th>Nome</th><th>Versão</th><th>Fabricante</th><th>Instalado em</th><th>Tamanho</th></tr></thead>
                <tbody>{sw_rows}</tbody>
            </table>
        </div>
    </div>

    <!-- ABA LICENÇAS -->
    <div id="licencas" class="tab-content">
        <div class="section-title">🪟 Licença Windows</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Edição</div><div class="card-value">{licenses.get('windows_edition','—')}</div></div>
            <div class="card"><div class="card-label">Build</div><div class="card-value">{licenses.get('windows_build','—')}</div></div>
            <div class="card"><div class="card-label">Status</div><div class="card-value"><span class="{win_badge_class}">{win_status}</span></div></div>
            <div class="card"><div class="card-label">Chave (parcial)</div><div class="card-value"><code>{licenses.get('windows_key_partial','—')}</code></div></div>
            <div class="card" style="grid-column:1/-1"><div class="card-label">Tipo de Licença</div><div class="card-value">{licenses.get('windows_license_type','—')}</div></div>
        </div>

        <div class="section-title">📊 Licença Office</div>
        <div class="cards-grid">
            <div class="card"><div class="card-label">Versão</div><div class="card-value">{licenses.get('office_version','—')}</div></div>
            <div class="card"><div class="card-label">Produto</div><div class="card-value">{licenses.get('office_product','—')}</div></div>
            <div class="card"><div class="card-label">Status</div><div class="card-value"><span class="{off_badge_class}">{off_status}</span></div></div>
            <div class="card"><div class="card-label">Chave (parcial)</div><div class="card-value"><code>{licenses.get('office_key_partial','—')}</code></div></div>
            <div class="card"><div class="card-label">Conta Vinculada</div><div class="card-value">{licenses.get('office_account','—')}</div></div>
            <div class="card" style="grid-column:1/-1">
                <div class="card-label">Aplicativos Instalados</div>
                <div class="card-value" style="margin-top:8px">{office_apps_html}</div>
            </div>
        </div>
    </div>

    <!-- ABA IMPRESSORAS -->
    <div id="impressoras" class="tab-content">
        <div class="section-title">🖨️ Impressoras Instaladas
            <span style="font-size:0.8rem;color:#7f8c8d;font-weight:normal">
                — {printers.get('printers_count',0)} encontradas | Padrão: {printers.get('default_printer','—')}
            </span>
        </div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Nome</th><th>Status</th><th>IP</th><th>Porta</th><th>Driver</th><th>Compartilhada</th></tr></thead>
                <tbody>{printer_rows}</tbody>
            </table>
        </div>
    </div>

    <!-- ABA MAPEAMENTOS -->
    <div id="mapeamentos" class="tab-content">
        <div class="section-title">📁 Drives de Rede Mapeados</div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Drive</th><th>Caminho UNC</th><th>Servidor</th><th>Tamanho</th><th>Livre</th><th>Status</th></tr></thead>
                <tbody>{mapping_rows}</tbody>
            </table>
        </div>

        <div class="section-title">🔗 Compartilhamentos Locais</div>
        <div class="table-wrap">
            <table>
                <thead><tr><th>Nome</th><th>Caminho</th><th>Descrição</th></tr></thead>
                <tbody>
                {''.join(f"<tr><td>{s.get('name','—')}</td><td>{s.get('path','—')}</td><td>{s.get('description','—')}</td></tr>" for s in mappings.get('local_shares',[]))}
                </tbody>
            </table>
        </div>
    </div>

</div>

<!-- FOOTER SPACE -->
<div class="footer-space"></div>

<!-- EXPORT BAR -->
<div class="export-bar">
    <button class="btn btn-xls" onclick="exportXLS()">📊 Exportar XLS</button>
    <button class="btn btn-pdf" onclick="exportPDF()">📄 Exportar PDF</button>
    <button class="btn btn-consolidate" onclick="consolidate()">🗂️ Consolidar Relatórios</button>
    <span class="export-label">ElysianOS © {datetime.now().year} | Desenvolvido por Luis Mesquitela</span>
</div>

<script>
    // Tabs
    function showTab(id) {{
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.getElementById(id).classList.add('active');
        event.target.classList.add('active');
    }}

    // Filtro de tabela
    function filterTable(inputId, tableId) {{
        const filter = document.getElementById(inputId).value.toLowerCase();
        const rows = document.getElementById(tableId).getElementsByTagName('tr');
        for (let i = 1; i < rows.length; i++) {{
            const text = rows[i].innerText.toLowerCase();
            rows[i].style.display = text.includes(filter) ? '' : 'none';
        }}
    }}

    // Exportar XLS — chama Python via protocolo customizado
    function exportXLS() {{
        window.location.href = 'elysianos://export-xls';
    }}

    // Exportar PDF
    function exportPDF() {{
        window.location.href = 'elysianos://export-pdf';
    }}

    // Consolidar
    function consolidate() {{
        window.location.href = 'elysianos://consolidate';
    }}
</script>
</body>
</html>"""

    return html


def save_and_open_html(data, output_dir=None):
    """Salva o HTML e abre no browser"""
    import webbrowser
    import tempfile

    html_content = generate_html(data)

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(output_dir, "..")

    output_dir = os.path.normpath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    machine = data.get("system", {}).get("machine_name", "PC")
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"inventario_{machine}_{date_str}.html"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_content)

    webbrowser.open(f"file:///{filepath.replace(os.sep, '/')}")

    return filepath