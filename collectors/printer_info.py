import subprocess


def run_powershell(command):
    """Executa comando PowerShell e retorna resultado"""
    try:
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", command],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore").strip()
        return result
    except:
        return ""


def get_printer_info():
    """Coleta informações de impressoras instaladas"""
    info = {}

    # Lista todas as impressoras
    try:
        result = run_powershell(
            "Get-WmiObject Win32_Printer | "
            "ForEach-Object { "
            "$_.Name + '|' + "
            "$_.Default + '|' + "
            "$_.PortName + '|' + "
            "$_.DriverName + '|' + "
            "$_.WorkOffline + '|' + "
            "$_.Shared + '|' + "
            "$_.ShareName "
            "}"
        )
        printers = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 4:
                name = parts[0].strip()

                # Filtra impressoras virtuais/sistema
                skip = ["Microsoft XPS", "Microsoft Print to PDF",
                        "Fax", "OneNote", "Send To OneNote",
                        "Adobe PDF"]
                if any(s.lower() in name.lower() for s in skip):
                    continue

                is_default = parts[1].strip().lower() == "true"
                port = parts[2].strip() if parts[2].strip() else "Não identificado"
                driver = parts[3].strip() if parts[3].strip() else "Não identificado"
                offline = parts[4].strip().lower() == "true" if len(parts) > 4 else False
                shared = parts[5].strip().lower() == "true" if len(parts) > 5 else False
                share_name = parts[6].strip() if len(parts) > 6 and parts[6].strip() else "—"

                # Status baseado em WorkOffline
                status = "Offline" if offline else "Online"

                printers.append({
                    "name": name,
                    "default": "Sim" if is_default else "Não",
                    "port": port,
                    "driver": driver,
                    "status": status,
                    "shared": "Sim" if shared else "Não",
                    "share_name": share_name
                })

        info["printers"] = printers
        info["printers_count"] = len(printers)

        # Impressora padrão
        default = next((p["name"] for p in printers if p["default"] == "Sim"), "Não definida")
        info["default_printer"] = default

    except:
        info["printers"] = []
        info["printers_count"] = 0
        info["default_printer"] = "Não identificado"

    # Tenta detectar IP das impressoras de rede
    try:
        result_ports = run_powershell(
            "Get-WmiObject Win32_TCPIPPrinterPort | "
            "ForEach-Object { $_.Name + '|' + $_.HostAddress + '|' + $_.PortNumber }"
        )
        port_map = {}
        for line in result_ports.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 2 and parts[1].strip():
                port_map[parts[0].strip()] = {
                    "ip": parts[1].strip(),
                    "port_number": parts[2].strip() if len(parts) > 2 else "9100"
                }

        # Adiciona IP às impressoras que tiverem porta TCP/IP
        for printer in info["printers"]:
            port_name = printer.get("port", "")
            if port_name in port_map:
                printer["ip"] = port_map[port_name]["ip"]
                printer["port_number"] = port_map[port_name]["port_number"]
            else:
                printer["ip"] = "Local / USB"
                printer["port_number"] = "—"

    except:
        for printer in info["printers"]:
            printer["ip"] = "Não identificado"
            printer["port_number"] = "—"

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_printer_info()
    print(f"Total de impressoras: {data['printers_count']}")
    print(f"Impressora padrão: {data['default_printer']}\n")
    for p in data["printers"]:
        print(f"  Nome:       {p['name']}")
        print(f"  Padrão:     {p['default']}")
        print(f"  Status:     {p['status']}")
        print(f"  Porta:      {p['port']}")
        print(f"  IP:         {p['ip']}")
        print(f"  Driver:     {p['driver']}")
        print(f"  Compartilh: {p['shared']}")
        print(f"  Nome comp.: {p['share_name']}")
        print()