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


def get_mapping_info():
    """Coleta drives de rede mapeados e conexões de rede"""
    info = {}

    # Drives de rede mapeados via PowerShell
    try:
        result = run_powershell(
            "Get-WmiObject Win32_MappedLogicalDisk | "
            "ForEach-Object { "
            "$_.DeviceID + '|' + "
            "$_.ProviderName + '|' + "
            "$_.VolumeName + '|' + "
            "$_.Size + '|' + "
            "$_.FreeSpace "
            "}"
        )
        mappings = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 2:
                drive = parts[0].strip()
                unc_path = parts[1].strip() if parts[1].strip() else "Não identificado"
                volume = parts[2].strip() if len(parts) > 2 and parts[2].strip() else "—"

                # Tamanho e espaço livre
                try:
                    size_gb = int(parts[3].strip()) / (1024**3) if len(parts) > 3 and parts[3].strip() else 0
                    free_gb = int(parts[4].strip()) / (1024**3) if len(parts) > 4 and parts[4].strip() else 0
                    size_str = f"{size_gb:.1f} GB" if size_gb > 0 else "—"
                    free_str = f"{free_gb:.1f} GB" if free_gb > 0 else "—"
                except:
                    size_str = "—"
                    free_str = "—"

                # Extrai servidor do caminho UNC
                server = "—"
                if unc_path.startswith("\\\\"):
                    parts_unc = unc_path.replace("\\\\", "").split("\\")
                    server = parts_unc[0] if parts_unc else "—"

                mappings.append({
                    "drive": drive,
                    "unc_path": unc_path,
                    "volume_name": volume,
                    "server": server,
                    "size": size_str,
                    "free": free_str,
                    "status": "Conectado"
                })

        info["mappings"] = mappings
        info["mappings_count"] = len(mappings)

    except:
        info["mappings"] = []
        info["mappings_count"] = 0

    # Conexões de rede via net use (captura desconectados também)
    try:
        result = subprocess.check_output(
            ["net", "use"],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore").strip()

        net_use_list = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            # Linhas com drives mapeados começam com OK, Desconectado ou letra
            if any(line.startswith(s) for s in ["OK", "Descon", "Disco", "Unavail", "Disconn"]):
                parts = line.split()
                if len(parts) >= 3:
                    status = parts[0]
                    drive = parts[1] if ":" in parts[1] else "—"
                    path = parts[2] if len(parts) > 2 else "—"
                    net_use_list.append({
                        "status": status,
                        "drive": drive,
                        "path": path
                    })

        info["net_use"] = net_use_list

    except:
        info["net_use"] = []

    # Compartilhamentos locais desta máquina
    try:
        result = run_powershell(
            "Get-WmiObject Win32_Share | "
            "Where-Object { $_.Type -eq 0 } | "
            "ForEach-Object { $_.Name + '|' + $_.Path + '|' + $_.Description }"
        )
        shares = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 2:
                share_name = parts[0].strip()
                # Filtra compartilhamentos administrativos padrão (C$, ADMIN$, etc)
                if share_name.endswith("$"):
                    continue
                shares.append({
                    "name": share_name,
                    "path": parts[1].strip() if parts[1].strip() else "—",
                    "description": parts[2].strip() if len(parts) > 2 and parts[2].strip() else "—"
                })
        info["local_shares"] = shares
        info["local_shares_count"] = len(shares)

    except:
        info["local_shares"] = []
        info["local_shares_count"] = 0

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_mapping_info()

    print(f"Drives mapeados: {data['mappings_count']}")
    if data["mappings"]:
        for m in data["mappings"]:
            print(f"  {m['drive']} → {m['unc_path']} | Servidor: {m['server']} | "
                  f"Tamanho: {m['size']} | Livre: {m['free']} | Status: {m['status']}")
    else:
        print("  Nenhum drive mapeado encontrado")

    print(f"\nNet Use:")
    if data["net_use"]:
        for n in data["net_use"]:
            print(f"  {n['status']} | {n['drive']} | {n['path']}")
    else:
        print("  Nenhuma conexão de rede ativa")

    print(f"\nCompartilhamentos locais: {data['local_shares_count']}")
    if data["local_shares"]:
        for s in data["local_shares"]:
            print(f"  {s['name']} → {s['path']} | {s['description']}")
    else:
        print("  Nenhum compartilhamento local encontrado")