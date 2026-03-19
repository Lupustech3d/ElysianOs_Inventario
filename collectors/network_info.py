import subprocess
import socket


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


def get_network_info():
    """Coleta informações de rede"""
    info = {}

    # Nome do computador na rede
    try:
        info["hostname"] = socket.gethostname()
    except:
        info["hostname"] = "Não identificado"

    # Adaptadores de rede com IP, MAC, Gateway, DNS
    try:
        result = run_powershell("""
            Get-WmiObject Win32_NetworkAdapterConfiguration |
            Where-Object { $_.IPEnabled -eq $true } |
            ForEach-Object {
                $_.Description + '|' +
                ($_.IPAddress -join ';') + '|' +
                ($_.MACAddress) + '|' +
                ($_.DefaultIPGateway -join ';') + '|' +
                ($_.DNSServerSearchOrder -join ';') + '|' +
                ($_.DHCPEnabled)
            }
        """)
        adapters = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 5:
                # Filtra apenas IPv4
                ips_all = parts[1].split(";") if parts[1] else []
                ipv4 = [ip for ip in ips_all if "." in ip and ":" not in ip]
                ipv6 = [ip for ip in ips_all if ":" in ip]
                adapters.append({
                    "description": parts[0].strip(),
                    "ipv4": ", ".join(ipv4) if ipv4 else "Não identificado",
                    "ipv6": ", ".join(ipv6) if ipv6 else "Não identificado",
                    "mac": parts[2].strip() if parts[2].strip() else "Não identificado",
                    "gateway": parts[3].strip() if parts[3].strip() else "Não identificado",
                    "dns": parts[4].strip().replace(";", ", ") if parts[4].strip() else "Não identificado",
                    "dhcp": "Sim" if parts[5].strip().lower() == "true" else "Não" if len(parts) > 5 else "?"
                })
        info["adapters"] = adapters
    except:
        info["adapters"] = []

    # IP principal (primeiro adaptador ativo)
    try:
        if info["adapters"]:
            info["main_ip"] = info["adapters"][0]["ipv4"]
            info["main_mac"] = info["adapters"][0]["mac"]
            info["main_gateway"] = info["adapters"][0]["gateway"]
        else:
            info["main_ip"] = "Não identificado"
            info["main_mac"] = "Não identificado"
            info["main_gateway"] = "Não identificado"
    except:
        info["main_ip"] = "Não identificado"
        info["main_mac"] = "Não identificado"
        info["main_gateway"] = "Não identificado"

    # DNS configurado no sistema
    try:
        dns_result = run_powershell(
            "Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object {$_.ServerAddresses} | ForEach-Object { $_.InterfaceAlias + '|' + ($_.ServerAddresses -join ';') }"
        )
        dns_list = []
        for line in dns_result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 2 and parts[1].strip():
                dns_list.append({
                    "interface": parts[0].strip(),
                    "servers": parts[1].strip().replace(";", ", ")
                })
        info["dns_servers"] = dns_list
    except:
        info["dns_servers"] = []

    # Status da conexão com internet
    try:
        result = run_powershell(
            "Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet"
        )
        info["internet"] = "Conectado" if result.strip().lower() == "true" else "Sem conexão"
    except:
        info["internet"] = "Não verificado"

    # Tipo de rede (WiFi ou Ethernet ativo)
    try:
        result = run_powershell("""
            Get-NetAdapter |
            Where-Object { $_.Status -eq 'Up' } |
            ForEach-Object { $_.Name + '|' + $_.InterfaceDescription + '|' + $_.LinkSpeed }
        """)
        active_adapters = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 3:
                active_adapters.append({
                    "name": parts[0].strip(),
                    "description": parts[1].strip(),
                    "speed": parts[2].strip()
                })
        info["active_adapters"] = active_adapters
    except:
        info["active_adapters"] = []

    # WiFi SSID conectado (se aplicável)
    try:
        ssid = run_powershell(
            "(netsh wlan show interfaces) -match 'SSID' | Select-Object -First 1"
        )
        if ssid and "SSID" in ssid:
            ssid_clean = ssid.split(":")[-1].strip()
            info["wifi_ssid"] = ssid_clean if ssid_clean else "Não conectado"
        else:
            info["wifi_ssid"] = "Não conectado / Ethernet"
    except:
        info["wifi_ssid"] = "Não identificado"

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_network_info()
    for key, value in data.items():
        if isinstance(value, list):
            print(f"\n{key}:")
            for item in value:
                print(f"  {item}")
        else:
            print(f"{key}: {value}")