import subprocess
import psutil
import platform


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


def get_hardware_info():
    """Coleta informações de hardware via psutil + PowerShell"""
    info = {}

    # CPU Nome via PowerShell
    try:
        cpu_name = run_powershell("(Get-WmiObject Win32_Processor).Name")
        info["cpu_name"] = cpu_name if cpu_name else platform.processor()
    except:
        info["cpu_name"] = platform.processor()

    # CPU Cores, velocidade e uso via psutil
    try:
        info["cpu_cores_physical"] = psutil.cpu_count(logical=False)
        info["cpu_cores_logical"] = psutil.cpu_count(logical=True)
        freq = psutil.cpu_freq()
        info["cpu_speed"] = f"{freq.max / 1000:.2f} GHz" if freq else "Não identificado"
        info["cpu_usage"] = f"{psutil.cpu_percent(interval=1)}%"
    except:
        info["cpu_cores_physical"] = "?"
        info["cpu_cores_logical"] = "?"
        info["cpu_speed"] = "Não identificado"
        info["cpu_usage"] = "?"

    # RAM Total via psutil
    try:
        ram = psutil.virtual_memory()
        info["ram_total"] = f"{ram.total / (1024**3):.1f} GB"
        info["ram_available"] = f"{ram.available / (1024**3):.1f} GB"
        info["ram_used"] = f"{ram.used / (1024**3):.1f} GB"
        info["ram_percent"] = f"{ram.percent}%"
    except:
        info["ram_total"] = "Não identificado"
        info["ram_available"] = "?"
        info["ram_used"] = "?"
        info["ram_percent"] = "?"

    # RAM Slots via PowerShell
    try:
        result = run_powershell(
            "Get-WmiObject Win32_PhysicalMemory | ForEach-Object { $_.DeviceLocator + '|' + $_.Capacity + '|' + $_.Speed }"
        )
        slots = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 3:
                try:
                    cap_gb = int(parts[1]) / (1024**3)
                    slots.append({
                        "locator": parts[0].strip(),
                        "capacity": f"{cap_gb:.0f} GB",
                        "speed": f"{parts[2].strip()} MHz"
                    })
                except:
                    pass
        info["ram_slots"] = slots
        info["ram_slots_used"] = len(slots)
    except:
        info["ram_slots"] = []
        info["ram_slots_used"] = 0

    # Discos físicos via PowerShell
    try:
        result = run_powershell(
            "Get-WmiObject Win32_DiskDrive | ForEach-Object { $_.Model + '|' + $_.Size }"
        )
        disks = []
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) >= 2:
                try:
                    size_gb = int(parts[1].strip()) / (1024**3)
                    disks.append({
                        "model": parts[0].strip(),
                        "size": f"{size_gb:.0f} GB"
                    })
                except:
                    pass
        info["disks"] = disks if disks else []
    except:
        info["disks"] = []

    # Partições via psutil
    try:
        partitions = []
        for part in psutil.disk_partitions():
            try:
                usage = psutil.disk_usage(part.mountpoint)
                partitions.append({
                    "drive": part.device,
                    "mountpoint": part.mountpoint,
                    "filesystem": part.fstype,
                    "total": f"{usage.total / (1024**3):.1f} GB",
                    "used": f"{usage.used / (1024**3):.1f} GB",
                    "free": f"{usage.free / (1024**3):.1f} GB",
                    "percent": f"{usage.percent}%"
                })
            except:
                pass
        info["partitions"] = partitions
    except:
        info["partitions"] = []

    # GPU via PowerShell
    try:
        result = run_powershell(
            "Get-WmiObject Win32_VideoController | ForEach-Object { $_.Name + '|' + $_.AdapterRAM }"
        )
        gpus = []
        skip_keywords = ["MS Idd", "IddSampleDriver", "Virtual", "Remote", "Indirect"]
        for line in result.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split("|")
            name = parts[0].strip()
            if not name:
                continue
            if any(kw.lower() in name.lower() for kw in skip_keywords):
                continue
            try:
                ram_bytes = int(parts[1].strip()) if len(parts) > 1 and parts[1].strip().isdigit() else 0
                ram_mb = ram_bytes / (1024**2)
                gpus.append({
                    "name": name,
                    "ram": f"{ram_mb:.0f} MB" if ram_mb > 0 else "Não identificado"
                })
            except:
                gpus.append({"name": name, "ram": "?"})
        info["gpus"] = gpus if gpus else []
    except:
        info["gpus"] = []

    # Placa mãe via PowerShell
    try:
        manufacturer = run_powershell("(Get-WmiObject Win32_BaseBoard).Manufacturer")
        product = run_powershell("(Get-WmiObject Win32_BaseBoard).Product")
        info["motherboard"] = f"{manufacturer} {product}".strip() if manufacturer or product else "Não identificado"
    except:
        info["motherboard"] = "Não identificado"

    # BIOS via PowerShell
    try:
        bios_maker = run_powershell("(Get-WmiObject Win32_BIOS).Manufacturer")
        bios_ver = run_powershell("(Get-WmiObject Win32_BIOS).SMBIOSBIOSVersion")
        info["bios"] = f"{bios_maker} - {bios_ver}".strip() if bios_maker or bios_ver else "Não identificado"
    except:
        info["bios"] = "Não identificado"

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_hardware_info()
    for key, value in data.items():
        if isinstance(value, list):
            print(f"\n{key}:")
            for item in value:
                print(f"  {item}")
        else:
            print(f"{key}: {value}")