import socket
import getpass
import platform
import subprocess
from datetime import datetime

def get_system_info(collaborator_name=""):
    """Coleta informações do sistema"""
    info = {}

    try:
        # Nome da máquina
        info["machine_name"] = socket.gethostname()
    except:
        info["machine_name"] = "Não identificado"

    try:
        # Usuário logado no Windows
        info["windows_user"] = getpass.getuser()
    except:
        info["windows_user"] = "Não identificado"

    # Nome do colaborador (informado na abertura)
    info["collaborator"] = collaborator_name if collaborator_name else "Não informado"

    try:
        # Sistema operacional
        info["os_name"] = platform.system()
        info["os_version"] = platform.version()
        info["os_release"] = platform.release()
        info["os_full"] = f"Windows {platform.release()} - Build {platform.version()}"
    except:
        info["os_full"] = "Não identificado"

    try:
        # Arquitetura
        info["architecture"] = platform.architecture()[0]
    except:
        info["architecture"] = "Não identificado"

    try:
        # Número de série via WMIC
        result = subprocess.check_output(
            ["wmic", "bios", "get", "serialnumber"],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        lines = [l.strip() for l in result.strip().splitlines() if l.strip()]
        info["serial_number"] = lines[1] if len(lines) > 1 else "Não identificado"
    except:
        info["serial_number"] = "Não identificado"

    try:
        # Fabricante e modelo via WMIC
        result_maker = subprocess.check_output(
            ["wmic", "computersystem", "get", "manufacturer"],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        lines_maker = [l.strip() for l in result_maker.strip().splitlines() if l.strip()]
        info["manufacturer"] = lines_maker[1] if len(lines_maker) > 1 else "Não identificado"

        result_model = subprocess.check_output(
            ["wmic", "computersystem", "get", "model"],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        lines_model = [l.strip() for l in result_model.strip().splitlines() if l.strip()]
        info["model"] = lines_model[1] if len(lines_model) > 1 else "Não identificado"
    except:
        info["manufacturer"] = "Não identificado"
        info["model"] = "Não identificado"

    try:
        # Tipo de máquina (Desktop ou Notebook)
        result_type = subprocess.check_output(
            ["wmic", "computersystem", "get", "pcsystemtype"],
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        lines_type = [l.strip() for l in result_type.strip().splitlines() if l.strip()]
        pc_type_code = lines_type[1] if len(lines_type) > 1 else "0"
        pc_types = {
            "1": "Desktop",
            "2": "Notebook",
            "3": "Workstation",
            "4": "Enterprise Server",
            "8": "Tablet",
            "9": "Notebook"
        }
        info["pc_type"] = pc_types.get(pc_type_code, "Não identificado")
    except:
        info["pc_type"] = "Não identificado"

    # Data e hora do inventário
    info["inventory_date"] = datetime.now().strftime("%d/%m/%Y")
    info["inventory_time"] = datetime.now().strftime("%H:%M:%S")
    info["inventory_datetime"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_system_info("Luis Mesquitela")
    for key, value in data.items():
        print(f"{key}: {value}")
