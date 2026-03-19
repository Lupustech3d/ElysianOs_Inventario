from collectors.system_info import get_system_info
from collectors.hardware_info import get_hardware_info
from collectors.network_info import get_network_info
from collectors.software_info import get_software_info
from collectors.license_info import get_license_info
from collectors.printer_info import get_printer_info
from collectors.mapping_info import get_mapping_info


def collect_all(collaborator_name=""):
    """Executa todos os coletores e retorna dicionário completo"""
    print("[1/7] Coletando informações do sistema...")
    system = get_system_info(collaborator_name)

    print("[2/7] Coletando informações de hardware...")
    hardware = get_hardware_info()

    print("[3/7] Coletando informações de rede...")
    network = get_network_info()

    print("[4/7] Coletando softwares instalados...")
    software = get_software_info()

    print("[5/7] Coletando licenças...")
    licenses = get_license_info()

    print("[6/7] Coletando impressoras...")
    printers = get_printer_info()

    print("[7/7] Coletando mapeamentos de rede...")
    mappings = get_mapping_info()

    return {
        "system": system,
        "hardware": hardware,
        "network": network,
        "software": software,
        "licenses": licenses,
        "printers": printers,
        "mappings": mappings
    }
