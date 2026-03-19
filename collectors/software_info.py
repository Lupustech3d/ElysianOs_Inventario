import subprocess
import winreg
from datetime import datetime


def get_software_info():
    """Coleta programas instalados via Registro do Windows"""
    info = {}
    software_list = []

    # Chaves do registro para buscar programas (32 e 64 bits)
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    seen = set()

    for hive, path in reg_paths:
        try:
            reg_key = winreg.OpenKey(hive, path)
            num_subkeys = winreg.QueryInfoKey(reg_key)[0]

            for i in range(num_subkeys):
                try:
                    subkey_name = winreg.EnumKey(reg_key, i)
                    subkey = winreg.OpenKey(reg_key, subkey_name)

                    def get_val(key, name):
                        try:
                            return winreg.QueryValueEx(key, name)[0]
                        except:
                            return ""

                    name = get_val(subkey, "DisplayName")
                    version = get_val(subkey, "DisplayVersion")
                    publisher = get_val(subkey, "Publisher")
                    install_date = get_val(subkey, "InstallDate")
                    install_location = get_val(subkey, "InstallLocation")
                    size_kb = get_val(subkey, "EstimatedSize")

                    # Pula entradas sem nome ou duplicadas
                    if not name or name in seen:
                        winreg.CloseKey(subkey)
                        continue

                    seen.add(name)

                    # Formata data de instalação
                    date_formatted = ""
                    if install_date and len(install_date) == 8:
                        try:
                            d = datetime.strptime(install_date, "%Y%m%d")
                            date_formatted = d.strftime("%d/%m/%Y")
                        except:
                            date_formatted = install_date

                    # Formata tamanho
                    size_formatted = ""
                    if size_kb:
                        try:
                            size_mb = int(size_kb) / 1024
                            if size_mb >= 1024:
                                size_formatted = f"{size_mb / 1024:.1f} GB"
                            else:
                                size_formatted = f"{size_mb:.0f} MB"
                        except:
                            size_formatted = ""

                    software_list.append({
                        "name": name.strip(),
                        "version": version.strip() if version else "—",
                        "publisher": publisher.strip() if publisher else "—",
                        "install_date": date_formatted if date_formatted else "—",
                        "size": size_formatted if size_formatted else "—"
                    })

                    winreg.CloseKey(subkey)
                except:
                    continue

            winreg.CloseKey(reg_key)
        except:
            continue

    # Ordena por nome
    software_list.sort(key=lambda x: x["name"].lower())

    info["software_list"] = software_list
    info["software_count"] = len(software_list)

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_software_info()
    print(f"\nTotal de programas encontrados: {data['software_count']}\n")
    for sw in data["software_list"]:
        print(f"  {sw['name']} | {sw['version']} | {sw['publisher']} | {sw['install_date']} | {sw['size']}")