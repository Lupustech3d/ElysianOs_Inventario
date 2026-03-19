import subprocess
import winreg


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


def get_license_info():
    """Coleta licenças do Windows e Office"""
    info = {}

    # ── WINDOWS ──────────────────────────────────────────

    # Edição do Windows
    try:
        edition = run_powershell("(Get-WmiObject Win32_OperatingSystem).Caption")
        info["windows_edition"] = edition if edition else "Não identificado"
    except:
        info["windows_edition"] = "Não identificado"

    # Build do Windows
    try:
        build = run_powershell("(Get-WmiObject Win32_OperatingSystem).Version")
        info["windows_build"] = build if build else "Não identificado"
    except:
        info["windows_build"] = "Não identificado"

    # Status de ativação do Windows
    try:
        status = run_powershell(
            "$sl = Get-WmiObject SoftwareLicensingProduct -Filter \"Name like 'Windows%'\" | "
            "Where-Object { $_.PartialProductKey } | Select-Object -First 1; "
            "switch ($sl.LicenseStatus) { "
            "0 { 'Nao licenciado' } "
            "1 { 'Ativado' } "
            "2 { 'Fora do periodo de cortesia' } "
            "3 { 'Fora do periodo de cortesia offline' } "
            "4 { 'Periodo de avaliacao' } "
            "5 { 'Periodo de notificacao' } "
            "6 { 'Periodo de cortesia estendido' } "
            "default { 'Status desconhecido' } }"
        )
        info["windows_status"] = status if status else "Não identificado"
    except:
        info["windows_status"] = "Não identificado"

    # Product Key parcial do Windows
    try:
        key_partial = run_powershell(
            "$sl = Get-WmiObject SoftwareLicensingProduct -Filter \"Name like 'Windows%'\" | "
            "Where-Object { $_.PartialProductKey } | Select-Object -First 1; "
            "$sl.PartialProductKey"
        )
        info["windows_key_partial"] = f"XXXXX-XXXXX-XXXXX-XXXXX-{key_partial}" if key_partial else "Não identificado"
    except:
        info["windows_key_partial"] = "Não identificado"

    # Tipo de licença do Windows
    try:
        lic_type = run_powershell(
            "$sl = Get-WmiObject SoftwareLicensingProduct -Filter \"Name like 'Windows%'\" | "
            "Where-Object { $_.PartialProductKey } | Select-Object -First 1; "
            "$sl.Description"
        )
        info["windows_license_type"] = lic_type if lic_type else "Não identificado"
    except:
        info["windows_license_type"] = "Não identificado"

    # ── OFFICE ───────────────────────────────────────────

    # Versão do Office via registro
    office_version = ""
    office_name = "Não encontrado"

    office_versions = {
        "16.0": "Office 2016 / 2019 / 2021 / 365",
        "15.0": "Office 2013",
        "14.0": "Office 2010",
        "12.0": "Office 2007",
    }

    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Office"),
    ]

    for hive, path in reg_paths:
        try:
            key = winreg.OpenKey(hive, path)
            num = winreg.QueryInfoKey(key)[0]
            for i in range(num):
                try:
                    sub = winreg.EnumKey(key, i)
                    if sub in office_versions:
                        office_version = sub
                        office_name = office_versions[sub]
                        break
                except:
                    continue
            winreg.CloseKey(key)
            if office_version:
                break
        except:
            continue

    info["office_version"] = office_name

    # Detecta licença Office via SoftwareLicensingProduct
    try:
        o365_check = run_powershell(
            "$sl = Get-WmiObject SoftwareLicensingProduct -Filter \"Name like 'Office%'\" | "
            "Where-Object { $_.PartialProductKey } | Select-Object -First 1; "
            "if ($sl) { $sl.Name + '|' + $sl.PartialProductKey + '|' + $sl.LicenseStatus }"
        )
        if o365_check and "|" in o365_check:
            parts = o365_check.split("|")
            info["office_product"] = parts[0].strip() if parts[0] else "Não identificado"
            info["office_key_partial"] = f"XXXXX-XXXXX-XXXXX-XXXXX-{parts[1].strip()}" if parts[1] else "Não identificado"
            status_map = {"0": "Não licenciado", "1": "Ativado", "2": "Fora do período", "5": "Período de notificação"}
            info["office_status"] = status_map.get(parts[2].strip() if len(parts) > 2 else "0", "Verificar manualmente")
        else:
            info["office_product"] = office_name
            info["office_key_partial"] = "Não identificado"
            info["office_status"] = "Não identificado"
    except:
        info["office_product"] = "Não identificado"
        info["office_key_partial"] = "Não identificado"
        info["office_status"] = "Não identificado"

    # Conta vinculada ao Office 365
    try:
        account = run_powershell(
            r"$path = $env:LOCALAPPDATA + '\Microsoft\Office\16.0\Licensing'; "
            "if (Test-Path $path) { "
            "Get-ChildItem $path -ErrorAction SilentlyContinue | "
            "Select-Object -First 1 -ExpandProperty Name }"
        )
        info["office_account"] = account if account else "Não identificado / Licença local"
    except:
        info["office_account"] = "Não identificado"

    # Detecta aplicativos Office instalados
    try:
        apps_result = run_powershell(
            "$officeApps = @('WINWORD','EXCEL','POWERPNT','OUTLOOK','MSACCESS','MSPUB','ONENOTE'); "
            "$found = @(); "
            "foreach ($app in $officeApps) { "
            r"$p1 = 'C:\Program Files\Microsoft Office\root\Office16\' + $app + '.EXE'; "
            r"$p2 = 'C:\Program Files (x86)\Microsoft Office\root\Office16\' + $app + '.EXE'; "
            "if ((Test-Path $p1) -or (Test-Path $p2)) { $found += $app } }; "
            "$found -join ','"
        )
        name_map = {
            "WINWORD": "Word", "EXCEL": "Excel", "POWERPNT": "PowerPoint",
            "OUTLOOK": "Outlook", "MSACCESS": "Access",
            "MSPUB": "Publisher", "ONENOTE": "OneNote"
        }
        if apps_result:
            info["office_apps"] = [name_map.get(a.strip(), a.strip()) for a in apps_result.split(",") if a.strip()]
        else:
            info["office_apps"] = []
    except:
        info["office_apps"] = []

    return info


# Teste direto do módulo
if __name__ == "__main__":
    data = get_license_info()
    for key, value in data.items():
        if isinstance(value, list):
            print(f"{key}: {', '.join(value)}")
        else:
            print(f"{key}: {value}")