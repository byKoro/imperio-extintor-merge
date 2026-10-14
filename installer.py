import os
import subprocess
import sys
import time
import winreg


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def restart_explorer():
    """Reinicia o Windows Explorer de forma segura para atualizar os menus."""
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], capture_output=True, check=False)
        time.sleep(0.5)
        subprocess.Popen(["explorer.exe"])
    except Exception as e:
        print(f"Aviso ao reiniciar explorer: {e}")


def install_context_menu(exe_path=None):
    """Registra a opção 'Unificar' no menu de contexto do Windows com o ícone personalizado."""
    app_dir = get_app_dir()
    if not exe_path:
        exe_path = os.path.join(app_dir, "unificar.exe")
        if not os.path.exists(exe_path):
            exe_path = os.path.abspath(sys.argv[0])

    ico_path = os.path.join(app_dir, "icone.ico")
    icon_reg = ico_path if os.path.exists(ico_path) else f'"{exe_path}",0'

    title = "Unificar"

    # 1. Registra em HKCU\Software\Classes\*\shell\Unificar com filtro para .xlsm e .xlsx
    key_star = r"Software\Classes\*\shell\Unificar"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_star) as k:
        winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, title)
        winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon_reg)
        winreg.SetValueEx(k, "AppliesTo", 0, winreg.REG_SZ, "System.FileExtension:=.xlsm OR System.FileExtension:=.xlsx")
        winreg.SetValueEx(k, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
        with winreg.CreateKey(k, "command") as cmd:
            winreg.SetValueEx(cmd, "", 0, winreg.REG_SZ, f'"{exe_path}" "%1"')

    # 2. Registra em SystemFileAssociations para .xlsm e .xlsx
    for ext in [".xlsm", ".xlsx"]:
        key_ext = rf"Software\Classes\SystemFileAssociations\{ext}\shell\Unificar"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_ext) as k:
            winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, title)
            winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon_reg)
            winreg.SetValueEx(k, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
            with winreg.CreateKey(k, "command") as cmd:
                winreg.SetValueEx(cmd, "", 0, winreg.REG_SZ, f'"{exe_path}" "%1"')

    # 3. Cria atalho no Enviar para (SendTo)
    sendto_dir = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\SendTo")
    if os.path.isdir(sendto_dir):
        sendto_bat = os.path.join(sendto_dir, "Unificar Planilhas.bat")
        with open(sendto_bat, "w", encoding="latin-1") as f:
            f.write(f'@start "" "{exe_path}" %*\n')

    print("Menu de contexto instalado com sucesso!")
    return True


def uninstall_context_menu():
    """Remove a opção 'Unificar' do menu de contexto."""
    keys_to_delete = [
        r"Software\Classes\*\shell\Unificar\command",
        r"Software\Classes\*\shell\Unificar",
        r"Software\Classes\SystemFileAssociations\.xlsm\shell\Unificar\command",
        r"Software\Classes\SystemFileAssociations\.xlsm\shell\Unificar",
        r"Software\Classes\SystemFileAssociations\.xlsx\shell\Unificar\command",
        r"Software\Classes\SystemFileAssociations\.xlsx\shell\Unificar",
    ]
    for key_path in keys_to_delete:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
        except OSError:
            pass

    sendto_dir = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\SendTo")
    sendto_bat = os.path.join(sendto_dir, "Unificar Planilhas.bat")
    if os.path.exists(sendto_bat):
        try:
            os.remove(sendto_bat)
        except OSError:
            pass

    print("Menu de contexto desinstalado com sucesso!")
    return True


def is_win11_classic_menu_enabled():
    """Verifica se o menu direto do Windows 11 já está ativado."""
    key_path = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
            return True
    except OSError:
        return False


def enable_win11_classic_menu(restart=True):
    """Ativa o menu de contexto direto no Windows 11 (exibe no 1º clique)."""
    key_path = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "")
    print("Menu direto ativado no Windows 11!")
    if restart:
        restart_explorer()
    return True


def restore_win11_modern_menu(restart=True):
    """Restaura o menu moderno padrão do Windows 11."""
    key_path = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r"\InprocServer32")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
        print("Menu moderno restaurado no Windows 11!")
    except OSError:
        pass
    if restart:
        restart_explorer()
    return True


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--uninstall":
        uninstall_context_menu()
    elif len(sys.argv) > 1 and sys.argv[1] == "--win11-direct":
        enable_win11_classic_menu()
    elif len(sys.argv) > 1 and sys.argv[1] == "--win11-restore":
        restore_win11_modern_menu()
    else:
        test_exe = os.path.abspath("unificar.exe")
        install_context_menu(test_exe)
