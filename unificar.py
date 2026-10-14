"""
Unificador de Planilhas - Império do Extintor
Permite selecionar 2 planilhas (.xlsm ou .xlsx), unificá-las com resolução
interativa de conflitos, indicador visual de loading e integração direta com
o menu de contexto do Windows 11.
"""
import datetime
import json
import logging
import msvcrt
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from merge_planilhas import execute_merge

SESSION_FILE = os.path.join(tempfile.gettempdir(), "unificar_session.json")
LOCK_FILE = os.path.join(tempfile.gettempdir(), "unificar_session.lock")


def get_app_dir():
    """Retorna o diretório onde o executável ou script está localizado."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# Configura arquivo de log
LOG_FILE = os.path.join(get_app_dir(), "unificar.log")


def log_msg(msg):
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{now}] [PID {os.getpid()}] {msg}\n")
    except Exception:
        pass


def apply_window_icon(window):
    """Aplica o ícone .ico estilizado na janela Tkinter."""
    ico_path = os.path.join(get_app_dir(), "icone.ico")
    if os.path.isfile(ico_path):
        try:
            window.iconbitmap(ico_path)
        except Exception:
            pass


def get_logo_image(size=(48, 48)):
    """Carrega o ícone PNG estilizado para exibição interna nas janelas."""
    png_path = os.path.join(get_app_dir(), "icone.png")
    if os.path.isfile(png_path):
        try:
            from PIL import Image, ImageTk
            im = Image.open(png_path).resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(im)
        except Exception:
            pass
    return None


def get_base_template():
    """Localiza BASE.xlsm ou MODELO.xlsm na pasta do aplicativo."""
    app_dir = get_app_dir()
    for name in ["BASE.xlsm", "MODELO.xlsm", "base.xlsm", "modelo.xlsm",
                 "BASE.xlsx", "MODELO.xlsx"]:
        path = os.path.join(app_dir, name)
        if os.path.isfile(path):
            return path

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        for name in ["BASE.xlsm", "MODELO.xlsm"]:
            path = os.path.join(sys._MEIPASS, name)
            if os.path.isfile(path):
                return path

    for name in ["BASE.xlsm", "MODELO.xlsm"]:
        if os.path.isfile(name):
            return os.path.abspath(name)

    return None


def safe_copy(src, dst, dst_name, max_retries=5):
    """Copia src para dst com tratamento de bloqueio do Excel e tentativa de repetição."""
    for attempt in range(max_retries):
        try:
            shutil.copy2(src, dst)
            return True
        except PermissionError:
            root = tk.Tk()
            root.withdraw()
            apply_window_icon(root)
            ans = messagebox.askretrycancel(
                "Arquivo Aberto no Excel",
                f"O arquivo '{dst_name}' não pôde ser atualizado porque está aberto no Excel ou em outro programa.\n\n"
                f"Por favor, FECHE O ARQUIVO NO EXCEL e clique em 'Repetir'."
            )
            root.destroy()
            if not ans:
                return False
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar '{dst_name}':\n{str(e)}")
            root.destroy()
            return False
    return False


def coordinate_arguments():
    """Coordena instâncias concorrentes disparadas pelo Windows Explorer."""
    log_msg(f"coordinate_arguments com argv: {sys.argv}")
    if len(sys.argv) >= 3:
        return [os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])]
    if len(sys.argv) <= 1:
        return []

    incoming_file = os.path.abspath(sys.argv[1])
    lock_fd = None

    for _ in range(25):
        try:
            lock_fd = os.open(LOCK_FILE, os.O_CREAT | os.O_RDWR)
            msvcrt.locking(lock_fd, msvcrt.LK_NBLCK, 1)
            break
        except (IOError, OSError):
            time.sleep(0.04)
    else:
        log_msg(f"Não conseguiu lock, retornando incoming_file: {incoming_file}")
        return [incoming_file]

    try:
        data = {"files": [], "time": 0}
        if os.path.exists(SESSION_FILE):
            try:
                with open(SESSION_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                pass

        now = time.time()
        if now - data.get("time", 0) > 3.0:
            data = {"files": [incoming_file], "time": now, "leader_pid": os.getpid()}
            is_leader = True
            log_msg(f"Novo grupo iniciado pelo lider PID {os.getpid()} com: {incoming_file}")
        else:
            if incoming_file not in data.get("files", []):
                data.setdefault("files", []).append(incoming_file)
            is_leader = (data.get("leader_pid") == os.getpid())
            log_msg(f"Processo PID {os.getpid()} adicionou {incoming_file}. Leader={is_leader}")

        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    finally:
        try:
            if lock_fd is not None:
                msvcrt.locking(lock_fd, msvcrt.LK_UNLCK, 1)
                os.close(lock_fd)
        except Exception:
            pass

    if not is_leader:
        log_msg(f"Instancia seguidora PID {os.getpid()} encerrando.")
        sys.exit(0)

    start_wait = time.time()
    while time.time() - start_wait < 1.5:
        time.sleep(0.05)
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                if len(d.get("files", [])) >= 2:
                    data = d
                    log_msg(f"Lider encontrou 2 arquivos apos {time.time() - start_wait:.2f}s: {data['files']}")
                    break
        except Exception:
            pass

    try:
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)
    except Exception:
        pass

    final_files = data.get("files", [incoming_file])
    log_msg(f"Lider retornando arquivos finais: {final_files}")
    return final_files


# -------------------------------------------------------------
# Janela Modal de Resolução Interativa de Conflitos
# -------------------------------------------------------------

def prompt_conflict_resolution(file1, file2, conflicts, log_lines, base_file):
    """
    Exibe janela modal detalhada permitindo ao usuário:
    - Ver a lista de todos os conflitos (localização, valor base, valor planilha 1, valor planilha 2)
    - Abrir o relatório completo no Bloco de Notas
    - Escolher priorizar Planilha 1 ('minha') ou Planilha 2 ('dela')
    - Ou ABORTAR a unificação sem modificar nenhum arquivo ('abort')
    """
    name1 = os.path.basename(file1)
    name2 = os.path.basename(file2)
    num_conflicts = len(conflicts)

    dialog = tk.Tk()
    dialog.title("⚠️ Resolução de Conflitos - Império do Extintor")
    dialog.geometry("740x600")
    dialog.minsize(620, 480)
    apply_window_icon(dialog)

    # Centraliza na tela
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (740 // 2)
    y = (dialog.winfo_screenheight() // 2) - (600 // 2)
    dialog.geometry(f"+{x}+{y}")

    style = ttk.Style(dialog)
    style.theme_use("clam")

    main_frame = ttk.Frame(dialog, padding=16)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Topo com logo e título
    top_frame = ttk.Frame(main_frame)
    top_frame.pack(fill=tk.X, pady=(0, 10))

    logo_img = get_logo_image(size=(48, 48))
    if logo_img:
        lbl_logo = ttk.Label(top_frame, image=logo_img)
        lbl_logo.image = logo_img
        lbl_logo.pack(side=tk.LEFT, padx=(0, 12))

    header_text_frame = ttk.Frame(top_frame)
    header_text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

    lbl_title = ttk.Label(
        header_text_frame,
        text=f"⚠️ {num_conflicts} Conflito(s) Detectado(s) na Unificação!",
        font=("Segoe UI", 12, "bold"),
        foreground="#b45309"
    )
    lbl_title.pack(anchor="w")

    desc_msg = (
        f"A mesma célula ou pedido foi editado com valores diferentes em cada planilha.\n"
        f"Escolha qual planilha você deseja PRIORIZAR nos campos com conflito, ou ABORTE a operação:"
    )
    lbl_desc = ttk.Label(header_text_frame, text=desc_msg, font=("Segoe UI", 9), wraplength=580)
    lbl_desc.pack(anchor="w", pady=(2, 0))

    # Tabela com preview dos conflitos (Treeview)
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 12))

    columns = ("where", "base", "val_a", "val_b")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

    tree.heading("where", text="Local / Célula")
    tree.heading("base", text="Valor Original (BASE)")
    tree.heading("val_a", text=f"Versão em {name1}")
    tree.heading("val_b", text=f"Versão em {name2}")

    tree.column("where", width=180, anchor="w")
    tree.column("base", width=140, anchor="center")
    tree.column("val_a", width=170, anchor="center")
    tree.column("val_b", width=170, anchor="center")

    scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    tree.grid(row=0, column=0, sticky="nsew")
    scroll_y.grid(row=0, column=1, sticky="ns")
    scroll_x.grid(row=1, column=0, sticky="ew")
    tree_frame.rowconfigure(0, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    for item in conflicts:
        w = item.get("where", "")
        b = str(item.get("base", ""))
        va = str(item.get("val_a", ""))
        vb = str(item.get("val_b", ""))
        tree.insert("", "end", values=(w, b, va, vb))

    # Opções de Seleção de Prioridade
    choice_frame = ttk.LabelFrame(main_frame, text=" Escolha a Prioridade para os Conflitos ", padding=12)
    choice_frame.pack(fill=tk.X, pady=(0, 12))

    selected_preference = tk.StringVar(value="minha")

    rb1 = ttk.Radiobutton(
        choice_frame,
        text=f"Priorizar '{name1}' (manter as edições de '{name1}' nos campos conflitantes)",
        variable=selected_preference,
        value="minha"
    )
    rb1.pack(anchor="w", pady=(2, 4))

    rb2 = ttk.Radiobutton(
        choice_frame,
        text=f"Priorizar '{name2}' (manter as edições de '{name2}' nos campos conflitantes)",
        variable=selected_preference,
        value="dela"
    )
    rb2.pack(anchor="w", pady=(2, 2))

    # Barra de Botões de Ação
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill=tk.X)

    user_decision = {"choice": "abort"}

    def open_full_report():
        temp_rep = os.path.join(tempfile.gettempdir(), f"conflitos_{int(time.time())}.txt")
        with open(temp_rep, "w", encoding="utf-8") as f:
            f.write("RELATÓRIO DETALHADO DE CONFLITOS - IMPÉRIO DO EXTINTOR\n")
            f.write(f"Arquivo 1: {file1}\n")
            f.write(f"Arquivo 2: {file2}\n")
            f.write(f"Base: {base_file}\n")
            f.write(f"Total de conflitos: {num_conflicts}\n")
            f.write("=" * 65 + "\n\n")
            for line in log_lines:
                f.write(line + "\n")
        try:
            os.startfile(temp_rep)
        except Exception:
            pass

    btn_report = ttk.Button(btn_frame, text="📄 Abrir no Bloco de Notas", command=open_full_report)
    btn_report.pack(side=tk.LEFT)

    def do_abort():
        user_decision["choice"] = "abort"
        dialog.destroy()

    def do_confirm():
        user_decision["choice"] = selected_preference.get()
        dialog.destroy()

    btn_abort = ttk.Button(btn_frame, text="❌ Abortar (Não Alterar Nada)", command=do_abort)
    btn_abort.pack(side=tk.RIGHT, padx=(6, 0))

    btn_confirm = ttk.Button(btn_frame, text="✔️ Confirmar e Aplicar", command=do_confirm)
    btn_confirm.pack(side=tk.RIGHT)

    dialog.protocol("WM_DELETE_WINDOW", do_abort)
    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after_idle(dialog.attributes, "-topmost", False)
    dialog.mainloop()

    return user_decision["choice"]


# -------------------------------------------------------------
# Janela de Sucesso e Diálogos de Notificação
# -------------------------------------------------------------

def show_success_dialog_with_choice(file1_name, file2_name, total_rows, num_conflicts, chosen_name, report_path):
    """Exibe mensagem de sucesso quando a unificação conclui com conflitos resolvidos."""
    root = tk.Tk()
    root.title("Unificação Concluída - Império do Extintor")
    root.geometry("540x280")
    root.resizable(False, False)
    apply_window_icon(root)

    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (540 // 2)
    y = (root.winfo_screenheight() // 2) - (280 // 2)
    root.geometry(f"+{x}+{y}")

    style = ttk.Style(root)
    style.theme_use("clam")

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    lbl_head = ttk.Label(
        frame,
        text="✅ Planilhas Unificadas com Sucesso!",
        font=("Segoe UI", 12, "bold"),
        foreground="#166534"
    )
    lbl_head.pack(anchor="w", pady=(0, 6))

    msg = (
        f"Ambos os arquivos foram substituídos pela versão combinada e sincronizada:\n"
        f"  • {file1_name}\n"
        f"  • {file2_name}\n\n"
        f"Total de linhas na BASE TROCA: {total_rows}\n"
        f"Conflitos resolvidos: {num_conflicts} (priorizado: '{chosen_name}')\n"
        f"Relatório salvo em: conflitos.txt"
    )
    lbl_body = ttk.Label(frame, text=msg, font=("Segoe UI", 9), wraplength=500)
    lbl_body.pack(anchor="w", pady=(0, 14))

    btn_box = ttk.Frame(frame)
    btn_box.pack(fill=tk.X)

    def open_rep():
        if os.path.exists(report_path):
            os.startfile(report_path)

    if os.path.exists(report_path):
        ttk.Button(btn_box, text="📄 Abrir conflitos.txt", command=open_rep).pack(side=tk.LEFT)

    ttk.Button(btn_box, text="Concluir", command=root.destroy).pack(side=tk.RIGHT)

    root.lift()
    root.attributes("-topmost", True)
    root.after_idle(root.attributes, "-topmost", False)
    root.mainloop()


def show_success_dialog(file1_name, file2_name, total_rows):
    """Exibe mensagem de sucesso quando não houve nenhum conflito."""
    root = tk.Tk()
    root.withdraw()
    apply_window_icon(root)
    msg = (
        f"✅ PLANILHAS UNIFICADAS COM SUCESSO!\n\n"
        f"Ambos os arquivos foram atualizados com a versão combinada:\n\n"
        f"  1. {file1_name}\n"
        f"  2. {file2_name}\n\n"
        f"Total de linhas na BASE TROCA: {total_rows}\n"
        f"Nenhum conflito encontrado. Versões 100% sincronizadas!"
    )
    messagebox.showinfo("Unificação Concluída", msg)
    root.destroy()


def show_error_dialog(title, message):
    """Exibe mensagem de erro."""
    root = tk.Tk()
    root.withdraw()
    apply_window_icon(root)
    messagebox.showerror(title, message)
    root.destroy()


# -------------------------------------------------------------
# Motor de Unificação com Indicador de Loading
# -------------------------------------------------------------

def run_unification(file1, file2):
    """
    Executa o fluxo completo com janela de loading em thread separada,
    resolução interativa de conflitos e substituição segura de ambos os arquivos.
    """
    name1 = os.path.basename(file1)
    name2 = os.path.basename(file2)
    log_msg(f"Iniciando run_unification: file1={file1}, file2={file2}")

    # 1. Valida existência dos arquivos selecionados
    if not os.path.isfile(file1) or not os.path.isfile(file2):
        show_error_dialog("Arquivo Não Encontrado", "Um dos arquivos selecionados não foi encontrado no disco.")
        return False

    # 2. Localiza arquivo base
    base_file = get_base_template()
    if not base_file:
        app_dir = get_app_dir()
        show_error_dialog(
            "Arquivo Base Não Encontrado",
            f"Não foi possível encontrar 'BASE.xlsm' ou 'MODELO.xlsm' na pasta:\n{app_dir}\n\n"
            f"Certifique-se de que BASE.xlsm ou MODELO.xlsm está presente na mesma pasta do executável."
        )
        return False

    # Container de resultados entre a thread de trabalho e a interface Tkinter
    worker_result = {
        "done": False,
        "success": False,
        "error": None,
        "total_rows": 0,
        "log_lines": [],
        "conflicts": [],
        "temp_out": None
    }

    # Janela de Loading
    loading_root = tk.Tk()
    loading_root.title("Unificando Planilhas - Império do Extintor")
    loading_root.geometry("520x250")
    loading_root.resizable(False, False)
    apply_window_icon(loading_root)

    loading_root.update_idletasks()
    x = (loading_root.winfo_screenwidth() // 2) - (520 // 2)
    y = (loading_root.winfo_screenheight() // 2) - (250 // 2)
    loading_root.geometry(f"+{x}+{y}")

    style = ttk.Style(loading_root)
    style.theme_use("clam")

    pad_frame = ttk.Frame(loading_root, padding=20)
    pad_frame.pack(fill=tk.BOTH, expand=True)

    header_frame = ttk.Frame(pad_frame)
    header_frame.pack(fill=tk.X, pady=(0, 12))

    logo_img = get_logo_image(size=(48, 48))
    if logo_img:
        lbl_ico = ttk.Label(header_frame, image=logo_img)
        lbl_ico.image = logo_img
        lbl_ico.pack(side=tk.LEFT, padx=(0, 12))

    title_box = ttk.Frame(header_frame)
    title_box.pack(side=tk.LEFT, fill=tk.X, expand=True)

    lbl_t = ttk.Label(title_box, text="Unificando Planilhas de Controle...", font=("Segoe UI", 12, "bold"), foreground="#0f172a")
    lbl_t.pack(anchor="w")

    lbl_sub = ttk.Label(
        title_box,
        text=f"Processando '{name1}' e '{name2}'",
        font=("Segoe UI", 9),
        foreground="#475569"
    )
    lbl_sub.pack(anchor="w", pady=(2, 0))

    # Barra de Progresso Animada (Indeterminada)
    prog_bar = ttk.Progressbar(pad_frame, mode="indeterminate", length=480)
    prog_bar.pack(fill=tk.X, pady=(8, 10))
    prog_bar.start(12)

    lbl_status = ttk.Label(
        pad_frame,
        text="Lendo dados, recalculando fórmulas e verificando divergências...",
        font=("Segoe UI", 9, "italic"),
        foreground="#334155"
    )
    lbl_status.pack(anchor="w")

    def run_merge_thread():
        ext = ".xlsm" if (file1.lower().endswith(".xlsm") or base_file.lower().endswith(".xlsm")) else ".xlsx"
        temp_dir = tempfile.gettempdir()
        temp_out = os.path.join(temp_dir, f"unificada_{int(time.time()*1000)}{ext}")

        try:
            log_msg(f"Iniciando thread de merge com base={base_file}")
            c_list = []
            rows, log_lines, structured = execute_merge(
                base_path=base_file,
                path_a=file1,
                path_b=file2,
                out_path=temp_out,
                prefer="minha",
                name_a=name1,
                name_b=name2,
                conflicts_list=c_list
            )
            worker_result["success"] = True
            worker_result["total_rows"] = rows
            worker_result["log_lines"] = log_lines
            worker_result["conflicts"] = structured
            worker_result["temp_out"] = temp_out
            log_msg(f"Thread de merge finalizada: {rows} linhas, {len(structured)} conflitos.")
        except Exception as e:
            log_msg(f"Erro na thread de merge: {e}")
            worker_result["success"] = False
            worker_result["error"] = str(e)
            if os.path.exists(temp_out):
                try: os.remove(temp_out)
                except Exception: pass
        finally:
            worker_result["done"] = True

    # Dispara worker em segundo plano
    t = threading.Thread(target=run_merge_thread, daemon=True)
    t.start()

    def check_worker():
        if worker_result["done"]:
            prog_bar.stop()
            loading_root.destroy()
        else:
            loading_root.after(80, check_worker)

    loading_root.after(100, check_worker)
    loading_root.lift()
    loading_root.attributes("-topmost", True)
    loading_root.after_idle(loading_root.attributes, "-topmost", False)
    loading_root.mainloop()

    # 3. Tratamento de Erros da Thread
    if not worker_result["success"]:
        show_error_dialog(
            "Erro na Unificação",
            f"Ocorreu um erro ao processar o merge das planilhas:\n\n{worker_result['error']}"
        )
        return False

    temp_output = worker_result["temp_out"]
    conflicts = worker_result["conflicts"]
    log_lines = worker_result["log_lines"]
    total_rows = worker_result["total_rows"]

    # 4. Caso haja conflitos: Pergunta ao usuário qual priorizar ou se deseja abortar!
    chosen_preference = "minha"
    work_dir = os.path.dirname(file1)
    report_file = os.path.join(work_dir, "conflitos.txt")

    if conflicts:
        decision = prompt_conflict_resolution(file1, file2, conflicts, log_lines, base_file)

        if decision == "abort":
            # Usuário optou por abortar: descarta temporários e sai sem alterar nada
            log_msg("Unificacao abortada pelo usuario na janela de conflitos.")
            if temp_output and os.path.exists(temp_output):
                try: os.remove(temp_output)
                except Exception: pass

            root = tk.Tk()
            root.withdraw()
            apply_window_icon(root)
            messagebox.showinfo(
                "Unificação Cancelada",
                "A unificação foi cancelada pelo usuário.\n\nNenhum arquivo foi alterado ou modificado."
            )
            root.destroy()
            return False

        chosen_preference = decision
        log_msg(f"Usuario escolheu priorizar: {chosen_preference}")

        # Se o usuário escolheu priorizar o Arquivo 2 (dela), refaz a saída com prefer='dela'
        if chosen_preference == "dela":
            try:
                execute_merge(
                    base_path=base_file,
                    path_a=file1,
                    path_b=file2,
                    out_path=temp_output,
                    prefer="dela",
                    name_a=name1,
                    name_b=name2
                )
            except Exception as e:
                log_msg(f"Erro ao regenerar com prioridade dela: {e}")
                show_error_dialog("Erro na Unificação", f"Não foi possível aplicar a prioridade escolhida:\n{e}")
                if os.path.exists(temp_output):
                    try: os.remove(temp_output)
                    except Exception: pass
                return False

    # 5. Cria backups de segurança (.bak)
    try:
        shutil.copy2(file1, file1 + ".bak")
        shutil.copy2(file2, file2 + ".bak")
        log_msg("Backups .bak criados com sucesso.")
    except Exception as e:
        log_msg(f"Aviso ao criar backups: {e}")

    # 6. Substitui AMBOS os arquivos com tratamento seguro e rollback
    f1_ok = safe_copy(temp_output, file1, name1)
    if not f1_ok:
        log_msg(f"Falha ao copiar para {name1}. Cancelado.")
        if os.path.exists(temp_output):
            try: os.remove(temp_output)
            except Exception: pass
        return False

    f2_ok = safe_copy(temp_output, file2, name2)
    if not f2_ok:
        log_msg(f"Falha ao copiar para {name2}. Fazendo rollback em {name1}.")
        if os.path.exists(file1 + ".bak"):
            try: shutil.copy2(file1 + ".bak", file1)
            except Exception: pass
        show_error_dialog(
            "Unificação Cancelada",
            f"O arquivo '{name2}' não pôde ser atualizado.\n"
            f"Para manter a consistência, as alterações foram revertidas.\n"
            f"Feche o Excel e tente novamente."
        )
        if os.path.exists(temp_output):
            try: os.remove(temp_output)
            except Exception: pass
        return False

    log_msg(f"Ambos os arquivos substituidos com sucesso: {name1} e {name2}")

    if os.path.exists(temp_output):
        try: os.remove(temp_output)
        except Exception: pass

    # 7. Grava relatório de conflitos se houver
    if conflicts:
        chosen_name = name1 if chosen_preference == "minha" else name2
        with open(report_file, "w", encoding="utf-8") as f:
            f.write("RELATÓRIO DE CONFLITOS NA UNIFICAÇÃO - IMPÉRIO DO EXTINTOR\n")
            f.write(f"Arquivo 1: {file1}\n")
            f.write(f"Arquivo 2: {file2}\n")
            f.write(f"Base utilizada: {base_file}\n")
            f.write(f"Prioridade escolhida pelo usuário: {chosen_name}\n")
            f.write(f"Total de linhas na BASE TROCA: {total_rows}\n")
            f.write(f"Total de conflitos resolvidos: {len(conflicts)}\n")
            f.write("=" * 65 + "\n\n")
            for c in log_lines:
                f.write(c + "\n")
        show_success_dialog_with_choice(name1, name2, total_rows, len(conflicts), chosen_name, report_file)
    else:
        if os.path.exists(report_file):
            try: os.remove(report_file)
            except Exception: pass
        show_success_dialog(name1, name2, total_rows)

    return True


# -------------------------------------------------------------
# Janela Principal Interativa (quando aberto sem argumentos)
# -------------------------------------------------------------

def open_main_window():
    """Janela principal exibida quando o usuário abre o aplicativo sem passar arquivos."""
    from installer import (
        enable_win11_classic_menu,
        install_context_menu,
        is_win11_classic_menu_enabled,
        restore_win11_modern_menu,
        uninstall_context_menu,
    )

    root = tk.Tk()
    root.title("Unificador de Planilhas - Império do Extintor")
    root.geometry("640x440")
    root.resizable(False, False)
    apply_window_icon(root)

    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (640 // 2)
    y = (root.winfo_screenheight() // 2) - (440 // 2)
    root.geometry(f"+{x}+{y}")

    style = ttk.Style(root)
    style.theme_use("clam")

    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Topo com logo e título
    top_frame = ttk.Frame(main_frame)
    top_frame.pack(fill=tk.X, pady=(0, 14))

    logo_img = get_logo_image(size=(52, 52))
    if logo_img:
        lbl_ico = ttk.Label(top_frame, image=logo_img)
        lbl_ico.image = logo_img
        lbl_ico.pack(side=tk.LEFT, padx=(0, 12))

    t_box = ttk.Frame(top_frame)
    t_box.pack(side=tk.LEFT, fill=tk.X, expand=True)

    lbl_header = ttk.Label(
        t_box,
        text="Unificador de Planilhas de Controle",
        font=("Segoe UI", 13, "bold"),
        foreground="#1e3a8a"
    )
    lbl_header.pack(anchor="w")

    lbl_sub = ttk.Label(
        t_box,
        text="Selecione as duas planilhas para combinar. Ambas serão sincronizadas.",
        font=("Segoe UI", 9),
        foreground="#475569"
    )
    lbl_sub.pack(anchor="w", pady=(2, 0))

    # Seleção Planilha 1
    f1_frame = ttk.Frame(main_frame)
    f1_frame.pack(fill=tk.X, pady=(0, 10))
    ttk.Label(f1_frame, text="Planilha 1:", width=12).pack(side=tk.LEFT)
    var_f1 = tk.StringVar()
    entry_f1 = ttk.Entry(f1_frame, textvariable=var_f1)
    entry_f1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_f1():
        p = filedialog.askopenfilename(
            title="Selecione a Planilha 1",
            filetypes=[("Planilhas Excel", "*.xlsm;*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        if p:
            var_f1.set(p)
    ttk.Button(f1_frame, text="Procurar...", command=choose_f1).pack(side=tk.RIGHT)

    # Seleção Planilha 2
    f2_frame = ttk.Frame(main_frame)
    f2_frame.pack(fill=tk.X, pady=(0, 16))
    ttk.Label(f2_frame, text="Planilha 2:", width=12).pack(side=tk.LEFT)
    var_f2 = tk.StringVar()
    entry_f2 = ttk.Entry(f2_frame, textvariable=var_f2)
    entry_f2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_f2():
        p = filedialog.askopenfilename(
            title="Selecione a Planilha 2",
            filetypes=[("Planilhas Excel", "*.xlsm;*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
        if p:
            var_f2.set(p)
    ttk.Button(f2_frame, text="Procurar...", command=choose_f2).pack(side=tk.RIGHT)

    # Botão Unificar
    def do_unify():
        p1 = var_f1.get().strip()
        p2 = var_f2.get().strip()
        if not p1 or not p2:
            messagebox.showwarning("Atenção", "Por favor, selecione as duas planilhas.")
            return
        if not os.path.isfile(p1) or not os.path.isfile(p2):
            messagebox.showerror("Erro", "Um dos arquivos selecionados não existe.")
            return
        if os.path.abspath(p1) == os.path.abspath(p2):
            messagebox.showwarning("Atenção", "Selecione dois arquivos diferentes.")
            return
        root.withdraw()
        run_unification(p1, p2)
        root.destroy()

    btn_unify = ttk.Button(main_frame, text="🚀 Iniciar Unificação", command=do_unify)
    btn_unify.pack(pady=(0, 14), ipady=5, fill=tk.X)

    # Divisor
    ttk.Separator(main_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 10))

    # Painel do Menu de Contexto
    ctx_frame = ttk.LabelFrame(main_frame, text=" Integração com o Botão Direito do Mouse ", padding=10)
    ctx_frame.pack(fill=tk.X)

    lbl_info_ctx = ttk.Label(
        ctx_frame,
        text="Permite selecionar 2 planilhas no Windows Explorer e clicar em 'Unificar':",
        font=("Segoe UI", 8)
    )
    lbl_info_ctx.pack(anchor="w", pady=(0, 8))

    row_btns = ttk.Frame(ctx_frame)
    row_btns.pack(fill=tk.X)

    def do_install():
        install_context_menu()
        messagebox.showinfo("Sucesso", "Menu 'Unificar' instalado com sucesso com o novo ícone!")

    def do_uninstall():
        uninstall_context_menu()
        messagebox.showinfo("Sucesso", "Menu 'Unificar' removido do botão direito.")

    def toggle_win11():
        if is_win11_classic_menu_enabled():
            restore_win11_modern_menu()
            messagebox.showinfo("Windows 11", "Menu padrão do Windows 11 restaurado.")
        else:
            enable_win11_classic_menu()
            messagebox.showinfo("Windows 11", "Menu direto ativado! Agora o menu completo abre no 1º clique.")
        update_win11_btn()

    btn_inst = ttk.Button(row_btns, text="📌 Instalar Menu", command=do_install)
    btn_inst.pack(side=tk.LEFT, padx=(0, 6))

    btn_win11 = ttk.Button(row_btns, text="", command=toggle_win11)
    btn_win11.pack(side=tk.LEFT, padx=(0, 6))

    def update_win11_btn():
        if is_win11_classic_menu_enabled():
            btn_win11.config(text="🔄 Menu Padrão W11")
        else:
            btn_win11.config(text="⚡ Menu Direto W11 (1º Clique)")

    update_win11_btn()

    btn_uninst = ttk.Button(row_btns, text="🗑️ Remover", command=do_uninstall)
    btn_uninst.pack(side=tk.RIGHT)

    root.mainloop()


# -------------------------------------------------------------
# Ponto de Entrada Principal
# -------------------------------------------------------------

def main():
    global show_success_dialog, show_error_dialog
    cli_mode = "--cli" in sys.argv
    if cli_mode:
        sys.argv.remove("--cli")

    if len(sys.argv) > 1 and sys.argv[1] == "--install":
        from installer import install_context_menu
        install_context_menu()
        if not cli_mode:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Menu Instalado", "Menu 'Unificar' registrado no botão direito com sucesso!")
            root.destroy()
        return

    if len(sys.argv) > 1 and sys.argv[1] == "--uninstall":
        from installer import uninstall_context_menu
        uninstall_context_menu()
        if not cli_mode:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Menu Removido", "Menu 'Unificar' removido do botão direito com sucesso!")
            root.destroy()
        return

    # Coordena arquivos recebidos do Windows Explorer
    files = coordinate_arguments()

    if len(files) == 0:
        open_main_window()
        return

    if len(files) == 1:
        if not cli_mode:
            show_error_dialog(
                "Seleção Incompleta",
                f"Você selecionou apenas 1 arquivo:\n{files[0]}\n\n"
                f"Para unificar, selecione EXATAMENTE 2 planilhas no Windows Explorer,\n"
                f"clique com o botão direito e escolha 'Unificar'."
            )
        else:
            print("Erro: Apenas 1 arquivo selecionado.")
        return

    if len(files) > 2:
        if not cli_mode:
            show_error_dialog(
                "Muitos Arquivos Selecionados",
                f"Foram selecionados {len(files)} arquivos.\n\n"
                f"Por favor, selecione EXATAMENTE 2 planilhas (.xlsm ou .xlsx) para unificar."
            )
        else:
            print("Erro: Mais de 2 arquivos selecionados.")
        return

    # Exatamente 2 arquivos: dispara o fluxo completo
    file1, file2 = files[0], files[1]
    if cli_mode:
        orig_success = show_success_dialog
        orig_err = show_error_dialog
        show_success_dialog = lambda f1, f2, rows: print(f"CLI SUCESSO: {f1} e {f2} unificados! Linhas: {rows}")
        show_error_dialog = lambda t, m: print(f"CLI ERRO [{t}]: {m}")
        try:
            run_unification(file1, file2)
        finally:
            show_success_dialog = orig_success
            show_error_dialog = orig_err
    else:
        run_unification(file1, file2)


if __name__ == "__main__":
    main()
