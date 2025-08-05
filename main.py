import os
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import shutil
import glob
from pypdf import PdfWriter
import requests
import sys
from datetime import datetime
import re
from ttkthemes import ThemedTk
# Módulos para concorrência e fila segura
import concurrent.futures
import queue
# Módulo para o novo formato de configuração
import json

# Constante para o nome do arquivo de configuração (extensão alterada)
CONFIG_FILE = "config.dat"
# Constante para o número de downloads paralelos.
# Aumentar este valor pode acelerar os downloads, mas depende do limite do servidor.
MAX_DOWNLOAD_WORKERS = 15

def get_asset_path(file_name):
    """Retorna o caminho absoluto para um arquivo de recurso."""
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(application_path, file_name)

def merge_pdfs_do_tecnico(tecnico_folder, tecnico_name, app_instance, delete_previous, newly_downloaded_files):
    """Junta os PDFs de um técnico específico."""
    if app_instance.cancel_event.is_set():
        app_instance.log_queue.put(f"Junção cancelada para {tecnico_name}.")
        return

    try:
        files_to_merge = []
        output_filename = ""
        
        if delete_previous:
            output_filename = f"{str(tecnico_name).lower()}_compilado.pdf"
            files_to_merge = [f for f in glob.glob(os.path.join(tecnico_folder, '*.pdf')) if os.path.basename(f) != output_filename]
        else:
            if not newly_downloaded_files:
                app_instance.log_queue.put(f"Nenhum PDF novo para compilar para {tecnico_name}.")
                return
            
            datestamp = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"{str(tecnico_name).lower()}_compilado_{datestamp}.pdf"
            files_to_merge = newly_downloaded_files
        
        files_to_merge.sort()
        
        if not files_to_merge:
            app_instance.log_queue.put(f"Nenhum PDF para compilar para {tecnico_name}.")
            return

        merger = PdfWriter()
        app_instance.log_queue.put(f"Iniciando junção de {len(files_to_merge)} PDFs para {tecnico_name}...")
        
        for pdf in files_to_merge:
            if app_instance.cancel_event.is_set():
                app_instance.log_queue.put("Processo de junção interrompido pelo usuário.")
                merger.close()
                return
            merger.append(pdf)
            
        output_path = os.path.join(tecnico_folder, output_filename)
        merger.write(output_path)
        merger.close()
        
        app_instance.log_queue.put(f"PDF compilado salvo como: {output_filename}")

    except Exception as e:
        error_msg = f"ERRO ao juntar PDFs para {tecnico_name}: {e}"
        print(error_msg)
        app_instance.log_queue.put(error_msg)

def download_rat(session, rat_id, tecnico_folder, log_queue, cancel_event):
    """
    Baixa um único arquivo PDF (RAT).
    Retorna uma tupla (status, valor), onde status é 'success' ou 'error'.
    """
    if cancel_event.is_set():
        return ('cancelled', rat_id)

    # O log de "Baixando" foi movido para a função principal para não poluir a tela
    rat_download = rat_id[3:] if len(rat_id) > 9 else rat_id
    url_pdf = f"https://assist.positivotecnologia.com.br/bin/at/comprovantes/gerarRatPdf.php?os_id={rat_download}"
    
    try:
        pdf_response = session.get(url_pdf, stream=True, timeout=30)
        pdf_response.raise_for_status()
        
        disposition = pdf_response.headers.get('content-disposition')
        if disposition and "filename" in disposition:
            fname = re.findall('filename="(.+)"', disposition)[0]
        else:
            fname = f"{rat_id}.pdf"
        
        output_path = os.path.join(tecnico_folder, fname)
        with open(output_path, 'wb') as f:
            for chunk in pdf_response.iter_content(chunk_size=8192):
                if cancel_event.is_set():
                    return ('cancelled', rat_id)
                f.write(chunk)
        
        if os.path.exists(output_path):
            return ('success', output_path)
        else:
            raise Exception("Arquivo não foi salvo no disco.")

    except Exception as e:
        log_queue.put(f"ERRO ao baixar RAT {rat_id}. Causa: {e}")
        return ('error', rat_id)


def iniciar_processo_requests(app_instance):
    """Função principal que executa o processo de download e junção de PDFs."""
    total_ignored_chamados = []
    total_failed_chamados = []
    main_download_dir = app_instance.diretorio_selecionado
    delete_previous = app_instance.delete_previous_var.get()
    log_queue = app_instance.log_queue
    
    try:
        if app_instance.cancel_event.is_set(): return

        log_queue.put("Iniciando sessão web...")
        
        session = requests.Session()
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.0.0 Safari/537.36'}
        session.headers.update(headers)

        response = session.get("https://assist.positivotecnologia.com.br/", timeout=15)
        response.raise_for_status()
        log_queue.put("Sessão iniciada com sucesso.")
        
        log_queue.put("Lendo planilha para iniciar downloads...")
        chamados_path = get_asset_path('chamados.xlsx')
        df = pd.read_excel(chamados_path, dtype=str)

        original_columns = df.columns.tolist()
        cleaned_tecnico_names = []
        unnamed_counter = 1
        for name in original_columns:
            name_str = str(name).strip()
            if not name_str or name_str.startswith('Unnamed:'):
                cleaned_tecnico_names.append(f"Sem Técnico Definido {unnamed_counter}")
                unnamed_counter += 1
            else:
                cleaned_tecnico_names.append(name_str)

        if delete_previous:
            log_queue.put("\nModo de Limpeza Ativado...")
            for tecnico_name in set(cleaned_tecnico_names):
                tecnico_folder = os.path.join(main_download_dir, str(tecnico_name).lower())
                if os.path.isdir(tecnico_folder):
                    shutil.rmtree(tecnico_folder)
        else:
            log_queue.put("\nModo Incremental Ativado...")

        for i, original_col_name in enumerate(original_columns):
            if app_instance.cancel_event.is_set():
                log_queue.put("\nProcesso interrompido pelo usuário.")
                return

            tecnico_name = cleaned_tecnico_names[i]
            log_queue.put(f"\n--- Processando técnico: {tecnico_name} ---")
            
            chamados_list = df[original_col_name].dropna().tolist()
            valid_chamados = [r.strip() for r in chamados_list if str(r).strip().startswith("500")]
            total_ignored_chamados.extend([r.strip() for r in chamados_list if not str(r).strip().startswith("500")])

            tecnico_folder = os.path.join(main_download_dir, str(tecnico_name).lower())
            
            rats_to_download = []
            if delete_previous or not os.path.isdir(tecnico_folder):
                rats_to_download = valid_chamados
            else:
                try:
                    existing_files = os.listdir(tecnico_folder)
                    existing_rats = {rat for rat in valid_chamados for filename in existing_files if rat in filename}
                    rats_to_download = [rat for rat in valid_chamados if rat not in existing_rats]
                    log_queue.put(f"-> {len(existing_rats)} PDFs já existem. {len(rats_to_download)} novos downloads necessários.")
                except FileNotFoundError:
                    rats_to_download = valid_chamados

            if not rats_to_download:
                if valid_chamados: log_queue.put("Nenhum download novo. Pulando.")
                else: log_queue.put("Nenhum chamado válido na lista.")
                if delete_previous and os.path.isdir(tecnico_folder): merge_pdfs_do_tecnico(tecnico_folder, tecnico_name, app_instance, delete_previous, [])
                continue
                
            os.makedirs(tecnico_folder, exist_ok=True)
            
            newly_downloaded_files = []
            log_queue.put(f"Baixando {len(rats_to_download)} RATs para {tecnico_name}...")
            with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_DOWNLOAD_WORKERS) as executor:
                future_to_rat = {
                    executor.submit(download_rat, session, rat_id, tecnico_folder, log_queue, app_instance.cancel_event): rat_id 
                    for rat_id in rats_to_download
                }

                for future in concurrent.futures.as_completed(future_to_rat):
                    if app_instance.cancel_event.is_set():
                        for f in future_to_rat: f.cancel()
                        break 
                    
                    try:
                        status, value = future.result()
                        if status == 'success':
                            newly_downloaded_files.append(value)
                        elif status == 'error':
                            total_failed_chamados.append(value)
                    except Exception as exc:
                        rat_id = future_to_rat[future]
                        log_queue.put(f"ERRO INESPERADO ao processar {rat_id}: {exc}")
                        total_failed_chamados.append(rat_id)
            
            app_instance.total_downloads += len(newly_downloaded_files)
            if newly_downloaded_files:
                log_queue.put(f"-> Sucesso: {len(newly_downloaded_files)} de {len(rats_to_download)} RATs baixadas.")

            if not app_instance.cancel_event.is_set():
                merge_pdfs_do_tecnico(tecnico_folder, tecnico_name, app_instance, delete_previous, newly_downloaded_files)

        if app_instance.cancel_event.is_set():
            log_queue.put("\n------------------------------------")
            log_queue.put("Processo cancelado pelo usuário.")
        else:
            log_queue.put("\n------------------------------------")
            log_queue.put("Processo finalizado!")
        
        if total_ignored_chamados: log_queue.put("\nChamados IGNORADOS (não iniciam com 500):")
        for chamado in total_ignored_chamados: log_queue.put(f"- {chamado}")
        
        if total_failed_chamados: log_queue.put("\nChamados que FALHARAM no processo:")
        for chamado in total_failed_chamados: log_queue.put(f"- {chamado}")

    except Exception as e:
        log_queue.put(f"\nERRO GERAL: {e}")
    finally:
        app_instance.root.after(0, app_instance.reabilitar_botoes)

def validar_planilha(app_instance):
    """Valida a planilha de chamados antes de iniciar o processo principal."""
    log_queue = app_instance.log_queue
    try:
        if app_instance.cancel_event.is_set(): return
        log_queue.put("Validando planilha 'chamados.xlsx'...")
        chamados_path = get_asset_path('chamados.xlsx')
        if not os.path.exists(chamados_path):
            raise FileNotFoundError("Arquivo 'chamados.xlsx' não encontrado! Tente reiniciar o programa.")
        
        df = pd.read_excel(chamados_path, dtype=str)
        
        warnings = []
        for header in df.columns:
            header_str = str(header).strip()
            if not header_str or header_str.startswith('Unnamed:'):
                warnings.append("AVISO: Foi encontrada uma ou mais colunas com cabeçalho em branco.")
                break
        
        for header in df.columns:
            header_str = str(header).strip()
            if header_str.isdigit() and len(header_str) >= 9:
                warnings.append(f"AVISO: O cabeçalho '{header_str}' parece ser um número de chamado, não um nome de técnico.")

        seen_rats = {}
        for tecnico in df.columns:
            for rat in df[tecnico].dropna():
                rat_id = str(rat).strip()
                if not rat_id.startswith("500"): continue
                
                if rat_id not in seen_rats:
                    seen_rats[rat_id] = []
                seen_rats[rat_id].append(str(tecnico))
        
        duplicates = {rat: tecnicos for rat, tecnicos in seen_rats.items() if len(tecnicos) > 1}
        
        if duplicates:
            warnings.append("\nAVISO: Foram encontrados os seguintes chamados duplicados:")
            for rat, tecnicos in duplicates.items():
                cleaned_tecnicos = [str(t).strip() if not (str(t).strip().startswith('Unnamed:')) else "'Cabeçalho Vazio'" for t in tecnicos]
                warnings.append(f"- Chamado {rat} nas colunas: {', '.join(cleaned_tecnicos)}")
        
        if not warnings:
            log_queue.put("Nenhum problema encontrado na planilha. Iniciando processo...")
            app_instance.process_thread = threading.Thread(target=iniciar_processo_requests, args=(app_instance,), daemon=True)
            app_instance.process_thread.start()
        else:
            log_queue.put("\nATENÇÃO: Verifique os seguintes problemas na planilha:")
            for warning in warnings:
                log_queue.put(warning)
            
            log_queue.put("\nSe estiver correto, clique em 'Continuar'. Caso contrário, corrija a planilha e inicie novamente.")
            app_instance.root.after(0, lambda: app_instance.continue_button.config(state=tk.NORMAL))
            app_instance.root.after(0, lambda: app_instance.select_button.config(state=tk.DISABLED))
            app_instance.root.after(0, lambda: app_instance.start_button.config(state=tk.DISABLED))
            app_instance.root.after(0, lambda: app_instance.cancel_button.config(state=tk.NORMAL))
            app_instance.root.after(0, lambda: app_instance.open_sheet_button.config(state=tk.NORMAL))

    except Exception as e:
        error_msg = f"ERRO na Validação: {e}"
        print(error_msg)
        log_queue.put(error_msg)
        app_instance.root.after(0, app_instance.reabilitar_botoes)

class App:
    def __init__(self, root):
        self.root = root
        
        config = self._load_config()
        self.diretorio_selecionado = config.get('directory', '')
        
        self.process_thread = None
        self.cancel_event = threading.Event()
        self.log_queue = queue.Queue()
        self.start_time = None
        self.total_downloads = 0
        
        self.root.title("Assist RAT Downloader by. Sojo & Cabelo (v2.3)")
        self.root.geometry("600x550")
        
        style = ttk.Style(self.root)
        style.theme_use('equilux') 
        
        BG_COLOR = "#3c3c3c"
        FG_COLOR = "#ffffff"
        LABEL_FRAME_BG = "#4a4a4a"
        BUTTON_ACCENT_BG = "#0078D7"
        BUTTON_CANCEL_BG = "#c42b1c"
        BUTTON_ACTIVE_BG = "#5a5a5a" 
        BUTTON_ACCENT_ACTIVE_BG = "#005a9e"
        BUTTON_CANCEL_ACTIVE_BG = "#a32417"

        self.root.configure(bg=BG_COLOR)
        style.configure("TLabel", font=("Segoe UI", 10), background=BG_COLOR, foreground=FG_COLOR)
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=5)
        style.configure("TCheckbutton", font=("Segoe UI", 10), background=BG_COLOR, foreground=FG_COLOR)
        style.map("TCheckbutton", background=[('active', BG_COLOR)])
        style.configure("TLabelframe", background=BG_COLOR, borderwidth=1)
        style.configure("TLabelframe.Label", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=FG_COLOR)
        
        style.map("TButton",
            background=[('!disabled', BG_COLOR), ('active', BUTTON_ACTIVE_BG)],
            foreground=[('!disabled', FG_COLOR)]
        )
        style.configure("Accent.TButton", foreground=FG_COLOR, background=BUTTON_ACCENT_BG)
        style.map("Accent.TButton",
            background=[('!disabled', BUTTON_ACCENT_BG), ('active', BUTTON_ACCENT_ACTIVE_BG)]
        )
        style.configure("Cancel.TButton", foreground=FG_COLOR, background=BUTTON_CANCEL_BG)
        style.map("Cancel.TButton",
            background=[('!disabled', BUTTON_CANCEL_BG), ('active', BUTTON_CANCEL_ACTIVE_BG)]
        )

        main_frame = ttk.Frame(self.root, padding="15 15 15 15", style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        config_frame = ttk.LabelFrame(main_frame, text="Configurações", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 10))

        self.path_label_var = tk.StringVar(value="Nenhuma pasta de destino selecionada.")
        self._update_path_label()
        path_label = ttk.Label(config_frame, textvariable=self.path_label_var, wraplength=500)
        path_label.pack(fill=tk.X, expand=True, pady=(5, 10))
        path_label.configure(background=LABEL_FRAME_BG)
        
        self.select_button = ttk.Button(config_frame, text="Selecionar Pasta de Destino", command=self.selecionar_pasta)
        self.select_button.pack(fill=tk.X, expand=True, pady=5)
        
        self.open_sheet_button = ttk.Button(config_frame, text="Abrir Planilha de Chamados", command=self.abrir_planilha)
        self.open_sheet_button.pack(fill=tk.X, expand=True, pady=5)

        initial_delete_state = config.get('delete_previous', False)
        self.delete_previous_var = tk.BooleanVar(value=initial_delete_state)
        
        self.delete_check = ttk.Checkbutton(
            config_frame, 
            text="Apagar arquivos anteriores (modo de limpeza total)", 
            variable=self.delete_previous_var,
            command=self._save_config
        )
        self.delete_check.pack(anchor='w', pady=5)
        self.delete_check.configure(style="TCheckbutton")
        style.configure("TCheckbutton", background=LABEL_FRAME_BG)

        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=(10, 10))
        controls_frame.columnconfigure((0, 1, 2), weight=1)

        self.start_button = ttk.Button(controls_frame, text="Iniciar", command=self.iniciar_automacao, style="Accent.TButton")
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.start_button.config(state=tk.DISABLED if not self.diretorio_selecionado else tk.NORMAL)

        self.continue_button = ttk.Button(controls_frame, text="Continuar", command=self.continuar_automacao, state=tk.DISABLED)
        self.continue_button.grid(row=0, column=1, sticky="ew", padx=(5, 5))

        self.cancel_button = ttk.Button(controls_frame, text="Cancelar", command=self.cancelar_execucao, state=tk.DISABLED, style="Cancel.TButton")
        self.cancel_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', height=10, wrap=tk.WORD, 
                                                  font=("Consolas", 10), relief="flat",
                                                  background="#2b2b2b", foreground="#dcdcdc")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self._verificar_e_criar_planilha()
        self.process_log_queue()

    def _verificar_e_criar_planilha(self):
        filepath = get_asset_path('chamados.xlsx')
        if not os.path.exists(filepath):
            try:
                self.log_queue.put("Planilha 'chamados.xlsx' não encontrada. Criando novo arquivo em branco...")
                df = pd.DataFrame()
                df.to_excel(filepath, index=False)
                self.log_queue.put("-> Planilha em branco 'chamados.xlsx' criada com sucesso!")
            except Exception as e:
                self.log_queue.put(f"ERRO CRÍTICO: Não foi possível criar a planilha: {e}")

    def process_log_queue(self):
        """Função que consome a fila de log e atualiza o widget de texto."""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, message + '\n')
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_log_queue)

    def abrir_planilha(self):
        self._verificar_e_criar_planilha()
        filepath = get_asset_path('chamados.xlsx')
        try:
            if not os.path.exists(filepath):
                self.log_queue.put(f"ERRO: Arquivo '{os.path.basename(filepath)}' não encontrado!")
                return
            os.startfile(filepath)
            self.log_queue.put(f"Abrindo '{os.path.basename(filepath)}'...")
        except Exception as e:
            self.log_queue.put(f"Falha ao abrir o arquivo: {e}")

    # ATUALIZADO: Função para carregar configurações de um arquivo JSON
    def _load_config(self):
        config_path = get_asset_path(CONFIG_FILE)
        config = {'directory': '', 'delete_previous': False}
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    loaded_config = json.load(f)
                config.update(loaded_config)
            except (json.JSONDecodeError, TypeError):
                self.log_queue.put(f"AVISO: Arquivo de configuração '{CONFIG_FILE}' corrompido ou inválido. Usando padrão.")
        return config

    # ATUALIZADO: Função para salvar configurações em um arquivo JSON
    def _save_config(self):
        config_path = get_asset_path(CONFIG_FILE)
        config_data = {
            'directory': self.diretorio_selecionado,
            'delete_previous': self.delete_previous_var.get()
        }
        try:
            with open(config_path, 'w') as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            self.log_queue.put(f"ERRO ao salvar configuração: {e}")

    def _update_path_label(self):
        if self.diretorio_selecionado:
            self.path_label_var.set(f"Pasta de Destino: {self.diretorio_selecionado}")
        else:
            self.path_label_var.set("Nenhuma pasta de destino selecionada.")

    def selecionar_pasta(self):
        diretorio = filedialog.askdirectory(title="Selecione a pasta para salvar os PDFs")
        if diretorio:
            self.diretorio_selecionado = diretorio
            self._update_path_label()
            self.start_button.config(state=tk.NORMAL)
            self._save_config()

    def iniciar_automacao(self):
        if not self.diretorio_selecionado:
            self.log_queue.put("ERRO: Por favor, selecione uma pasta de destino primeiro.")
            return

        self.start_time = time.monotonic()
        self.total_downloads = 0

        self.select_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        self.open_sheet_button.config(state=tk.DISABLED)
        self.continue_button.config(state=tk.DISABLED)
        self.delete_check.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)

        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')
        
        self.cancel_event.clear()
        
        self.process_thread = threading.Thread(target=validar_planilha, args=(self,), daemon=True)
        self.process_thread.start()
        
    def continuar_automacao(self):
        self.continue_button.config(state=tk.DISABLED)
        self.delete_check.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.log_queue.put("\nRecarregando planilha e continuando com os downloads...")
        
        self.cancel_event.clear()
        self.process_thread = threading.Thread(target=iniciar_processo_requests, args=(self,), daemon=True)
        self.process_thread.start()

    def cancelar_execucao(self):
        if self.process_thread and self.process_thread.is_alive():
            self.log_queue.put("------------------------------------")
            self.log_queue.put("CANCELAMENTO SOLICITADO... Aguardando tarefas atuais.")
            self.cancel_event.set()
            self.cancel_button.config(state=tk.DISABLED)
            self.continue_button.config(state=tk.DISABLED)

    def reabilitar_botoes(self):
        """Reabilita os botões da interface e exibe o sumário final."""
        if self.start_time:
            end_time = time.monotonic()
            elapsed_seconds = end_time - self.start_time
            minutes, seconds = divmod(elapsed_seconds, 60)
            
            summary_lines = [
                "------------------------------------",
                f"Total de RATs baixadas com sucesso: {self.total_downloads}",
                f"Tempo de execução: {int(minutes)} minutos e {int(seconds)} segundos.",
            ]
            for line in summary_lines:
                self.log_queue.put(line)

            self.start_time = None

        self.select_button.config(state=tk.NORMAL)
        self.start_button.config(state=tk.NORMAL if self.diretorio_selecionado else tk.DISABLED)
        self.open_sheet_button.config(state=tk.NORMAL)
        self.continue_button.config(state=tk.DISABLED)
        self.delete_check.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.process_thread = None

if __name__ == "__main__":
    try:
        import pandas, openpyxl, pypdf, requests, ttkthemes
    except ImportError:
        print("INFO: Módulos necessários não encontrados. Instalando...")
        os.system(f'"{sys.executable}" -m pip install pandas openpyxl pypdf requests ttkthemes')
    
    root = ThemedTk(theme="equilux")
    app = App(root)
    root.mainloop()
