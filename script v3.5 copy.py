# -*- coding: utf-8 -*-
"""
Automatizador de Tarefas com Controle de Mouse e Teclado v3.11 (Versão Final Estável)

Este script usa OpenCV para reconhecimento de imagem, PyDirectInput para controle do mouse/teclado
e Pillow para capturas de tela.

MODIFICAÇÃO v3.11 (Correção de Bug Crítico):
1. Corrigido um erro de sintaxe na função `update_row_count` e outras funções auxiliares
   que foi introduzido na versão anterior por compactação indevida do código.
2. O código foi reformatado para garantir legibilidade e estabilidade.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import time
import threading
import os
# Bibliotecas para automação real
import pydirectinput
import cv2
import numpy as np
from PIL import ImageGrab

class DesktopAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatizador Desktop v3.11 (Estável)")
        self.root.geometry("700x720")
        pydirectinput.PAUSE = 0.1

        # Variáveis para os caminhos dos arquivos
        self.excel_path = tk.StringVar()
        self.imagem_rotulo_contrato = tk.StringVar()
        self.imagem_botao_pesquisar = tk.StringVar()
        self.imagem_botao_liquidacao_nativo = tk.StringVar()
        self.imagem_botao_liquidacao_hover = tk.StringVar()
        self.imagem_botao_final = tk.StringVar()
        self.automation_thread = None
        self.stop_event = threading.Event()
        self.total_rows = 0
        self.create_widgets()

    def create_widgets(self):
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        excel_config_frame = tk.LabelFrame(main_frame, text="1. Configuração do Excel", padx=10, pady=10)
        excel_config_frame.pack(fill=tk.X, pady=5)
        tk.Label(excel_config_frame, text="Arquivo Excel:").grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(excel_config_frame, textvariable=self.excel_path, width=60, state='readonly').grid(row=0, column=1, sticky="ew", padx=5)
        tk.Button(excel_config_frame, text="Procurar...", command=self.browse_excel_file).grid(row=0, column=2)
        tk.Label(excel_config_frame, text="Nome da Coluna de Contratos:").grid(row=1, column=0, sticky="w", pady=5)
        self.coluna_contrato_entry = tk.Entry(excel_config_frame, width=60)
        self.coluna_contrato_entry.insert(0, "Numero do Contrato")
        self.coluna_contrato_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5)
        self.coluna_contrato_entry.bind("<KeyRelease>", self.update_row_count)
        self.total_rows_label = tk.Label(excel_config_frame, text="Total de Contratos a processar: 0", fg="navy", font=('Helvetica', 9, 'bold'))
        self.total_rows_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=5)
        excel_config_frame.columnconfigure(1, weight=1)

        image_config_frame = tk.LabelFrame(main_frame, text="2. Configuração das Imagens (.png)", padx=10, pady=10)
        image_config_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(image_config_frame, text="Imagem do RÓTULO do Campo:").grid(row=0, column=0, sticky="w", pady=2)
        tk.Entry(image_config_frame, textvariable=self.imagem_rotulo_contrato, state='readonly').grid(row=0, column=1, sticky="ew")
        tk.Button(image_config_frame, text="Procurar...", command=lambda: self.browse_image_file(self.imagem_rotulo_contrato, "Rótulo do Campo (Âncora)")).grid(row=0, column=2, padx=5)

        offset_frame = tk.Frame(image_config_frame)
        offset_frame.grid(row=1, column=0, columnspan=3, sticky='w', pady=2)
        tk.Label(offset_frame, text="Offset do Clique (em pixels):   X:").pack(side=tk.LEFT)
        self.offset_x_entry = tk.Entry(offset_frame, width=5)
        self.offset_x_entry.insert(0, "200")
        self.offset_x_entry.pack(side=tk.LEFT)
        tk.Label(offset_frame, text="Y:").pack(side=tk.LEFT, padx=(10,0))
        self.offset_y_entry = tk.Entry(offset_frame, width=5)
        self.offset_y_entry.insert(0, "0")
        self.offset_y_entry.pack(side=tk.LEFT)
        
        tk.Label(image_config_frame, text="Imagem do Botão 'Pesquisar':").grid(row=2, column=0, sticky="w", pady=2)
        tk.Entry(image_config_frame, textvariable=self.imagem_botao_pesquisar, state='readonly').grid(row=2, column=1, sticky="ew")
        tk.Button(image_config_frame, text="Procurar...", command=lambda: self.browse_image_file(self.imagem_botao_pesquisar, "Botão Pesquisar")).grid(row=2, column=2, padx=5)

        tk.Label(image_config_frame, text="Botão Crédito Recebido (Nativo):").grid(row=3, column=0, sticky="w", pady=2)
        tk.Entry(image_config_frame, textvariable=self.imagem_botao_liquidacao_nativo, state='readonly').grid(row=3, column=1, sticky="ew")
        tk.Button(image_config_frame, text="Procurar...", command=lambda: self.browse_image_file(self.imagem_botao_liquidacao_nativo, "Crédito Recebido (Nativo)")).grid(row=3, column=2, padx=5)
        
        tk.Label(image_config_frame, text="Botão Crédito Recebido (Hover):").grid(row=4, column=0, sticky="w", pady=2)
        tk.Entry(image_config_frame, textvariable=self.imagem_botao_liquidacao_hover, state='readonly').grid(row=4, column=1, sticky="ew")
        tk.Button(image_config_frame, text="Procurar...", command=lambda: self.browse_image_file(self.imagem_botao_liquidacao_hover, "Crédito Recebido (Hover)")).grid(row=4, column=2, padx=5)
        
        tk.Label(image_config_frame, text="Imagem do Botão Final (Confirmar):").grid(row=5, column=0, sticky="w", pady=2)
        tk.Entry(image_config_frame, textvariable=self.imagem_botao_final, state='readonly').grid(row=5, column=1, sticky="ew")
        tk.Button(image_config_frame, text="Procurar...", command=lambda: self.browse_image_file(self.imagem_botao_final, "Botão Final")).grid(row=5, column=2, padx=5)
        image_config_frame.columnconfigure(1, weight=1)

        control_frame = tk.LabelFrame(main_frame, text="3. Controle", padx=10, pady=10)
        control_frame.pack(fill=tk.X, pady=10)
        self.btn_start = tk.Button(control_frame, text="Iniciar Automação", command=self.start_automation, state=tk.DISABLED, bg="#cccccc", fg="white", font=('Helvetica', 10, 'bold'))
        self.btn_start.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=4)
        self.btn_stop = tk.Button(control_frame, text="Parar Automação", command=self.stop_automation, state=tk.DISABLED, bg="#f44336", fg="white", font=('Helvetica', 10, 'bold'))
        self.btn_stop.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=4)
        self.progress_label = tk.Label(control_frame, text="Progresso: N/A", font=('Helvetica', 10))
        self.progress_label.pack(side=tk.RIGHT, padx=10)

        log_frame = tk.LabelFrame(main_frame, text="Status e Log", padx=10, pady=10)
        log_frame.pack(expand=True, fill=tk.BOTH)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', height=10)
        self.log_text.pack(expand=True, fill=tk.BOTH)

    def log(self, message):
        self.root.after(0, self._log_thread_safe, message)

    def _log_thread_safe(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def update_progress_label(self, message):
        self.root.after(0, self._update_progress_label_thread_safe, message)

    def _update_progress_label_thread_safe(self, message):
        self.progress_label.config(text=message)

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if file_path:
            self.excel_path.set(file_path)
            self.log(f"Arquivo selecionado: {file_path}")
            self.update_row_count()
            self.check_start_button_state()

    def browse_image_file(self, string_var, image_name):
        file_path = filedialog.askopenfilename(title=f"Selecione a imagem para '{image_name}'", filetypes=[("Arquivos de Imagem", "*.png")])
        if file_path:
            string_var.set(file_path)
            self.log(f"Imagem '{image_name}' selecionada: {os.path.basename(file_path)}")
            self.check_start_button_state()

    def check_start_button_state(self):
        if (self.excel_path.get() and
                self.imagem_rotulo_contrato.get() and
                self.imagem_botao_pesquisar.get() and
                self.imagem_botao_liquidacao_nativo.get() and
                self.imagem_botao_liquidacao_hover.get() and
                self.imagem_botao_final.get()):
            self.btn_start.config(state=tk.NORMAL, bg="#4CAF50")
        else:
            self.btn_start.config(state=tk.DISABLED, bg="#cccccc")

    def update_row_count(self, event=None):
        excel_file = self.excel_path.get()
        column_name = self.coluna_contrato_entry.get()
        
        if not excel_file or not column_name:
            self.total_rows = 0
        else:
            try:
                df = pd.read_excel(excel_file)
                if column_name in df.columns:
                    self.total_rows = df[column_name].dropna().count()
                    self.log(f"Planilha carregada. Total de {self.total_rows} contratos a serem processados.")
                else:
                    self.log(f"AVISO: A coluna '{column_name}' não foi encontrada na planilha.")
                    self.total_rows = 0
            except Exception as e:
                self.log(f"Erro ao ler o arquivo Excel ou coluna: {e}")
                self.total_rows = 0
        
        self.total_rows_label.config(text=f"Total de Contratos a processar: {self.total_rows}")
        self.check_start_button_state()

    def start_automation(self):
        if self.total_rows == 0:
            messagebox.showwarning("Aviso", "Não há contratos para processar.")
            return
        self.btn_start.config(state=tk.DISABLED, bg="#cccccc")
        self.btn_stop.config(state=tk.NORMAL)
        self.stop_event.clear()
        self.automation_thread = threading.Thread(target=self.run_automation_logic, daemon=True)
        self.automation_thread.start()

    def stop_automation(self):
        self.log("Sinal de parada enviado...")
        self.stop_event.set()
        self.btn_stop.config(state=tk.DISABLED)

    def on_automation_finish(self, final_message):
        self.log(final_message)
        self.update_progress_label("Progresso: Finalizado")
        self.check_start_button_state()
        self.btn_stop.config(state=tk.DISABLED)

    def run_automation_logic(self):
        try:
            caminho_excel = self.excel_path.get()
            coluna_contrato = self.coluna_contrato_entry.get()
            for i in range(5, 0, -1):
                self.log(f"A automação real começará em {i} segundos...")
                time.sleep(1)
            
            df = pd.read_excel(caminho_excel).dropna(subset=[coluna_contrato])
            processed_count = 0
            for index, row in df.iterrows():
                if self.stop_event.is_set():
                    self.log("Processo interrompido pelo usuário.")
                    break
                contrato = row[coluna_contrato]
                contrato_str = str(int(contrato)) if isinstance(contrato, (int, float)) else str(contrato).strip()
                processed_count += 1
                if contrato_str:
                    self.update_progress_label(f"Processando {processed_count} de {self.total_rows}...")
                    success = self.processar_contrato(contrato_str)
                    if success:
                        self.log(f"SUCESSO: Contrato {contrato_str} (Linha {index + 2}) processado.")
                    else:
                        self.log(f"FALHA: Contrato {contrato_str} (Linha {index + 2}) não processado.\n")
            
            final_message = "PROCESSO FINALIZADO." if not self.stop_event.is_set() else "PROCESSO PARADO."
            self.root.after(0, self.on_automation_finish, final_message)
        except Exception as e:
            error_message = f"ERRO CRÍTICO: {e}"
            self.root.after(0, self.on_automation_finish, error_message)

    def paste_text(self, text):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            pydirectinput.keyDown('ctrl')
            pydirectinput.press('v')
            pydirectinput.keyUp('ctrl')
        except Exception as e:
            self.log(f"   ERRO ao tentar colar o texto: {e}")

    def encontrar_e_clicar_offset(self, image_path, description, offset_x, offset_y, timeout=10, confidence=0.8):
        self.log(f" - Procurando âncora '{description}' para clicar com offset...")
        start_time = time.time()
        
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if template is None:
            raise FileNotFoundError(f"Arquivo de imagem da âncora não encontrado em '{image_path}'")

        while time.time() - start_time < timeout:
            screenshot = ImageGrab.grab()
            main_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            result = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= confidence:
                h, w = template.shape[:2]
                center_x = max_loc[0] + w // 2
                center_y = max_loc[1] + h // 2
                
                click_x = center_x + offset_x
                click_y = center_y + offset_y
                
                self.log(f"   - Âncora '{description}' encontrada. Clicando em ({click_x}, {click_y})...")
                pydirectinput.moveTo(click_x, click_y, duration=0.25)
                pydirectinput.click()
                return True
            time.sleep(0.5)
            
        self.log(f"   - ERRO: Não foi possível encontrar a âncora '{description}' na tela após {timeout} segundos.")
        return False

    def encontrar_e_clicar_duas_imagens(self, image_path_nativo, image_path_hover, description, timeout=10, confidence=0.8):
        self.log(f" - Procurando '{description}' (estados nativo e hover)...")
        start_time = time.time()
        template_nativo = cv2.imread(image_path_nativo, cv2.IMREAD_GRAYSCALE)
        template_hover = cv2.imread(image_path_hover, cv2.IMREAD_GRAYSCALE)
        if template_nativo is None or template_hover is None:
            raise FileNotFoundError(f"Uma ou ambas as imagens para '{description}' não foram encontradas.")
        
        h_nativo, w_nativo = template_nativo.shape
        h_hover, w_hover = template_hover.shape
        while time.time() - start_time < timeout:
            screenshot = ImageGrab.grab()
            main_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
            result_nativo = cv2.matchTemplate(main_image, template_nativo, cv2.TM_CCOEFF_NORMED)
            _, max_val_nativo, _, max_loc_nativo = cv2.minMaxLoc(result_nativo)
            if max_val_nativo >= confidence:
                self.log(f"   - '{description} (Nativo)' encontrado com {max_val_nativo*100:.2f}% de confiança. Clicando...")
                center_x = max_loc_nativo[0] + w_nativo // 2
                center_y = max_loc_nativo[1] + h_nativo // 2
                pydirectinput.moveTo(center_x, center_y, duration=0.25)
                pydirectinput.click()
                return True
            result_hover = cv2.matchTemplate(main_image, template_hover, cv2.TM_CCOEFF_NORMED)
            _, max_val_hover, _, max_loc_hover = cv2.minMaxLoc(result_hover)
            if max_val_hover >= confidence:
                self.log(f"   - '{description} (Hover)' encontrado com {max_val_hover*100:.2f}% de confiança. Clicando...")
                center_x = max_loc_hover[0] + w_hover // 2
                center_y = max_loc_hover[1] + h_hover // 2
                pydirectinput.moveTo(center_x, center_y, duration=0.25)
                pydirectinput.click()
                return True
            time.sleep(0.5)
        self.log(f"   - ERRO: Não foi possível encontrar a imagem '{description}' em nenhum dos estados após {timeout} segundos.")
        return False
    
    def encontrar_e_clicar_imagem(self, image_path, description, timeout=10, confidence=0.8, usar_escala_de_cinza=False):
        self.log(f" - Procurando '{description}'...")
        start_time = time.time()
        template_color = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if template_color is None:
            raise FileNotFoundError(f"Arquivo de imagem não encontrado em '{image_path}'")
        
        template = cv2.cvtColor(template_color, cv2.COLOR_BGR2GRAY) if usar_escala_de_cinza else template_color
        while time.time() - start_time < timeout:
            screenshot = ImageGrab.grab()
            main_image_color = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            main_image = cv2.cvtColor(main_image_color, cv2.COLOR_BGR2GRAY) if usar_escala_de_cinza else main_image_color
            result = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            if max_val >= confidence:
                h, w = template.shape[:2]
                center_x = max_loc[0] + w // 2
                center_y = max_loc[1] + h // 2
                self.log(f"   - '{description}' encontrado com {max_val*100:.2f}% de confiança. Clicando...")
                pydirectinput.moveTo(center_x, center_y, duration=0.25)
                pydirectinput.click()
                return True
            time.sleep(0.5)
        self.log(f"   - ERRO: Não foi possível encontrar a imagem '{description}' na tela após {timeout} segundos.")
        return False

    def processar_contrato(self, numero_contrato):
        try:
            offset_x = int(self.offset_x_entry.get())
            offset_y = int(self.offset_y_entry.get())
            
            # Etapa 1: Encontrar o rótulo e clicar no campo de contrato
            if not self.encontrar_e_clicar_offset(self.imagem_rotulo_contrato.get(), "Rótulo do Campo", offset_x, offset_y): return False
            time.sleep(0.5)

            # Etapa 2: Colar o número do contrato
            self.log(f"   - Colando contrato (Ctrl+V): {numero_contrato}")
            self.paste_text(numero_contrato)

            # Etapa 3: Clicar no botão de pesquisar
            if not self.encontrar_e_clicar_imagem(self.imagem_botao_pesquisar.get(), "Botão Pesquisar"): return False
            time.sleep(1) 

            # Etapa 4: Clicar no botão de liquidação/crédito
            if not self.encontrar_e_clicar_duas_imagens(
                self.imagem_botao_liquidacao_nativo.get(), 
                self.imagem_botao_liquidacao_hover.get(),
                "Crédito Recebido"
            ): return False
            
            # Etapa 5: Rolar a tela para baixo
            self.log("   - Rolando a tela para baixo...")
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')
            pydirectinput.press('down')

            time.sleep(0.5)

            # Etapa 6: Clicar no botão final
            if not self.encontrar_e_clicar_imagem(self.imagem_botao_final.get(), "Botão Final"): return False
            
            # Etapa 7: Confirmar a operação com Enter
            self.log("   - Confirmando a operação (Enter, aguarda 2s, Enter)...")
            pydirectinput.press('enter')
            time.sleep(2)
            pydirectinput.press('enter')
            time.sleep(1)
            
            # Etapa 8: Limpar o campo para o próximo contrato
            self.log("   - Preparando para o próximo ciclo...")
            if not self.encontrar_e_clicar_offset(self.imagem_rotulo_contrato.get(), "Rótulo do Campo", offset_x, offset_y):
                self.log("   - ERRO: Não foi possível retornar ao campo de contrato para limpá-lo.")
                return False

            time.sleep(0.3) 
            
            self.log("   - Limpando o campo (Ctrl+A -> Backspace)...")
            pydirectinput.keyDown('ctrl')
            pydirectinput.press('a')
            pydirectinput.keyUp('ctrl')
            time.sleep(0.1)
            pydirectinput.press('backspace')
            return True

        except Exception as e:
            self.log(f"   ERRO inesperado durante o processamento do contrato {numero_contrato}: {e}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = DesktopAutomationApp(root)
    root.mainloop()