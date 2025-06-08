#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
App Massa - Processador de NF-e em Massa
Interface gráfica para processamento em lote de XMLs
Desenvolvido por Thucosta
"""

import os
import sys
import glob
import time
import threading
import queue
from datetime import datetime
from pathlib import Path
from typing import List
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

# Importar o processador individual
from processar_nfe import ProcessadorNFe

# Configuração do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ProcessadorMassa:
    """Classe responsável pelo processamento em massa de arquivos XML de NF-e"""
    
    def __init__(self):
        self.pasta_xmls = None
        self.pasta_saida = None
        self.template_path = None
        
        # Cache para otimização
        self.template_cache = None
        self.processador_cache = None
        
        # Estatísticas
        self.total_arquivos = 0
        self.processados = 0
        self.sucessos = 0
        self.erros = 0
        self.inicio_processamento = None
        self.processando = False
        self.parar_solicitado = False
    
    def descobrir_xmls(self, pasta_xmls: str) -> List[str]:
        """Descobrir todos os XMLs na pasta"""
        if not os.path.exists(pasta_xmls):
            return []
        
        # Padrões de busca para XMLs de NF-e
        padroes = [
            os.path.join(pasta_xmls, "*.xml"),
            os.path.join(pasta_xmls, "**", "*.xml")
        ]
        
        xmls_encontrados = set()
        for padrao in padroes:
            xmls_encontrados.update(glob.glob(padrao, recursive=True))
        
        # Filtrar apenas arquivos válidos
        xmls_validos = [xml for xml in xmls_encontrados if os.path.isfile(xml)]
        return sorted(xmls_validos)

class AppMassa(ctk.CTk):
    """Interface gráfica para processamento em massa de NF-e"""
    
    def __init__(self):
        super().__init__()
        
        self.title("Processador de NF-e em Massa")
        self.geometry("900x700")
        
        # Variáveis de interface
        self.pasta_xmls_var = tk.StringVar()
        self.pasta_saida_var = tk.StringVar(value="./massa_output")
        self.template_var = tk.StringVar(value="nfe_vertical.html")
        
        # Processador
        self.processador = ProcessadorMassa()
        
        # Queue para comunicação entre threads
        self.message_queue = queue.Queue()
        self.processando = False
        
        # Configurar interface
        self.setup_ui()
        
        # Iniciar verificação de mensagens
        self.after(100, self.check_message_queue)
    
    def setup_ui(self):
        """Configurar elementos da interface"""
        
        # Título principal
        title_label = ctk.CTkLabel(
            self,
            text="🔥 Processador de NF-e em Massa",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=20)
        
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        # === SEÇÃO DE CONFIGURAÇÕES ===
        config_label = ctk.CTkLabel(
            main_frame,
            text="📁 Configurações",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        config_label.pack(pady=(20, 15))
        
        # Pasta XMLs
        xmls_frame = ctk.CTkFrame(main_frame)
        xmls_frame.pack(padx=20, pady=5, fill="x")
        
        ctk.CTkLabel(
            xmls_frame,
            text="Pasta com XMLs:",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w", padx=15, pady=(10, 5))
        
        xmls_input_frame = ctk.CTkFrame(xmls_frame)
        xmls_input_frame.pack(padx=15, pady=(0, 15), fill="x")
        
        self.xmls_entry = ctk.CTkEntry(
            xmls_input_frame,
            textvariable=self.pasta_xmls_var,
            placeholder_text="Selecione a pasta contendo os XMLs de NF-e...",
            height=35
        )
        self.xmls_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        xmls_btn = ctk.CTkButton(
            xmls_input_frame,
            text="Procurar",
            command=self.select_xmls_folder,
            width=100
        )
        xmls_btn.pack(side="right", padx=5)
        
        # Pasta Saída
        output_frame = ctk.CTkFrame(main_frame)
        output_frame.pack(padx=20, pady=5, fill="x")
        
        ctk.CTkLabel(
            output_frame,
            text="Pasta de Saída:",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w", padx=15, pady=(10, 5))
        
        output_input_frame = ctk.CTkFrame(output_frame)
        output_input_frame.pack(padx=15, pady=(0, 15), fill="x")
        
        self.output_entry = ctk.CTkEntry(
            output_input_frame,
            textvariable=self.pasta_saida_var,
            height=35
        )
        self.output_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        output_btn = ctk.CTkButton(
            output_input_frame,
            text="Procurar",
            command=self.select_output_folder,
            width=100
        )
        output_btn.pack(side="right", padx=5)
        
        # Template
        template_frame = ctk.CTkFrame(main_frame)
        template_frame.pack(padx=20, pady=5, fill="x")
        
        ctk.CTkLabel(
            template_frame,
            text="Template HTML:",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w", padx=15, pady=(10, 5))
        
        template_input_frame = ctk.CTkFrame(template_frame)
        template_input_frame.pack(padx=15, pady=(0, 15), fill="x")
        
        self.template_entry = ctk.CTkEntry(
            template_input_frame,
            textvariable=self.template_var,
            height=35
        )
        self.template_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        template_btn = ctk.CTkButton(
            template_input_frame,
            text="Procurar",
            command=self.select_template_file,
            width=100
        )
        template_btn.pack(side="right", padx=5)
        

        
        # === CONTROLES ===
        controls_frame = ctk.CTkFrame(main_frame)
        controls_frame.pack(padx=20, pady=20, fill="x")
        
        self.start_btn = ctk.CTkButton(
            controls_frame,
            text="🚀 Iniciar Processamento",
            command=self.iniciar_processamento,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="green",
            hover_color="darkgreen"
        )
        self.start_btn.pack(side="left", padx=10, expand=True, fill="x")
        
        self.stop_btn = ctk.CTkButton(
            controls_frame,
            text="⏹️ Parar",
            command=self.parar_processamento,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="red",
            hover_color="darkred",
            state="disabled"
        )
        self.stop_btn.pack(side="right", padx=10)
        
        # === PROGRESSO ===
        progress_label = ctk.CTkLabel(
            main_frame,
            text="📊 Progresso",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        progress_label.pack(pady=(20, 10))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(main_frame, width=400, height=20)
        self.progress_bar.pack(padx=20, pady=10)
        self.progress_bar.set(0)
        
        # Estatísticas
        self.stats_label = ctk.CTkLabel(
            main_frame,
            text="Aguardando início do processamento...",
            font=ctk.CTkFont(size=14)
        )
        self.stats_label.pack(pady=10)
        
        # Log de mensagens
        log_label = ctk.CTkLabel(
            main_frame,
            text="📝 Log de Mensagens",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        log_label.pack(pady=(20, 5))
        
        self.message_text = ctk.CTkTextbox(main_frame, height=150)
        self.message_text.pack(padx=20, pady=(0, 20), fill="both", expand=True)
        
        # Mensagem inicial
        self.add_message("🔥 Processador de NF-e em Massa inicializado!")
        self.add_message("📁 Selecione a pasta com XMLs para começar")
    

    
    def select_xmls_folder(self):
        """Selecionar pasta com XMLs"""
        folder = filedialog.askdirectory(title="Selecionar pasta com XMLs de NF-e")
        if folder:
            self.pasta_xmls_var.set(folder)
            self.verificar_xmls()
    
    def select_output_folder(self):
        """Selecionar pasta de saída"""
        folder = filedialog.askdirectory(title="Selecionar pasta de saída")
        if folder:
            self.pasta_saida_var.set(folder)
    
    def select_template_file(self):
        """Selecionar arquivo de template"""
        file_path = filedialog.askopenfilename(
            title="Selecionar template HTML",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        if file_path:
            self.template_var.set(file_path)
    
    def verificar_xmls(self):
        """Verificar quantos XMLs existem na pasta selecionada"""
        if not self.pasta_xmls_var.get():
            return
        
        try:
            xmls = self.processador.descobrir_xmls(self.pasta_xmls_var.get())
            count = len(xmls)
            self.add_message(f"📊 Encontrados {count:,} XMLs na pasta selecionada")
            
        except Exception as e:
            self.add_message(f"❌ Erro ao verificar XMLs: {str(e)}", is_error=True)
    
    def add_message(self, message: str, is_error: bool = False):
        """Adicionar mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.message_text.insert("end", formatted_message)
        self.message_text.see("end")
        
        if is_error:
            # Destacar erros em vermelho (se possível)
            pass
    
    def iniciar_processamento(self):
        """Iniciar processamento em massa"""
        # Validações
        if not self.pasta_xmls_var.get():
            messagebox.showerror("Erro", "Selecione a pasta com XMLs!")
            return
        
        if not os.path.exists(self.pasta_xmls_var.get()):
            messagebox.showerror("Erro", "Pasta de XMLs não existe!")
            return
        
        if not os.path.exists(self.template_var.get()):
            messagebox.showerror("Erro", "Template HTML não encontrado!")
            return
        
        # Descobrir XMLs
        xmls = self.processador.descobrir_xmls(self.pasta_xmls_var.get())
        if not xmls:
            messagebox.showerror("Erro", "Nenhum XML encontrado na pasta!")
            return
        
        # Confirmar processamento
        resposta = messagebox.askyesno(
            "Confirmar Processamento",
            f"Processar {len(xmls):,} XMLs sequencialmente?\n\n"
            f"Isso pode levar alguns minutos dependendo da quantidade.\n"
            f"O processamento será feito de forma sequencial para estabilidade."
        )
        
        if not resposta:
            return
        
        # Iniciar processamento
        self.processando = True
        self.processador.parar_solicitado = False
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        
        # Executar em thread separada
        thread = threading.Thread(target=self.executar_processamento, args=(xmls,), daemon=True)
        thread.start()
    
    def executar_processamento(self, xmls: List[str]):
        """Executar processamento em massa (roda em thread separada)"""
        try:
            # Configurar processador
            self.processador.pasta_xmls = self.pasta_xmls_var.get()
            self.processador.pasta_saida = self.pasta_saida_var.get()
            self.processador.template_path = self.template_var.get()
            self.processador.total_arquivos = len(xmls)
            self.processador.processados = 0
            self.processador.sucessos = 0
            self.processador.erros = 0
            self.processador.inicio_processamento = time.time()
            
            # Criar pasta de saída
            os.makedirs(self.processador.pasta_saida, exist_ok=True)
            
            # ⚡ OTIMIZAÇÃO: Carregar template uma única vez
            if self.processador.template_cache is None:
                try:
                    with open(self.processador.template_path, 'r', encoding='utf-8') as f:
                        self.processador.template_cache = f.read()
                    self.message_queue.put(("message", "⚡ Template carregado em cache"))
                except Exception as e:
                    self.message_queue.put(("message", f"❌ Erro ao carregar template: {e}"))
                    return
            
            self.message_queue.put(("message", f"🚀 Iniciando processamento de {len(xmls):,} XMLs"))
            self.message_queue.put(("message", f"📁 Pasta de saída: {self.processador.pasta_saida}"))
            self.message_queue.put(("message", "⚡ Modo otimizado: template em cache + processador reutilizado"))
            
            # Processar cada XML sequencialmente
            for i, xml_path in enumerate(xmls):
                if self.processador.parar_solicitado:
                    self.message_queue.put(("message", "⚠️ Processamento interrompido pelo usuário"))
                    break
                
                try:
                    # Nome do arquivo PDF baseado no XML
                    nome_base = Path(xml_path).stem
                    pdf_filename = f"{nome_base}.pdf"
                    pdf_path = os.path.join(self.processador.pasta_saida, pdf_filename)
                    
                    # Reutilizar processador (cache)
                    if self.processador.processador_cache is None:
                        self.processador.processador_cache = ProcessadorNFe()
                    
                    # Processar NF-e gerando apenas PDF (com cache)
                    resposta = self.processador.processador_cache.processar_nfe_optimized(
                        xml_path=xml_path,
                        template_content=self.processador.template_cache,
                        output_dir=self.processador.pasta_saida,
                        pdf_filename=pdf_filename
                    )
                    
                    self.processador.processados += 1
                    
                    if resposta.get('success', False):
                        self.processador.sucessos += 1
                    else:
                        self.processador.erros += 1
                        erro_msg = resposta.get('error', 'Erro desconhecido')
                        self.message_queue.put(("message", f"❌ {os.path.basename(xml_path)}: {erro_msg}"))
                    
                    # Calcular progresso
                    progresso = self.processador.processados / self.processador.total_arquivos
                    tempo_decorrido = time.time() - self.processador.inicio_processamento
                    velocidade = self.processador.processados / tempo_decorrido if tempo_decorrido > 0 else 0
                    tempo_restante = (self.processador.total_arquivos - self.processador.processados) / velocidade if velocidade > 0 else 0
                    
                    # Enviar atualização de progresso
                    self.message_queue.put(("progress", {
                        'valor': progresso,
                        'processados': self.processador.processados,
                        'total': self.processador.total_arquivos,
                        'sucessos': self.processador.sucessos,
                        'erros': self.processador.erros,
                        'velocidade': velocidade,
                        'tempo_restante': tempo_restante
                    }))
                    
                    # Log otimizado: menos frequente para massa
                    if self.processador.processados % max(50, self.processador.total_arquivos // 20) == 0:
                        self.message_queue.put(("message", f"📊 Processados: {self.processador.processados}/{self.processador.total_arquivos} ({progresso*100:.1f}%)"))
                    
                except Exception as e:
                    self.processador.erros += 1
                    self.message_queue.put(("message", f"❌ Erro em {os.path.basename(xml_path)}: {str(e)}"))
            
            # Estatísticas finais
            tempo_total = time.time() - self.processador.inicio_processamento
            velocidade_media = self.processador.total_arquivos / tempo_total if tempo_total > 0 else 0
            
            self.message_queue.put(("message", "🎉 PROCESSAMENTO CONCLUÍDO!"))
            self.message_queue.put(("message", f"📊 Total: {self.processador.total_arquivos:,} XMLs"))
            self.message_queue.put(("message", f"✅ Sucessos: {self.processador.sucessos:,} ({self.processador.sucessos/self.processador.total_arquivos*100:.1f}%)"))
            self.message_queue.put(("message", f"❌ Erros: {self.processador.erros:,} ({self.processador.erros/self.processador.total_arquivos*100:.1f}%)"))
            self.message_queue.put(("message", f"⏱️ Tempo total: {tempo_total/60:.1f} minutos"))
            self.message_queue.put(("message", f"⚡ Velocidade média: {velocidade_media:.1f} XMLs/segundo"))
            self.message_queue.put(("message", f"📁 Arquivos salvos em: {self.processador.pasta_saida}"))
            
            self.message_queue.put(("finish", None))
            
        except Exception as e:
            self.message_queue.put(("message", f"❌ Erro geral no processamento: {str(e)}"))
            self.message_queue.put(("finish", None))
    
    def parar_processamento(self):
        """Parar processamento em massa"""
        self.processador.parar_solicitado = True
        self.processando = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        
        self.add_message("⚠️ Solicitação de parada enviada...")
    
    def check_message_queue(self):
        """Verificar mensagens na queue (executa na thread principal)"""
        try:
            while True:
                msg_type, msg_data = self.message_queue.get_nowait()
                
                if msg_type == "message":
                    if isinstance(msg_data, tuple):
                        message, is_error = msg_data
                        self.add_message(message, is_error)
                    else:
                        self.add_message(msg_data)
                
                elif msg_type == "progress":
                    # Atualizar barra de progresso
                    self.progress_bar.set(msg_data['valor'])
                    
                    # Atualizar estatísticas
                    stats_text = (
                        f"🚀 {msg_data['processados']:,}/{msg_data['total']:,} "
                        f"({msg_data['valor']*100:.1f}%) | "
                        f"✅ {msg_data['sucessos']:,} | "
                        f"❌ {msg_data['erros']:,} | "
                        f"⚡ {msg_data['velocidade']:.1f}/s | "
                        f"⏱️ {msg_data['tempo_restante']/60:.1f}min restante"
                    )
                    self.stats_label.configure(text=stats_text)
                
                elif msg_type == "finish":
                    # Finalizar processamento
                    self.processando = False
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.progress_bar.set(1.0)
                    
                    # Mostrar resultado final
                    messagebox.showinfo(
                        "Processamento Concluído!",
                        f"Processamento finalizado com sucesso!\n\n"
                        f"✅ Sucessos: {self.processador.sucessos:,}\n"
                        f"❌ Erros: {self.processador.erros:,}\n"
                        f"📁 Arquivos salvos em:\n{self.processador.pasta_saida}"
                    )
                    
        except queue.Empty:
            pass
        
        # Agendar próxima verificação
        self.after(100, self.check_message_queue)

def main():
    """Função principal"""
    try:
        # Executar aplicativo
        app = AppMassa()
        app.mainloop()
        
    except Exception as e:
        print(f"Erro ao inicializar aplicativo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 
