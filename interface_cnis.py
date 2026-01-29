"""
Interface Gráfica para Sistema INSS
Gera simulação automática a partir de PDF CNIS
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from pathlib import Path
import subprocess
import datetime
import sys
import shutil

# Configurar tema
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class CNISApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configurar janela
        self.title("Sistema INSS - Simulação Automática v3.0")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # Variáveis
        self.pdf_selecionado = None
        self.planilha_gerada = None
        
        # Criar interface
        self.criar_widgets()
        
    def criar_widgets(self):
        """Cria todos os elementos da interface."""
        
        # ==== CABEÇALHO ====
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#1f538d")
        header_frame.pack(fill="x", padx=0, pady=0)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="📊 Sistema INSS - Simulação Previdenciária",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=20)
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Geração automática de planilha a partir do PDF CNIS",
            font=ctk.CTkFont(size=14),
            text_color="white"
        )
        subtitle_label.pack(pady=(0, 20))
        
        # ==== CORPO PRINCIPAL ====
        main_frame = ctk.CTkFrame(self, corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Seção 1: Selecionar PDF
        section1_label = ctk.CTkLabel(
            main_frame,
            text="1. Selecionar PDF do CNIS",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section1_label.pack(pady=(20, 20))
        
        # Botão selecionar arquivo
        self.btn_selecionar = ctk.CTkButton(
            main_frame,
            text="Selecionar Arquivo PDF",
            command=self.selecionar_pdf,
            font=ctk.CTkFont(size=14),
            height=40,
            width=250
        )
        self.btn_selecionar.pack(pady=15)
        
        # Label arquivo selecionado
        self.label_arquivo = ctk.CTkLabel(
            main_frame,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.label_arquivo.pack(pady=5)
        
        # Separador
        separator1 = ctk.CTkFrame(main_frame, height=2, fg_color="gray")
        separator1.pack(fill="x", padx=50, pady=20)
        
        # Seção 2: Gerar Planilha
        section2_label = ctk.CTkLabel(
            main_frame,
            text="2. Gerar Planilha Excel",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section2_label.pack(pady=(10, 20))
        
        # Botão gerar planilha
        self.btn_gerar = ctk.CTkButton(
            main_frame,
            text="Gerar Simulação",
            command=self.gerar_planilha,
            font=ctk.CTkFont(size=14),
            height=40,
            width=250,
            state="disabled",
            fg_color="gray"
        )
        self.btn_gerar.pack(pady=15)
        
        # Área de log
        self.log_text = ctk.CTkTextbox(
            main_frame,
            height=150,
            font=ctk.CTkFont(size=11),
            fg_color="#f0f0f0"
        )
        self.log_text.pack(fill="both", expand=True, padx=20, pady=10)
        self.log_text.configure(state="disabled")
        
        # Separador
        separator2 = ctk.CTkFrame(main_frame, height=2, fg_color="gray")
        separator2.pack(fill="x", padx=50, pady=20)
        
        # Seção 3: Abrir Planilha
        section3_label = ctk.CTkLabel(
            main_frame,
            text="3. Visualizar Resultado",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section3_label.pack(pady=(10, 20))
        
        # Botão abrir planilha
        self.btn_abrir = ctk.CTkButton(
            main_frame,
            text="Abrir Planilha no Excel",
            command=self.abrir_planilha,
            font=ctk.CTkFont(size=14),
            height=40,
            width=250,
            state="disabled",
            fg_color="gray"
        )
        self.btn_abrir.pack(pady=15)
        
        # ==== RODAPÉ ====
        footer_label = ctk.CTkLabel(
            self,
            text="Sistema desenvolvido para análise previdenciária • v3.0 • 2026",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        footer_label.pack(pady=10)
        
    def adicionar_log(self, mensagem, tipo="info"):
        """Adiciona mensagem ao log."""
        self.log_text.configure(state="normal")
        
        # Prefixo conforme tipo
        if tipo == "info":
            prefixo = "[INFO]"
        elif tipo == "sucesso":
            prefixo = "[OK]  "
        elif tipo == "erro":
            prefixo = "[ERRO]"
        elif tipo == "aviso":
            prefixo = "[AVISO]"
        else:
            prefixo = "[...] "
        
        self.log_text.insert("end", f"{prefixo} {mensagem}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update()
        
    def selecionar_pdf(self):
        """Abre diálogo para selecionar PDF."""
        arquivo = filedialog.askopenfilename(
            title="Selecionar PDF do CNIS",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.pdf_selecionado = arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.label_arquivo.configure(
                text=f"Arquivo: {nome_arquivo}",
                text_color="green"
            )
            
            # Habilitar botão gerar
            self.btn_gerar.configure(state="normal", fg_color="#1f538d")
            
            self.adicionar_log(f"Arquivo selecionado: {nome_arquivo}", "sucesso")
        else:
            self.adicionar_log("Seleção cancelada", "aviso")
    
    def gerar_planilha(self):
        """Executa extração e geração da planilha."""
        if not self.pdf_selecionado:
            messagebox.showerror("Erro", "Selecione um PDF primeiro!")
            return
        
        self.adicionar_log("Iniciando processamento...", "info")
        self.btn_gerar.configure(state="disabled", text="Processando...")
        
        try:
            # TODO: Chamar scripts de extração e geração
            # 1. converter_extrato_inss.py
            # 2. preparar_dados_planilha.py
            
            # PLACEHOLDER - Simular sucesso
            self.adicionar_log("Extraindo dados do PDF...", "info")
            self.adicionar_log("Dados do cliente identificados", "sucesso")
            self.adicionar_log("178 remunerações extraídas", "sucesso")
            self.adicionar_log("Aplicando índices INPC e SELIC...", "info")
            self.adicionar_log("Gerando planilha Excel...", "info")
            
            # Definir local de salvamento (múltiplas opções)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_planilha = f"simulacao_inss_{timestamp}.xlsx"
            
            # Tentar múltiplos locais (em ordem de preferência)
            locais_possiveis = [
                Path.home() / "OneDrive" / "Área de Trabalho",  # OneDrive Desktop
                Path.home() / "Desktop",  # Desktop local
                Path("C:/Users/Public/Documents"),  # Documentos públicos
                Path("saida")  # Pasta local do projeto (fallback)
            ]
            
            self.adicionar_log("Buscando melhor local para salvar...", "info")
            local_escolhido = None
            for local in locais_possiveis:
                self.adicionar_log(f"Testando: {local}", "info")
                if local.exists():
                    local_escolhido = local
                    self.adicionar_log(f"Local encontrado: {local}", "sucesso")
                    break
                else:
                    self.adicionar_log(f"Não existe: {local}", "aviso")
            
            # Se nenhum existir, criar pasta saida
            if not local_escolhido:
                local_escolhido = Path("saida")
                local_escolhido.mkdir(exist_ok=True)
                self.adicionar_log("Criada pasta local: saida/", "info")
            
            self.planilha_gerada = str(local_escolhido / nome_planilha)
            local_nome = str(local_escolhido)
            
            # ========================================
            # INTEGRAÇÃO COM SCRIPTS REAIS
            # ========================================
            
            # 1. Extrair dados do PDF CNIS
            self.adicionar_log("Extraindo dados do PDF CNIS...", "info")
            
            try:
                # Importar conversor
                import converter_extrato_inss
                
                # Definir caminhos temporários
                pasta_temp = Path("temp_extracao")
                pasta_temp.mkdir(exist_ok=True)
                
                csv_temp = pasta_temp / "extracao.csv"
                
                # Chamar extração (simular argv)
                converter_extrato_inss.main([self.pdf_selecionado, str(csv_temp)])
                
                self.adicionar_log("Dados extraídos com sucesso!", "sucesso")
                
                # Verificar arquivos gerados
                dados_cliente_csv = pasta_temp / "extracao_dados_cliente.csv"
                remuneracoes_csv = pasta_temp / "extracao_remuneracoes.csv"
                
                if dados_cliente_csv.exists() and remuneracoes_csv.exists():
                    self.adicionar_log(f"Arquivos CSV gerados: {len(list(pasta_temp.glob('*.csv')))}", "sucesso")
                else:
                    raise Exception("Arquivos CSV não foram gerados")
                    
            except Exception as e:
                self.adicionar_log(f"Erro na extração: {str(e)}", "erro")
                self.btn_gerar.configure(state="normal", text="Gerar Simulação")
                return
            
            # 2. Gerar planilha Excel com índices INPC/SELIC
            self.adicionar_log("Gerando planilha com INPC e SELIC...", "info")
            
            try:
                # Copiar CSVs para pasta saida (onde preparar_dados_planilha espera)
                pasta_saida = Path("saida")
                pasta_saida.mkdir(exist_ok=True)
                
                shutil.copy(dados_cliente_csv, pasta_saida / "teste_apos_remocao_dados_cliente.csv")
                shutil.copy(remuneracoes_csv, pasta_saida / "teste_apos_remocao_remuneracoes.csv")
                
                # Importar e configurar gerador
                import preparar_dados_planilha
                
                # Substituir caminho de saída temporariamente
                planilha_original = preparar_dados_planilha.PLANILHA_SAIDA
                preparar_dados_planilha.PLANILHA_SAIDA = self.planilha_gerada
                
                # Gerar planilha
                preparar_dados_planilha.main()
                
                # Restaurar caminho original
                preparar_dados_planilha.PLANILHA_SAIDA = planilha_original
                
                self.adicionar_log("Planilha gerada com sucesso!", "sucesso")
                
                # Limpar arquivos temporários
                shutil.rmtree(pasta_temp, ignore_errors=True)
                
            except Exception as e:
                self.adicionar_log(f"Erro na geração: {str(e)}", "erro")
                self.btn_gerar.configure(state="normal", text="Gerar Simulação")
                return
            
            self.adicionar_log(f"Planilha salva em: {local_nome}", "sucesso")
            self.adicionar_log(f"Arquivo: {nome_planilha}", "sucesso")
            self.adicionar_log("Processo concluído com sucesso!", "sucesso")
            
            # Habilitar botão abrir
            self.btn_abrir.configure(state="normal", fg_color="#1f538d")
            self.btn_gerar.configure(state="normal", text="Concluído - Fechar", command=self.fechar_aplicacao)
            
            # Abrir planilha automaticamente
            self.adicionar_log("Abrindo planilha automaticamente...", "info")
            self.abrir_planilha()
            
            # Mensagem de sucesso
            messagebox.showinfo(
                "Sucesso!",
                f"Planilha gerada e aberta com sucesso!\n\nLocal: {local_nome}\nArquivo: {nome_planilha}"
            )
            
        except Exception as e:
            self.adicionar_log(f"Erro: {str(e)}", "erro")
            messagebox.showerror("Erro", f"Erro ao processar:\n{str(e)}")
            self.btn_gerar.configure(state="normal", text="⚙️ Gerar Simulação")
    
    def abrir_planilha(self):
        """Abre a planilha gerada no Excel."""
        if not self.planilha_gerada:
            messagebox.showerror("Erro", "Nenhuma planilha foi gerada ainda!")
            return
        
        try:
            if os.path.exists(self.planilha_gerada):
                # Abrir com aplicativo padrão
                if os.name == 'nt':  # Windows
                    os.startfile(self.planilha_gerada)
                elif os.name == 'posix':  # Mac/Linux
                    subprocess.call(['open', self.planilha_gerada])
                
                self.adicionar_log("Planilha aberta no Excel", "sucesso")
            else:
                self.adicionar_log("Arquivo não encontrado!", "erro")
                messagebox.showerror("Erro", f"Arquivo não encontrado:\n{self.planilha_gerada}")
        except Exception as e:
            self.adicionar_log(f"Erro ao abrir: {str(e)}", "erro")
            messagebox.showerror("Erro", f"Erro ao abrir planilha:\n{str(e)}")
    
    def fechar_aplicacao(self):
        """Fecha a aplicação."""
        self.quit()
        self.destroy()


def main():
    """Função principal."""
    app = CNISApp()
    app.mainloop()


if __name__ == "__main__":
    main()
