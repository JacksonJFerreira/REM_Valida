import customtkinter
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from datetime import datetime
import pandas as pd

# Configuração do tema do customtkinter
customtkinter.set_appearance_mode("System")  # Modos: "System" (padrão), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Temas: "blue" (padrão), "green", "dark-blue"

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da janela principal
        self.title("Comparador de Dados REM vs. SIFAC")
        self.geometry("700x600")
        self.minsize(600, 500)

        # Variáveis para armazenar os caminhos selecionados
        self.rem_file_path = None
        self.sifac_folder_path = None
        
        # Variáveis para armazenar os resultados
        self.resultados_comparacao = []

        # Frame principal
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Título
        self.lbl_titulo = customtkinter.CTkLabel(
            self.main_frame, 
            text="Comparador de Dados REM vs. SIFAC", 
            font=customtkinter.CTkFont(size=20, weight="bold")
        )
        self.lbl_titulo.pack(pady=10)

        # Frame para seleção de arquivos
        self.frame_selecao = customtkinter.CTkFrame(self.main_frame)
        self.frame_selecao.pack(fill="x", padx=10, pady=10)

        # Seleção de Arquivo REM
        self.frame_rem = customtkinter.CTkFrame(self.frame_selecao)
        self.frame_rem.pack(fill="x", padx=5, pady=5)
        
        self.lbl_rem = customtkinter.CTkLabel(self.frame_rem, text="Arquivo REM:")
        self.lbl_rem.pack(side="left", padx=5)
        
        self.entry_rem = customtkinter.CTkEntry(self.frame_rem, width=400)
        self.entry_rem.pack(side="left", padx=5, fill="x", expand=True)
        
        self.btn_select_rem = customtkinter.CTkButton(
            self.frame_rem, 
            text="Selecionar", 
            command=self.selecionar_rem,
            width=100
        )
        self.btn_select_rem.pack(side="right", padx=5)

        # Seleção de Pasta SIFAC
        self.frame_sifac = customtkinter.CTkFrame(self.frame_selecao)
        self.frame_sifac.pack(fill="x", padx=5, pady=5)
        
        self.lbl_sifac = customtkinter.CTkLabel(self.frame_sifac, text="Pasta SIFAC:")
        self.lbl_sifac.pack(side="left", padx=5)
        
        self.entry_sifac = customtkinter.CTkEntry(self.frame_sifac, width=400)
        self.entry_sifac.pack(side="left", padx=5, fill="x", expand=True)
        
        self.btn_select_sifac = customtkinter.CTkButton(
            self.frame_sifac, 
            text="Selecionar", 
            command=self.selecionar_sifac,
            width=100
        )
        self.btn_select_sifac.pack(side="right", padx=5)

        # Botão para iniciar a comparação
        self.btn_comparar = customtkinter.CTkButton(
            self.main_frame, 
            text="Comparar Dados", 
            command=self.comparar_dados,
            height=40,
            font=customtkinter.CTkFont(size=14, weight="bold")
        )
        self.btn_comparar.pack(pady=10)

        # Caixa de texto para exibição dos resultados
        self.frame_resultados = customtkinter.CTkFrame(self.main_frame)
        self.frame_resultados.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.lbl_resultados = customtkinter.CTkLabel(
            self.frame_resultados, 
            text="Resultados da Comparação:", 
            font=customtkinter.CTkFont(size=14, weight="bold")
        )
        self.lbl_resultados.pack(anchor="w", padx=5, pady=5)
        
        self.txt_resultados = customtkinter.CTkTextbox(
            self.frame_resultados, 
            width=650, 
            height=350,
            font=customtkinter.CTkFont(size=12)
        )
        self.txt_resultados.pack(fill="both", expand=True, padx=5, pady=5)

        # Botão para exportar resultados
        self.btn_exportar = customtkinter.CTkButton(
            self.main_frame, 
            text="Exportar Resultados", 
            command=self.exportar_resultados
        )
        self.btn_exportar.pack(pady=10)

    def selecionar_rem(self):
        # Abre a janela para o usuário escolher o arquivo REM (Excel)
        self.rem_file_path = filedialog.askopenfilename(
            title="Selecione o arquivo REM",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if self.rem_file_path:
            self.entry_rem.delete(0, tk.END)
            self.entry_rem.insert(0, self.rem_file_path)
            self.txt_resultados.delete("0.0", tk.END)
            self.txt_resultados.insert("0.0", f"Arquivo REM selecionado:\n{self.rem_file_path}\n\n")
        else:
            self.txt_resultados.insert("end", "Nenhum arquivo REM selecionado.\n\n")

    def selecionar_sifac(self):
        # Abre a janela para o usuário selecionar a pasta que contém os arquivos SIFAC
        self.sifac_folder_path = filedialog.askdirectory(title="Selecione a pasta SIFAC")
        if self.sifac_folder_path:
            self.entry_sifac.delete(0, tk.END)
            self.entry_sifac.insert(0, self.sifac_folder_path)
            self.txt_resultados.insert("end", f"Pasta SIFAC selecionada:\n{self.sifac_folder_path}\n\n")
        else:
            self.txt_resultados.insert("end", "Nenhuma pasta SIFAC selecionada.\n\n")

    def comparar_dados(self):
        self.txt_resultados.delete("0.0", tk.END)
        self.txt_resultados.insert("end", "Iniciando comparação de dados...\n\n")
        self.resultados_comparacao = []  # Limpa os resultados anteriores
        
        # Verifica se os caminhos foram selecionados
        if not self.entry_rem.get() or not self.entry_sifac.get():
            self.txt_resultados.insert("end", "Por favor, selecione o arquivo REM e a pasta SIFAC.\n")
            return
        
        # Atualiza as variáveis de caminho a partir dos campos de entrada
        self.rem_file_path = self.entry_rem.get()
        self.sifac_folder_path = self.entry_sifac.get()

        # Lê o arquivo REM e extrai os dados das células Q11 e Q7
        try:
            wb_rem = openpyxl.load_workbook(self.rem_file_path, data_only=True)
            ws_rem = wb_rem.active  # utiliza a primeira planilha
            rem_contrato = ws_rem["Q11"].value
            rem_competencia = ws_rem["Q7"].value
            
            # Formata os dados para melhor visualização
            if isinstance(rem_competencia, datetime):
                rem_competencia_str = rem_competencia.strftime("%d/%m/%Y")
            else:
                rem_competencia_str = str(rem_competencia)
                
            self.txt_resultados.insert("end", f"Arquivo REM:\n")
            self.txt_resultados.insert("end", f"  Contrato (Q11): {rem_contrato}\n")
            self.txt_resultados.insert("end", f"  Competência/Data (Q7): {rem_competencia_str}\n\n")
            
            # Adiciona informações do REM aos resultados
            self.resultados_comparacao.append({
                "tipo": "REM",
                "arquivo": os.path.basename(self.rem_file_path),
                "contrato": rem_contrato,
                "competencia": rem_competencia_str
            })
            
        except Exception as e:
            self.txt_resultados.insert("end", f"Erro ao ler arquivo REM: {e}\n")
            return

        # Variáveis para estatísticas
        total_arquivos = 0
        arquivos_ok = 0
        arquivos_divergentes = 0

        # Percorre os arquivos da pasta SIFAC
        self.txt_resultados.insert("end", "Arquivos SIFAC analisados:\n\n")
        for arquivo in os.listdir(self.sifac_folder_path):
            if arquivo.endswith(('.xlsx', '.xls')):
                total_arquivos += 1
                caminho_arquivo = os.path.join(self.sifac_folder_path, arquivo)
                try:
                    wb_sifac = openpyxl.load_workbook(caminho_arquivo, data_only=True)
                    ws_sifac = wb_sifac.active
                    sifac_contrato = ws_sifac["C5"].value
                    sifac_competencia = ws_sifac["F5"].value
                    
                    # Formata os dados para melhor visualização
                    if isinstance(sifac_competencia, datetime):
                        sifac_competencia_str = sifac_competencia.strftime("%d/%m/%Y")
                    else:
                        sifac_competencia_str = str(sifac_competencia)

                    self.txt_resultados.insert("end", f"Arquivo: {arquivo}\n")
                    self.txt_resultados.insert("end", f"  Contrato (C5): {sifac_contrato}\n")
                    self.txt_resultados.insert("end", f"  Competência/Data (F5): {sifac_competencia_str}\n")

                    # Comparação entre os dados do REM e do arquivo SIFAC
                    result = {}
                    result["tipo"] = "SIFAC"
                    result["arquivo"] = arquivo
                    result["contrato"] = sifac_contrato
                    result["competencia"] = sifac_competencia_str
                    
                    # Verificação de contrato
                    if str(rem_contrato).strip() == str(sifac_contrato).strip():
                        contrato_ok = True
                        result["contrato_ok"] = True
                    else:
                        contrato_ok = False
                        result["contrato_ok"] = False
                        
                    # Verificação de competência/data - pode precisar de ajustes dependendo do formato
                    if self.comparar_datas(rem_competencia, sifac_competencia):
                        competencia_ok = True
                        result["competencia_ok"] = True
                    else:
                        competencia_ok = False
                        result["competencia_ok"] = False
                    
                    if contrato_ok and competencia_ok:
                        self.txt_resultados.insert("end", "  Resultado: ✓ OK ✓\n\n")
                        result["status"] = "OK"
                        arquivos_ok += 1
                    else:
                        msg_erro = "  Resultado: ✗ Divergência encontrada ✗\n"
                        if not contrato_ok:
                            msg_erro += "    • Contrato não confere\n"
                        if not competencia_ok:
                            msg_erro += "    • Data/Competência não confere\n"
                        self.txt_resultados.insert("end", msg_erro + "\n")
                        result["status"] = "Divergência"
                        arquivos_divergentes += 1
                    
                    self.resultados_comparacao.append(result)
                    
                except Exception as e:
                    self.txt_resultados.insert("end", f"Erro ao ler arquivo {arquivo}: {e}\n\n")

        # Exibe o resumo da comparação
        self.txt_resultados.insert("end", "\n=== Resumo da Comparação ===\n")
        self.txt_resultados.insert("end", f"Total de arquivos analisados: {total_arquivos}\n")
        self.txt_resultados.insert("end", f"Arquivos com dados OK: {arquivos_ok}\n")
        self.txt_resultados.insert("end", f"Arquivos com divergências: {arquivos_divergentes}\n")
        self.txt_resultados.insert("end", "==========================\n\n")
        self.txt_resultados.insert("end", "Comparação concluída.\n")

    def comparar_datas(self, data1, data2):
        """
        Compara duas datas em diferentes formatos possíveis
        """
        # Se ambas forem objetos datetime
        if isinstance(data1, datetime) and isinstance(data2, datetime):
            return data1.date() == data2.date()
        
        # Tenta converter para string e comparar
        str_data1 = str(data1).strip()
        str_data2 = str(data2).strip()
        
        # Se forem strings idênticas
        if str_data1 == str_data2:
            return True
            
        # Tenta converter ambas para datetime
        try:
            # Tenta vários formatos possíveis
            formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"]
            
            for formato in formatos:
                try:
                    dt1 = datetime.strptime(str_data1, formato).date()
                    for f2 in formatos:
                        try:
                            dt2 = datetime.strptime(str_data2, f2).date()
                            if dt1 == dt2:
                                return True
                        except:
                            continue
                except:
                    continue
        except:
            pass
            
        # Se chegou até aqui, não conseguiu confirmar que são iguais
        return False

    def exportar_resultados(self):
        """
        Exporta os resultados da comparação para um arquivo Excel
        """
        if not self.resultados_comparacao:
            messagebox.showinfo("Exportar Resultados", "Não há resultados para exportar. Execute uma comparação primeiro.")
            return
            
        # Solicita o local para salvar o arquivo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            title="Salvar Resultados"
        )
        
        if not file_path:
            return
            
        try:
            # Cria um DataFrame com os resultados
            df = pd.DataFrame(self.resultados_comparacao)
            
            # Exporta para Excel
            df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Exportar Resultados", f"Resultados exportados com sucesso para:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro ao Exportar", f"Ocorreu um erro ao exportar os resultados:\n{str(e)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
