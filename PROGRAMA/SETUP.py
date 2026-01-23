import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import sqlite3
import os

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class BancoDados:
    """Gerencia exclusivamente a comunicação com o arquivo .db existente."""
    def __init__(self):
        self.db_name = "dados_contratos.db"
        # Verifica se o arquivo realmente existe na pasta antes de prosseguir
        if not os.path.exists(self.db_name):
            messagebox.showerror("Erro de Arquivo", f"O banco de dados '{self.db_name}' não foi encontrado.")

    def get_itens_nomes(self):
        """Busca os nomes dos itens para preencher a seleção (ComboBox)."""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("SELECT nome FROM itens")
            nomes = [row[0] for row in cursor.fetchall()]
            conn.close()
            return nomes
        except sqlite3.Error as e:
            print(f"Erro ao carregar nomes: {e}")
            return []

    def get_detalhes_por_nome(self, nome_item):
        """Recupera ID, descrição e preço do item selecionado."""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("SELECT id, descricao, preco FROM itens WHERE nome = ?", (nome_item,))
            resultado = cursor.fetchone()
            conn.close()
            return resultado
        except sqlite3.Error as e:
            print(f"Erro ao buscar detalhes: {e}")
            return None

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = BancoDados()
        self.title("RickSheet - Sistema de Orçamentos")
        self.state("zoomed")
        self.geometry("1000x850")
        
        self.carrinho = []

        # --- Título ---
        ctk.CTkLabel(self, text="GERADOR DE ORÇAMENTOS", font=("Arial", 30, "bold")).pack(pady=15)

        # --- Seleção ---
        ctk.CTkLabel(self, text="Selecione o Equipamento:", font=("Arial", 12)).pack()
        self.combo_itens = ctk.CTkComboBox(self, values=self.db.get_itens_nomes(), width=450, command=self.atualizar_info)
        self.combo_itens.pack(pady=5)
        
        self.lbl_cod_item = ctk.CTkLabel(self, text="Cód: --", font=("Arial", 14, "bold"), text_color="gray")
        self.lbl_cod_item.pack()

        self.txt_desc = ctk.CTkTextbox(self, width=750, height=120, font=("Arial", 12))
        self.txt_desc.pack(pady=10)

        # --- Entradas (Qtd/Dias) ---
        frame_inputs = ctk.CTkFrame(self)
        frame_inputs.pack(pady=10)
        self.ent_qtd = ctk.CTkEntry(frame_inputs, placeholder_text="Quantidade", width=140)
        self.ent_qtd.grid(row=0, column=0, padx=10)
        self.ent_dias = ctk.CTkEntry(frame_inputs, placeholder_text="Qtd. Dias", width=140)
        self.ent_dias.grid(row=0, column=1, padx=10)

        # --- Botões ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=15)
        ctk.CTkButton(btn_frame, text="Adicionar Item", command=self.add_item).grid(row=0, column=0, padx=10)
        ctk.CTkButton(btn_frame, text="Gerar Excel", fg_color="#27ae60", hover_color="#219150", command=self.finalizar).grid(row=0, column=1, padx=10)

        # --- Tabela ---
        self.tabela = ctk.CTkTextbox(self, width=950, height=300, font=("Courier New", 12))
        self.tabela.pack(pady=10)
        
        self.lbl_total = ctk.CTkLabel(self, text="TOTAL GERAL: R$ 0,00", font=("Arial", 24, "bold"), text_color="#2ecc71")
        self.lbl_total.pack(pady=10)

    def atualizar_info(self, escolha):
        """Atualiza a interface com os dados do banco."""
        dados = self.db.get_detalhes_por_nome(escolha)
        if dados:
            id_item, descricao, preco = dados
            self.lbl_cod_item.configure(text=f"Código: {id_item}")
            self.txt_desc.delete("1.0", "end")
            self.txt_desc.insert("1.0", descricao)
            self.preco_atual = preco
            self.id_atual = id_item

    def add_item(self):
        """Adiciona o item configurado ao carrinho."""
        try:
            desc = self.txt_desc.get("1.0", "end-1c").strip()
            q = int(self.ent_qtd.get())
            d = int(self.ent_dias.get())
            
            if not hasattr(self, 'preco_atual'):
                return

            subtotal = (q * self.preco_atual) * d
            
            self.carrinho.append({
                "ID": self.id_atual,
                "Descrição": desc,
                "Unitário": self.preco_atual,
                "Qtd": q,
                "Dias": d,
                "Subtotal": subtotal
            })
            
            self.render_tabela()
            self.ent_qtd.delete(0, 'end')
            self.ent_dias.delete(0, 'end')
            
        except ValueError:
            messagebox.showerror("Erro", "Quantidade e Dias devem ser números inteiros.")

    def render_tabela(self):
        """Exibe o carrinho na tela."""
        self.tabela.delete("1.0", "end")
        total = sum(i['Subtotal'] for i in self.carrinho)
        header = f"{'ID':<10} | {'DESCRIÇÃO':<35} | {'UNIT.':<10} | {'QTD':<5} | {'DIAS':<5} | {'SUBTOTAL'}\n"
        self.tabela.insert("end", header + "-"*105 + "\n")
        for i in self.carrinho:
            linha = f"{i['ID']:<10} | {i['Descrição'][:33]:<35} | {i['Unitário']:>10.2f} | {i['Qtd']:^5} | {i['Dias']:^5} | R$ {i['Subtotal']:>12.2f}\n"
            self.tabela.insert("end", linha)
        self.lbl_total.configure(text=f"TOTAL GERAL: R$ {total:,.2f}")

    def finalizar(self):
        """Abre uma janela para o usuário escolher onde salvar o arquivo Excel."""
        if not self.carrinho:
            messagebox.showwarning("Aviso", "O carrinho está vazio!")
            return
            
        # Sugestão de nome de arquivo com data e hora
        nome_sugerido = f"Orcamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Abre a janela de "Salvar Como"
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
            initialfile=nome_sugerido,
            title="Escolha onde salvar seu orçamento"
        )
        
        # Se o usuário não cancelar a operação
        if caminho_arquivo:
            df = pd.DataFrame(self.carrinho)
            total_geral = df['Subtotal'].sum()
            
            # Adiciona linha de total
            linha_total = pd.DataFrame([{
                "ID": "", 
                "Descrição": "VALOR TOTAL GERAL", 
                "Unitário": "", 
                "Qtd": "", 
                "Dias": "", 
                "Subtotal": total_geral
            }])
            df_final = pd.concat([df, linha_total], ignore_index=True)
            
            try:
                # Salva o arquivo no caminho escolhido
                df_final.to_excel(caminho_arquivo, index=False)
                messagebox.showinfo("Sucesso", f"Planilha salva com sucesso!\nTotal: R$ {total_geral:,.2f}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()