import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("SISTEMAS DE PLANILHAS")
        self.geometry("1920x1080")

      # Banco de dados com os nomes como chaves para aparecerem primeiro na seleção
        self.banco_de_dados = {
            "Arquibancada Coberta (1000 pess)": {"nome": "Item 01", "desc": "Estrutura tubular, altura 2,20m, piso em perfis de aço.", "preco": 1500.0},
            "Arquibancada 04 Níveis": {"nome": "Item 02", "desc": "Módulos cavalete, proteção corrimãos, compensado naval 18mm.", "preco": 800.0},
            "Camarim Montado (5x5m)": {"nome": "Item 03", "desc": "Toldo piramidal, piso carpete, ar-condicionado, iluminação.", "preco": 500.0},
            "Camarote Escalonado": {"nome": "Item 04", "desc": "Níveis de 1m e 2m, estrutura galvanizada, 24m x 6m.", "preco": 3000.0},
            "Fechamento Metálico (3x2,2m)": {"nome": "Item 05", "desc": "Estruturado metálico cor prata.", "preco": 50.0},
            "Disciplinador Alambrado": {"nome": "Item 06", "desc": "Ferro 1,50m de altura para contenção de público.", "preco": 25.0},
            "Praticável Pantográfico": {"nome": "Item 07", "desc": "Alumínio, altura máx 1m, piso emborrachado.", "preco": 120.0},
            "Módulo Elevado Policiamento": {"nome": "Item 08", "desc": "Aço galvanizado, 3x2m, 3 pisos, lona branca.", "preco": 450.0},
            "Palco Modular 10x8m": {"nome": "Item 09", "desc": "BoxTruss alumínio, 6m pé direito, House Mix 4x4m.", "preco": 4500.0},
            "Palco Modular 6x6m": {"nome": "Item 10", "desc": "BoxTruss alumínio, 6m pé direito, carpete, House Mix.", "preco": 2800.0},
            "Palco Modular 12x10m": {"nome": "Item 11", "desc": "8m pé direito, House Mix 2 andares, 3 camarins, 2 torres PA.", "preco": 8500.0},
            "Tablado com Cobertura 4x4m": {"nome": "Item 12", "desc": "Estrutura tubular, carpete, OSB 12mm.", "preco": 400.0},
            "Passarela Tubular 10x2,2m": {"nome": "Item 13", "desc": "Chapa de aço, revestida com lycra ou lona.", "preco": 700.0},
            "Tablado 6x4m Elevado": {"nome": "Item 14", "desc": "Toldo piramidal, altura 1,10m, painéis Octanorm.", "preco": 650.0},
            "Pórtico Alumínio K30/K50": {"nome": "Item 15", "desc": "10m comprimento, 5m altura, testeira para painel.", "preco": 1200.0},
            "Pórtico Duplo Tubular": {"nome": "Item 16", "desc": "Torres 2x2m, vãos de 16m, altura 7m, pintura acrílica.", "preco": 2500.0},
            "Toldo 12x12m": {"nome": "Item 17", "desc": "Aço galvanizado, 4 águas, lona Sanlux 4 UV.", "preco": 800.0},
            "Toldo 10x10m": {"nome": "Item 18", "desc": "Aço galvanizado, 4 águas, lona vulcanizada.", "preco": 600.0},
            "Toldo 8x8m": {"nome": "Item 19", "desc": "Aço galvanizado, lona branca translúcida.", "preco": 450.0},
            "Toldo 6x6m": {"nome": "Item 20", "desc": "Estrutura 4 águas, lona Sanlux 4, antimofo.", "preco": 300.0},
            "Toldo 5x5m": {"nome": "Item 21", "desc": "Estrutura aço galvanizado, lona UV.", "preco": 200.0},
            "Toldo 4x4m": {"nome": "Item 22", "desc": "Estrutura aço galvanizado, lona vulcanizada.", "preco": 150.0},
            "Toldo 3x3m": {"nome": "Item 23", "desc": "Estrutura aço galvanizado, lona branca.", "preco": 100.0}
        }
        
        self.carrinho = []

        # --- UI Setup ---
        ctk.CTkLabel(self, text="PLANILHAS RÁPIDAS", font=("Arial", 50, "bold")).pack(pady=10)

        # Seleção
        self.combo_itens = ctk.CTkComboBox(self, values=list(self.banco_de_dados.keys()), width=200, command=self.atualizar_info)
        self.combo_itens.pack(pady=5)
        
        self.lbl_nome_item = ctk.CTkLabel(self, text="Selecione um item acima", font=("Arial", 14, "italic"))
        self.lbl_nome_item.pack()

        self.txt_nome = ctk.CTkTextbox(self, width=700, height=70)
        self.txt_nome.pack(pady=10)

        # Inputs
        frame = ctk.CTkFrame(self)
        frame.pack(pady=10)
        self.ent_qtd = ctk.CTkEntry(frame, placeholder_text="Quantidade", width=120)
        self.ent_qtd.grid(row=0, column=0, padx=5)
        self.ent_dias = ctk.CTkEntry(frame, placeholder_text="Dias", width=120)
        self.ent_dias.grid(row=0, column=1, padx=5)

        # Botões
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="Adicionar Item", command=self.add_item).grid(row=0, column=0, padx=5)
        ctk.CTkButton(btn_frame, text="Gerar Planilha e Gráfico", fg_color="green", command=self.finalizar).grid(row=0, column=1, padx=5)

        # Tabela
        self.tabela = ctk.CTkTextbox(self, width=900, height=300)
        self.tabela.pack(pady=10)
        
        self.lbl_total = ctk.CTkLabel(self, text="TOTAL GERAL: R$ 0,00", font=("Arial", 22, "bold"), text_color="green")
        self.lbl_total.pack(pady=10)

    def atualizar_info(self, escolha):
        # Agora a 'escolha' é o nome do item (ex: Arquibancada Coberta)
        item = self.banco_de_dados[escolha]
        # O label agora mostrará o "Item XX" correspondente
        self.lbl_nome_item.configure(text=item['nome'])
        self.txt_desc.configure(state="normal")
        self.txt_desc.delete("1.0", "end")
        self.txt_desc.insert("1.0", item['desc'])
        self.txt_desc.configure(state="disabled")

    def add_item(self):
        try:
            # Limpa espaços em branco para evitar o erro de conversão
            qtd_val = self.ent_qtd.get().strip()
            dias_val = self.ent_dias.get().strip()
            
            if not qtd_val or not dias_val:
                raise ValueError
                
            q = int(qtd_val)
            d = int(dias_val)
            
            # 'escolha' agora é o nome técnico (Arquibancada...)
            escolha = self.combo_itens.get()
            item = self.banco_de_dados[escolha]
            sub = (q * item['preco']) * d
            
            self.carrinho.append({
                "ID": item['nome'],     # Salva como "Item XX" na planilha
                "Descrição": item['desc'],   # Salva o nome por extenso
                "Unitário": item['preco'],
                "Qtd": q,
                "Dias": d,
                "Subtotal": sub
            })
            self.render_tabela()
        except ValueError:
            messagebox.showerror("Erro", "Verifique se a quantidade e dias são números inteiros.")

    def render_tabela(self):
        self.tabela.delete("1.0", "end")
        total = 0
        header = f"{'CÓD':<8} | {'DESCRIÇÃO':<40} | {'VALOR.':<10} | {'QTD':<5} | {'DIAS':<5} | {'SUBTOTAL'}\n"
        self.tabela.insert("end", header + "-"*100 + "\n")
        for i in self.carrinho:
            linha = f"{i['ID']:<8} | {i['Descrição'][:38]:<40} | {i['Unitário']:>10.2f} | {i['Qtd']:^5} | {i['Dias']:^5} | R$ {i['Subtotal']:>12.2f}\n"
            self.tabela.insert("end", linha)
            total += i['Subtotal']
        self.lbl_total.configure(text=f"TOTAL GERAL: R$ {total:,.2f}")
        

    def finalizar(self):
        if not self.carrinho:
            return
        
        df = pd.DataFrame(self.carrinho)
        
        # 1. Salvar Excel
        filename = f"Orcamento_{datetime.now().strftime('%H%M%S')}.xlsx"
        df.to_excel(filename, index=False)
        
        # 2. Gerar Gráfico
        plt.figure(figsize=(10, 6))
        plt.bar(df['ID'], df['Subtotal'], color='skyblue')
        plt.xlabel('Itens do Contrato')
        plt.ylabel('Valor Total (R$)')
        plt.title('Distribuição de Custos por Item')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig("grafico_custos.png")
        plt.show()

        messagebox.showinfo("Sucesso", f"Planilha '{filename}' e Gráfico salvos com sucesso!")

if __name__ == "__main__":
    app = App()
    app.mainloop()