import customtkinter as ctk
from PIL import Image
import pandas as pd
import os
import time

# --- CONFIGURAÇÃO GERAL ---
USAR_DADOS_PROVISORIOS = True

# --- Configurações de Arquivos ---
CAMINHO_PLANILHA_PRIORIDADES = r"D:\PainelEXE\Dados\Fila de prioridades do laboratório.xlsx"
CAMINHO_PLANILHA_STATUS = r"D:\PainelEXE\Dados\Status dos Pedidos.xlsx" 

# --- Nomes das colunas ---
# <<< MUDANÇA: Adicionada a coluna de Prioridade >>>
COLUNA_PRIORIDADE = 'Prioridade'
COLUNA_PEDIDO_ID = 'Cotação de Venda  ↑'
COLUNA_QTD = 'Quantidade Solicitada'
COLUNA_SERVICO = 'Produto / Serviço: Descrição do Produto/Serviço'
COLUNA_PV = 'Pv de Transferência'
COLUNA_STATUS_PEDIDO_ID = 'Pedido'
COLUNA_STATUS = 'Status'

# --- Constantes de Status ---
STATUS_PENDENTE = 'Pendente'
STATUS_AGUARDANDO = 'Aguardando Montagem'
STATUS_CONCLUIDO = 'Concluído'
STATUS_CANCELADO = 'Cancelado'

# --- Intervalo de verificação em milissegundos ---
INTERVALO_CHECK_MS = 5000 

# --- LÓGICA DE DADOS ---

def carregar_e_mesclar_dados():
    if not os.path.exists(CAMINHO_PLANILHA_PRIORIDADES) or not os.path.exists(CAMINHO_PLANILHA_STATUS):
        raise FileNotFoundError("Uma ou ambas as planilhas não foram encontradas.")
    
    df_prioridades = pd.read_excel(CAMINHO_PLANILHA_PRIORIDADES)
    df_status = pd.read_excel(CAMINHO_PLANILHA_STATUS)
    
    # <<< MUDANÇA: Inclui a leitura da nova coluna 'Prioridade' >>>
    colunas_necessarias = [COLUNA_PRIORIDADE, COLUNA_PEDIDO_ID, COLUNA_QTD, COLUNA_SERVICO, COLUNA_PV]
    df_prioridades = df_prioridades[colunas_necessarias]
    df_prioridades.columns = ['Prioridade', 'Pedido', 'Maquinas', 'Servico', 'PV']
    
    df_status = df_status[[COLUNA_STATUS_PEDIDO_ID, COLUNA_STATUS]]
    df_status.columns = ['Pedido', 'Status']

    df_prioridades = df_prioridades.dropna(subset=['Pedido', 'Prioridade']).copy()
    df_prioridades['Maquinas'] = pd.to_numeric(df_prioridades['Maquinas'], errors='coerce').fillna(0).astype(int)
    
    df_merged = pd.merge(df_prioridades, df_status, on='Pedido', how='left')
    df_merged['Status'] = df_merged['Status'].fillna(STATUS_PENDENTE)

    # <<< MUDANÇA: Ordena o DataFrame com base na nova coluna Prioridade >>>
    df_merged = df_merged.sort_values(by='Prioridade').reset_index(drop=True)
    
    return df_merged

def criar_dados_provisorios():
    dados_prioridades = {
        'Prioridade': [1, 2, 3, 4, 5, 6, 7], # <<< MUDANÇA: Coluna de prioridade adicionada
        'Pedido': ['CV-2025-001', 'PV-10580', 'OP-98765', 'CV-2025-002', 'OP-98770', 'PV-10588', 'CV-2025-003'],
        'Maquinas': [5, 1, 10, 2, 1, 3, 8],
        'Servico': [
            'Montagem de painel elétrico completo com disjuntores Siemens para cliente industrial.',
            'Usinagem de flange de precisão em aço inox 316.',
            'Calibração e ajuste de 10 sensores de pressão modelo XPT-5.',
            'Desenvolvimento de firmware para CLP de controle de esteira.',
            'Manutenção corretiva em sistema de bombeamento.',
            'Corte e dobra de 3 chapas de alumínio.',
            'Impressão 3D de protótipo de carcaça para dispositivo eletrônico.'
        ],
        'PV': ['PV-T101', 'PV-T102', 'PV-T103', 'PV-T104', 'PV-T105', 'PV-T106', 'PV-T107']
    }
    df_prioridades = pd.DataFrame(dados_prioridades)
    
    dados_status = { 'Pedido': ['CV-2025-001', 'OP-98765', 'CV-2025-002', 'PV-10588', 'OP-98770', 'PV-10580'], 'Status': [STATUS_AGUARDANDO, STATUS_CANCELADO, STATUS_CONCLUIDO, STATUS_PENDENTE, STATUS_AGUARDANDO, STATUS_PENDENTE] }
    df_status = pd.DataFrame(dados_status)
    
    df_merged = pd.merge(df_prioridades, df_status, on='Pedido', how='left')
    df_merged['Status'] = df_merged['Status'].fillna(STATUS_PENDENTE)
    
    df_merged = df_merged.sort_values(by='Prioridade').reset_index(drop=True)
    return df_merged

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Painel de Produção MTEC")
        self.geometry("1600x900") # Aumentado para comportar mais informação
        self.configurar_tema_mtec()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=5) # Mais peso para a área de detalhes
        self.grid_rowconfigure(1, weight=1)

        self.dados_df = pd.DataFrame()
        self.timestamp_prioridades, self.timestamp_status = 0, 0

        self.criar_widgets()
        self.atualizar_dados_e_ui()
        self.iniciar_loop_de_verificacao()

    def configurar_tema_mtec(self):
        self.COLOR_ORANGE = "#FF6600"
        self.COLOR_BLACK = "#1C1C1C"
        self.COLOR_DARK_GRAY = "#2E2E2E"
        self.COLOR_LIGHT_GRAY = "#424242"
        self.COLOR_WHITE = "#FFFFFF"
        ctk.set_appearance_mode("dark")
        self.configure(fg_color=self.COLOR_BLACK)
        self.status_colors = {'pendente': '#D4AC0D', 'aguardando montagem': '#3498DB'}

    def carregar_imagens(self):
        # <<< MUDANÇA: Carrega o logo da MTEC e o ícone da lixeira >>>
        try:
            self.logo_image = ctk.CTkImage(Image.open("logoMtec.jpeg"), size=(150, 50))
        except FileNotFoundError:
            self.logo_image = None
        
        try:
            self.trash_icon = ctk.CTkImage(Image.open("trash_icon.png"), size=(24, 24))
        except FileNotFoundError:
            self.trash_icon = None

    def criar_widgets(self):
        self.carregar_imagens()
        
        # --- CABEÇALHO ---
        self.cabecalho_frame = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color=self.COLOR_DARK_GRAY)
        self.cabecalho_frame.grid(row=0, column=0, columnspan=2, sticky="new")
        
        # <<< MUDANÇA: Adiciona o logo no cabeçalho >>>
        if self.logo_image:
            logo_label = ctk.CTkLabel(self.cabecalho_frame, image=self.logo_image, text="")
            logo_label.pack(side="left", padx=20, pady=10)
        else:
            ctk.CTkLabel(self.cabecalho_frame, text="MTEC", font=ctk.CTkFont(size=24, weight="bold"), text_color=self.COLOR_ORANGE).pack(side="left", padx=20, pady=10)

        self.botao_lixeira = ctk.CTkButton(self.cabecalho_frame, text="", image=self.trash_icon, command=self.mostrar_tela_lixeira, width=40, fg_color="transparent", hover_color=self.COLOR_LIGHT_GRAY)
        self.botao_lixeira.pack(side="right", padx=(10, 20), pady=10)
        
        self.botao_atualizar = ctk.CTkButton(self.cabecalho_frame, text="Atualização Manual Necessária", command=self.atualizar_dados_e_ui, fg_color=self.COLOR_ORANGE, hover_color="#D15600")
        self.botao_atualizar.pack_forget()

        # --- FRAME PRINCIPAL ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=5)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        self.lista_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="FILA", label_fg_color=self.COLOR_DARK_GRAY, label_text_color=self.COLOR_ORANGE)
        self.lista_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # <<< MUDANÇA: Área de detalhes agora tem um grid 2x2 para 4 cards >>>
        self.detalhes_area_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.detalhes_area_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.detalhes_area_frame.grid_rowconfigure(0, weight=1)
        self.detalhes_area_frame.grid_rowconfigure(1, weight=1)
        self.detalhes_area_frame.grid_columnconfigure(0, weight=1)
        self.detalhes_area_frame.grid_columnconfigure(1, weight=1)

    def atualizar_dados_e_ui(self):
        print("Atualizando dados...")
        try:
            if USAR_DADOS_PROVISORIOS:
                self.dados_df = criar_dados_provisorios()
            else:
                self.dados_df = carregar_e_mesclar_dados()
            
            self.desenhar_interface_com_dados()
            self.botao_atualizar.pack_forget()
        except Exception as e:
            self.mostrar_erro(str(e))
            self.botao_atualizar.pack(side="right", padx=10, pady=10)

    def desenhar_interface_com_dados(self):
        # Limpa widgets antigos
        for widget in self.lista_frame.winfo_children(): widget.destroy()
        for widget in self.detalhes_area_frame.winfo_children(): widget.destroy()

        df_ativos = self.dados_df[~self.dados_df['Status'].isin([STATUS_CONCLUIDO, STATUS_CANCELADO])]
        df_pendentes = self.dados_df[self.dados_df['Status'] == STATUS_PENDENTE]

        # --- Popula a lista da esquerda com seções (Prioridades e Pendentes) ---
        ctk.CTkLabel(self.lista_frame, text="TOP 5 PRIORIDADES", font=ctk.CTkFont(weight="bold"), text_color=self.COLOR_ORANGE).pack(fill="x", padx=10, pady=(5,5))
        if df_ativos.empty:
            ctk.CTkLabel(self.lista_frame, text="Nenhum pedido ativo.").pack(pady=5, padx=10)
        else:
            # Mostra apenas os 5 primeiros na seção de prioridade
            for index, row in df_ativos.head(5).iterrows():
                texto_item = f"{int(row['Prioridade'])}. {row['Pedido']}"
                label = ctk.CTkLabel(self.lista_frame, text=texto_item, anchor="w")
                label.pack(fill="x", padx=10, pady=2)
        
        # --- Seção de Pendentes ---
        ctk.CTkFrame(self.lista_frame, height=2, fg_color=self.COLOR_LIGHT_GRAY).pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(self.lista_frame, text="PENDENTES", font=ctk.CTkFont(weight="bold"), text_color=self.status_colors['pendente']).pack(fill="x", padx=10, pady=(5,5))
        if df_pendentes.empty:
            ctk.CTkLabel(self.lista_frame, text="Nenhum pedido pendente.").pack(pady=5, padx=10)
        else:
            for index, row in df_pendentes.iterrows():
                texto_item = f"P{int(row['Prioridade'])}: {row['Pedido']}"
                label = ctk.CTkLabel(self.lista_frame, text=texto_item, anchor="w", text_color="gray")
                label.pack(fill="x", padx=10, pady=2)

        # --- Cria os cards de detalhes para os 4 primeiros ---
        posicoes = [(0, 0), (0, 1), (1, 0), (1, 1)]
        for i in range(min(len(df_ativos), 4)):
            pedido = df_ativos.iloc[i]
            r, c = posicoes[i]
            self.criar_card_pedido(self.detalhes_area_frame, pedido).grid(row=r, column=c, sticky="nsew", pady=5, padx=5)
            
    def criar_card_pedido(self, parent, pedido_data):
        card_frame = ctk.CTkFrame(parent, fg_color=self.COLOR_DARK_GRAY, border_color=self.COLOR_ORANGE, border_width=1, corner_radius=10)
        status = pedido_data['Status']
        cor_status = self.status_colors.get(status.lower(), self.COLOR_WHITE)

        titulo_texto = f"{int(pedido_data['Prioridade'])}º: {pedido_data['Pedido']}"
        ctk.CTkLabel(card_frame, text=titulo_texto, font=ctk.CTkFont(size=22, weight="bold"), text_color=self.COLOR_ORANGE).pack(pady=(10, 2), padx=15, anchor="w")
        ctk.CTkLabel(card_frame, text=status.upper(), font=ctk.CTkFont(size=16, weight="bold"), text_color=cor_status).pack(pady=0, padx=15, anchor="w")
        
        ctk.CTkFrame(card_frame, height=1, fg_color=self.COLOR_LIGHT_GRAY).pack(fill="x", padx=15, pady=8)

        ctk.CTkLabel(card_frame, text="Descrição:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(5, 2), padx=15, anchor="w")
        ctk.CTkLabel(card_frame, text=pedido_data['Servico'], wraplength=350, justify="left").pack(pady=2, padx=15, anchor="w")
        ctk.CTkLabel(card_frame, text=f"QTD: {pedido_data['Maquinas']}", font=ctk.CTkFont(size=14)).pack(pady=(10, 2), padx=15, anchor="w")
        ctk.CTkLabel(card_frame, text=f"PV: {pedido_data['PV']}", font=ctk.CTkFont(size=14)).pack(pady=2, padx=15, anchor="w")
        
        return card_frame

    def mostrar_tela_lixeira(self):
        lixeira_window = ctk.CTkToplevel(self)
        lixeira_window.title("Pedidos Cancelados")
        lixeira_window.geometry("800x600")
        lixeira_window.transient(self)
        lixeira_window.configure(fg_color=self.COLOR_BLACK)
        
        scroll_frame = ctk.CTkScrollableFrame(lixeira_window, label_text="Itens Cancelados", label_fg_color=self.COLOR_DARK_GRAY, label_text_color=self.COLOR_ORANGE)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        df_cancelados = self.dados_df[self.dados_df['Status'] == STATUS_CANCELADO]

        if df_cancelados.empty:
            ctk.CTkLabel(scroll_frame, text="A lixeira está vazia.").pack(pady=10)
        else:
            for index, row in df_cancelados.iterrows():
                frame_item = ctk.CTkFrame(scroll_frame, fg_color=self.COLOR_DARK_GRAY)
                item_text = f"Pedido: {row['Pedido']}  |  Serviço: {str(row['Servico'])[:50]}..."
                ctk.CTkLabel(frame_item, text=item_text, anchor="w").pack(fill="x", padx=10, pady=10)
                frame_item.pack(fill="x", padx=5, pady=3)
    
    def mostrar_erro(self, mensagem):
        for widget in self.main_frame.winfo_children(): widget.destroy()
        ctk.CTkLabel(self.main_frame, text=f"ERRO:\n{mensagem}", text_color="red", font=ctk.CTkFont(size=18)).pack(pady=50, padx=10)

    def iniciar_loop_de_verificacao(self):
        if not USAR_DADOS_PROVISORIOS:
            try:
                mod_prio = os.path.getmtime(CAMINHO_PLANILHA_PRIORIDADES)
                mod_stat = os.path.getmtime(CAMINHO_PLANILHA_STATUS)
                if (self.timestamp_prioridades != 0 and mod_prio != self.timestamp_prioridades) or \
                   (self.timestamp_status != 0 and mod_stat != self.timestamp_status):
                    print("Detecção de mudança nas planilhas. Atualizando...")
                    self.atualizar_dados_e_ui()
                self.timestamp_prioridades, self.timestamp_status = mod_prio, mod_stat
            except FileNotFoundError:
                if not self.botao_atualizar.winfo_viewable():
                    self.botao_atualizar.pack(side="right", padx=10, pady=10)
        self.after(INTERVALO_CHECK_MS, self.iniciar_loop_de_verificacao)

if __name__ == "__main__":
    app = App()
    app.mainloop()