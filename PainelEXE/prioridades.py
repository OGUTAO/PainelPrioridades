import customtkinter as ctk
import pandas as pd
import os
import time

# --- CONFIGURAÇÃO GERAL ---
# Alterne para False para usar seus arquivos Excel reais
USAR_DADOS_PROVISORIOS = True

# --- Configurações de Arquivos (só são usadas se USAR_DADOS_PROVISORIOS for False) ---
CAMINHO_PLANILHA_PRIORIDADES = r"D:\PainelEXE\Dados\Fila de prioridades do laboratório.xlsx"
CAMINHO_PLANILHA_STATUS = r"D:\PainelEXE\Dados\Status dos Pedidos.xlsx" 

# Nomes das colunas
COLUNA_PEDIDO_ID = 'Cotação de Venda  ↑'
COLUNA_QTD = 'Quantidade Solicitada'
COLUNA_SERVICO = 'Produto / Serviço: Descrição do Produto/Serviço'
COLUNA_PV = 'Pv de Transferência'
COLUNA_STATUS_PEDIDO_ID = 'Pedido'
COLUNA_STATUS = 'Status'

# Constantes de Status
STATUS_PENDENTE = 'Pendente'
STATUS_AGUARDANDO = 'Aguardando Montagem'
STATUS_CONCLUIDO = 'Concluído'
STATUS_CANCELADO = 'Cancelado'

# Intervalo de verificação em milissegundos para o loop da interface
INTERVALO_CHECK_MS = 5000 

# --- LÓGICA DE DADOS (Não muda) ---

def carregar_e_mesclar_dados():
    # Esta função é a mesma de antes
    if not os.path.exists(CAMINHO_PLANILHA_PRIORIDADES) or not os.path.exists(CAMINHO_PLANILHA_STATUS):
        raise FileNotFoundError("Uma ou ambas as planilhas não foram encontradas.")
    df_prioridades = pd.read_excel(CAMINHO_PLANILHA_PRIORIDADES)
    df_status = pd.read_excel(CAMINHO_PLANILHA_STATUS)
    df_prioridades = df_prioridades[[COLUNA_PEDIDO_ID, COLUNA_QTD, COLUNA_SERVICO, COLUNA_PV]]
    df_prioridades.columns = ['Pedido', 'Maquinas', 'Servico', 'PV']
    df_status = df_status[[COLUNA_STATUS_PEDIDO_ID, COLUNA_STATUS]]
    df_status.columns = ['Pedido', 'Status']
    df_prioridades = df_prioridades.dropna(subset=['Pedido']).copy()
    df_prioridades['Maquinas'] = pd.to_numeric(df_prioridades['Maquinas'], errors='coerce').fillna(0).astype(int)
    df_merged = pd.merge(df_prioridades, df_status, on='Pedido', how='left')
    df_merged['Status'] = df_merged['Status'].fillna(STATUS_PENDENTE)
    return df_merged

def criar_dados_provisorios():
    # Esta função é a mesma de antes
    dados_prioridades = {
        'Pedido': ['CV-2025-001', 'PV-10580', 'OP-98765', 'CV-2025-002', 'OP-98770', 'PV-10588', 'CV-2025-003'],
        'Maquinas': [5, 1, 10, 2, 1, 3, 8],
        'Servico': [
            'Montagem de painel elétrico completo com disjuntores Siemens para cliente industrial. Inclui teste de isolamento.',
            'Usinagem de flange de precisão em aço inox 316. Peça crítica para sistema de vácuo.',
            'Calibração e ajuste de 10 sensores de pressão modelo XPT-5. Serviço no laboratório Teravix.',
            'Desenvolvimento de firmware para CLP de controle de esteira transportadora. Requer teste em bancada.',
            'Manutenção corretiva em sistema de bombeamento. Substituir selo mecânico e rolamentos.',
            'Corte e dobra de 3 chapas de alumínio conforme desenho técnico 123-A.',
            'Impressão 3D de protótipo de carcaça para dispositivo eletrônico em material ABS de alta resistência.'
        ],
        'PV': ['PV-T101', 'PV-T102', 'PV-T103', 'PV-T104', 'PV-T105', 'PV-T106', 'PV-T107']
    }
    df_prioridades = pd.DataFrame(dados_prioridades)
    dados_status = { 'Pedido': ['CV-2025-001', 'OP-98765', 'CV-2025-002', 'PV-10588'], 'Status': [STATUS_AGUARDANDO, STATUS_CANCELADO, STATUS_CONCLUIDO, STATUS_PENDENTE] }
    df_status = pd.DataFrame(dados_status)
    df_merged = pd.merge(df_prioridades, df_status, on='Pedido', how='left')
    df_merged['Status'] = df_merged['Status'].fillna(STATUS_PENDENTE)
    return df_merged


# <<< MUDANÇA: APLICAÇÃO INTEIRA REESCRITA USANDO CUSTOMTKINTER >>>

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- CONFIGURAÇÃO DA JANELA PRINCIPAL ---
        self.title("Painel de Produção")
        self.geometry("1280x720") # Tamanho inicial, mas a janela é redimensionável
        self.configurar_tema()
        
        # --- CONFIGURAÇÃO DO GRID RESPONSIVO ---
        # A janela terá 2 colunas: a lista (peso 1) e os detalhes (peso 3)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(1, weight=1)

        # --- DADOS E ESTADO ---
        self.dados_df = pd.DataFrame()
        self.timestamp_prioridades, self.timestamp_status = 0, 0
        self.pedido_selecionado = None

        # --- CRIAÇÃO DOS WIDGETS ---
        self.criar_widgets_cabecalho()
        self.criar_widgets_painel_principal()
        
        # --- INICIALIZAÇÃO ---
        self.atualizar_dados_e_ui() # Primeira carga de dados
        self.iniciar_loop_de_verificacao() # Inicia o auto-refresh

    def configurar_tema(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.colors = {'pendente': '#D4AC0D', 'aguardando': '#2471A3', 'cancelado': '#C0392B'}

    def criar_widgets_cabecalho(self):
        # Frame para o cabeçalho
        self.cabecalho_frame = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.cabecalho_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=0, pady=0)
        
        self.label_titulo = ctk.CTkLabel(self.cabecalho_frame, text="PAINEL DE PRODUÇÃO", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_titulo.pack(side="left", padx=20)

        self.botao_lixeira = ctk.CTkButton(self.cabecalho_frame, text="Ver Cancelados", command=self.mostrar_tela_lixeira)
        self.botao_lixeira.pack(side="right", padx=20)
        
        self.botao_atualizar = ctk.CTkButton(self.cabecalho_frame, text="Atualizar", command=self.atualizar_dados_e_ui)
        self.botao_atualizar.pack(side="right", padx=5)

    def criar_widgets_painel_principal(self):
        # Frame principal que conterá a lista e os detalhes
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=3)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # Frame com scroll para a lista de pedidos (esquerda)
        self.lista_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="Pedidos Ativos")
        self.lista_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Frame para os detalhes do pedido (direita)
        self.detalhes_frame = ctk.CTkFrame(self.main_frame, fg_color=("gray85", "gray17"))
        self.detalhes_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

    def atualizar_dados_e_ui(self):
        """Função central que carrega os dados e atualiza toda a interface."""
        print("Atualizando dados...")
        try:
            if USAR_DADOS_PROVISORIOS:
                self.dados_df = criar_dados_provisorios()
            else:
                self.dados_df = carregar_e_mesclar_dados()
            
            self.mostrar_tela_principal() # Garante que a tela principal seja exibida
        except Exception as e:
            self.mostrar_erro(str(e))

    def mostrar_tela_principal(self):
        """Limpa e recria a lista de pedidos ativos."""
        # Limpa widgets antigos da lista
        for widget in self.lista_frame.winfo_children():
            widget.destroy()
            
        # Limpa os detalhes
        self.limpar_detalhes()

        df_ativos = self.dados_df[~self.dados_df['Status'].isin([STATUS_CONCLUIDO, STATUS_CANCELADO])].reset_index()

        if df_ativos.empty:
            ctk.CTkLabel(self.lista_frame, text="Nenhum pedido ativo.").pack(pady=10)
            return

        # Cria um botão para cada pedido ativo
        for index, row in df_ativos.iterrows():
            pedido_id = row['Pedido']
            # O comando lambda é crucial para passar o 'row' correto para a função
            btn = ctk.CTkButton(self.lista_frame, text=pedido_id,
                                command=lambda r=row: self.mostrar_detalhes(r))
            btn.pack(fill="x", padx=5, pady=3)
        
        # Mostra os detalhes do primeiro item da lista por padrão
        self.mostrar_detalhes(df_ativos.iloc[0])

    def mostrar_detalhes(self, pedido_data):
        """Mostra os detalhes do pedido selecionado no painel direito."""
        self.limpar_detalhes()
        
        status = pedido_data['Status']
        cor = self.colors.get(status.lower().replace(" ", "_"), "gray")

        ctk.CTkLabel(self.detalhes_frame, text=pedido_data['Pedido'], font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(10, 5), padx=20, anchor="w")
        ctk.CTkLabel(self.detalhes_frame, text=status.upper(), font=ctk.CTkFont(size=18, weight="bold"), text_color=cor).pack(pady=5, padx=20, anchor="w")
        
        separator = ctk.CTkFrame(self.detalhes_frame, height=2, fg_color="gray50")
        separator.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(self.detalhes_frame, text="Descrição do Serviço:", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(10, 2), padx=20, anchor="w")
        ctk.CTkLabel(self.detalhes_frame, text=pedido_data['Servico'], wraplength=self.detalhes_frame.winfo_width() - 60, justify="left").pack(pady=2, padx=20, anchor="w")

        ctk.CTkLabel(self.detalhes_frame, text=f"Quantidade: {pedido_data['Maquinas']}", font=ctk.CTkFont(size=16)).pack(pady=(20, 2), padx=20, anchor="w")
        ctk.CTkLabel(self.detalhes_frame, text=f"PV de Transferência: {pedido_data['PV']}", font=ctk.CTkFont(size=16)).pack(pady=2, padx=20, anchor="w")
        
    def limpar_detalhes(self):
        for widget in self.detalhes_frame.winfo_children():
            widget.destroy()
            
    def mostrar_tela_lixeira(self):
        """Abre uma nova janela (Toplevel) para mostrar os pedidos cancelados."""
        lixeira_window = ctk.CTkToplevel(self)
        lixeira_window.title("Pedidos Cancelados")
        lixeira_window.geometry("800x600")
        lixeira_window.transient(self) # Mantém a janela sobre a principal
        
        scroll_frame = ctk.CTkScrollableFrame(lixeira_window, label_text="Itens Cancelados")
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        df_cancelados = self.dados_df[self.dados_df['Status'] == STATUS_CANCELADO]

        if df_cancelados.empty:
            ctk.CTkLabel(scroll_frame, text="A lixeira está vazia.").pack(pady=10)
        else:
            for index, row in df_cancelados.iterrows():
                item_text = f"Pedido: {row['Pedido']}  |  Serviço: {str(row['Servico'])[:50]}..."
                ctk.CTkLabel(scroll_frame, text=item_text, anchor="w").pack(fill="x", padx=5, pady=2)

    def mostrar_erro(self, mensagem):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.main_frame, text=f"ERRO:\n{mensagem}", text_color="red", font=ctk.CTkFont(size=18)).pack(pady=50)

    def iniciar_loop_de_verificacao(self):
        """Verifica se os arquivos foram modificados e agenda a próxima verificação."""
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
                pass # Erro já será tratado na atualização manual
        
        # Agenda a próxima chamada desta função
        self.after(INTERVALO_CHECK_MS, self.iniciar_loop_de_verificacao)

if __name__ == "__main__":
    app = App()
    app.mainloop()