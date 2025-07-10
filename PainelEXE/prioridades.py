import sys
import os
import time
import locale
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout,
                               QDialog)
from PySide6.QtGui import QPixmap, QFont, QColor
from PySide6.QtCore import QTimer, Qt, QSize

# --- CONFIGURAÇÃO GERAL ---
# Nomes dos arquivos Excel
NOME_ARQUIVO_PRIORIDADES = r"C:\Prioridades-Excel\Fila_de_prioridades_do_laboratório.xlsx"
NOME_ARQUIVO_STATUS = r"C:\Prioridades-Excel\Status_dos_pedidos.xlsx"

# Caminho para a pasta que contém os arquivos Excel
# Altere este caminho se as suas planilhas estiverem em outro lugar
CAMINHO_PASTA_EXCEL = r"C:\Prioridades-Excel"

# Caminhos completos para os arquivos
CAMINHO_PLANILHA_PRIORIDADES = os.path.join(CAMINHO_PASTA_EXCEL, NOME_ARQUIVO_PRIORIDADES)
CAMINHO_PLANILHA_STATUS = os.path.join(CAMINHO_PASTA_EXCEL, NOME_ARQUIVO_STATUS)

# --- Nomes das colunas ---
COLUNA_DATA_PRIORIDADE = 'Data Entrega Prorrogada'
COLUNA_PEDIDO_ID = 'Cotação de Venda  ↑' # Nome com dois espaços
COLUNA_QTD = 'Quantidade Solicitada'
COLUNA_SERVICO = 'Produto / Serviço: Descrição do Produto/Serviço'
COLUNA_PV = 'Pv de Transferência'
COLUNA_COD_ITEM = 'Código Item Integração'
COLUNA_VALOR_TOTAL = 'Valor Total'
COLUNA_STATUS_PEDIDO_ID = 'Pedido'
COLUNA_STATUS = 'Status'

# --- Constantes de Status ---
STATUS_PENDENTE = 'Pendente'
STATUS_AGUARDANDO = 'Aguardando Montagem'
STATUS_EM_MONTAGEM = 'Em Montagem'
STATUS_CONCLUIDO = 'Concluído'
STATUS_CANCELADO = 'Cancelado'

INTERVALO_CHECK_MS = 3000

# --- LÓGICA DE DADOS (Permanece a mesma) ---
def carregar_e_mesclar_dados():
    if not os.path.exists(CAMINHO_PLANILHA_PRIORIDADES):
        raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_PLANILHA_PRIORIDADES}")
    if not os.path.exists(CAMINHO_PLANILHA_STATUS):
        raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_PLANILHA_STATUS}")

    df_prioridades = pd.read_excel(CAMINHO_PLANILHA_PRIORIDADES)
    colunas_necessarias = [
        COLUNA_DATA_PRIORIDADE, COLUNA_PEDIDO_ID, COLUNA_QTD, COLUNA_SERVICO, 
        COLUNA_PV, COLUNA_COD_ITEM, COLUNA_VALOR_TOTAL
    ]
    
    for col in colunas_necessarias:
        if col not in df_prioridades.columns:
            raise ValueError(f"Coluna '{col}' não encontrada na planilha '{NOME_ARQUIVO_PRIORIDADES}'.")

    df_prioridades = df_prioridades[colunas_necessarias]
    df_prioridades[COLUNA_DATA_PRIORIDADE] = pd.to_datetime(df_prioridades[COLUNA_DATA_PRIORIDADE], errors='coerce')
    df_prioridades.dropna(subset=[COLUNA_DATA_PRIORIDADE, COLUNA_PEDIDO_ID], inplace=True)
    df_prioridades = df_prioridades.sort_values(by=COLUNA_DATA_PRIORIDADE)
    df_prioridades['Prioridade'] = range(1, len(df_prioridades) + 1)

    df_prioridades.rename(columns={
        COLUNA_PEDIDO_ID: 'Pedido', COLUNA_QTD: 'Maquinas',
        COLUNA_SERVICO: 'Servico', COLUNA_PV: 'PV',
        COLUNA_COD_ITEM: 'CodItem', COLUNA_VALOR_TOTAL: 'Valor'
    }, inplace=True)
    
    df_status = pd.read_excel(CAMINHO_PLANILHA_STATUS)
    df_status = df_status[[COLUNA_STATUS_PEDIDO_ID, COLUNA_STATUS]]
    df_status.columns = ['Pedido', 'Status']

    df_merged = pd.merge(df_prioridades, df_status, on='Pedido', how='left')
    df_merged['Status'] = df_merged['Status'].fillna(STATUS_PENDENTE)
    
    df_merged['Maquinas'] = pd.to_numeric(df_merged['Maquinas'], errors='coerce').fillna(0).astype(int)
    df_merged['Valor'] = pd.to_numeric(df_merged['Valor'], errors='coerce').fillna(0)
    
    return df_merged

# --- NOVA INTERFACE GRÁFICA COM PYSIDE6 ---

# Estilos da aplicação (similar a CSS)
STYLESHEET = """
    QMainWindow, QDialog {
        background-color: #1C1C1C;
    }
    QLabel {
        color: #FFFFFF;
    }
    #Header {
        background-color: #2E2E2E;
    }
    #TitleLabel, #SectionLabel {
        color: #FF6600;
    }
    #Card {
        background-color: #2E2E2E;
        border: 2px solid #FF6600;
        border-radius: 10px;
    }
    #LixeiraButton {
        background-color: #424242;
        color: white;
        padding: 8px 16px;
        border-radius: 5px;
        border: none;
    }
    #LixeiraButton:hover {
        background-color: #555555;
    }
    #ErrorLabel {
        color: #E74C3C;
    }
"""

class PainelMtec(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Painel de Produção MTEC")
        self.setGeometry(100, 100, 1600, 900)
        self.setStyleSheet(STYLESHEET)

        # Configura a localização para o formato de moeda brasileiro (R$)
        try:
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        except locale.Error:
            print("Aviso: Local 'pt_BR.UTF-8' não encontrado. Usando formatação padrão.")

        self.dados_df = pd.DataFrame()
        self.timestamp_prioridades, self.timestamp_status = 0, 0

        self.setup_ui()
        self.atualizar_dados_e_ui()

        # Timer para atualização automática
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.verificar_atualizacoes)
        self.timer.start(INTERVALO_CHECK_MS)

    def setup_ui(self):
        # Widget Central
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # 1. Cabeçalho
        header = QWidget()
        header.setObjectName("Header")
        header.setFixedHeight(70)
        header_layout = QHBoxLayout(header)
        
        logo_label = QLabel()
        pixmap = QPixmap("logoMtec.jpeg")
        if not pixmap.isNull():
            logo_label.setPixmap(pixmap.scaled(150, 50, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            logo_label.setText("MTEC")
            logo_label.setObjectName("TitleLabel")
            logo_label.setFont(QFont("Inter", 24, QFont.Bold))
        
        lixeira_btn = QLabel("<a href='#'>Lixeira</a>")
        lixeira_btn.setObjectName("LixeiraButton")
        lixeira_btn.setFont(QFont("Inter", 12, QFont.Bold))
        lixeira_btn.linkActivated.connect(self.mostrar_lixeira)

        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        header_layout.addWidget(lixeira_btn)
        self.main_layout.addWidget(header)

        # 2. Corpo Principal
        body_widget = QWidget()
        self.body_layout = QHBoxLayout(body_widget)
        self.body_layout.setContentsMargins(20, 20, 20, 20)
        
        # Coluna da Esquerda (Listas)
        left_column = QFrame()
        left_column.setFixedWidth(400)
        self.left_layout = QVBoxLayout(left_column)
        
        # Coluna da Direita (Cards)
        self.right_grid_layout = QGridLayout()
        self.right_grid_layout.setSpacing(20)

        self.body_layout.addWidget(left_column)
        self.body_layout.addLayout(self.right_grid_layout)
        self.main_layout.addWidget(body_widget)

    def atualizar_dados_e_ui(self):
        try:
            self.dados_df = carregar_e_mesclar_dados()
            self.desenhar_interface_com_dados()
        except Exception as e:
            self.mostrar_erro(str(e))

    def desenhar_interface_com_dados(self):
        # Limpa a interface antiga
        for i in reversed(range(self.left_layout.count())): 
            self.left_layout.itemAt(i).widget().setParent(None)
        for i in reversed(range(self.right_grid_layout.count())):
            self.right_grid_layout.itemAt(i).widget().setParent(None)

        # Filtra os dados
        df_ativos = self.dados_df[~self.dados_df['Status'].isin([STATUS_CONCLUIDO, STATUS_CANCELADO])]
        df_pendentes = self.dados_df[self.dados_df['Status'] == STATUS_PENDENTE]
        df_top_prioridades = df_ativos[df_ativos['Status'] != STATUS_PENDENTE]

        # Desenha a lista da esquerda
        self.desenhar_lista_esquerda(df_top_prioridades, df_pendentes)

        # Desenha os cards da direita
        posicoes = [(0, 0), (0, 1), (1, 0), (1, 1)]
        for i, (index, row) in enumerate(df_top_prioridades.head(4).iterrows()):
            if i < 4:
                card = self.criar_card(row)
                self.right_grid_layout.addWidget(card, posicoes[i][0], posicoes[i][1])

    def desenhar_lista_esquerda(self, top_prioridades, pendentes):
        # Top 5
        top5_label = QLabel("TOP 5 PRIORIDADES")
        top5_label.setObjectName("SectionLabel")
        top5_label.setFont(QFont("Inter", 16, QFont.Bold))
        self.left_layout.addWidget(top5_label)
        
        for index, row in top_prioridades.head(5).iterrows():
            item_label = QLabel(f"{row['Prioridade']}. {row['Pedido']}")
            item_label.setFont(QFont("Inter", 12))
            self.left_layout.addWidget(item_label)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setStyleSheet("color: #424242;")
        self.left_layout.addWidget(separator)

        # Pendentes
        pendentes_label = QLabel("PENDENTES")
        pendentes_label.setObjectName("SectionLabel")
        pendentes_label.setStyleSheet("color: #D4AC0D;")
        pendentes_label.setFont(QFont("Inter", 16, QFont.Bold))
        self.left_layout.addWidget(pendentes_label)

        for index, row in pendentes.iterrows():
            item_label = QLabel(f"P{row['Prioridade']}: {row['Pedido']}")
            item_label.setFont(QFont("Inter", 12))
            item_label.setStyleSheet("color: #888888;")
            self.left_layout.addWidget(item_label)
            
        self.left_layout.addStretch()

    def criar_card(self, data):
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        
        # Formatação dos dados
        data_entrega_str = data[COLUNA_DATA_PRIORIDADE].strftime('%d/%m/%Y')
        valor_formatado = locale.currency(data['Valor'], grouping=True, symbol=True)
        
        # Conteúdo do Card
        titulo = QLabel(f"{data['Prioridade']}º: {data['Pedido']}")
        titulo.setFont(QFont("Inter", 22, QFont.Bold))
        titulo.setObjectName("TitleLabel")

        status = QLabel(data['Status'].upper())
        status.setFont(QFont("Inter", 14, QFont.Bold))
        
        # Cores dos status
        if data['Status'] == STATUS_PENDENTE: status.setStyleSheet("color: #D4AC0D;")
        elif data['Status'] == STATUS_AGUARDANDO: status.setStyleSheet("color: #3498DB;")
        elif data['Status'] == STATUS_EM_MONTAGEM: status.setStyleSheet("color: #F39C12;")
        
        servico = QLabel(data['Servico'])
        servico.setWordWrap(True)
        servico.setFont(QFont("Inter", 12))
        
        # Grid para detalhes
        details_grid = QGridLayout()
        details_grid.addWidget(QLabel(f"<b>Entrega:</b> {data_entrega_str}"), 0, 0)
        details_grid.addWidget(QLabel(f"<b>Valor:</b> {valor_formatado}"), 0, 1)
        details_grid.addWidget(QLabel(f"<b>Cód. Item:</b> {data['CodItem']}"), 1, 0)
        details_grid.addWidget(QLabel(f"<b>QTD:</b> {data['Maquinas']}"), 1, 1)

        layout.addWidget(titulo)
        layout.addWidget(status)
        layout.addWidget(servico)
        layout.addStretch()
        layout.addLayout(details_grid)
        return card

    def mostrar_lixeira(self):
        df_cancelados = self.dados_df[self.dados_df['Status'] == STATUS_CANCELADO]
        dialog = LixeiraDialog(df_cancelados, self)
        dialog.exec()

    def mostrar_erro(self, mensagem):
        # Limpa a interface e mostra o erro
        for i in reversed(range(self.body_layout.count())): 
            item = self.body_layout.takeAt(i)
            if item.widget():
                item.widget().setParent(None)
            elif item.layout():
                 # Limpa sub-layouts se existirem
                while item.layout().count() > 0:
                    sub_item = item.layout().takeAt(0)
                    if sub_item.widget():
                        sub_item.widget().setParent(None)

        error_label = QLabel(f"ERRO AO CARREGAR DADOS:\n\n{mensagem}")
        error_label.setObjectName("ErrorLabel")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setFont(QFont("Inter", 18, QFont.Bold))
        self.body_layout.addWidget(error_label)
    
    def verificar_atualizacoes(self):
        try:
            mod_prio = os.path.getmtime(CAMINHO_PLANILHA_PRIORIDADES)
            mod_stat = os.path.getmtime(CAMINHO_PLANILHA_STATUS)
            
            if (self.timestamp_prioridades != 0 and mod_prio != self.timestamp_prioridades) or \
               (self.timestamp_status != 0 and mod_stat != self.timestamp_status):
                print("Detecção de mudança nas planilhas. Atualizando...")
                self.atualizar_dados_e_ui()
            
            self.timestamp_prioridades, self.timestamp_status = mod_prio, mod_stat
        except FileNotFoundError:
            pass # A função de erro já lida com isso visualmente

class LixeiraDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Lixeira - Pedidos Cancelados")
        self.setMinimumSize(800, 600)
        self.setStyleSheet(STYLESHEET)
        
        layout = QVBoxLayout(self)
        
        title = QLabel("Pedidos Cancelados")
        title.setObjectName("TitleLabel")
        title.setFont(QFont("Inter", 20, QFont.Bold))
        layout.addWidget(title)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        if data.empty:
            scroll_layout.addWidget(QLabel("A lixeira está vazia."))
        else:
            for index, row in data.iterrows():
                item_text = f"<b>{row['Pedido']}</b>: {row['Servico']}"
                label = QLabel(item_text)
                label.setWordWrap(True)
                scroll_layout.addWidget(label)
        
        scroll_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PainelMtec()
    window.show()
    sys.exit(app.exec())
