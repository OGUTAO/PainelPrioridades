import sys
import os
import time
import locale
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout,
                               QDialog, QPushButton)
from PySide6.QtGui import QPixmap, QFont
from PySide6.QtCore import QTimer, Qt

# --- CONFIGURAÇÃO GERAL ---
try:
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    script_dir = os.getcwd()

CAMINHO_PASTA_EXCEL = os.path.join(script_dir, "dados")
NOME_ARQUIVO_STATUS = "Status_dos_pedidos.xlsx"
CAMINHO_PLANILHA_STATUS = os.path.join(CAMINHO_PASTA_EXCEL, NOME_ARQUIVO_STATUS)

# --- Nomes das colunas ---
COLUNA_PEDIDO_ID_STATUS = 'Pedido'
COLUNA_STATUS = 'Status'

# --- Constantes de Status ---
STATUS_PENDENTE = 'Pendente'
STATUS_AGUARDANDO = 'Aguardando Montagem'
STATUS_AGUARDANDO_CHEGADA = 'Aguardando Chegada'
STATUS_EM_MONTAGEM = 'Em Montagem'
STATUS_CONCLUIDO = 'Concluído'
STATUS_CANCELADO = 'Cancelado'

INTERVALO_CHECK_MS = 3000

# --- LÓGICA DE DADOS (SIMPLIFICADA) ---
def carregar_e_mesclar_dados():
    if not os.path.exists(CAMINHO_PLANILHA_STATUS):
        raise FileNotFoundError(f"Arquivo de status não encontrado: {CAMINHO_PLANILHA_STATUS}")

    df_status = pd.read_excel(CAMINHO_PLANILHA_STATUS)
    df_status.columns = df_status.columns.str.strip()
    
    df_status.rename(columns={COLUNA_PEDIDO_ID_STATUS: 'Pedido'}, inplace=True)
    df_status['Pedido'] = df_status['Pedido'].astype(str)
    df_status = df_status[df_status['Pedido'].str.startswith('CV-')].copy()

    df_status.reset_index(drop=True, inplace=True)
    df_status['Prioridade'] = df_status.index + 1

    df_status['Servico'] = "Detalhe não encontrado"
    df_status['CodItem'] = "N/A"
    df_status['Maquinas'] = 0
    
    df_final = df_status[~df_status[COLUNA_STATUS].isin([STATUS_CONCLUIDO, STATUS_CANCELADO])].copy()

    return df_final

# --- INTERFACE GRÁFICA ---
STYLESHEET = """
    QMainWindow, QDialog { background-color: #1C1C1C; }
    QLabel { color: #FFFFFF; }
    #Header { background-color: #2E2E2E; }
    #TitleLabel { color: #FF6600; }
    #TopPrioLabel { color: #FF6600; border-bottom: 2px solid #FF6600; padding-bottom: 5px; margin-bottom: 5px;}
    #SectionLabel { color: #3498DB; border-bottom: 2px solid #3498DB; padding-bottom: 5px; margin-bottom: 5px;}
    #PendenteLabel { color: #D4AC0D; border-bottom: 2px solid #D4AC0D; padding-bottom: 5px; margin-bottom: 5px;}
    #AguardandoChegadaLabel { color: #9B59B6; border-bottom: 2px solid #9B59B6; padding-bottom: 5px; margin-bottom: 5px;}
    #Card {
        background-color: #2E2E2E; border: 1px solid #FF6600;
        border-radius: 10px; padding: 15px;
    }
    #HeaderButton {
        background-color: #424242; color: white;
        padding: 8px 16px; border-radius: 5px; border: none; font-weight: bold;
    }
    #HeaderButton:hover { background-color: #555555; }
    #ErrorLabel { color: #E74C3C; }
    QScrollArea { border: none; }
"""

class PainelMtec(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Painel de Produção MTEC")
        self.setGeometry(100, 100, 1600, 900)
        self.setStyleSheet(STYLESHEET)
        try:
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        except locale.Error:
            print("Aviso: Local 'pt_BR.UTF-8' não encontrado.")
        self.dados_df = pd.DataFrame()
        self.timestamp_status = 0
        self.setup_ui()
        self.atualizar_dados_e_ui()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.verificar_atualizacoes)
        self.timer.start(INTERVALO_CHECK_MS)

    def setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        header = QWidget()
        header.setObjectName("Header")
        header.setFixedHeight(70)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 0, 20, 0)
        logo_label = QLabel()
        logo_path = os.path.join(script_dir, "logoMtec.jpeg")
        pixmap = QPixmap(logo_path)
        if not pixmap.isNull():
            logo_label.setPixmap(pixmap.scaled(150, 50, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            logo_label.setText("MTEC")
            logo_label.setObjectName("TitleLabel")
            logo_label.setFont(QFont("Inter", 24, QFont.Bold))
        
        cancelados_btn = QPushButton("Cancelados")
        cancelados_btn.setObjectName("HeaderButton")
        cancelados_btn.setCursor(Qt.PointingHandCursor)
        cancelados_btn.clicked.connect(self.mostrar_cancelados)

        atualizar_btn = QPushButton("Atualizar")
        atualizar_btn.setObjectName("HeaderButton")
        atualizar_btn.setCursor(Qt.PointingHandCursor)
        atualizar_btn.clicked.connect(self.forcar_atualizacao)
        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        header_layout.addWidget(atualizar_btn)
        header_layout.addWidget(cancelados_btn)
        self.main_layout.addWidget(header)
        body_widget = QWidget()
        self.body_layout = QHBoxLayout(body_widget)
        self.body_layout.setContentsMargins(20, 20, 20, 20)
        scroll_area_left = QScrollArea()
        scroll_area_left.setWidgetResizable(True)
        scroll_area_left.setFixedWidth(400)
        self.scroll_content_left = QWidget()
        self.left_layout = QVBoxLayout(self.scroll_content_left)
        self.left_layout.setContentsMargins(10, 10, 10, 10)
        self.left_layout.setSpacing(10)
        scroll_area_left.setWidget(self.scroll_content_left)
        self.right_grid_layout = QGridLayout()
        self.right_grid_layout.setSpacing(20)
        self.body_layout.addWidget(scroll_area_left)
        self.body_layout.addLayout(self.right_grid_layout)
        self.main_layout.addWidget(body_widget)

    def forcar_atualizacao(self):
        print("Atualização manual solicitada.")
        self.atualizar_dados_e_ui()

    def atualizar_dados_e_ui(self):
        try:
            self.dados_df = carregar_e_mesclar_dados()
            self.desenhar_interface_com_dados()
        except Exception as e:
            self.mostrar_erro(str(e))

    def desenhar_interface_com_dados(self):
        while self.left_layout.count():
            child = self.left_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()
        while self.right_grid_layout.count():
            child = self.right_grid_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()
        
        df_pendentes = self.dados_df[self.dados_df['Status'] == STATUS_PENDENTE]
        df_aguardando = self.dados_df[self.dados_df['Status'] == STATUS_AGUARDANDO]
        df_aguardando_chegada = self.dados_df[self.dados_df['Status'] == STATUS_AGUARDANDO_CHEGADA]
        df_para_cards = self.dados_df[self.dados_df['Status'].isin([STATUS_AGUARDANDO, STATUS_EM_MONTAGEM])]
        
        self.desenhar_lista_esquerda(df_para_cards, df_pendentes, df_aguardando, df_aguardando_chegada)
        self.desenhar_cards_direita(df_para_cards)

    # << ALTERAÇÃO: Reordenada a exibição das listas
    def desenhar_lista_esquerda(self, df_para_cards, df_pendentes, df_aguardando, df_aguardando_chegada):
        # 1. PRIORIDADES
        top4_label = QLabel("PRIORIDADES")
        top4_label.setObjectName("TopPrioLabel")
        top4_label.setFont(QFont("Inter", 14, QFont.Bold))
        self.left_layout.addWidget(top4_label)

        if df_para_cards.empty:
            self.left_layout.addWidget(QLabel("Nenhuma prioridade para exibir."))
        else:
            for pos_lista, (index, row) in enumerate(df_para_cards.head(4).iterrows(), 1):
                item_label = QLabel(f"{pos_lista}º (P{row['Prioridade']}): {row['Pedido']}")
                item_label.setFont(QFont("Inter", 11))
                self.left_layout.addWidget(item_label)
        
        # 2. PENDENTES (Movido para cima)
        pendentes_label = QLabel("PENDENTES")
        pendentes_label.setObjectName("PendenteLabel")
        pendentes_label.setFont(QFont("Inter", 14, QFont.Bold))
        self.left_layout.addWidget(pendentes_label)
        if df_pendentes.empty:
            self.left_layout.addWidget(QLabel("Nenhum pedido pendente."))
        else:
            for index, row in df_pendentes.iterrows():
                item_label = QLabel(f"P{row['Prioridade']}: {row['Pedido']}")
                item_label.setFont(QFont("Inter", 11))
                self.left_layout.addWidget(item_label)

        # 3. AGUARDANDO MONTAGEM
        aguardando_label = QLabel("AGUARDANDO MONTAGEM")
        aguardando_label.setObjectName("SectionLabel")
        aguardando_label.setFont(QFont("Inter", 14, QFont.Bold))
        self.left_layout.addWidget(aguardando_label)
        if df_aguardando.empty:
            self.left_layout.addWidget(QLabel("Nenhum pedido aguardando montagem."))
        else:
            for index, row in df_aguardando.iterrows():
                item_label = QLabel(f"P{row['Prioridade']}: {row['Pedido']}")
                item_label.setFont(QFont("Inter", 11))
                self.left_layout.addWidget(item_label)

        # 4. AGUARDANDO CHEGADA
        aguardando_chegada_label = QLabel("AGUARDANDO CHEGADA")
        aguardando_chegada_label.setObjectName("AguardandoChegadaLabel")
        aguardando_chegada_label.setFont(QFont("Inter", 14, QFont.Bold))
        self.left_layout.addWidget(aguardando_chegada_label)
        if df_aguardando_chegada.empty:
            self.left_layout.addWidget(QLabel("Nenhum pedido aguardando chegada."))
        else:
            for index, row in df_aguardando_chegada.iterrows():
                item_label = QLabel(f"P{row['Prioridade']}: {row['Pedido']}")
                item_label.setFont(QFont("Inter", 11))
                self.left_layout.addWidget(item_label)
        
        self.left_layout.addStretch()

    def desenhar_cards_direita(self, df_para_cards):
        posicoes = [(0, 0), (0, 1), (1, 0), (1, 1)]
        if df_para_cards.empty:
            aviso = QLabel("Não há pedidos para exibir nos cards.")
            aviso.setFont(QFont("Inter", 16))
            aviso.setAlignment(Qt.AlignCenter)
            self.right_grid_layout.addWidget(aviso, 0, 0, 2, 2)
        else:
            for pos_lista, (index, row) in enumerate(df_para_cards.head(4).iterrows(), 1):
                card = self.criar_card(row, pos_lista)
                self.right_grid_layout.addWidget(card, posicoes[pos_lista-1][0], posicoes[pos_lista-1][1])

    def criar_card(self, data, pos_lista):
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        
        texto_titulo = f"{pos_lista}º (P{data['Prioridade']}): {data['Pedido']}"
        titulo = QLabel(texto_titulo)
        titulo.setFont(QFont("Inter", 20, QFont.Bold))
        titulo.setObjectName("TitleLabel")
        
        status_text = str(data.get('Status', 'N/A')).upper()
        status = QLabel(status_text)
        status.setFont(QFont("Inter", 14, QFont.Bold))
        if data['Status'] == STATUS_AGUARDANDO: status.setStyleSheet("color: #3498DB;")
        elif data['Status'] == STATUS_EM_MONTAGEM: status.setStyleSheet("color: #F39C12;")
        elif data['Status'] == STATUS_AGUARDANDO_CHEGADA: status.setStyleSheet("color: #9B59B6;")

        servico_texto = str(data.get('Servico', 'N/A'))
        servico = QLabel(servico_texto)
        servico.setWordWrap(True)
        servico.setFont(QFont("Inter", 12))
        
        details_grid = QGridLayout()
        details_grid.addWidget(QLabel(f"<b>Cód. Item:</b> {data.get('CodItem', 'N/A')}"), 0, 0)
        details_grid.addWidget(QLabel(f"<b>QTD:</b> {int(data.get('Maquinas', 0))}"), 0, 1)
        
        layout.addWidget(titulo)
        layout.addWidget(status)
        layout.addWidget(servico)
        layout.addStretch()
        layout.addLayout(details_grid)
        return card

    def mostrar_cancelados(self):
        try:
            df_status_completo = pd.read_excel(CAMINHO_PLANILHA_STATUS)
            df_status_completo.columns = df_status_completo.columns.str.strip()
            df_cancelados = df_status_completo[df_status_completo[COLUNA_STATUS] == STATUS_CANCELADO]
            dialog = LixeiraDialog(df_cancelados, self)
            dialog.exec()
        except Exception as e:
            self.mostrar_erro(f"Erro ao carregar pedidos cancelados:\n{e}")

    def mostrar_erro(self, mensagem):
        for i in reversed(range(self.body_layout.count())):
            item = self.body_layout.itemAt(i)
            widget = item.widget()
            if widget: widget.deleteLater()
            layout = item.layout()
            if layout:
                while layout.count():
                    sub_item = layout.takeAt(0)
                    if sub_item.widget(): sub_item.widget().deleteLater()
        error_label = QLabel(f"ERRO AO CARREGAR DADOS:\n\n{mensagem}")
        error_label.setObjectName("ErrorLabel")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setWordWrap(True)
        error_label.setFont(QFont("Inter", 18, QFont.Bold))
        self.body_layout.addWidget(error_label)
    
    def verificar_atualizacoes(self):
        try:
            if not os.path.exists(CAMINHO_PLANILHA_STATUS):
                return

            mod_stat = os.path.getmtime(CAMINHO_PLANILHA_STATUS)
            
            if self.timestamp_status != 0 and mod_stat != self.timestamp_status:
                print("Detecção de mudança na planilha de status. Atualizando...")
                self.forcar_atualizacao()
                
            self.timestamp_status = mod_stat
        except FileNotFoundError:
            pass

class LixeiraDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pedidos Cancelados")
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
            scroll_layout.addWidget(QLabel("Não há pedidos cancelados."))
        else:
            for index, row in data.iterrows():
                item_text = f"<b>{row.get('Pedido', 'N/A')}</b>"
                label = QLabel(item_text)
                label.setWordWrap(True)
                scroll_layout.addWidget(label)
        scroll_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PainelMtec()
    window.showMaximized()
    sys.exit(app.exec())