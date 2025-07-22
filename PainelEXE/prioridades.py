import sys
import os
import locale
import random
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QLabel, QFrame, QProgressBar)
from PySide6.QtGui import QFont
from PySide6.QtCore import QTimer, Qt

# --- CONFIGURAÇÃO GERAL E DE DADOS ---
META_SEMANAL = 500 # A meta agora é baseada na QUANTIDADE DE MÁQUINAS

FRASES_MOTIVACIONAIS = [
    "A qualidade do nosso trabalho hoje é a garantia do nosso sucesso amanhã.", "O único lugar onde o sucesso vem antes do trabalho é no dicionário.",
    "Grandes coisas em negócios nunca são feitas por uma pessoa. São feitas por uma equipe.", "A persistência realiza o impossível.",
    "Foco, força e fé: os três pilares para um dia produtivo.", "A perfeição não é alcançável, mas se buscarmos a perfeição, podemos alcançar a excelência.",
    "O talento vence jogos, mas o trabalho em equipe ganha campeonatos.", "Não observe o relógio; faça o que ele faz. Continue em frente.",
    "A disciplina é a ponte entre metas e realizações.", "Sua dedicação de hoje está construindo a reputação de amanhã."
]

FRASE_DO_DIA_ATUAL = ""; ULTIMO_DIA_FRASE = None

try:
    if getattr(sys, 'frozen', False): script_dir = os.path.dirname(sys.executable)
    else: script_dir = os.path.dirname(os.path.abspath(__file__))
except NameError: script_dir = os.getcwd()

CAMINHO_PASTA_EXCEL = os.path.join(script_dir, "dados")
NOME_ARQUIVO_STATUS = "Status_dos_pedidos.xlsm"
CAMINHO_PLANILHA_STATUS = os.path.join(CAMINHO_PASTA_EXCEL, NOME_ARQUIVO_STATUS)

COLUNA_PEDIDO_ID, COLUNA_PV, COLUNA_SERVICO, COLUNA_STATUS, COLUNA_DATA_STATUS, COLUNA_QTD = 'Pedido', 'PV', 'Servico', 'Status', 'Data Status', 'Qtd Maquinas'
STATUS_PENDENTE, STATUS_AGUARDANDO, STATUS_AGUARDANDO_CHEGADA, STATUS_EM_MONTAGEM, STATUS_CONCLUIDO, STATUS_CANCELADO = 'Pendente', 'Aguardando Montagem', 'Aguardando Chegada', 'Em Montagem', 'Concluído', 'Cancelado'
INTERVALO_CHECK_MS = 5000

# --- LÓGICA DE DADOS ---
def carregar_dados():
    if not os.path.exists(CAMINHO_PLANILHA_STATUS):
        raise FileNotFoundError(f"Arquivo de dados não encontrado: {CAMINHO_PLANILHA_STATUS}")
    df = pd.read_excel(CAMINHO_PLANILHA_STATUS, parse_dates=[COLUNA_DATA_STATUS])
    df.columns = df.columns.str.strip()
    
    for col, default_val in [(COLUNA_PV, "TERAVIX"), (COLUNA_SERVICO, "Detalhe não disponível"), (COLUNA_QTD, 0)]:
        if col not in df.columns: df[col] = default_val
        else:
            if col == COLUNA_QTD:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = df[col].astype(object).fillna(default_val)
    df[COLUNA_QTD] = df[COLUNA_QTD].astype(int)

    df[COLUNA_PEDIDO_ID] = df[COLUNA_PEDIDO_ID].astype(str)
    df = df[df[COLUNA_PEDIDO_ID].str.startswith('CV-')].copy()
    df.reset_index(drop=True, inplace=True); df['Prioridade'] = df.index + 2
    df_principal = df[~df[COLUNA_STATUS].isin([STATUS_CONCLUIDO, STATUS_CANCELADO])].copy()
    hoje = datetime.now().date()
    df_concluidos_hoje = df[(df[COLUNA_STATUS] == STATUS_CONCLUIDO) & (pd.to_datetime(df[COLUNA_DATA_STATUS]).dt.date == hoje)].sort_values(by=COLUNA_DATA_STATUS, ascending=False)
    df_cancelados_hoje = df[(df[COLUNA_STATUS] == STATUS_CANCELADO) & (pd.to_datetime(df[COLUNA_DATA_STATUS]).dt.date == hoje)].sort_values(by=COLUNA_DATA_STATUS, ascending=False)
    
    teravix_concluidos = len(df_concluidos_hoje[df_concluidos_hoje[COLUNA_PV] == 'TERAVIX'])
    pv_concluidos = len(df_concluidos_hoje[df_concluidos_hoje[COLUNA_PV] != 'TERAVIX'])
    total_concluidos = len(df_concluidos_hoje)
    teravix_concluidos_qtd = df_concluidos_hoje[df_concluidos_hoje[COLUNA_PV] == 'TERAVIX'][COLUNA_QTD].sum()
    pv_concluidos_qtd = df_concluidos_hoje[df_concluidos_hoje[COLUNA_PV] != 'TERAVIX'][COLUNA_QTD].sum()
    total_concluidos_qtd = df_concluidos_hoje[COLUNA_QTD].sum()

    teravix_cancelados = len(df_cancelados_hoje[df_cancelados_hoje[COLUNA_PV] == 'TERAVIX'])
    pv_cancelados = len(df_cancelados_hoje[df_cancelados_hoje[COLUNA_PV] != 'TERAVIX'])
    total_cancelados = len(df_cancelados_hoje)
    teravix_cancelados_qtd = df_cancelados_hoje[df_cancelados_hoje[COLUNA_PV] == 'TERAVIX'][COLUNA_QTD].sum()
    pv_cancelados_qtd = df_cancelados_hoje[df_cancelados_hoje[COLUNA_PV] != 'TERAVIX'][COLUNA_QTD].sum()
    total_cancelados_qtd = df_cancelados_hoje[COLUNA_QTD].sum()

    return df, df_principal, df_concluidos_hoje, df_cancelados_hoje, \
           (teravix_concluidos, pv_concluidos, total_concluidos, teravix_concluidos_qtd, pv_concluidos_qtd, total_concluidos_qtd), \
           (teravix_cancelados, pv_cancelados, total_cancelados, teravix_cancelados_qtd, pv_cancelados_qtd, total_cancelados_qtd)

def obter_frase_do_dia():
    global FRASE_DO_DIA_ATUAL, ULTIMO_DIA_FRASE
    hoje = datetime.now().date()
    if ULTIMO_DIA_FRASE != hoje: FRASE_DO_DIA_ATUAL = random.choice(FRASES_MOTIVACIONAIS); ULTIMO_DIA_FRASE = hoje
    return FRASE_DO_DIA_ATUAL

def calcular_metricas_dashboard(df_full):
    hoje = datetime.now()
    inicio_mes_atual = hoje.replace(day=1, hour=0, minute=0, second=0)
    df_concluidos_mes_atual = df_full[(df_full[COLUNA_STATUS] == STATUS_CONCLUIDO) & (pd.to_datetime(df_full[COLUNA_DATA_STATUS]) >= inicio_mes_atual) & (pd.to_datetime(df_full[COLUNA_DATA_STATUS]) <= hoje)]
    total_mes_atual_pedidos = len(df_concluidos_mes_atual)
    total_mes_atual_qtd = df_concluidos_mes_atual[COLUNA_QTD].sum()
    dias_uteis_mes_atual = np.busday_count(inicio_mes_atual.strftime('%Y-%m-%d'), (hoje + timedelta(days=1)).strftime('%Y-%m-%d'))
    media_diaria_atual = total_mes_atual_pedidos / dias_uteis_mes_atual if dias_uteis_mes_atual > 0 else 0
    media_diaria_qtd = total_mes_atual_qtd / dias_uteis_mes_atual if dias_uteis_mes_atual > 0 else 0
    
    fim_mes_anterior = inicio_mes_atual - timedelta(days=1); inicio_mes_anterior = fim_mes_anterior.replace(day=1)
    df_concluidos_mes_anterior = df_full[(df_full[COLUNA_STATUS] == STATUS_CONCLUIDO) & (pd.to_datetime(df_full[COLUNA_DATA_STATUS]) >= inicio_mes_anterior) & (pd.to_datetime(df_full[COLUNA_DATA_STATUS]) <= fim_mes_anterior)]
    total_mes_anterior = len(df_concluidos_mes_anterior)
    dias_uteis_mes_anterior = np.busday_count(inicio_mes_anterior.strftime('%Y-%m-%d'), (fim_mes_anterior + timedelta(days=1)).strftime('%Y-%m-%d'))
    media_diaria_anterior = len(df_concluidos_mes_anterior) / dias_uteis_mes_anterior if dias_uteis_mes_anterior > 0 else 0
    
    recorde_dia_valor = 0; recorde_dia_data = ""; recorde_dia_qtd = 0
    if not df_concluidos_mes_atual.empty:
        producao_diaria = df_concluidos_mes_atual.groupby(pd.to_datetime(df_concluidos_mes_atual[COLUNA_DATA_STATUS]).dt.date).size()
        if not producao_diaria.empty:
            recorde_dia_valor = producao_diaria.max()
            recorde_dia_data_obj = producao_diaria.idxmax()
            recorde_dia_data = recorde_dia_data_obj.strftime('%d/%m')
            recorde_dia_qtd = df_concluidos_mes_atual[pd.to_datetime(df_concluidos_mes_atual[COLUNA_DATA_STATUS]).dt.date == recorde_dia_data_obj][COLUNA_QTD].sum()

    return {"total_mes_atual": total_mes_atual_pedidos, "total_mes_atual_qtd": total_mes_atual_qtd, "media_diaria_atual": media_diaria_atual, "media_diaria_qtd": media_diaria_qtd,
            "total_mes_anterior": total_mes_anterior, "media_diaria_anterior": media_diaria_anterior,
            "recorde_dia_valor": recorde_dia_valor, "recorde_dia_data": recorde_dia_data, "recorde_dia_qtd": recorde_dia_qtd}

def calcular_dados_grafico(df_full):
    df_concluidos = df_full.dropna(subset=[COLUNA_DATA_STATUS]).copy()
    df_concluidos = df_concluidos[df_concluidos[COLUNA_STATUS] == STATUS_CONCLUIDO].copy()
    if df_concluidos.empty: return []
    df_concluidos['Semana'] = pd.to_datetime(df_concluidos[COLUNA_DATA_STATUS]).dt.to_period('W-SUN').dt.start_time
    semanal = df_concluidos.groupby('Semana')[COLUNA_QTD].sum()
    semanas_recentes = pd.date_range(end=datetime.now(), periods=4, freq='W-MON').normalize()
    semanal = semanal.reindex(semanas_recentes, fill_value=0)
    return list(semanal.items())

STYLESHEET = """
    QMainWindow { background-color: #1C1C1C; } QLabel { color: #E0E0E0; }
    #Header { background-color: #2E2E2E; border-bottom: 2px solid #FF6600; }
    #LogoLabel { padding: 5px; } .SectionTitle { border-bottom: 2px solid; padding-bottom: 8px; margin-bottom: 10px; }
    #PrioridadesTitle, #PendentesTitle, #AguardandoMontagemTitle, #AguardandoChegadaTitle { color: #FF6600; border-bottom-color: #FF6600; }
    #ConcluidosTitle { color: #2ECC71; border-bottom-color: #2ECC71; } #CanceladosTitle { color: #E74C3C; border-bottom-color: #E74C3C; }
    #CounterLabel { color: #888888; font-style: italic; padding-top: 10px; }
    #SideColumnFrame { background-color: #252525; border-radius: 8px; } #ErrorLabel { color: #E74C3C; }
    #Card { background-color: #2E2E2E; border: 1px solid #FF6600; border-radius: 8px; padding: 12px; margin-bottom: 10px; }
    #CardTitle { color: #FF8C33; } #CardStatus_Aguardando { color: #3498DB; } #CardStatus_EmMontagem { color: #F39C12; }
    #TotalLabel { color: #BDBDBD; padding-top: 15px; border-top: 1px solid #444444; margin-top: 10px; }
    #DashboardFrame { border-top: 1px solid #444; margin-top: 10px; padding: 10px; }
    #MetricaTitle, #KpiTitle { color: #FFFFFF; font-weight: bold; } #MetricaValue, #KpiValue { color: #FF6600; }
    #FraseMotivacional { color: #DDD; font-style: italic; } #KpiRecorde { color: #3498DB; }
    QProgressBar { border: 1px solid #555; border-radius: 5px; text-align: center; background-color: #2E2E2E; }
    QProgressBar::chunk { background-color: #FF6600; border-radius: 4px; }
    QProgressBar#currentWeek::chunk { background-color: #FFAA33; }
"""

class PainelMtec(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Painel de Produção MTEC"); self.setGeometry(100, 100, 1920, 1080); self.setStyleSheet(STYLESHEET)
        self.main_container = QWidget(); self.error_container = QWidget(); self.is_showing_error = False
        self.setup_ui()
        self.timer = QTimer(self); self.timer.timeout.connect(self.atualizar_dados_e_ui)
        self.atualizar_dados_e_ui(); self.timer.start(INTERVALO_CHECK_MS)

    def setup_ui(self):
        self.central_widget = QWidget(); self.setCentralWidget(self.central_widget); layout = QVBoxLayout(self.central_widget); layout.setContentsMargins(0,0,0,0); layout.setSpacing(0)
        main_layout = QVBoxLayout(self.main_container); main_layout.setContentsMargins(0, 0, 0, 0); main_layout.setSpacing(0)
        header = QWidget(); header.setObjectName("Header"); header.setFixedHeight(60); header_layout = QHBoxLayout(header); header_layout.setContentsMargins(20, 0, 20, 0)
        logo_label = QLabel("mtec."); logo_label.setObjectName("LogoLabel"); logo_label.setFont(QFont("Inter", 22, QFont.Bold)); header_layout.addWidget(logo_label); header_layout.addStretch(); main_layout.addWidget(header)
        self.body_widget = QWidget(); self.body_layout = QHBoxLayout(self.body_widget); self.body_layout.setContentsMargins(15, 15, 15, 15); self.body_layout.setSpacing(20); main_layout.addWidget(self.body_widget, 1)
        dashboard_frame = QFrame(); dashboard_frame.setObjectName("DashboardFrame"); dashboard_frame.setFixedHeight(280); self.dashboard_layout = QHBoxLayout(dashboard_frame); main_layout.addWidget(dashboard_frame)
        self.setup_ui_columns()
        error_page_layout = QVBoxLayout(self.error_container); self.error_label = QLabel(); self.error_label.setObjectName("ErrorLabel"); self.error_label.setAlignment(Qt.AlignCenter); self.error_label.setWordWrap(True); self.error_label.setFont(QFont("Inter", 18, QFont.Bold)); error_page_layout.addWidget(self.error_label)
        layout.addWidget(self.main_container); layout.addWidget(self.error_container); self.error_container.hide()
        
    def setup_ui_columns(self):
        self.limpar_layout(self.body_layout); self.limpar_layout(self.dashboard_layout)
        self.prioridades_layout = QVBoxLayout(); self.pendentes_layout = QVBoxLayout(); self.aguardando_montagem_layout = QVBoxLayout(); self.aguardando_chegada_layout = QVBoxLayout()
        side_column_frame = QFrame(); side_column_frame.setObjectName("SideColumnFrame"); side_column_frame.setFixedWidth(300)
        self.side_layout = QVBoxLayout(side_column_frame); self.concluidos_layout = QVBoxLayout(); self.cancelados_layout = QVBoxLayout()
        self.side_layout.addLayout(self.concluidos_layout); self.side_layout.addSpacing(20); self.side_layout.addLayout(self.cancelados_layout); self.side_layout.addStretch()
        self.body_layout.addLayout(self.prioridades_layout, 2); self.body_layout.addLayout(self.aguardando_montagem_layout, 1); self.body_layout.addLayout(self.aguardando_chegada_layout, 1); self.body_layout.addLayout(self.pendentes_layout, 1); self.body_layout.addWidget(side_column_frame)
        self.metricas_layout = QVBoxLayout(); self.grafico_layout = QVBoxLayout(); self.kpi_layout = QVBoxLayout()
        self.dashboard_layout.addLayout(self.metricas_layout, 1); self.dashboard_layout.addLayout(self.grafico_layout, 2); self.dashboard_layout.addStretch(1); self.dashboard_layout.addLayout(self.kpi_layout, 1)

    def atualizar_dados_e_ui(self):
        try:
            df_full, df_principal, df_concluidos, df_cancelados, totais_concluidos, totais_cancelados = carregar_dados()
            if self.is_showing_error: self.clear_error_message()
            self.desenhar_colunas(df_principal, df_concluidos, df_cancelados, totais_concluidos, totais_cancelados)
            metricas = calcular_metricas_dashboard(df_full)
            dados_grafico = calcular_dados_grafico(df_full)
            frase = obter_frase_do_dia()
            self.desenhar_dashboard(metricas, dados_grafico, frase)
        except Exception as e:
            self.mostrar_erro(str(e))
    
    def desenhar_colunas(self, df_principal, df_concluidos, df_cancelados, totais_concluidos, totais_cancelados):
        font_titulo = QFont("Inter", 16, QFont.Bold); font_item = QFont("Inter", 11); font_contador = QFont("Inter", 10); font_total = QFont("Inter", 9)
        df_prioridades = df_principal[df_principal[COLUNA_STATUS].isin([STATUS_AGUARDANDO, STATUS_EM_MONTAGEM])]
        pedidos_em_prioridade_ids = df_prioridades.head(4)[COLUNA_PEDIDO_ID].tolist()
        
        self.desenhar_cards_prioridade(self.prioridades_layout, df_prioridades, font_titulo)
        df_pendentes = df_principal[df_principal[COLUNA_STATUS] == STATUS_PENDENTE]
        self.desenhar_lista_vertical(self.pendentes_layout, df_pendentes, "PENDENTES", font_titulo, font_item, font_contador)
        
        df_aguardando_base = df_principal[df_principal[COLUNA_STATUS] == STATUS_AGUARDANDO]
        df_aguardando_filtrado = df_aguardando_base[~df_aguardando_base[COLUNA_PEDIDO_ID].isin(pedidos_em_prioridade_ids)]
        self.desenhar_lista_vertical(self.aguardando_montagem_layout, df_aguardando_filtrado, "AGUARDANDO MONTAGEM", font_titulo, font_item, font_contador)

        df_aguardando_chegada = df_principal[df_principal[COLUNA_STATUS] == STATUS_AGUARDANDO_CHEGADA]
        self.desenhar_lista_vertical(self.aguardando_chegada_layout, df_aguardando_chegada, "AGUARDANDO CHEGADA", font_titulo, font_item, font_contador)
        
        self.desenhar_lista_lateral(self.concluidos_layout, df_concluidos, "CONCLUÍDOS DO DIA", font_titulo, font_item, font_contador, font_total, totais_concluidos, limit=5)
        self.desenhar_lista_lateral(self.cancelados_layout, df_cancelados, "CANCELADOS DO DIA", font_titulo, font_item, font_contador, font_total, totais_cancelados, limit=5)

    def desenhar_lista_lateral(self, layout, df, titulo_texto, font_titulo, font_item, font_contador, font_total, totais, limit=None):
        self.limpar_layout(layout); object_name = f"{titulo_texto.replace(' ', '')}Title"; layout.addWidget(self.criar_titulo(titulo_texto, object_name, font_titulo))
        if df.empty: layout.addWidget(QLabel("Nenhum."))
        else:
            df_display = df.head(limit) if limit is not None else df
            for _, row in df_display.iterrows():
                texto = f"<b>{row[COLUNA_PEDIDO_ID]}</b> ({row[COLUNA_PV]}) <font color='#2ECC71'>\"{row[COLUNA_QTD]}\"</font>"
                label = QLabel(texto); label.setFont(font_item); layout.addWidget(label)
            if limit is not None and len(df) > limit:
                restantes = len(df) - limit; contador_label = QLabel(f"+{restantes}..."); contador_label.setObjectName("CounterLabel"); contador_label.setFont(font_contador); layout.addWidget(contador_label)
        teravix, pv, total, teravix_qtd, pv_qtd, total_qtd = totais
        texto_total = f"TERAVIX: {teravix} ({teravix_qtd}) | PV: {pv} ({pv_qtd})<br><b>TOTAL DIA: {total} ({total_qtd})</b>"
        total_label = QLabel(texto_total); total_label.setObjectName("TotalLabel"); total_label.setFont(font_total); layout.addWidget(total_label)
        layout.addStretch(1)

    def desenhar_dashboard(self, metricas, dados_grafico, frase_do_dia):
        self.limpar_layout(self.metricas_layout); self.limpar_layout(self.grafico_layout); self.limpar_layout(self.kpi_layout)
        titulo_metrica_font = QFont("Inter", 14, QFont.Bold); valor_metrica_font = QFont("Inter", 38, QFont.Bold)
        total_mes_titulo = QLabel("Total Concluído no Mês"); total_mes_titulo.setObjectName("MetricaTitle"); total_mes_titulo.setFont(titulo_metrica_font)
        total_mes_valor = QLabel(f"{metricas['total_mes_atual']:.0f} <font color='#999' style='font-size:18px;'>({metricas['total_mes_atual_qtd']:.0f} máquinas)</font>"); total_mes_valor.setObjectName("MetricaValue"); total_mes_valor.setFont(valor_metrica_font)
        media_diaria_titulo = QLabel("Média Diária no Mês"); media_diaria_titulo.setObjectName("MetricaTitle"); media_diaria_titulo.setFont(titulo_metrica_font)
        media_diaria_valor = QLabel(f"{metricas['media_diaria_atual']:.1f} <font color='#999' style='font-size:18px;'>({metricas['media_diaria_qtd']:.1f} máquinas)</font>"); media_diaria_valor.setObjectName("MetricaValue"); media_diaria_valor.setFont(valor_metrica_font)
        self.metricas_layout.addWidget(total_mes_titulo); self.metricas_layout.addWidget(total_mes_valor); self.metricas_layout.addStretch(1); self.metricas_layout.addWidget(media_diaria_titulo); self.metricas_layout.addWidget(media_diaria_valor); self.metricas_layout.addStretch(1)
        titulo_grafico = QLabel(f"Desempenho Semanal (Meta: {META_SEMANAL} máquinas)"); titulo_grafico.setFont(titulo_metrica_font); self.grafico_layout.addWidget(titulo_grafico)
        start_of_current_week = datetime.now().date() - timedelta(days=datetime.now().weekday())
        for data, valor in dados_grafico:
            fim_semana = data + timedelta(days=6); texto_semana = f"Semana {data.strftime('%d/%m')} a {fim_semana.strftime('%d/%m')}"
            is_current_week = data.date() == start_of_current_week
            if is_current_week: texto_semana = f"<b>▶ {texto_semana}</b>"
            label_semana = QLabel(f"{texto_semana}: <b>{int(valor)}</b>"); label_semana.setFont(QFont("Inter", 10))
            progress_bar = QProgressBar(); progress_bar.setRange(0, META_SEMANAL); progress_bar.setValue(min(int(valor), META_SEMANAL)); progress_bar.setTextVisible(False); progress_bar.setFixedHeight(18); progress_bar.setMaximumWidth(550)
            if is_current_week: progress_bar.setObjectName("currentWeek")
            self.grafico_layout.addWidget(label_semana); self.grafico_layout.addWidget(progress_bar)
        self.grafico_layout.addStretch()
        kpi_titulo_font = QFont("Inter", 12, QFont.Bold); kpi_valor_font = QFont("Inter", 16, QFont.Bold)
        frase_titulo = QLabel("Frase do Dia"); frase_titulo.setObjectName("KpiTitle"); frase_titulo.setFont(kpi_titulo_font)
        frase_texto = QLabel(f'"{frase_do_dia}"'); frase_texto.setObjectName("FraseMotivacional"); frase_texto.setWordWrap(True); frase_texto.setFont(QFont("Inter", 11, italic=True))
        self.kpi_layout.addWidget(frase_titulo); self.kpi_layout.addWidget(frase_texto); self.kpi_layout.addStretch(1)
        recorde_titulo = QLabel("Recorde de Produção do Mês"); recorde_titulo.setObjectName("KpiTitle"); recorde_titulo.setFont(kpi_titulo_font)
        recorde_texto = QLabel(f"🏆 <font color='#3498DB'>{metricas['recorde_dia_valor']} pedidos ({metricas['recorde_dia_qtd']} máq.)</font> em {metricas['recorde_dia_data']}"); recorde_texto.setFont(kpi_valor_font)
        self.kpi_layout.addWidget(recorde_titulo); self.kpi_layout.addWidget(recorde_texto); self.kpi_layout.addStretch(1)
        comp_titulo = QLabel("Comparativo Mensal"); comp_titulo.setObjectName("KpiTitle"); comp_titulo.setFont(kpi_titulo_font)
        comp_texto_str = (f"Produção (Mês Atual): <b>{metricas['total_mes_atual']:.0f}</b> (vs. {metricas['total_mes_anterior']:.0f} Mês Anterior)<br>"
                          f"Média Diária (Mês Atual): <b>{metricas['media_diaria_atual']:.1f}</b> (vs. {metricas['media_diaria_anterior']:.1f} Mês Anterior)")
        comp_texto = QLabel(comp_texto_str); comp_texto.setFont(QFont("Inter", 11))
        self.kpi_layout.addWidget(comp_titulo); self.kpi_layout.addWidget(comp_texto); self.kpi_layout.addStretch(2)

    def limpar_layout(self, layout):
        if layout is None: return
        while layout.count():
            item = layout.takeAt(0); widget = item.widget()
            if widget: widget.deleteLater()
            else:
                layout_item = item.layout()
                if layout_item: self.limpar_layout(layout_item)
            
    def criar_titulo(self, texto, object_name, font):
        titulo_label = QLabel(texto); titulo_label.setObjectName(object_name); titulo_label.setProperty("class", "SectionTitle"); titulo_label.setFont(font)
        return titulo_label

    def desenhar_cards_prioridade(self, layout, df, font_titulo):
        self.limpar_layout(layout); layout.addWidget(self.criar_titulo("PRIORIDADES", "PrioridadesTitle", font_titulo))
        if df.empty: layout.addWidget(QLabel("Nenhuma prioridade para exibir."))
        else:
            for index, (_, row) in enumerate(df.head(4).iterrows()):
                card = self.criar_card_widget(row, index + 1); layout.addWidget(card)
        layout.addStretch()

    def criar_card_widget(self, data, pos_lista):
        card = QFrame(); card.setObjectName("Card"); layout = QVBoxLayout(card); layout.setSpacing(4)
        titulo_texto = f"<b>{pos_lista}º (P{data['Prioridade']}): {data[COLUNA_PEDIDO_ID]}</b> ({data[COLUNA_PV]})"
        titulo = QLabel(titulo_texto); titulo.setObjectName("CardTitle"); titulo.setFont(QFont("Inter", 12, QFont.Bold))
        status = QLabel(str(data.get(COLUNA_STATUS, 'N/A')).upper()); status.setFont(QFont("Inter", 10, QFont.Bold))
        if data[COLUNA_STATUS] == STATUS_AGUARDANDO: status.setObjectName("CardStatus_Aguardando")
        elif data[COLUNA_STATUS] == STATUS_EM_MONTAGEM: status.setObjectName("CardStatus_EmMontagem")
        servico = QLabel(str(data.get(COLUNA_SERVICO, 'N/A'))); servico.setWordWrap(True); servico.setFont(QFont("Inter", 10))
        qtd = QLabel(f"<b>QTD. MÁQUINAS:</b> {data[COLUNA_QTD]}"); qtd.setFont(QFont("Inter", 10, QFont.Bold)); qtd.setStyleSheet("color: #2ECC71;")
        layout.addWidget(titulo); layout.addWidget(status); layout.addSpacing(5); layout.addWidget(servico); layout.addStretch(); layout.addWidget(qtd)
        return card

    def desenhar_lista_vertical(self, layout, df, titulo_texto, font_titulo, font_item, font_contador):
        self.limpar_layout(layout); object_name = f"{titulo_texto.replace(' ', '')}Title"; layout.addWidget(self.criar_titulo(titulo_texto, object_name, font_titulo))
        if df.empty: layout.addWidget(QLabel("Nenhum pedido para exibir."))
        else:
            for _, row in df.head(5).iterrows():
                texto = f"<b>P{row['Prioridade']}: {row[COLUNA_PEDIDO_ID]}</b> ({row[COLUNA_PV]}) <font color='#2ECC71'>\"{row[COLUNA_QTD]}\"</font>"
                label = QLabel(texto); label.setFont(font_item); layout.addWidget(label)
            if len(df) > 5:
                restantes = len(df) - 5; contador_label = QLabel(f"+{restantes} pedidos..."); contador_label.setObjectName("CounterLabel"); contador_label.setFont(font_contador); layout.addWidget(contador_label)
        layout.addStretch()

    def mostrar_erro(self, mensagem):
        print(f"ERRO CRÍTICO: {mensagem}")
        if not self.is_showing_error:
            self.is_showing_error = True
            self.error_label.setText(f"ERRO AO ATUALIZAR DADOS:\n\n{mensagem}\n\nVerifique a planilha ou a conexão. Tentando novamente...")
            self.main_container.hide(); self.error_container.show()
    
    def clear_error_message(self):
        self.is_showing_error = False; self.error_container.hide(); self.main_container.show(); self.setup_ui_columns()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    try: locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    except locale.Error: print("Aviso: Local 'pt_BR.UTF-8' não pôde ser definido.")
    window = PainelMtec()
    window.showMaximized()
    sys.exit(app.exec())