# Functions.py
# Funções utilitárias para a interface gráfica (GUI) do sistema OTC.
# Contém a lógica para integração do painel de commodities (preenchimento de combos, busca, limpeza),
# gerenciamento de temas, e conexão de eventos dos menus laterais e central.
# Todos os blocos estão comentados detalhadamente para facilitar a manutenção e entendimento.

from Custom_Widgets import *
from Custom_Widgets.QAppSettings import QAppSettings
from Custom_Widgets.QCustomTipOverlay import QCustomTipOverlay
from Custom_Widgets.QCustomLoadingIndicators import QCustom3CirclesLoader
from PySide6.QtCore import QSettings, QTimer, Qt, QSize
from PySide6.QtGui import QColor, QFont, QFontDatabase, QIcon
from PySide6.QtWidgets import (
    QGraphicsDropShadowEffect, QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QSizePolicy, QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QMessageBox
)
import src.commodity_search  # Certifique-se de que esse import está correto conforme estrutura do seu projeto

# Classe para representar um tema do app (apenas dois exemplos: Light e Dark)
class Theme:
    def __init__(self, name, defaultTheme, backgroundColor, textColor, accentColor, iconsColor, createNewIcons):
        self.name = name
        self.defaultTheme = defaultTheme
        self.backgroundColor = backgroundColor
        self.textColor = textColor
        self.accentColor = accentColor
        self.iconsColor = iconsColor
        self.createNewIcons = createNewIcons

class GuiFunctions():
    """
    Classe central para manipulação e controle dos widgets da interface principal.
    Conecta eventos, inicializa combos, gerencia temas e lida com a busca de commodities.
    """
    def __init__(self, MainWindow):
        # Instâncias principais da janela e da interface
        self.main = MainWindow
        self.ui = MainWindow.ui

        # Carrega e aplica a fonte Product Sans globalmente para toda a aplicação
        self.loadProductSansFont()
        # Inicializa tema (carrega config salva, popula lista de temas, conecta troca)
        self.initializeAppTheme()
        # Conecta eventos de busca global (search do topo)
        self.ui.searchBtn.clicked.connect(self.showSearchResults)
        self.ui.returnCenterMenuBtn.setVisible(False)
        # Conecta todos botões dos menus laterais/centrais
        self.connectMenuButtons()
        # Inicialização do painel de commodities
        self.inicializar_interface()

        # Inicializa referência do botão voltar (para controle seguro do ciclo de vida)
        self.backSearchCommoditiesBtn = None

    #########################
    # MÉTODOS DE MENU
    #########################
    def setCenterMenuLabel(self, text):
        """
        Atualiza o texto do label do menu central superior.
        O label é utilizado para indicar ao usuário em qual seção do sistema ele se encontra.
        """
        self.ui.label.setText(text)

    def connectMenuButtons(self):
        """
        Conecta todos os botões dos menus laterais/esquerdo/centro/direito às funções corretas,
        garantindo navegação entre páginas e abertura/fechamento dos menus deslizantes.
        Esta função garante toda a navegação do app e previne que múltiplas janelas sejam abertas.
        """
        # Menu lateral esquerdo (funções fixas)
        self.ui.homeBtn.clicked.connect(self.ui.centerMenu.collapseMenu)
        self.ui.settingsBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.settingsPage),
            self.setCenterMenuLabel("Settings"),
            self.ui.returnCenterMenuBtn.setVisible(False)
        ])
        self.ui.infoBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.informationPage),
            self.setCenterMenuLabel("Information"),
            self.ui.returnCenterMenuBtn.setVisible(False)
        ])
        self.ui.helpBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.helpPage),
            self.setCenterMenuLabel("Help"),
            self.ui.returnCenterMenuBtn.setVisible(False)
        ])
        # Commodities: botão "Search" chama o painel de busca central
        self.ui.commoditiesSearchBtn.clicked.connect(self.showCommoditiesSearchArea)

        # Botões principais/subpáginas (sempre escondem return)
        for btn, page, label in [
            (self.ui.ndfBtn, self.ui.ndfSubPage, "NDF"),
            (self.ui.optBtn, self.ui.opcaoSubPage, "Option"),
            (self.ui.swaptBtn, self.ui.swapSubPage, "Swap"),
            (self.ui.counterpartyBtn, self.ui.counterpartySubPage, "Counterparty"),
            (self.ui.commodiBtn, self.ui.commoditiesSubPage, "Commodities"),
            (self.ui.metricsBtn, self.ui.metricsSubPage, "Metrics"),
            (self.ui.confirmationsBtn, self.ui.confirmationsSubPage, "Confirmations"),
            (self.ui.intragBtn, self.ui.intragSubPage, "Intrag"),
        ]:
            btn.clicked.connect(lambda _, p=page, l=label: [
                self.ui.centerMenu.expandMenu(),
                self.ui.centerMenuPages.setCurrentWidget(p),
                self.ui.returnCenterMenuBtn.setVisible(False),
                self.setCenterMenuLabel(l)
            ])

        # Botões de registro levam para subsubpages e deixam return visível
        self.ui.ndfRegisterBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.ndfSubSubPage),
            self.ui.returnCenterMenuBtn.setVisible(True),
            self.setCenterMenuLabel("NDF > Register")
        ])
        self.ui.optRegisterBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.opcaoSubSubPage),
            self.ui.returnCenterMenuBtn.setVisible(True),
            self.setCenterMenuLabel("Option > Register")
        ])
        self.ui.swapRegisterBtn.clicked.connect(lambda: [
            self.ui.centerMenu.expandMenu(),
            self.ui.centerMenuPages.setCurrentWidget(self.ui.swapSubSubPage),
            self.ui.returnCenterMenuBtn.setVisible(True),
            self.setCenterMenuLabel("Swap > Register")
        ])
        self.ui.returnCenterMenuBtn.clicked.connect(self.handleReturnFromSubSubPage)
        self.ui.closeCenterMenuBtn.clicked.connect(lambda: self.ui.centerMenu.collapseMenu())

        # Demais subpáginas apenas mudam o label
        self.ui.ndfSettleBtn.clicked.connect(lambda: self.setCenterMenuLabel("NDF > Settlement"))
        self.ui.ndfSearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("NDF > Search"))
        self.ui.ndfNewBtn.clicked.connect(lambda: self.setCenterMenuLabel("NDF > Register > New Deals"))
        self.ui.ndfUnwindBtn.clicked.connect(lambda: self.setCenterMenuLabel("NDF > Register > Unwinds"))
        self.ui.optSettleBtn.clicked.connect(lambda: self.setCenterMenuLabel("Option > Settlement"))
        self.ui.optSearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("Option > Search"))
        self.ui.optNewBtn.clicked.connect(lambda: self.setCenterMenuLabel("Option > Register > New Deals"))
        self.ui.optUnwindBtn.clicked.connect(lambda: self.setCenterMenuLabel("Option > Register > Unwinds"))
        self.ui.swapSettleBtn.clicked.connect(lambda: self.setCenterMenuLabel("Swap > Settlement"))
        self.ui.swapSearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("Swap > Search"))
        self.ui.swapNewBtn.clicked.connect(lambda: self.setCenterMenuLabel("Swap > Register > New Deals"))
        self.ui.swapUnwindBtn.clicked.connect(lambda: self.setCenterMenuLabel("Swap > Register > Unwinds"))
        self.ui.kpiBtn.clicked.connect(lambda: self.setCenterMenuLabel("Metrics > KPI"))
        self.ui.adhocBtn.clicked.connect(lambda: self.setCenterMenuLabel("Metrics > Ad-Hoc"))
        self.ui.NDFConfirmationBtn.clicked.connect(lambda: self.setCenterMenuLabel("Confirmations > NDF"))
        self.ui.optionConfirmationBtn.clicked.connect(lambda: self.setCenterMenuLabel("Confirmations > Option"))
        self.ui.swapConfirmationBtn.clicked.connect(lambda: self.setCenterMenuLabel("Confirmations > Swap"))
        self.ui.intragSearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("Intrag > Search"))
        self.ui.intragBoletaBtn.clicked.connect(lambda: self.setCenterMenuLabel("Intrag > Boleta"))

        # Counterparty/Commodities sub-abas (caso existam)
        if hasattr(self.ui, "counterpartyRegisterBtn"):
            self.ui.counterpartyRegisterBtn.clicked.connect(lambda: self.setCenterMenuLabel("Counterparty > Register"))
        if hasattr(self.ui, "counterpartySearchBtn"):
            self.ui.counterpartySearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("Counterparty > Search"))
        if hasattr(self.ui, "commoditiesRegisterBtn"):
            self.ui.commoditiesRegisterBtn.clicked.connect(lambda: self.setCenterMenuLabel("Commodities > Register"))
        if hasattr(self.ui, "commoditiesSearchBtn"):
            self.ui.commoditiesSearchBtn.clicked.connect(lambda: self.setCenterMenuLabel("Commodities > Search"))

        # Menu lateral direito (Notificações, Mais, Perfil)
        self.ui.notificationBtn.clicked.connect(lambda: self.ui.rightMenu.expandMenu())
        self.ui.moreBtn.clicked.connect(lambda: self.ui.rightMenu.expandMenu())
        self.ui.profileBtn.clicked.connect(lambda: self.ui.rightMenu.expandMenu())
        self.ui.closeRightMenuBtn.clicked.connect(lambda: self.ui.rightMenu.collapseMenu())

    def showCommoditiesSearchArea(self):
        """
        Exibe o painel de busca de commodities NA MESMA PÁGINA principal da interface.
        - Troca para a página correta usando QStackedWidget principal (mainPages).
        - Garante que todos os widgets de filtro estejam visíveis.
        - Remove/destrói a tabela de resultados e o botão voltar, caso existam.
        - Preenche os combos com dados do JSON.
        """
        # Troca para a página principal do painel de commodities
        self.ui.mainPages.setCurrentWidget(self.ui.commoditiesSearchPage)

        # Garante que todos os widgets do filtro estejam visíveis
        self.showCommoditiesFilterWidgets()

        # Remove/destrói tabela de resultados e botão voltar se existirem
        self.remover_tabela_resultados()
        if self.backSearchCommoditiesBtn:
            layout = self.ui.verticalLayout_commoditiesSearchPage
            layout.removeWidget(self.backSearchCommoditiesBtn)
            self.backSearchCommoditiesBtn.deleteLater()
            self.backSearchCommoditiesBtn = None

        # Garante atualização dos combos com os dados JSON
        self.preencher_combos_commodities()

    def handleReturnFromSubSubPage(self):
        """
        Retorna da subsubpágina (cadastro) para a subpágina principal.
        Atualiza label e oculta botão de retorno.
        """
        current = self.ui.centerMenuPages.currentWidget()
        if current == self.ui.ndfSubSubPage:
            self.ui.centerMenuPages.setCurrentWidget(self.ui.ndfSubPage)
            self.setCenterMenuLabel("NDF")
        elif current == self.ui.opcaoSubSubPage:
            self.ui.centerMenuPages.setCurrentWidget(self.ui.opcaoSubPage)
            self.setCenterMenuLabel("Option")
        elif current == self.ui.swapSubSubPage:
            self.ui.centerMenuPages.setCurrentWidget(self.ui.swapSubPage)
            self.setCenterMenuLabel("Swap")
        self.ui.returnCenterMenuBtn.setVisible(False)

    #########################
    # PAINEL COMMODITIES: SETUP & INTERFACE
    #########################
    def inicializar_interface(self):
        """
        Método principal que configura o painel de busca de commodities:
        - Preenche combos (carregando do JSON)
        - Conecta botões (pesquisar/limpar)
        - Aplica estilos e tooltips
        """
        self.setupCommoditiesSearch()
        self.configurar_estilo()
        self.configurar_tooltips()

    def setupCommoditiesSearch(self):
        """
        Organiza a lógica de inicialização do painel de busca:
        - Preenche combos ao iniciar
        - Conecta botões de pesquisa e limpeza
        """
        self.preencher_combos_commodities()
        self.ui.searchFilterCommoditiesBtn.clicked.connect(self.pesquisar_commodities)
        self.ui.deletFilterCommoditiesBtn.clicked.connect(self.limpar_commodities)

    def preencher_combos_commodities(self):
        """
        Preenche todos os ComboBoxes do painel de busca de commodities
        com as opções extraídas do JSON base de commodities.
        Garante que os filtros estejam sempre atualizados para o usuário.
        """
        base = src.commodity_search.load_commodities_base('/Users/giullianoaccarinideluccia/Desktop/Cortex OTC/json-cache/commodities_base.json')
        if not base:
            print("Aviso: Nenhuma commodity foi carregada. Verifique o arquivo JSON commodities_base.json.")
            return

        # Preenche ComboBox Código (campo "Ativo Subjacente")
        self.ui.combo_codigo.clear()
        self.ui.combo_codigo.addItem("Selecione o código")
        opcoes_codigo = src.commodity_search.get_combo_options(base, "Ativo Subjacente")
        for opcao in opcoes_codigo:
            if opcao:
                self.ui.combo_codigo.addItem(str(opcao))

        # Preenche ComboBox Bolsa
        self.ui.combo_bolsa.clear()
        self.ui.combo_bolsa.addItem("Selecione a bolsa")
        opcoes_bolsa = src.commodity_search.get_combo_options(base, "Bolsa")
        for opcao in opcoes_bolsa:
            if opcao:
                self.ui.combo_bolsa.addItem(str(opcao))

        # Preenche ComboBox Mercadoria
        self.ui.combo_commodity.clear()
        self.ui.combo_commodity.addItem("Selecione a Mercadoria")
        opcoes_mercadoria = src.commodity_search.get_combo_options(base, "Mercadoria")
        for opcao in opcoes_mercadoria:
            if opcao:
                self.ui.combo_commodity.addItem(str(opcao))

        # Preenche ComboBox Mês
        self.ui.combo_mes.clear()
        self.ui.combo_mes.addItem("Selecione o mês")
        opcoes_mes = src.commodity_search.get_combo_options(base, "Mes Vencimento")
        meses_numericos = []
        for opcao in opcoes_mes:
            try:
                mes_int = int(opcao)
                meses_numericos.append(mes_int)
            except (ValueError, TypeError):
                continue
        meses_numericos.sort()
        for mes in meses_numericos:
            self.ui.combo_mes.addItem(str(mes))

        # Preenche ComboBox Ano
        self.ui.combo_ano.clear()
        self.ui.combo_ano.addItem("Selecione o ano")
        opcoes_ano = src.commodity_search.get_combo_options(base, "Ano Vencimento")
        for opcao in opcoes_ano:
            if opcao:
                self.ui.combo_ano.addItem(str(opcao))

        # Preenche ComboBox Status
        self.ui.combo_status.clear()
        self.ui.combo_status.addItem("Selecione o status")
        opcoes_status = src.commodity_search.get_combo_options(base, "Status")
        for opcao in opcoes_status:
            if opcao:
                self.ui.combo_status.addItem(str(opcao))

    def pesquisar_commodities(self):
        """
        Executa a busca de commodities conforme filtros dos combos e exibe resultados na tabela.
        Quando o botão "Pesquisar" é clicado, todos os widgets de filtro somem e a tabela ocupa toda a página,
        com um botão 'Voltar' logo abaixo para restaurar os filtros e nova busca.
        """
        # Coleta filtros dos combos (pegando os textos dos widgets de filtro)
        filtros = {
            "Ativo Subjacente": self.ui.combo_codigo.currentText() if self.ui.combo_codigo.currentIndex() > 0 else "",
            "Bolsa": self.ui.combo_bolsa.currentText() if self.ui.combo_bolsa.currentIndex() > 0 else "",
            "Mercadoria": self.ui.combo_commodity.currentText() if self.ui.combo_commodity.currentIndex() > 0 else "",
            "Mes Vencimento": int(self.ui.combo_mes.currentText()) if self.ui.combo_mes.currentIndex() > 0 else "",
            "Ano Vencimento": int(self.ui.combo_ano.currentText()) if self.ui.combo_ano.currentIndex() > 0 else "",
            "Status": self.ui.combo_status.currentText() if self.ui.combo_status.currentIndex() > 0 else "",
        }
        # Remove filtros vazios
        filtros = {k: v for k, v in filtros.items() if v != ""}

        # Busca na base de dados (lê o arquivo JSON de commodities)
        base = src.commodity_search.load_commodities_base('/Users/giullianoaccarinideluccia/Desktop/Cortex OTC/json-cache/commodities_base.json')
        resultados = src.commodity_search.search_commodities(base, filtros)

        # Esconde todos os widgets de filtro e botões do painel de busca
        self.hideCommoditiesFilterWidgets()
        # Exibe resultados na página inteira, com botão voltar
        self.exibir_tabela_resultados(resultados, fullscreen=True)

    def exibir_tabela_resultados(self, dados, fullscreen=False):
        """
        Exibe os resultados da busca de commodities na área apropriada da interface.
        Não cria layouts ou botões -- apenas usa os já definidos no UI.
        """

        # Remove qualquer tabela antiga
        self.remover_tabela_resultados()

        # Se não há dados, mostra aviso e retorna
        if not dados:
            self.ui.widget_commoditiesTableArea.show()
            msg = QLabel("Nenhum resultado encontrado.")
            msg.setObjectName("tabela_msg")
            msg.setStyleSheet("color: #900; font-weight: bold; padding: 12px;")
            self.ui.vbox_commoditiesTableArea.insertWidget(0, msg)
            self.tabela_msg = msg
            return

        # Cria a nova tabela de resultados e preenche
        colunas = list(dados[0].keys())
        tabela = QTableWidget(len(dados), len(colunas), self.ui.widget_commoditiesTableArea)
        tabela.setObjectName("tabela_resultados")
        tabela.setHorizontalHeaderLabels(colunas)
        tabela.setSelectionBehavior(QTableWidget.SelectRows)
        tabela.setEditTriggers(QTableWidget.NoEditTriggers)
        tabela.setAlternatingRowColors(True)
        tabela.setStyleSheet("QTableWidget { background: #fff; color: #333; font-size: 12px; }")
        tabela.horizontalHeader().setStretchLastSection(True)
        tabela.verticalHeader().setVisible(False)
        tabela.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Preenche cada célula: None vira "", exibe sempre como string
        for i, linha in enumerate(dados):
            for j, chave in enumerate(colunas):
                valor = linha.get(chave, "")
                if valor is None:
                    valor = ""
                item = QTableWidgetItem(str(valor))
                item.setTextAlignment(Qt.AlignCenter)
                tabela.setItem(i, j, item)

        # Ajusta larguras das colunas
        header = tabela.horizontalHeader()
        for coluna in range(len(colunas)):
            header.setSectionResizeMode(coluna, QHeaderView.ResizeToContents)

        # Adiciona a tabela no topo do layout da área de resultados (ANTES dos botões)
        self.ui.vbox_commoditiesTableArea.insertWidget(0, tabela)
        self.tabela_resultados = tabela

        # Mostra a área de resultados (isso exibe a tabela + botões)
        self.ui.widget_commoditiesTableArea.show()

        # Conecta os botões (se não estiverem conectados)
        self.ui.backSearchCommoditiesBtn.clicked.disconnect()
        self.ui.backSearchCommoditiesBtn.clicked.connect(self.onVoltarParaFiltros)
        self.ui.exportCommoditiesSearchBtn.clicked.disconnect()
        self.ui.exportCommoditiesSearchBtn.clicked.connect(lambda: self.exportar_resultado_busca(dados))

        
    def exportar_resultado_busca(self, dados):
        """
        Exporta os resultados exibidos na tabela para um arquivo CSV na pasta Downloads.
        O nome do arquivo segue o padrão:
        - Search_Commodities(Bolsa=ICE NYBOT; Ano=2026).csv
        - Search_Commodities.csv (caso nenhum filtro tenha sido utilizado)
        Os filtros utilizados são obtidos do painel de filtros (comboboxes).
        O CSV será exportado utilizando ponto e vírgula (;) como separador.
        """
        import os
        import csv
        from datetime import datetime

        # Lista para armazenar filtros ativos na busca (ex: "Bolsa=ICE NYBOT").
        filtros = []
        # Lista para armazenar os nomes dos filtros ativos (não usado, mas mantido do código original).
        nome_filtros = []

        # Dicionário mapeando o nome do filtro (como aparece na interface) para o widget ComboBox correspondente.
        # É daqui que serão extraídas as opções selecionadas pelo usuário na interface gráfica.
        combo_filtros = {
            "Código": self.ui.combo_codigo,
            "Bolsa": self.ui.combo_bolsa,
            "Mercadoria": self.ui.combo_commodity,
            "Mês": self.ui.combo_mes,
            "Ano": self.ui.combo_ano,
            "Status": self.ui.combo_status,
        }

        # Percorre todos os combos e extrai apenas os filtros que estão ativos (ou seja, diferentes do placeholder "Selecione ...").
        for nome, combo in combo_filtros.items():
            valor = combo.currentText()
            if valor and not valor.lower().startswith("selecione"):
                filtros.append(f"{nome}={valor}")
                nome_filtros.append(f"{nome}={valor}")

        # Monta o nome do arquivo conforme os filtros utilizados.
        # Exemplo: Search_Commodities(Bolsa=ICE NYBOT; Ano=2026).csv
        if filtros:
            filtros_str = "; ".join(filtros) # separa os filtros por "; " para refletir o CSV
            nome_arquivo = f"Search_Commodities({filtros_str}).csv"
        else:
            nome_arquivo = "Search_Commodities.csv"

        # Remove caracteres inválidos para nome de arquivo em Windows, Linux ou Mac.
        nome_arquivo = "".join(c for c in nome_arquivo if c not in r'\/:*?"<>|')

        # Descobre o caminho absoluto para a pasta Downloads do usuário, de forma multiplataforma.
        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo)

        # Só prossegue se houver dados para exportar.
        if not dados:
            return

        # Extrai os nomes das colunas a partir do primeiro registro dos dados.
        # Isso garantirá que o cabeçalho do CSV corresponda à estrutura da tabela.
        colunas = list(dados[0].keys())
        try:
            # Abre (ou cria) o arquivo CSV no caminho determinado, com codificação UTF-8.
            # newline='' evita linhas em branco extras em alguns ambientes.
            with open(caminho_arquivo, mode="w", encoding="utf-8", newline="") as f:
                # Cria o writer do CSV, especificando o delimitador como ponto e vírgula.
                writer = csv.DictWriter(f, fieldnames=colunas, delimiter=';')
                writer.writeheader()      # Escreve o cabeçalho do CSV.
                writer.writerows(dados)   # Escreve cada linha de dados.

            # Exibe mensagem de sucesso ao usuário após salvar o arquivo.
            QMessageBox.information(
                self.ui.commoditiesSearchPage,
                "Exportação realizada",
                f"Arquivo salvo em:\n{caminho_arquivo}"
            )
        except Exception as e:
            # Em caso de erro ao exportar, exibe mensagem de erro detalhada ao usuário.
            QMessageBox.critical(
                self.ui.commoditiesSearchPage,
                "Erro ao exportar",
                f"Erro ao exportar arquivo:\n{str(e)}"
            )

    def onVoltarParaFiltros(self):
        """
        Ao clicar em 'Voltar':
        - Remove a tabela de resultados e o botão voltar.
        - Mostra novamente todos os widgets de filtro no layout.
        - Preenche combos de filtro com dados atualizados.
        """
        self.remover_tabela_resultados()
        if self.backSearchCommoditiesBtn:
            layout = self.ui.verticalLayout_commoditiesSearchPage
            layout.removeWidget(self.backSearchCommoditiesBtn)
            self.backSearchCommoditiesBtn.deleteLater()
            self.backSearchCommoditiesBtn = None

        # Mostra widgets de filtro novamente
        self.showCommoditiesFilterWidgets()
        # Preenche combos para garantir atualização
        self.preencher_combos_commodities()

    def hideCommoditiesFilterWidgets(self):
        """
        Esconde todos os widgets de filtro do painel (sem remover do layout).
        Essencial para garantir que ao mostrar a tabela, os filtros não sejam removidos do layout.
        """
        if hasattr(self.ui, "searchCommoditiesWidget"):
            self.ui.searchCommoditiesWidget.setVisible(False)
        else:
            widgets_to_hide = [
                "label_commodities_title", "label_codigo", "combo_codigo",
                "label_bolsa", "combo_bolsa", "label_commodity", "combo_commodity",
                "label_mes", "combo_mes", "label_ano", "combo_ano",
                "label_status", "combo_status", "searchFilterCommoditiesBtn", "deletFilterCommoditiesBtn"
            ]
            for name in widgets_to_hide:
                if hasattr(self.ui, name):
                    getattr(self.ui, name).setVisible(False)

    def showCommoditiesFilterWidgets(self):
        # O painel correto é searchCommoditiesWidget
        if hasattr(self.ui, "searchCommoditiesWidget"):
            self.ui.searchCommoditiesWidget.setVisible(True)
        else:
            widgets_to_show = [
                "label_searchHeader", "label_codigo", "combo_codigo",
                "label_bolsa", "combo_bolsa", "label_commodity", "combo_commodity",
                "label_mes", "combo_mes", "label_ano", "combo_ano",
                "label_status", "combo_status", "searchFilterCommoditiesBtn", "deletFilterCommoditiesBtn"
            ]
            for name in widgets_to_show:
                if hasattr(self.ui, name):
                    getattr(self.ui, name).setVisible(True)

    def remover_tabela_resultados(self):
        """
        Remove qualquer tabela de resultados ou mensagens antigas do painel de commodities.
        Não remove widgets de filtro do layout, apenas a tabela e mensagens.
        """
        if hasattr(self, "tabela_msg") and self.tabela_msg:
            self.ui.verticalLayout_commoditiesSearchPage.removeWidget(self.tabela_msg)
            self.tabela_msg.deleteLater()
            self.tabela_msg = None
        if hasattr(self, "tabela_resultados") and self.tabela_resultados:
            self.ui.verticalLayout_commoditiesSearchPage.removeWidget(self.tabela_resultados)
            self.tabela_resultados.deleteLater()
            self.tabela_resultados = None

    def limpar_commodities(self):
        """
        Limpa todos os filtros dos combos e remove a tabela de resultados do painel de commodities.
        Também faz os widgets de filtro aparecerem novamente (caso a tabela esteja fullscreen).
        """
        self.ui.combo_codigo.setCurrentIndex(0)
        self.ui.combo_bolsa.setCurrentIndex(0)
        self.ui.combo_commodity.setCurrentIndex(0)
        self.ui.combo_mes.setCurrentIndex(0)
        self.ui.combo_ano.setCurrentIndex(0)
        self.ui.combo_status.setCurrentIndex(0)
        self.remover_tabela_resultados()
        self.showCommoditiesFilterWidgets()

    #########################
    # ESTILOS E TOOLTIP
    #########################
    def configurar_estilo(self):
        """
        Aplica estilos visuais customizados ao painel de busca de commodities, como sombra,
        cor de fundo, e outros efeitos gráficos para melhor UX.
        """
        sombra = QGraphicsDropShadowEffect()
        sombra.setBlurRadius(12)
        sombra.setColor(QColor(0, 0, 0, 70))
        sombra.setOffset(0, 4)
        self.ui.searchCommoditiesWidget.setGraphicsEffect(sombra)

    def configurar_tooltips(self):
        """
        Sets informative tooltips for each field in the commodity search panel.
        These tooltips help the user understand the purpose of each ComboBox filter.
        """
        self.ui.combo_codigo.setToolTip("Choose the code of the underlying asset for the commodity.")
        self.ui.combo_bolsa.setToolTip("Select the exchange where the asset is traded.")
        self.ui.combo_commodity.setToolTip("Select the commodity associated with the asset.")
        self.ui.combo_mes.setToolTip("Choose the expiration month of the contract.")
        self.ui.combo_ano.setToolTip("Choose the expiration year of the contract.")
        self.ui.combo_status.setToolTip("Filter contracts by the selected status.")

    #########################
    # GERENCIAMENTO DE TEMA
    #########################
    def loadProductSansFont(self):
        """
        Carrega a fonte Product Sans do arquivo na pasta resources/fonts, 
        tornando-a disponível para toda a aplicação.
        """
        font_id = QFontDatabase.addApplicationFont("resources/fonts/ProductSans-Regular.ttf")
        if font_id != -1:
            families = QFontDatabase.applicationFontFamilies(font_id)
            if families:
                app_font = QFont(families[0])
                app_font.setPointSize(10)
                self.main.setFont(app_font)
        else:
            print("Fonte Product Sans não encontrada!")

    def initializeAppTheme(self):
        """
        Inicializa o tema da aplicação a partir das configurações (QSettings).
        Preenche a lista de temas disponíveis e conecta o evento de troca de tema.
        """
        settings = QSettings()  # Acessa as configurações persistentes do app
        current_theme = settings.value("THEME")  # Recupera o tema atual salvo

        self.populateThemeList(current_theme)
        self.ui.themeList.currentTextChanged.connect(self.changeAppTheme)

    def populateThemeList(self, current_theme):
        """
        Preenche a lista dropdown de temas disponíveis no app.
        Seleciona o tema padrão ou o tema atualmente salvo.
        """
        self.ui.themeList.clear()
        themes = [
            Theme("Light", True, "#f0ffff", "#000000", "#7fffd4", "#000080", True),
            Theme("Dark", False, "#21272a", "#fefefe", "#fba43b", "#fefefe", True)
        ]
        self.ui.themes = themes
        for i, theme in enumerate(themes):
            self.ui.themeList.addItem(theme.name, theme.name)
            if theme.defaultTheme or theme.name == current_theme:
                self.ui.themeList.setCurrentIndex(i)

    def changeAppTheme(self):
        """
        Troca o tema do app conforme a seleção do usuário na lista.
        Atualiza as configurações e recarrega as preferências.
        """
        settings = QSettings()
        selected_theme = self.ui.themeList.currentData()  # Nome do tema selecionado
        current_theme = settings.value("THEME")  # Tema salvo atualmente

        if current_theme != selected_theme:
            settings.setValue("THEME", selected_theme)
            QAppSettings.updateAppSettings(self.main, reloadJson=True)
        if hasattr(self.main.theme, "reloadJsonStyles"):
            self.main.theme.reloadJsonStyles(update=True)

    #########################
    # BUSCA GLOBAL
    #########################
    def showSearchResults(self):
        """
        Exibe resultados da busca global (campo search do topo), 
        pode ser customizado para buscar em todas as entidades do sistema.
        """
        termo = self.ui.searchEdit.text().strip()
        if not termo:
            QCustomTipOverlay.showTip(self.ui.searchEdit, "Digite um termo para buscar.", 3000)
            return
        loader = QCustom3CirclesLoader(self.ui.centralwidget)
        loader.setGeometry(self.ui.searchEdit.geometry())
        loader.show()
        QTimer.singleShot(1000, loader.close)
        QTimer.singleShot(
            1200,
            lambda: QCustomTipOverlay.showTip(self.ui.searchEdit, f"Busca por '{termo}' não implementada!", 3500),
        )