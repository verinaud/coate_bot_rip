import estilos as estilo
import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QListWidget, QListWidgetItem, QComboBox, QHBoxLayout, \
    QHBoxLayout, QLineEdit, QDesktopWidget, QWidget, QVBoxLayout, QLabel, QPushButton, QMainWindow, QAction
from PyQt5.QtCore import QThread
from PyQt5.QtGui import QFont
from coate_bot import MyApp
from time import sleep
from configuracao import Config
from registro_log import Logger
import pyperclip
from manual import Manual
import os
import subprocess


class EventosThread(QThread):
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.evento = None
        
        
    def run(self):
        
        while self.parent.flag_thread :
            if self.parent.coate_app.texto_eventos != None :
                
                self.parent.info_evento(self.parent.coate_app.texto_eventos)
                self.parent.coate_app.texto_eventos = None
                
            sleep(0.1)         

        print("terminou eventos")

class CoateThread(QThread):
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
                

    def run(self):       
        
        self.parent.coate_app.start(self.parent.config)  # Passa o config como parâmetro para o método iniciar da classe MyApp
        
        self.parent.button_iniciar.setText("Iniciar")
        self.parent.button_iniciar.setEnabled(True)

        self.parent.button_iniciar2.setText("Iniciar")
        self.parent.button_iniciar2.setEnabled(True)     


class UI(QMainWindow) :  # inserir QMainWindow no parâmetro
    
    def __init__(self) :
        super(UI, self).__init__()       
        
        self.configuracao = Config()
        
        self.config = self.configuracao.get_config()
        
        self.criar_manual()
         
        self.log = Logger(self.config["diretorio_log"], None)       
        
        self.lista_exception = []      
        
        self.initUI()   
            
                
    def iniciar_coate(self):
        self.log.salvar_log("--")
        self.log.salvar_log("--")
        
        self.flag_thread = True      
        
        self.config["ultimo_acesso_user"] = self.login_input.text()
        
        self.config["ultimo_orgao"] = self.unidade_combo_box.currentText()      

        self.senha = self.password_input.text()
        
        if str(self.login_input.text()) == "" or str(self.password_input.text()) == "":
            self.aviso("Aviso!", "Usuário ou senha SEI não foi informado!")
        else:
            self.aviso("Atenção","Assim que solicitado, selecione o certificado e digite o código PIN.")

            self.configuracao.set_config(self.config)
        
            self.config["senha"] = self.password_input.text() #A senha digitada é setada no config porém não altera o config para que não fique gravada
        
            self.coate_app = MyApp()
                                 
            self.coate_thread = CoateThread(self)
            self.coate_thread.start()  # Inicie a thread para executar o programa coate            

            self.eventos_thread = EventosThread(self)       
            self.eventos_thread.start()
            
            self.button_iniciar.setText("Aguarde...")
            self.button_iniciar.setEnabled(False)

            self.button_iniciar2.setText("Aguarde...")
            self.button_iniciar2.setEnabled(False)
        
            self.show_password_button.setEnabled(False)
        
            self.show_password_button.setText("Exibir")
            self.password_input.setEchoMode(QLineEdit.Password)        
        
            self.showPanel4()
            
    def criar_manual(self):
        Manual()
        print()
        
    def abrir_manual(self):
        self.criar_manual()
        
        if os.name == 'nt':  # Verifica se o sistema operacional é Windows
            os.startfile("Manual-CoateBotRip.pdf")
        elif os.name == 'posix':  # Verifica se o sistema operacional é Unix/Linux/Mac
            subprocess.call(['xdg-open', nome_arquivo])
        else:
            print("Sistema operacional não suportado.")
        
                    
    def initUI(self) :
        # Criar um único painel para ambos os widgets
        self.panel_container = QWidget(self)

        # Criar 8 widgets (painéis)
        self.panel1 = QWidget(self.panel_container)

        self.panel2 = QWidget(self.panel_container)
        self.panel2.setMaximumHeight(200)

        self.panel3 = QWidget(self.panel_container)
        self.panel3.setMaximumHeight(200)
        
        self.panel4 = QWidget(self.panel_container)
        self.panel4.setMaximumHeight(350)       
        
        self.panel5 = QWidget(self.panel_container)
        self.panel5.setMaximumHeight(200)
        
        self.panel6 = QWidget(self.panel_container)
        self.panel6.setMaximumHeight(200)

        self.panel7 = QWidget(self.panel_container)
        self.panel7.setMaximumHeight(200)
        
        self.panel8 = QWidget(self.panel_container)
        self.panel8.setMaximumHeight(350)

        # Adicionar widgets ao panel1
        # Campo de Login
        self.login_label_m = QLabel("Usuário e senha SEI.", self.panel1)
        
        self.login_label = QLabel("Usuário:", self.panel1)
        self.login_input = QLineEdit(self.config["ultimo_acesso_user"], self.panel1)

        # Campo de senha
        self.password_label = QLabel("Senha:", self.panel1)
        self.password_input = QLineEdit(self.panel1)
        self.password_input.setEchoMode(QLineEdit.Password)

        self.show_password_button = QPushButton("Exibir", self.panel1)
        self.show_password_button.setCheckable(True)
        self.show_password_button.clicked.connect(self.toggle_password_visibility)
        self.show_password_button.setFixedWidth(100)

        # Campo de Órgão
        self.unidade_label = QLabel("Unidade:", self.panel1)

        # Criar um QComboBox
        self.unidade_combo_box = QComboBox(self.panel1)

        # Adicionar itens ao QComboBox
        self.unidade_combo_box.addItem(self.config["ultimo_orgao"])
        lista_unidade = self.config["lista_orgaos"]
        for unidades in lista_unidade:
             if self.unidade_combo_box.currentText() == unidades:
                print()
             else:
              self.unidade_combo_box.addItem(unidades)

        #botão iniciar
        self.button_iniciar = QPushButton('Iniciar', self.panel1)
        self.button_iniciar.clicked.connect(self.iniciar_coate)


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        self.novo_marcador_label = QLabel("Digite um novo marcador:", self.panel2)
        self.novo_marcador_input = QLineEdit(self.panel2)

        self.novo_marcador_button = QPushButton('Inserir', self.panel2)
        self.novo_marcador_button.clicked.connect(self.inserir_marcador)
        self.novo_marcador_button.setFixedWidth(100)

        self.excluir_marcador_label = QLabel("Selecione um marcador para excluir:", self.panel2)

        # Criar um QComboBox
        self.excluir_marcador_combobox = QComboBox(self.panel2)

        # Adicionar itens ao QComboBox
        lista_marcador = self.config["lista_marcadores_substitutos"]
        for marcadores in lista_marcador:
             self.excluir_marcador_combobox.addItem(marcadores)

        self.excluir_marcador_button = QPushButton('Excluir', self.panel2)
        self.excluir_marcador_button.clicked.connect(self.excluir_marcador)
        self.excluir_marcador_button.setFixedWidth(100)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        # Adicionar widgets ao panel3
        self.novo_vinculo_label = QLabel("Digite novo vínculo Decipex:", self.panel3)
        self.novo_vinculo_input = QLineEdit(self.panel3)

        self.novo_vinculo_button = QPushButton('Inserir', self.panel3)
        self.novo_vinculo_button.clicked.connect(self.inserir_vinculo)
        self.novo_vinculo_button.setFixedWidth(100)

        self.excluir_vinculo_label = QLabel("Selecione um vínculo para excluir:", self.panel3)

        # Criar um QComboBox
        self.excluir_vinculo_combobox = QComboBox(self.panel3)

        # Adicionar itens ao QComboBox
        lista_vinculo = self.config["vinculos_decipex"]
        for vinculos in lista_vinculo:
             self.excluir_vinculo_combobox.addItem(vinculos)

        self.excluir_vinculo_button = QPushButton('Excluir', self.panel3)
        self.excluir_vinculo_button.clicked.connect(self.excluir_vinculo)
        self.excluir_vinculo_button.setFixedWidth(100)

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        # Criação de um QLabel
        self.painel_eventos_label = QLabel("Painel de Eventos")

        # Criação do QListWidget
        self.lista_widget = QListWidget(self)
        
        self.lista_widget.itemClicked.connect(self.abrir_janela_item)

        #botão iniciar
        self.button_iniciar2 = QPushButton('Iniciar', self.panel4)
        self.button_iniciar2.clicked.connect(self.iniciar_coate)


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        self.novo_excluido_batimento_label = QLabel("Digite novo código para Excluído Batimento:", self.panel5)
        self.novo_excluido_batimento_input = QLineEdit(self.panel5)
        
        self.novo_excluido_batimento_button = QPushButton('Inserir', self.panel5)
        self.novo_excluido_batimento_button.clicked.connect(self.inserir_excluido_batimento)
        self.novo_excluido_batimento_button.setFixedWidth(100)
        
        self.excluir_excluido_batimento_label = QLabel("Selecione um código para excluir:", self.panel5)
        
        # Criar um QComboBox
        self.excluir_excluido_batimento_combobox = QComboBox(self.panel5)

        # Adicionar itens ao QComboBox
        lista_excluido_batimento = self.config["lista_excluido_batimento"]
        for batimentos in lista_excluido_batimento:
             self.excluir_excluido_batimento_combobox.addItem(batimentos)

        self.excluir_excluido_batimento_button = QPushButton('Excluir', self.panel5)
        self.excluir_excluido_batimento_button.clicked.connect(self.excluir_excluido_batimento)
        self.excluir_excluido_batimento_button.setFixedWidth(100)
        

        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        self.novo_excluido_falecimento_label = QLabel("Digite novo código para Excluído Falecimento:")
        self.novo_excluido_falecimento_input = QLineEdit(self.panel6)
        
        self.novo_excluido_falecimento_button = QPushButton('Inserir', self.panel6)
        self.novo_excluido_falecimento_button.clicked.connect(self.inserir_excluido_falecimento)
        self.novo_excluido_falecimento_button.setFixedWidth(100)
        
        self.excluir_excluido_falecimento_label = QLabel("Selecione um código para excluir:", self.panel6)
        
        # Criar um QComboBox
        self.excluir_excluido_falecimento_combobox = QComboBox(self.panel6)

        # Adicionar itens ao QComboBox
        lista_excluido_falecimento = self.config["lista_excluido_falecimento"]
        for falecimentos in lista_excluido_falecimento:
             self.excluir_excluido_falecimento_combobox.addItem(falecimentos)

        self.excluir_excluido_falecimento_button = QPushButton('Excluir', self.panel6)
        self.excluir_excluido_falecimento_button.clicked.connect(self.excluir_excluido_falecimento)
        self.excluir_excluido_falecimento_button.setFixedWidth(100)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        self.novo_marcador_substituto_label = QLabel("Digite novo marcador substituto:")
        self.novo_marcador_substituto_input = QLineEdit(self.panel7)
        
        
        self.novo_marcador_substituto_button = QPushButton('Inserir', self.panel7)
        self.novo_marcador_substituto_button.clicked.connect(self.inserir_marcador_substituto)
        self.novo_marcador_substituto_button.setFixedWidth(100)
        
        
        self.excluir_marcador_substituto_label = QLabel("Selecione um marcador para excluir:", self.panel7)
        
        # Criar um QComboBox
        self.excluir_marcador_substituto_combobox = QComboBox(self.panel7)

        # Adicionar itens ao QComboBox
        lista_marcador_substituto = self.config["lista_marcadores_substitutos"]
        for marcadores_substitutos in lista_marcador_substituto:
             self.excluir_marcador_substituto_combobox.addItem(marcadores_substitutos)

        self.excluir_marcador_substituto_button = QPushButton('Excluir', self.panel7)
        self.excluir_marcador_substituto_button.clicked.connect(self.excluir_marcador_substituto)
        self.excluir_marcador_substituto_button.setFixedWidth(100)
        
               
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        self.novo_url_sei_label = QLabel("Alterar URL SEI:")
        self.novo_url_sei_input = QLineEdit(self.panel8)
        self.novo_url_sei_input.setText(self.config["url_sei"])
        
        self.novo_url_siape_label = QLabel("Alterar URL SiapeNet:")
        self.novo_url_siape_input = QLineEdit(self.panel8)
        self.novo_url_siape_input.setText(self.config["url_siapenet"])
        
        self.novo_unidade_sei_label = QLabel("Alterar Unidade Sei:")
        self.novo_unidade_sei_input = QLineEdit(self.panel8)
        self.novo_unidade_sei_input.setText(self.config["unidade_sei"])        
        
        self.novo_outros_button = QPushButton('Alterar', self.panel8)
        self.novo_outros_button.clicked.connect(self.alterar_outros)
        self.novo_outros_button.setFixedWidth(100)        
               
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        # Criar layouts verticais para os painéis
        layout1 = QVBoxLayout(self.panel1)
        layout2 = QVBoxLayout(self.panel2)
        layout3 = QVBoxLayout(self.panel3)
        layout4 = QVBoxLayout(self.panel4)
        layout5 = QVBoxLayout(self.panel5)
        layout6 = QVBoxLayout(self.panel6)
        layout7 = QVBoxLayout(self.panel7)
        layout8 = QVBoxLayout(self.panel8)

        # Adicionar os widgets aos layouts dos painéis
        layout1.addWidget(self.login_label_m)
        layout1.addWidget(self.login_label)
        layout1.addWidget(self.login_input)
        layout1.addSpacing(10)
        layout1.addWidget(self.password_label)
        layoutSenha = QHBoxLayout(self.panel1)
        layoutSenha.addWidget(self.password_input)
        layoutSenha.addWidget(self.show_password_button)
        layout1.addLayout(layoutSenha)
        layout1.addWidget(self.unidade_label)
        layout1.addWidget(self.unidade_combo_box)
        layout1.addSpacing(70)
        layout1.addWidget(self.button_iniciar)

        layout2.addWidget(self.novo_marcador_label)
        lauout_add_marcador = QHBoxLayout(self.panel2)
        lauout_add_marcador.addWidget(self.novo_marcador_input)
        lauout_add_marcador.addWidget(self.novo_marcador_button)
        layout2.addLayout(lauout_add_marcador)
        layout2.addSpacing(70)
        layout2.addWidget(self.excluir_marcador_label)
        layout_exc_marcador = QHBoxLayout(self.panel2)
        layout_exc_marcador.addWidget(self.excluir_marcador_combobox)
        layout_exc_marcador.addWidget(self.excluir_marcador_button)
        layout2.addLayout(layout_exc_marcador)

        layout3.addWidget(self.novo_vinculo_label)
        layout_add_vinculo = QHBoxLayout(self.panel3)
        layout_add_vinculo.addWidget(self.novo_vinculo_input)
        layout_add_vinculo.addWidget(self.novo_vinculo_button)
        layout3.addLayout(layout_add_vinculo)
        layout3.addSpacing(70)
        layout3.addWidget(self.excluir_vinculo_label)
        layout_exc_vinculo = QHBoxLayout(self.panel3)
        layout_exc_vinculo.addWidget(self.excluir_vinculo_combobox)
        layout_exc_vinculo.addWidget(self.excluir_vinculo_button)
        layout3.addLayout(layout_exc_vinculo)
        #############################################################################################
        layout4.addWidget(self.painel_eventos_label)
        layout4.addWidget(self.lista_widget)
        layout4.addSpacing(20)
        layout4.addWidget(self.button_iniciar2)
        #############################################################################################
        layout5.addWidget(self.novo_excluido_batimento_label)
        layout_add_excluido_batimento = QHBoxLayout(self.panel5)
        layout_add_excluido_batimento.addWidget(self.novo_excluido_batimento_input)
        layout_add_excluido_batimento.addWidget(self.novo_excluido_batimento_button)
        layout5.addLayout(layout_add_excluido_batimento)
        layout5.addSpacing(70)
        
        layout5.addWidget(self.excluir_excluido_batimento_label)
        layout_exc_excluido_batimento = QHBoxLayout(self.panel5)
        layout_exc_excluido_batimento.addWidget(self.excluir_excluido_batimento_combobox)
        layout_exc_excluido_batimento.addWidget(self.excluir_excluido_batimento_button)
        layout5.addLayout(layout_exc_excluido_batimento)
        #############################################################################################
        layout6.addWidget(self.novo_excluido_falecimento_label)
        layout_add_excluido_falecimento = QHBoxLayout(self.panel6)
        layout_add_excluido_falecimento.addWidget(self.novo_excluido_falecimento_input)
        layout_add_excluido_falecimento.addWidget(self.novo_excluido_falecimento_button)
        layout6.addLayout(layout_add_excluido_falecimento)
        layout6.addSpacing(70)
        
        layout6.addWidget(self.excluir_excluido_falecimento_label)
        layout_exc_excluido_falecimento = QHBoxLayout(self.panel6)
        layout_exc_excluido_falecimento.addWidget(self.excluir_excluido_falecimento_combobox)
        layout_exc_excluido_falecimento.addWidget(self.excluir_excluido_falecimento_button)
        layout6.addLayout(layout_exc_excluido_falecimento)
        ############################################################################################# 
        layout7.addWidget(self.novo_marcador_substituto_label)
        layout_add_marcador_substituto = QHBoxLayout(self.panel7)
        layout_add_marcador_substituto.addWidget(self.novo_marcador_substituto_input)
        layout_add_marcador_substituto.addWidget(self.novo_marcador_substituto_button)
        layout7.addLayout(layout_add_marcador_substituto)
        layout7.addSpacing(70)
        
        layout7.addWidget(self.excluir_marcador_substituto_label)
        layout_exc_marcador_substituto = QHBoxLayout(self.panel7)
        layout_exc_marcador_substituto.addWidget(self.excluir_marcador_substituto_combobox)
        layout_exc_marcador_substituto.addWidget(self.excluir_marcador_substituto_button)
        layout7.addLayout(layout_exc_marcador_substituto)
        #############################################################################################
        layout8.addWidget(self.novo_url_sei_label)
        layout_novo_url_sei = QHBoxLayout(self.panel8)
        layout_novo_url_sei.addWidget(self.novo_url_sei_input)
        layout8.addLayout(layout_novo_url_sei)
        
        layout8.addSpacing(50)
        
        layout8.addWidget(self.novo_url_siape_label)
        layout_novo_url_siape = QHBoxLayout(self.panel8)
        layout_novo_url_siape.addWidget(self.novo_url_siape_input)
        layout8.addLayout(layout_novo_url_siape)
        
        layout8.addSpacing(50)
        
        layout8.addWidget(self.novo_unidade_sei_label)
        layout_novo_unidade_sei = QHBoxLayout(self.panel8)
        layout_novo_unidade_sei.addWidget(self.novo_unidade_sei_input)
        layout8.addLayout(layout_novo_unidade_sei)
        
        layout8.addSpacing(50)
        layout8.addWidget(self.novo_outros_button)
        
        #############################################################################################

        # Criar um layout vertical
        layout = QVBoxLayout(self.panel_container)

        # Adicionar os painéis ao layout vertical
        layout.addWidget(self.panel1)
        layout.addWidget(self.panel2)
        layout.addWidget(self.panel3)
        layout.addWidget(self.panel4)
        layout.addWidget(self.panel5)
        layout.addWidget(self.panel6)
        layout.addWidget(self.panel7)
        layout.addWidget(self.panel8)

        # Esconder inicialmente o Painel 2
        self.panel1.show()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()

        # Adicionar o layout à janela principal
        self.panel_container.setLayout(layout)
        self.setCentralWidget(self.panel_container)

        menubar = self.menuBar()

        # Menu "Arquivo"
        exibir = menubar.addMenu('Exibir')

        # Item de menu "Janela principal"
        menu_jan_principal = QAction('Painel Home', self)
        menu_jan_principal.triggered.connect(self.showPanel1)
        exibir.addAction(menu_jan_principal)
        
        menu_eventos = QAction('Painel Eventos', self)
        menu_eventos.triggered.connect(self.showPanel4)
        exibir.addAction(menu_eventos)
        
        # Menu "Arquivo"
        editar = menubar.addMenu('Editar')       

        menu_vinculos = QAction('Vínculos Decipex', self)
        menu_vinculos.triggered.connect(self.showPanel3)
        editar.addAction(menu_vinculos)
        
        menu_excluido_batimento = QAction('Lista Excluído Batimento', self)
        menu_excluido_batimento.triggered.connect(self.showPanel5)
        editar.addAction(menu_excluido_batimento)
        
        menu_excluido_falecimento = QAction('Lista Excluído Falecimento', self)
        menu_excluido_falecimento.triggered.connect(self.showPanel6)
        editar.addAction(menu_excluido_falecimento)
        
        menu_marcadores_substitutos = QAction('Marcadores Substituros', self)
        menu_marcadores_substitutos.triggered.connect(self.showPanel7)
        editar.addAction(menu_marcadores_substitutos)
        
        menu_outros = QAction('Outros', self)
        menu_outros.triggered.connect(self.showPanel8)
        editar.addAction(menu_outros)
        
        # Menu "Arquivo"
        ajuda = menubar.addMenu('Ajuda')       
        
        # Item de menu "Sobre"
        sobre_action = QAction('Sobre', self)
        sobre_action.triggered.connect(self.show_description)
        ajuda.addAction(sobre_action)
        
        # Item de menu "Sobre"
        manual = QAction('Manual do usuário', self)
        manual.triggered.connect(self.abrir_manual)
        ajuda.addAction(manual)

        # Exibir a janela
        altura = 400
        self.setGeometry(10, 40, 500, altura)

        self.panel_container.setMaximumHeight(altura-30)

        self.setWindowTitle('Coate Bot')

        screen = QDesktopWidget().screenGeometry()

        window = self.geometry()

        self.move(int((screen.width() - window.width()) / 2), int((screen.height() - window.height()) / 2))

        self.show()       
    
    def abrir_janela_item(self, item):
        conteudo = item.text()
        texto = str(conteudo)
        
        pyperclip.copy(conteudo)#comentar esta linha
        
        partes = texto.split(' ')
        
        for pt in partes:
            if str(pt.strip())=="CPF:":
                pyperclip.copy( str( partes[1] ) )
                conteudo = "O CPF "+ str(partes[1]) +" foi copiado para a área de tranferência. Utilize Ctrl + V para colar."
                self.aviso("Área de transferência", conteudo )      

            if str(pt.strip())=="Processo":
                pyperclip.copy( str( partes[2] ) )
                conteudo = "O Processo "+ str(partes[2]) +" foi copiado para a área de tranferência. Utilize Ctrl + V para colar."
                self.aviso("Área de transferência", conteudo )
    
    def inserir_marcador_substituto(self):
        print()
        texto = self.novo_marcador_substituto_input.text()
        texto = texto.upper()
        self.config["lista_marcadores_substitutos"].append(texto)

        self.configuracao.set_config(self.config)

        self.novo_marcador_substituto_input.clear()

        self.atualiza_excluir_marcador_substituto_combobox()

        desc_text = """
        Marcador substituto inserido com sucesso!
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Mensagem!")
        msgBox.setText(desc_text)
        msgBox.exec_()
    
    def atualiza_excluir_marcador_substituto_combobox(self):
        self.excluir_marcador_substituto_combobox.clear()
        for marcador_sub in list(self.config["lista_marcadores_substitutos"]):
            self.excluir_marcador_substituto_combobox.addItem(marcador_sub)
        
    def excluir_marcador_substituto(self):
        print()
        mensagem = "Desesa realmente excluir\n"+str(self.excluir_marcador_substituto_combobox.currentText())

        op = self.showMessageBox("Exclusão",mensagem)
        if op:
            self.config["lista_marcadores_substitutos"].remove(self.excluir_marcador_substituto_combobox.currentText())

            self.configuracao.set_config(self.config)

            self.atualiza_excluir_marcador_substituto_combobox()
            desc_text = """
            Marcador substituto excluido com sucesso!
            """
            msgBox = QMessageBox(self)
            msgBox.setWindowTitle("Mensagem!")
            msgBox.setText(desc_text)
            msgBox.exec_() 
        
    def alterar_outros(self):
        print() 
        print(self.novo_url_sei_input.text())  
        self.config["url_sei"]  = self.novo_url_sei_input.text()
        self.config["url_siapenet"] = self.novo_url_siape_input.text()
        self.config["unidade_sei"]  = self.novo_unidade_sei_input.text()
        
        self.configuracao.set_config(self.config)
                
    def atualiza_excluir_excluido_batimento_combobox(self):
        self.excluir_excluido_batimento_combobox.clear()
        for batimentos in list(self.config["lista_excluido_batimento"]):
            self.excluir_excluido_batimento_combobox.addItem(batimentos)
    
    def inserir_excluido_batimento(self):
        texto = self.novo_excluido_batimento_input.text()
        texto = texto.upper()
        self.config["lista_excluido_batimento"].append(texto)

        self.configuracao.set_config(self.config)

        self.novo_excluido_batimento_input.clear()

        self.atualiza_excluir_excluido_batimento_combobox()

        desc_text = """
        Código inserido com sucesso!
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Mensagem!")
        msgBox.setText(desc_text)
        msgBox.exec_()
        
    def excluir_excluido_batimento(self):
        mensagem = "Desesa realmente excluir\n"+str(self.excluir_excluido_batimento_combobox.currentText())

        op = self.showMessageBox("Exclusão",mensagem)
        if op:
            self.config["lista_excluido_batimento"].remove(self.excluir_excluido_batimento_combobox.currentText())

            self.configuracao.set_config(self.config)

            self.atualiza_excluir_excluido_batimento_combobox()
            desc_text = """
            Código excluido com sucesso!
            """
            msgBox = QMessageBox(self)
            msgBox.setWindowTitle("Mensagem!")
            msgBox.setText(desc_text)
            msgBox.exec_()         
            
    def atualiza_excluir_excluido_falecimento_combobox(self):
        self.excluir_excluido_falecimento_combobox.clear()
        for batimentos in list(self.config["lista_excluido_falecimento"]):
            self.excluir_excluido_falecimento_combobox.addItem(batimentos)
    
    def inserir_excluido_falecimento(self):
        texto = self.novo_excluido_falecimento_input.text()
        texto = texto.upper()
        self.config["lista_excluido_falecimento"].append(texto)

        self.configuracao.set_config(self.config)

        self.novo_excluido_falecimento_input.clear()

        self.atualiza_excluir_excluido_falecimento_combobox()

        desc_text = """
        Código inserido com sucesso!
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Mensagem!")
        msgBox.setText(desc_text)
        msgBox.exec_()
        
    def excluir_excluido_falecimento(self):
        mensagem = "Desesa realmente excluir\n"+str(self.excluir_excluido_falecimento_combobox.currentText())

        op = self.showMessageBox("Exclusão",mensagem)
        if op:
            self.config["lista_excluido_falecimento"].remove(self.excluir_excluido_falecimento_combobox.currentText())

            self.configuracao.set_config(self.config)

            self.atualiza_excluir_excluido_falecimento_combobox()
            desc_text = """
            Código excluido com sucesso!
            """
            msgBox = QMessageBox(self)
            msgBox.setWindowTitle("Mensagem!")
            msgBox.setText(desc_text)
            msgBox.exec_()
        
    def atualiza_excluir_vinculo_combobox(self):
        self.excluir_vinculo_combobox.clear()
        for vinculos in list(self.config["vinculos_decipex"]):
            self.excluir_vinculo_combobox.addItem(vinculos)
        

    def inserir_vinculo(self):
        texto = self.novo_vinculo_input.text()
        texto = texto.upper()
        self.config["vinculos_decipex"].append(texto)

        self.configuracao.set_config(self.config)

        self.novo_vinculo_input.clear()

        self.atualiza_excluir_vinculo_combobox()

        desc_text = """
        Unidade inserida com sucesso!
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Mensagem!")
        msgBox.setText(desc_text)
        msgBox.exec_()

    def excluir_vinculo(self):
        mensagem = "Desesa realmente excluir\n"+str(self.excluir_vinculo_combobox.currentText())

        op = self.showMessageBox("Exclusão",mensagem)
        if op:
            self.config["vinculos_decipex"].remove(self.excluir_vinculo_combobox.currentText())

            self.configuracao.set_config(self.config)

            self.atualiza_excluir_vinculo_combobox()
            desc_text = """
            Unidade excluida com sucesso!
            """
            msgBox = QMessageBox(self)
            msgBox.setWindowTitle("Mensagem!")
            msgBox.setText(desc_text)
            msgBox.exec_()
        
    def inserir_marcador(self):
        self.config["lista_marcadores_substitutos"].append(self.novo_marcador_input.text())

        self.configuracao.set_config(self.config)

        self.novo_marcador_input.clear()

        self.atualiza_excluir_marcador_combobox()

        desc_text = """
        Marcador inserido com sucesso!
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Mensagem!")
        msgBox.setText(desc_text)
        msgBox.exec_()
        
    def excluir_marcador(self):
        mensagem = "Desesa realmente excluir\n"+str(self.excluir_marcador_combobox.currentText())

        op = self.showMessageBox("Exclusão",mensagem)
        if op:
            self.config["lista_marcadores_substitutos"].remove(self.excluir_marcador_combobox.currentText())
            
            self.configuracao.set_config(self.config)

            self.atualiza_excluir_marcador_combobox()

            desc_text = """
            Marcador Excluido com sucesso!
            """
            msgBox = QMessageBox(self)
            msgBox.setWindowTitle("Mensagem!")
            msgBox.setText(desc_text)
            msgBox.exec_()
            
    def atualiza_excluir_marcador_combobox(self):
        self.excluir_marcador_combobox.clear()
        for marcadores in list(self.config["lista_marcadores_substitutos"]):
            self.excluir_marcador_combobox.addItem(marcadores)
            
    def showMessageBox(self, titulo_mensagem, mensagem):
        # Exemplo de uso do QMessageBox
        reply = QMessageBox.question(self, titulo_mensagem, mensagem, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            return True
        else:
            return False

    def toggle_password_visibility(self):
        if self.show_password_button.isChecked():
            self.show_password_button.setText("Ocultar")
            self.password_input.setEchoMode(QLineEdit.Normal)
        else:
            self.show_password_button.setText("Exibir")
            self.password_input.setEchoMode(QLineEdit.Password)
    
    def show_description(self):
        desc_text = """
        Sobre esta Automação:

        A presente automação tem por objetivo substituir os marcadores 'CDCONVINC' dos processos
de comunicado de falecimento no portal SEI - Sistema Eletrônico de Informações - por marcadores
'CADASTRO ATIVO', 'EXCLUÍDO BATIMENTO', 'EXCLUÍDO FALECIMENTO', 'SEM VÍNCULO' e
'VERIFICAÇÃO MANUAL' e anexar um documento PDF ao processo com as telas necessárias copiadas
do SIAPE - Terminal 3270.
 Para substituir os marcadores a automação acessa os processos no portal SEI, copia o CPF do
servidor/pensionista no marcador CDCONVINC, acessa o Terminal 3270 e verifica vínculo, código de exclusão
e certidão de óbito.

        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Descrição")
        msgBox.setText(desc_text)
        msgBox.exec_()
    
    def showPanel1(self):
        self.panel1.show()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide() 
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()       
        self.setWindowTitle('Painel Home')

    def showPanel2(self):
        self.panel1.hide()
        self.panel2.show()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()       
        self.setWindowTitle('Editar Lista de Marcadores')

    def showPanel3(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.show()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()
        self.setWindowTitle('Vínculos Decipex')
        
        
    def showPanel4(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.show()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()
        self.setWindowTitle('Painel de Eventos')
    
    def showPanel5(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.show()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.hide()
        self.setWindowTitle('Excluído Batimento')
        
    def showPanel6(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.show()
        self.panel7.hide()
        self.panel8.hide()
        self.setWindowTitle('Excluído Falecimento')
        
    def showPanel7(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.show()
        self.panel8.hide()
        self.setWindowTitle('Marcadores')
        
    def showPanel8(self):
        self.panel1.hide()
        self.panel2.hide()
        self.panel3.hide()
        self.panel4.hide()
        self.panel5.hide()
        self.panel6.hide()
        self.panel7.hide()
        self.panel8.show()
        self.setWindowTitle('Outros')
    
    def aviso(self, titulo, mensagem):
        desc_text = """
        Assim que solicitado, selecione o certificado e digite o código PIN.        
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle(titulo)
        msgBox.setText(mensagem)
        msgBox.exec_()        
        

    def info_evento(self, msg):
        """
        Adiciona uma mensagem à lista de exceções, exibe as exceções em ordem inversa em um widget QListWidget.

        Parâmetros:
            msg (str): A mensagem a ser adicionada à lista de exceções e exibida no widget.

        Exemplo de Uso:
            Para utilizar esta função, chame-a com a mensagem que deseja adicionar à lista de exceções.
            Por exemplo:
            info_evento("Erro ao processar os dados.")
        """
        self.lista_exception.append("\n"+msg)
        
        self.lista_invertida = self.lista_exception[::-1]

        self.lista_widget.clear()

        for lista_i in self.lista_invertida:
            print(lista_i)
            texto = lista_i
            fonte = QFont("Helvetica", 9)
            item = QListWidgetItem()
            texto = str(texto)
            item.setText(texto)
            item.setFont(fonte)
            self.lista_widget.addItem(item)
            
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(estilo.dark_stylesheet)
    ex = UI()
    sys.exit(app.exec_())