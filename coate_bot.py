from time import sleep
import time
import random
import json
from selenium.common.exceptions import StaleElementReferenceException
from web_automation import WebAutomation
import re
import os
import keyboard as kb
from datetime import datetime
#from selenium.webdriver.common.action_chains import ActionChains
from modSiape import IniciaModSiape
from web_automation import WebAutomation
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from pywinauto.application import Application
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import controle_terminal3270
from reportlab.lib.pagesizes import letter
from pywinauto import *
from pywinauto.findwindows import ElementNotFoundError
from seleciona_unidade_sei import SelecionaUnidade
import sys
from tela_mensagem import Mensagem as msg
import estilos as estilo
from selenium_setup import SeleniumSetup
import PyPDF2
from pdf_siape import PDFSiape as pdf
from registro_log import Logger
from PyQt5.QtWidgets import QMessageBox

class MyApp:
    def __init__(self):
        self.wauto                          = None
        self.browser                        = None
        self.data_atual                     = None
        self.app                            = None
        self.dlg                            = None
        self.acesso                         = None
        self.cpf                            = None
        self.numero_processo                = None
        self.arquivo                        = None
        self.qt_pagina_vinculo              = None
        self.pagina_linha_orgao             = []
        self.vinculos_decipex               = []
        self.conta_impressao                = 0
        self.servidor                       = False
        self.config                         = None
        self.impressora_corrente            = None
        self.lista_marcadores_substitutos   = []
        self.texto_eventos                  = None
        self.popup                          = False
        self.flag_terminar_manualmente      = False
        
        #Salva a impressora corrente para, ao final do programa voltar ao estado definido pelo usuário
        self.impressora_corrente = pdf.impressora_corrente()

        print("Impressora corrente: " + str(self.impressora_corrente) )

        #modifica a impressora padrão do windows para "Microsoft Print to PDF"
        pdf.print_to_pdf()
        
    def start(self, config):
        self.config = config
        
        self.log = Logger(self.config["diretorio_log"], None)
        
        try:
            self.iniciar()
        except Exception as erro:
            self.trata_exception(str(erro))
            
            
    def trata_exception(self,txt):
        txt = str(txt)
        self.log.salvar_log(txt)       
        
        erro1 = "Alert Text: Usuário ou Senha Inválida."
        erro2 = "Message: no such window: target window already closed"
        erro3 = "{'title_re': '^SEI - Marcadores.*'"
        erro4 = "{'title_re': '^Terminal 3270.*'"
        erro5 = "invalid literal for int()"
        erro6 = "'NoneType' object has no attribute 'is_displayed'"
        
                        
        if txt[0:len(erro1)] == erro1:
            self.texto_eventos ="Senha ou usuário inválido!\nDigite seu usuário e senha SEI corretamente\ne reinicie o processo."     
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
            
            self.sair_sei()
            self.sair_siape()                       
            
               
        elif txt[0:len(erro2)] == erro2:
            self.texto_eventos = "A página foi fechada!\nReinicie o programa."
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
                        
            self.sair_siape()            
            
            
        elif txt[0:len(erro3)] == erro3:
            self.texto_eventos = "A página foi fechada!\nReinicie o programa."
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
                        
            self.sair_siape()            
            
            
        elif txt[0:len(erro4)] == erro4:
            self.texto_eventos = "A página foi fechada!\nReinicie o programa."
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
                        
            self.sair_sei()           
            
        elif txt[0:len(erro5)] == erro5:
            self.texto_eventos = "Siapenet não está respondendo!\nReinicie o programa."
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
                        
            self.sair_sei() 
        
        elif txt[0:len(erro6)] == erro6:
            self.texto_eventos = "A página foi fechada!\nReinicie o programa."
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
                       
                 
            
        else:
            if self.flag_terminar_manualmente :
                self.texto_eventos = "Ocorreu um erro inesperado e o processo não pode ser concluído.\nComo o marcador CDCONVINC já foi excluído\ntermine o processo manualmente.\n\nSe precisar clique no processo abaixo para copiá-lo\npara a área de transferência."
            else :
                self.texto_eventos = txt
            
        
        try :
            self.app = Application().connect(title_re="^Painel de Eventos.*")
            self.dlg = self.app.window(title_re="^Painel de Eventos.*") 
            sleep(1)
            self.dlg.type_keys(" ")
        except Exception as erro :
            print("")    
    
    def iniciar(self):
        
        self.log.salvar_log("Início")
        
        self.texto_eventos = "Iniciando." 
        
        self.lista_marcadores_substitutos = list( self.config["lista_marcadores_substitutos"] )
       
        self.vinculos_decipex = list( self.config["vinculos_decipex"] )        

        # Obtenha a data e hora atuais
        self.data_atual = datetime.now()        

        # Formata a data no formato "dd/mm/yyyy"
        self.data_atual = self.data_atual.strftime("%d/%m/%Y")
        
        #se o siape (ela preta) estiver aberto fecha
        try :
            self.app = Application().connect(title_re="^Painel de controle.*")
            self.sair_siape()
            
        except Exception as e :
            print("continua")
            
        #se o SEI estiver aberto em SEI - Marcadores fecha o SEI
        try :
            app_sei = Application().connect(title_re=f"^SEI - Marcadores.*")
            self.sair_sei()           
        except Exception as e :
            print("continua")
            
         #se o SEI estiver aberto em SEI - Processo fecha o SEI
        try :
            app_sei = Application().connect(title_re=f"^SEI - Processo.*")
            self.sair_sei()           
        except Exception as e :
            print("continua")
    
        self.texto_eventos = "A seguir siga as instruções na página do SiapeNet"
        
        sleep(3)
        
        siape = SeleniumSetup()            
        IniciaModSiape(siape, self.config["url_siapenet"]).executar_siape()
        self.log.salvar_log("Abrindo siape net")        
        
        #o bloco abaixo fica aguardando a tela preta aparecer por 30 segundos
        contador_telapreta = 0
        flag_telapreta = True

        self.texto_eventos = "Abrindo Terminal 3270."
        
        while flag_telapreta :
            try :
                self.app = Application().connect(title_re="^Terminal 3270.*")
                self.dlg = self.app.window(title_re="^Terminal 3270.*")
                self.acesso = controle_terminal3270.Janela3270()
                flag_telapreta = False
            except Exception as e :
                contador_telapreta+=1
                print("Tentativa",contador_telapreta)
                if contador_telapreta >= 30 :
                    flag_telapreta = False
                    self.texto_eventos = "Falha ao abrir o SiapeNet - Reinicie o programa."
                    self.log("Falha ao abrir o SiapeNet.")
                    
            self.sleepal(1,3)
    
        service = Service()
        options = webdriver.ChromeOptions()
        self.browser = webdriver.Chrome(service=service, options=options)
        
        self.texto_eventos = "Abrindo página SEI."
        
        
        #Entra no SEI
        self.log.salvar_log("Abrindo página SEI")                
        self.wauto = WebAutomation(self.browser)
        
        self.log.salvar_log("Digita usuário, senha e órgão")            
        self.wauto.entrar_no_sei(self.config["url_sei"],self.config["ultimo_acesso_user"], self.config["senha"], self.config["ultimo_orgao"])
        self.browser.implicitly_wait(3)
    
        # fecha tela de aviso
        fecha_tela_aviso = '/html/body/div[7]/div[2]/div[1]/div[3]/img'
        self.wauto.tela_aviso(fecha_tela_aviso)

        # Seleciona unidade
        self.log.salvar_log("Seleciona Unidade SEI 'Caixa': "+self.config["unidade_sei"])
        SelecionaUnidade(self.browser, self.config["unidade_sei"])

        # Clica em "Ver por por Marcadores" na janela de Controle de Processos
        self.browser.find_element('xpath', '//*[@id="divFiltro"]/div[3]/a').click()  # Clica em ver por marcadores        
        
        #procura o marcador CDCONVINC e clica
        flag_marcador = self.clica_marcador()
        
        if flag_marcador :    
            self.log.salvar_log("Marcador "+self.config["marcador"]+": True")
            
            # pega o total de registros
            self.sleepal(1,2)
            total_registros = self.browser.find_element('xpath', '//*[@id="tblProcessosRecebidos"]/caption').text
            
            # REGEX para extrair só números do total de registros
            total_registros = re.findall(r'\d+', total_registros)
            total_registros = int(total_registros[0])
            
            # Pega qt de linhas geradas no marcador CDCONVINC e faz o laço pra fazer as consultas CDCONVINC de todas as linhas
            self.sleepal(1,2)
            
            indice_da_linha = 1
    
            while indice_da_linha <= total_registros and flag_marcador:
                self.flag_terminar_manualmente      = False
                
                if self.config["funcionando_teste"]:msg("Homologação assistida " + str(self.data_atual) + "\nProcesso "+str(indice_da_linha) + "de "+ str(total_registros))
                inicio_evento = time.time()
                try:            

                    #Clica na caixa de seleção
                    self.log.salvar_log("--")
                    self.wauto.clicar_caixaselecao_e_marcador()            

                    self.cpf = self.browser.find_element('xpath', '//*[@id="tdAndamentoMarcador78958"]').text  # cpf que está no texto do marcador CDCONVINC
                    self.log.salvar_log("CPF "+self.cpf)
                    
                    self.numero_processo = self.browser.find_element('xpath', '//*[@id="divInfraBarraLocalizacao"]').text[-20:]
                    self.log.salvar_log("Processo "+self.numero_processo)
                    
                    self.texto_eventos = "CPF: "+str(self.cpf)                    
                    sleep(0.4)
                    
                    self.texto_eventos = "Processo nº "+str(self.numero_processo)                    
                    sleep(0.4)
                    
                    
                   
                    # Entra no Terminal 3270 - Siape Hod (tela preta)
                    self.app = Application().connect(title_re="^Terminal 3270.*")
                    self.dlg = self.app.window(title_re="^Terminal 3270.*")
                    self.acesso = controle_terminal3270.Janela3270()                   

                    self.acessa_terminal_3270()                   
                    
                    #testa condição acesso negado à tela preta
                    flag_acesso_negado = True

                    while flag_acesso_negado :
                        try:
                            cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 9, 24, 44).strip()
                            texto_dif_cpf = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 2, 24, 16).strip()
                            texto_cpf_invalido = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 9, 24, 21).strip()
                            flag_acesso_negado = False
                            sleep(2)               
                        except Exception as e :
                            print(e)

                    #################################################################################################
                    '''
                    Testa 4 condições: Se cpf não tem vinculo algum; se cpf não existe; se cpf tem vinculos mas não
                    com decipex; se cpf tem vinculo com decipex
                    '''
                    #################################################################################################
                    if texto_cpf_invalido == "CPF INVALIDO":
                        self.log.salvar_log("CPF inválido")
                        marcador = self.lista_marcadores_substitutos[3]
                        self.cpf_invalido(marcador)
                        
                    elif texto_dif_cpf == "INFORME APENAS":
                        self.log.salvar_log("CPF inválido")
                        marcador = self.lista_marcadores_substitutos[3]
                        self.cpf_invalido(marcador)

                    elif cpf_mensagem == "DIGITO VERIFICADOR INVALIDO":
                        self.log.salvar_log("CPF inválido")
                        marcador = self.lista_marcadores_substitutos[3]
                        self.cpf_invalido(marcador)

                    elif cpf_mensagem == "NAO EXISTEM DADOS PARA ESTA CONSULTA":
                        self.log.salvar_log("Não existem dados para a consulta do CPF")
                        marcador = self.lista_marcadores_substitutos[4]
                        self.sem_vinculo_algum(marcador)

                    else:
                        pdf.coletar_tela(self)
                        
                        self.dlg.type_keys('x')  # marca o nome

                        self.sleepal(1,2)

                        kb.press("Enter")

                        #verifica se tem vinculo com decipex
                        tem_vinculo_decipex = self.verifica_vinculo(self.vinculos_decipex.copy())
                        
                        if tem_vinculo_decipex:
                            self.log.salvar_log("CPF com vínculo Decipex")
                            self.com_vinculo_decipex()

                        else:
                            self.log.salvar_log("CPF sem vínculo Decipex")
                            self.sem_vinculo_decipex()


                except StaleElementReferenceException as e:
                    print(e)
                    continue
                except IndexError:
                    print("Fim da tabela.")
                    break

                indice_da_linha+=1
                self.pagina_linha_orgao.clear()
            
                tempo_decorrido = time.time() - inicio_evento  # Calcula o tempo decorrido em segundos
                minutos = int(tempo_decorrido // 60)  # Converte os segundos em minutos
                segundos = int(tempo_decorrido % 60)  # Calcula o restante dos segundos                
                
                self.texto_eventos = "Tempo: "+ str(minutos) + "m " + str(segundos) + "s."
                
                self.log.salvar_log( "Tempo " + str(minutos) + "m " + str(segundos) + "s." )
            

            self.sair_siape()
            self.sair_sei()
            self.texto_eventos = "Terminado."
            pdf.impressora_padrao(self.impressora_corrente)
            
        else :
            self.log.salvar_log("Marcador "+self.config["marcador"]+": False")

    def com_vinculo_decipex(self):
        
        vinculos_decipex = self.vinculos_decipex.copy()
        
        # o codigo abaixo gera uma tupla de vinculos decipex
        tupla_decipex = []

        # página_linha_orgao é a tupla que contém todos os vinculos, é a tupla principal
        # tupla_decipex é a tupla que contém apenas as paginas, linhas e orgãos com vinculo decipex

        for i in range(self.qt_pagina_vinculo):#esse for varre todas as páginas

            #Recebe uma tupla de lista com pagina, linha, orgão e tipo(True=Pensionista ou False=Servidor) de todos os vinculos que o serv/pens possui
            tupla_integral = [(pagina, linha, orgao, tipo) for pagina, linha, orgao, tipo in self.pagina_linha_orgao if pagina == i]

            '''exemplo: tupla_integral = [(0,0,40801,False),(0,1,40806,True),(0,2,40801,False)]'''
            
            #Recebe a ultima lista da tupla_integral
            ultima_lista_da_tupla = tupla_integral[-1]
            
            '''ultima_lista_da_tupla = (0,2,40806,True)'''

            #Recebe o segundo elemento da ultima lista
            #O segundo elemento da ultima lista indica quantas linhas a página comtém
            qt_linha = int( ultima_lista_da_tupla[1] )
            
            '''qt_linha = 2'''        

            #varre todas as linhas da página
            for j in range(qt_linha + 1):            
                
                #recebe cada lista da tupla_integral da pagina atual (i)
                lista_j = tupla_integral[j]

                #varre a lista de orgaos decipex e compara com os orgaos da lista_J
                #monta uma tupla com pagina, linha e orgao somente com vinculos decipex
                for vinc_d in vinculos_decipex:

                    if lista_j[2] == vinc_d:
                        tupla_decipex.append(lista_j)
                        #msg("tupla_decipex: " + str(tupla_decipex) )

        ativo               = False
        batimento           = False
        tem_certidao        = False

        for lista_decipex in tupla_decipex:
            
            '''(0,1,40806,False)'''

            # o metodo verifica_exclusao recebe a página e a linha com orgao decipex e verifica se encontra o codigo decipex
            pagina  = lista_decipex[0]
            linha   = lista_decipex[1]
            tipo    = lista_decipex[3]          
            
            lista_codigo_obito = self.verifica_exclusao_obito(pagina, linha, tipo)
            
            marcador = None
            
            lista_ativo = list( self.config["lista_ativo"] )

            lista_excluido_batimento = list( self.config["lista_excluido_batimento"] )

            lista_excluido_falecimento = list( self.config["lista_excluido_falecimento"] )        


            if self.consulta_cod_lista(lista_codigo_obito[0], lista_ativo, tipo):
                ativo = True                

            elif self.consulta_cod_lista(lista_codigo_obito[0], lista_excluido_batimento, tipo):
                batimento = True                

            elif self.consulta_cod_lista(lista_codigo_obito[0], lista_excluido_falecimento, tipo):
                if not tem_certidao : 
                    tem_certidao = lista_codigo_obito[1]#1 é excluído batimento
                    if tem_certidao == 0:
                        marcador = self.lista_marcadores_substitutos[3]
                
            else:
                marcador = self.lista_marcadores_substitutos[3]#3 é verinficação manual

        # fim do for      


        #o bloco abaixo define a precedência : Ativo, Batimento e Falecimento
        if ativo :
            #marcador = "CADASTRO ATIVO"
            marcador = self.lista_marcadores_substitutos[0]#0 é cadatstro ativo
        elif batimento :
            #marcador = "EXCLUÍDO BATIMENTO"
            marcador = self.lista_marcadores_substitutos[1]#1 é excluído batimento
        else :
            if tem_certidao : 
                if marcador == "VERIFICAÇÃO MANUAL" : #isso garante a precedência da verificação manual. Se o marcado já vier das condicionais como verificação manual, continua mesmo que haja obito em algum dos órgaos, caso haja mais de um órgão
                    print("Marcador já estava como VERIFICAÇÃO MANUAL")
                else :
                    marcador = self.lista_marcadores_substitutos[2]#2 é excluído falecimento
            else :
                marcador = self.lista_marcadores_substitutos[3]#3 é verinficação manual

        self.salva_arquivo()
        self.exclui_marcador(self.config["marcador"] )
        self.add_marcador(marcador,None)
        self.add_documento()
    
    def consulta_cod_lista(self, codigo, lista, tipo):
        flag = False
        codigo_na_lista = ""
        for item_lista in lista :

            if tipo :
                codigo_na_lista = str(item_lista).replace("/", "")

            else :
                codigo_na_lista = str(item_lista)

            if str(codigo_na_lista) == str(codigo) :
                flag = True
                break

        return flag    


    def verifica_exclusao_obito(self, pag, lnh, tp) :              
        lista_retorno = []    

        self.dlg.type_keys("{F8 " + str(pag) + "}")  # pagina
        sleep(0.5)
        
        self.dlg.type_keys("{TAB " + str(lnh) + "}")  # linha
        
        self.dlg.type_keys('x')
        sleep(0.5)
        
        kb.press("Enter")

        sleep(1.5)

        lista_linha_pagina_servidor = []

        flag = True
        obito = False
        codigo_exclusao = "00/000" # se não encontrar cpalavra exclusao retorna ativo
        achou_obito = False
        achou_exclusao = False
        

        # o laço abaixo varre todas as paginas do servidor/pensionista em busca de codigo de exclusão e registro óbito
        while flag :        
            pdf.coletar_tela(self)
            
            if not achou_obito and not achou_exclusao :
                linha = 7
                for i in range(24) :
                    lista_linha_pagina_servidor.append( self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 3, linha, 80).strip() )  # Pega o texto na linha                texto_na_tela.append(pega_texto)
                    linha+=1
                
                for indice, elemento in enumerate(lista_linha_pagina_servidor) :

                
                    if "OBITO" in elemento :
                        achou_obito = True
                        L = indice + 8
                        certidao_obito = self.acesso.pega_texto_siape(self.acesso.copia_tela(), L, 16, L, 80).strip()
                        if certidao_obito == "" :
                            print()
                        else:
                            obito = True
                            

                    if "EXCLUSAO" in elemento :
                        achou_exclusao = True
                        L = indice + 8
                        if tp :#verifica se o tipo e pensionista ou servidor True=Pensionista False=Servidor
                            codigo_exclusao = self.acesso.pega_texto_siape(self.acesso.copia_tela(), L, 19, L, 23).strip()
                            achou_obito = True
                        else :
                            codigo_exclusao = self.acesso.pega_texto_siape(self.acesso.copia_tela(), L, 23, L, 29).strip()
                            if codigo_exclusao == "02/227" or codigo_exclusao == "02/237": achou_obito = True                   


            if self.ultima_pagina_vinculo():
                self.dlg.type_keys("{F12}")
                flag = False
            else :
                self.avanca()
                lista_linha_pagina_servidor.clear()
            
        # se for pensionista não vem informação de obito então coloca como verdadeiro
        # tp True = Pensionista, tp False = Servidor    
        if tp : obito = True

        lista_retorno.append(codigo_exclusao)
        lista_retorno.append(obito)
        return lista_retorno

    def avanca(self):        
        #printa_tela antes de avançar
        a = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 1, 1, 25, 80)
        print("=================================================================================")
        print(a)
        print("=================================================================================")
                
        #avança
        self.dlg.type_keys("{F8}")            
            
        #entra num laço para verificar se a tela avançou ou não
        flag_repete = True
        while flag_repete :
            #printa_tela depois de avançar
            b = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 1, 1, 25, 80)
                    
            if a == b:
                print("Telas iguais, não avançou! Repete")
                sleep(0.5)
            else :
                flag_repete = False
        

    def sem_vinculo_decipex(self):
        for i in range(self.qt_pagina_vinculo):
            
            #Recebe uma tupla de lista com pagina, linha, orgão e tipo(True=Pensionista ou False=Servidor) de todos os vinculos que o serv/pens possui
            #tipo aqui não é utilizado
            tupla_integral = [(pagina, linha, orgao, tipo) for pagina, linha, orgao, tipo in self.pagina_linha_orgao if pagina == i]

            #Recebe a ultima lista da tupla_integral
            ultima_lista_da_tupla = tupla_integral[-1]

            #Recebe o segundo elemento da ultima lista
            #O segundo elemento da ultima lista indica quantas linhas a página comtém
            qt_linha = int(ultima_lista_da_tupla[1])
            self.conta_impressao = 0
            for j in range(qt_linha + 1):                  

                self.dlg.type_keys("{TAB " + str(j) + "}")
                self.dlg.type_keys('x')
                self.sleepal(1,2)
                kb.press("Enter")

                lista_linha_pagina_servidor = []                
                flag = True
                
                # o laço abaixo varre todas as paginas do servidor/pensionista em busca de exclusão e ingresso
                while flag :
                    
                    linha = 7

                    for i in range(24) :
                        lista_linha_pagina_servidor.append( self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 3, linha, 80).strip() )  # Pega o texto na linha                texto_na_tela.append(pega_texto)
                        linha+=1

                    printa_tela = True
                    for indice, elemento in enumerate(lista_linha_pagina_servidor) :

                        
                        if "EXCLUSAO" in elemento :
                            if printa_tela :
                                pdf.coletar_tela(self)
                                printa_tela = False

                        if "INGRESSO NO" in elemento :
                            if printa_tela :
                                pdf.coletar_tela(self)
                                printa_tela = False


                    if self.ultima_pagina_vinculo() :
                        flag = False
                        self.dlg.type_keys("{F12}")
                        lista_linha_pagina_servidor.clear()
                    else :
                        #self.dlg.type_keys("{F8}")
                        self.avanca()
                        lista_linha_pagina_servidor.clear() 



                #fim do while

            #fim do for orgao

        #fim do for página

        self.salva_arquivo()
        marcador = self.config["marcador"]
        self.exclui_marcador(marcador )
        marcador_substituto = self.lista_marcadores_substitutos[4] # marcador SEM VÍNCULO
        self.add_marcador(marcador_substituto,None)
        self.add_documento()

    def ultima_pagina_vinculo(self) :
        ultima_pagina = False
        
        #condições que indicam ultima página:
        #---------------------------------------------------------------------------------------------------------#
        linha_comando = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 22, 56, 22, 77).strip()
        if linha_comando == "ULTIMA TELA DO VINCULO" :
            ultima_pagina = True
            
        avanca = False
        linha_comando = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 23, 1, 23, 80).strip()
        partes = linha_comando.split(" ")
        for p in partes :
            if ( str(p) == str("PF8=AVANCA") ) or ( str(p) == str("PROXIMO") ) :
                avanca = True
        if not avanca :
            ultima_pagina = True
            
        return ultima_pagina
    
      
    
    def salva_arquivo(self):
        # cria a pasta para receber o print do vinculo
        try:
            os.makedirs('c:\pdf')
            print('Diretório criado com sucesso!')
        except FileExistsError:
            print('O diretório já existe!')
        except Exception as e:
            print('Deu certo não. O diretório não foi criado...')
            
        self.sleepal(1,2)

        #limpa o diretorio dos vinculos.pdf
        arquivos = os.listdir(self.config["diretorio_vinculos"])
        for arquivo in arquivos:
            caminho_completo = os.path.join(self.config["diretorio_vinculos"], arquivo)
            # Verifica se o caminho é um arquivo
            if os.path.isfile(caminho_completo):
                os.remove(caminho_completo)
                
        self.sleepal(1,2)
        
        pdf.salva_telas_coletadas(self)
        
        # Obtém a janela de impressão e tecla ENTER
        flag = True
        while flag:
            try:
                app_imprimir = Application().connect(title_re="Imprimir")
                window = app_imprimir.window(title_re="Imprimir")
                window.set_focus()
                kb.press("Enter")
                flag = False
            except:
                sleep(0.5)
                
        flag = True
        #obtém a janela Saída de Impressão Como
        while flag:
            try:
                app_salvar = Application().connect(title_re='Salvar Saída de Impressão como')
                window = app_salvar.window(title_re='Salvar Saída de Impressão como')
                window.set_focus()
                self.dlg_salvar = app_salvar[u'Salvar Saída de Impressão como']                
                self.arquivo = "C:\pdf\Vinculo" + self.cpf + ".pdf"
                self.dlg_salvar.type_keys(self.arquivo)
                sleep(0.5)
                kb.press("Enter")
                flag = False
            except:
                sleep(0.5)             
    
    def verifica_vinculo(self, lista_orgaos_decipex):
        
        # verifica a quantidade de páginas que o serv/pens possui de vinculos    
        self.sleepal(1,2)
        self.qt_pagina_vinculo = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 9, 78, 9, 80).strip()            


        self.qt_pagina_vinculo = int(self.qt_pagina_vinculo)
        sleep(0.5)

        for qtp in range(self.qt_pagina_vinculo):

            if self.qt_pagina_vinculo > 1: self.dlg.type_keys("{F8}")  # Se quantidade de páginas for maior que 1 tecla F8 para alternar e coletar outras páginas

            sleep(0.3)

            pdf.coletar_tela(self)

        self.popula_tupla()

        flag = False

        for org_decipex in lista_orgaos_decipex:
            
            flag = bool([(pagina, linha, orgao, tipo) for pagina, linha, orgao, tipo in self.pagina_linha_orgao if orgao == org_decipex])
            if flag : break
            
        return flag


    def popula_tupla(self):
        tipos = []

        for qtp in range(self.qt_pagina_vinculo):
            linha = 10
            linha_representativa = 0

            if self.qt_pagina_vinculo > 1: self.dlg.type_keys(
                "{F8}")  # Se quantidade de páginas for maior que 1 tecla F8 para alternar e coletar outras páginas

            for a in range(13):  # trocar depois pra 13 que pesquisa até a linha 22
                pens = False

                pega_tipo = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 7, linha,
                                                        18).strip()  # Pega se é servidor ou pensionista
                tipos.append(pega_tipo)  # cria uma lista com os tipos de vínculos

                if pega_tipo == "SERVIDOR":
                    pens = False

                    pega_orgaoservidor = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 25, linha,
                                                                    29).strip()  # Pega o órgão do servidor
                    # orgaos_na_tela.append(str(qtp)+"-"+pega_orgaoservidor)
                    self.pagina_linha_orgao.append((qtp, linha_representativa, pega_orgaoservidor, pens))

                elif pega_tipo == "PENSIONISTA":
                    pens = True

                    pega_orgaopensionista = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 40, linha,
                                                                        44).strip()  # Pega o órgão do pensionista
                    # orgaos_na_tela.append(str(qtp)+"-"+pega_orgaopensionista)
                    self.pagina_linha_orgao.append((qtp, linha_representativa, pega_orgaopensionista, pens))

                linha += 1
                linha_representativa += 1
                # print(tipos)
            # print(orgaos_na_tela)
            print(self.pagina_linha_orgao)

    def cpf_invalido(self, marcador):
        #pdf.coletar_tela(self)
        #self.salva_arquivo()
        self.exclui_marcador(str(self.config["marcador"]))
        info_add = "CPF informado no marcador CDCONVINC ("+self.cpf+") é inválido. Nenhum documento foi adicionado ao processo."
        self.add_marcador(marcador, info_add)        
        # clica botão processos e continua para fazer o próximo processo
        self.browser.find_element(By.XPATH, '//*[@id="lnkControleProcessos"]/img').click()   
        
    def sem_vinculo_algum(self, marcador):
        pdf.coletar_tela(self)
        self.salva_arquivo()
        self.exclui_marcador(str(self.config["marcador"]))
        info_add = "CPF ("+self.cpf+") não tem vínculo algum. Verificar documento anexo."
        self.add_marcador(marcador, info_add)
        self.add_documento()
        
    def sair_siape(self) :
        try:
            self.app = Application().connect(title_re="^Painel de controle.*")
            self.dlg = self.app.window(title_re="^Painel de controle.*")
            self.dlg.type_keys('%{F4}')
            self.log.salvar_log("Fechou Terminal 3270")
            sleep(1)
            kb.press("Enter")
        except Exception as e:
            print(e)
            
    def sair_sei(self):
        try:
            self.app = Application().connect(title_re="^SEI *")
            self.dlg = self.app.window(title_re="^SEI *")
            #sleep(1)
            #try:            
                #self.browser.find_element('xpath', '//*[@id="lnkInfraSairSistema"]/img').click()
            #except Exception as erro:
                #self.dlg.type_keys('%{F4}')
            sleep(1)
            self.dlg.type_keys('%{F4}')
            self.log.salvar_log('Fechou página do SEI')
            
        except Exception as e:
            print(e)
        

    def clica_marcador(self):
        ite = 2
        sem_marcador = True
        marcador_procurado = str(self.config["marcador"])#"CDCONVINC"
        self.texto_eventos = "Procurando marcador "+str(marcador_procurado)
        rows = self.browser.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/form/div/div[5]/div/table/tbody/tr')
        
        for row in rows[1:]:
            coluna_marcadores = row.find_element(By.XPATH, "./td[3]").text
            if coluna_marcadores == marcador_procurado:
                sem_marcador = False
                elemento = '//*[@id="tblMarcadores"]/tbody/tr[' + str(ite) + ']/td[1]/a[2]'
                self.wauto.clica_elemento_by_xpath(None, elemento)
                break
            ite += 1

        if sem_marcador:
            self.texto_eventos = "Não há processo marcado com "+ marcador_procurado + "."
            self.sair_siape()
            self.sair_sei()
        else:
            return True
            
    
    def aviso(self, titulo, mensagem):
        desc_text = """
        Assim que solicitado, selecione o certificado e digite o código PIN.        
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle(titulo)
        msgBox.setText(mensagem)
        msgBox.exec_()        


    def add_marcador(self, marcador, info_add):
        self.log.salvar_log("Marcador adicionado: "+marcador)
        self.texto_eventos = marcador
        
        self.sleepal(1,2)

        self.browser.find_element('xpath', '//*[@id="btnAdicionar"]').click()  # botão adicionar marcador

        if info_add == None:
            self.browser.find_element('xpath', '//*[@id="txaTexto"]').send_keys(marcador)  # preenche o texto com o código pego no hod
        else:
            self.browser.find_element('xpath', '//*[@id="txaTexto"]').send_keys(str(marcador) +" - "+ str(info_add) )  # preenche o texto com o código pego no hod
            
        
        self.sleepal(1,2)

        # Localize o combo "marcador"
        combo_marcador = WebDriverWait(self.browser, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="selMarcador"]/div/span'))
        )

        # Clique no combo de pesquisa "marcador"
        combo_marcador.click()

        # Aguarde até que os elementos dentro do combo "marcador" estejam presentes
        elementos_labels_assinado = WebDriverWait(self.browser, 3).until(
            EC.presence_of_all_elements_located((By.XPATH, f'//*[@id="selMarcador"]//ul/li/a/label[text()="{marcador}"]')))

        self.sleepal(1,2)
        # Se a lista for longa, pode ser necessário rolar até o elemento antes de clicar
        for elemento_label_assinado in elementos_labels_assinado:
            self.browser.execute_script("arguments[0].scrollIntoView(true);", elemento_label_assinado)
            elemento_label_assinado.click()
            break  # Para após clicar no primeiro elemento "ASSINADO" encontrado

        self.sleepal(1,2)
        self.browser.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()  # botão salvar
        self.sleepal(2,4)
       

    def exclui_marcador(self, marcador):
        #msg("metodo exclui_marcador: ")
        # Assume a janela SEI - Marcadores
        app_sei = Application().connect(title_re=f"^SEI - Marcadores.*")
        self.sleepal(1,2)
        window_sei = app_sei.window(title_re=f"^SEI - Marcadores.*")
        window_sei.set_focus()
        self.sleepal(1,2)

        ############################################################################################################
        # Marca a linha com a etiqueta CDCONVINC e exclui a mesma
        ############################################################################################################

        ite = 2

        rows = self.browser.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/form/div[3]/table/tbody/tr')

        for row in rows[1:]:
            coluna_marcadores = row.find_element(By.XPATH, "./td[2]")

            texto1 = str(coluna_marcadores.text)
            texto2 = str(marcador)        


            if texto1.strip() == texto2.strip() :
                elemento = '//*[@id="tblMarcadores"]/tbody/tr['+str(ite)+']/td[6]/a[2]/img'            
                self.browser.find_element(By.XPATH, elemento).click() 

                # Aguarde a janela de confirmação aparecer
                self.sleepal(1,2)
                confirmation = self.browser.switch_to.alert

                # Aceitar a confirmação (clicar em "OK")
                confirmation.accept()
                break       

            ite += 1
            
        self.flag_terminar_manualmente = True   

    def is_pdf_empty(file_path):
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                num_pages = len(pdf_reader.pages)
                return num_pages == 0
        except Exception as e:
            print(f"Erro ao verificar o arquivo PDF: {e}")
            return False

    def add_documento(self):
        
        self.browser.find_element(By.XPATH, '//*[@id="txtPesquisaRapida"]').send_keys(
            self.numero_processo)  # Campo pesquisar
        self.browser.find_element(By.XPATH,
                                '//*[@id="spnInfraUnidade"]/img').click()  # clica na lupa para pesquisar o processo

        self.sleepal(1,2)
        # msg("Clica botão novo doc")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao',
                                                '//*[@id="divArvoreAcoes"]/a[1]/img')  # Clica botão novo doc
        sleep(1)
        # msg("Digita Externo no campo tipo documento")
        self.wauto.inserir_texto_enter_by_id_in_iframe('ifrVisualizacao', 'txtFiltro',
                                                    'Externo')  # Digita Externo no campo tipo documento
        sleep(1)
        # msg("Clica na opção Externo")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao',
                                                '//*[@id="tblSeries"]/tbody/tr[1]/td/a[2]/span')  # Clica na opção Externo
        sleep(1)
        # msg("Clica no tipo de documento (opção escolhida é Anexo)")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao',
                                                '//*[@id="selSerie"]')  # Clica no tipo de documento (opção escolhida é Anexo)
        sleep(1)
        # msg("Digita texto 'Anexo'")
        self.wauto.inserir_texto_enter_by_id_in_iframe('ifrVisualizacao', 'selSerie', 'Anexo')  # Digita texto 'Anexo'
        sleep(1)
        # msg("data_atual")
        # //*[@id="txtDataElaboracao"]
        self.wauto.inserir_texto_enter_by_id_in_iframe('ifrVisualizacao', 'txtDataElaboracao', self.data_atual)
        sleep(1)
        # msg("Marca opção Nato-Digital")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao', '//*[@id="divOptNato"]')  # Marca opção Nato-Digital
        sleep(1)
        # msg("Digita nome que aparecerá na árvore")
        self.wauto.inserir_texto_enter_by_id_in_iframe('ifrVisualizacao', 'txtNomeArvore',
                                                    'Comprovante de Vínculo')  # Digita nome que aparecerá na árvore
        sleep(1)
        # msg("Marca Nivel Restrito")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao', '//*[@id="divOptRestrito"]')  # Marca Nivel Restrito
        sleep(1)
        
       
        # msg("Digita texto para filtrar a Hipótese legal no Combolist Hipótese legal")
        self.wauto.inserir_texto_enter_by_id_in_iframe('ifrVisualizacao', 'selHipoteseLegal','Informação Pessoal (Art. 31 da Lei')  # Digita texto para filtrar a Hipótese legal no Combolist "Hipótese legal"
                
            
        
        # msg("Clica botão Anexar Arquivo")
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao', '//*[@id="lblArquivo"]')  # Clica botão Anexar Arquivo
        sleep(0.8)

        #o bloco abaixo seleciona o arquivo pdf e dá enter
        app_abrir = Application().connect(title_re='Abrir')
        dlg_abrir = app_abrir[u'Abrir']
        self.sleepal(1,2)
        window = app_abrir.window(title_re='Abrir')
        window.set_focus()
        self.sleepal(1,2)
        dlg_abrir.type_keys(self.arquivo)
        sleep(0.5)
        kb.press("Enter")
        
        self.sleepal(2,3)  # Aguarda terminar preenchimento do anexo
        self.wauto.default_clica_elemento_by_xpath('ifrVisualizacao', '//*[@id="btnSalvar"]')  # Clica botão Salvar
        
        # clica botão processos e continua para fazer o próximo processo
        self.browser.find_element(By.XPATH, '//*[@id="lnkControleProcessos"]/img').click()   
    
    
    def acessa_terminal_3270(self) :
        tela = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 1, 1, 24, 80).strip()
                    
        #Copia a tela e verifica se há mansagens cadastradas
        partes_tela = tela.split(" ")
                    
        txt1 = False
        txt2 = False
                    
        for pt in partes_tela :
            if pt == "MENSAGEM(NS)" :
                txt1 = True
            if pt == "CADASTRADA(S):" :
                txt2 = True
                            
        if txt1 and txt2 :
            self.dlg.type_keys('{F3}')
            self.dlg.type_keys('{F2}')                   
            time.sleep(2)                    
            self.dlg.type_keys(">"+str(self.config["marcador"]))
            time.sleep(2)
            kb.press("Enter")                    
            time.sleep(2)
        else :
            txt = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 6, 6, 6, 61).strip()
            if txt == "POSICIONE O CURSOR NA OPCAO DESEJADA E PRESSIONE <ENTER>" :
                flag_ultima_pagina = False
                while not flag_ultima_pagina :
                    self.dlg.type_keys('{F8}')    
                    txt = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 23, 8, 23, 20).strip()
                    if txt == "ULTIMA PAGINA" :
                        time.sleep(2)                    
                        self.dlg.type_keys(">"+str(self.config["marcador"]))
                        flag_ultima_pagina = True
                    
                        time.sleep(2)
                        kb.press("Enter")                    
                        time.sleep(2)
            else:
                self.dlg.type_keys('{F3}')
                self.dlg.type_keys('{F2}')                   
                time.sleep(2)                    
                self.dlg.type_keys(">"+str(self.config["marcador"]))
                time.sleep(2)
                kb.press("Enter")                    
                time.sleep(2)
                    
        self.dlg.type_keys('{TAB}')
        self.dlg.type_keys(self.cpf)
                    
        self.sleepal(1,2)
        kb.press("Enter")
        
    def sleepal(self,i,f) :
        sleep( random.randint(i, f) )

