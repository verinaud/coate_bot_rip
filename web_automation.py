from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import random
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchFrameException
from selenium.webdriver.common.action_chains import ActionChains
import time
from pywinauto.application import Application
from pywinauto import clipboard
import controle_terminal3270
import keyboard as kb
import os
import glob
import shutil
from tela_mensagem import Mensagem as msg


class WebAutomation:
    def __init__(self, browser):
        self.browser = browser
        
           

    def default_clica_elemento_by_xpath(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        self.browser.switch_to.default_content()
        self.browser.switch_to.frame(iframe)
        self.browser.find_element(By.XPATH, path).click()
        self.browser.switch_to.default_content()

    def achar_elemento_by_xpath(self, iframe, path):
        """
        Retorna um elemento localizado por XPath dentro de um iframe.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            WebElement: O elemento encontrado.
        """
        try:
            self.browser.switch_to.frame(iframe)
            elemento = self.browser.find_element(By.XPATH, path)
            return elemento
        except (NoSuchFrameException, NoSuchElementException) as e:
            print(f"Erro ao encontrar elemento ou iframe: {e}")
        finally:
            self.browser.switch_to.default_content()

    def entrar_no_sei(self, url, login, senha, orgao): ##inclui orgao##aqui
        """
        Acessa um site usando o browser Chrome, maximiza a janela e aguarda pelo tempo especificado.
        
        Parâmetros:
        - url (str): Endereço do site que será acessado.
        - tempo_espera (int): Tempo em segundos que o método deve aguardar após acessar o site.
        
        Retorna:
        - Nenhum.
        """         
        #url = 'https://sei.economia.gov.br/sei/controlador.php?acao=andamento_marcador_cadastrar&acao_origem=andamento_marcador_gerenciar&acao_retorno=andamento_marcador_gerenciar&id_procedimento=42182876&infra_sistema=100000100&infra_unidade_atual=110008369&infra_hash=7b37406f57b4c25e2633a41fbfef355d214a023d1752e33a1af857110b9e4462'
        
        try:
            self.browser.get(url)
            self.browser.maximize_window()
            sleep(5)
            self.browser.find_element(by=By.ID, value='selOrgao').send_keys(orgao) ##aqui
            sleep(1)
            self.browser.find_element(by=By.ID, value='txtUsuario').send_keys(login)
            sleep(1)
            self.browser.find_element(by=By.ID, value='pwdSenha').send_keys(senha)
            sleep(1)
            try: # nem sempre precisa clicar no botao "Acessar" para que o login aconteca
                self.browser.find_element(by=By.ID, value='Acessar').click()
                sleep(1)
            except:
                sleep(1)
        except Exception as e:
            print(f"Erro ao acessar o site {url}: {e}")

    def clica_elemento_mouse_by_xpath_com_scroll(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe após realizar um scroll.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        # Voltar para o contexto padrão
        self.browser.switch_to.default_content()
        
        # Trocar para o iframe especificado
        self.browser.switch_to.frame(iframe)
        
        # Encontrar o elemento no iframe
        element = self.browser.find_element(By.XPATH, path)
        
        # Executar um scroll para que o elemento fique visível
        self.browser.execute_script("arguments[0].scrollIntoView();", element)
        
        # Aguardar um momento para a página rolar até o elemento
        sleep(2)
        
        # Clicar no elemento
        element.click()
        sleep(2)
        
        # Voltar para o contexto padrão
        self.browser.switch_to.default_content()

    def clica_elemento_enter_by_xpath_com_scroll(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe após realizar um scroll.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        # Voltar para o contexto padrão
        self.browser.switch_to.default_content()
        
        # Trocar para o iframe especificado
        self.browser.switch_to.frame(iframe)
        
        # Encontrar o elemento no iframe
        element = self.browser.find_element(By.XPATH, path)
        
        # Executar um scroll para que o elemento fique visível
        self.browser.execute_script("arguments[0].scrollIntoView();", element)
        
        # Aguardar um momento para a página rolar até o elemento
        sleep(2)
        
        # Clicar no elemento
        element.click()
        sleep(2)

        # Simular a tecla Enter
        action = ActionChains(self.browser)
        action.send_keys(Keys.ENTER).perform()
        
        # Voltar para o contexto padrão
        self.browser.switch_to.default_content()

    def edita_elemento_in_iframe(self, iframe_xpath, element_xpath, new_text):
        """
        Edita o texto de um elemento HTML dentro de um iframe usando Selenium.

        Args:
            iframe_xpath (str): XPath do iframe que contém o elemento.
            element_xpath (str): XPath do elemento a ser editado.
            new_text (str): Novo texto para ser definido no elemento.

        Returns:
            None
        """
        # Mudar para o iframe
        self.browser.switch_to.frame(self.browser.find_element(By.XPATH, iframe_xpath))
        
        # Encontrar o elemento
        element = self.browser.find_element(By.XPATH, element_xpath)             
        # Executar o script para alterar o texto
        self.browser.execute_script("arguments[0].innerText = arguments[1];", element, new_text)
        
        # Voltar ao conteúdo principal (opcional)
        self.browser.switch_to.default_content()

    def clica_elemento_text_by_xpath(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.default_content()
            self.browser.switch_to.frame(iframe)
            elemento = self.browser.find_element(By.XPATH, path)
            actions = ActionChains(self.browser)
            actions.move_to_element(elemento).click().perform()
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException, NoSuchFrameException) as e:
            return f"Erro ao clicar no elemento: {e}"
    
    def clica_elemento_by_xpath(self, iframe, path):
        """
        Clica em um elemento localizado por XPath dentro de um iframe.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.default_content()
            self.browser.switch_to.frame(iframe)
            self.browser.find_element(By.XPATH, path).click()
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao clicar no elemento: {e}")

    def inserir_texto_enter_by_id_in_iframe(self, iframe, element_id, text):
        """
        Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            element_id (str): ID do elemento.
            text (str): Texto a ser inserido no elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.frame(iframe)
            wait = WebDriverWait(self.browser, 15)
            # element = wait.until(EC.presence_of_element_located((By.ID, element_id)))
            # Espera explícita antes de enviar as teclas (opcional)
            element = wait.until(EC.element_to_be_clickable((By.ID, element_id)))  # Espera até que o elemento possa ser clicado
            element.send_keys(text)     
            element.send_keys(Keys.ENTER)  
            sleep(0.5)
            # element.send_keys(Keys.ENTER)  
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao inserir texto: {e}")

    def inserir_texto_enter_by_xpath_in_iframe(self, iframe, xpath, text):
        """
        Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            xpath (str): xpath do elemento.
            text (str): Texto a ser inserido no elemento.
        Returns:
            None
        """
        try:
            self.browser.switch_to.frame(iframe)
            wait = WebDriverWait(self.browser, 10)
            element = wait.until(EC.presence_of_element_located((By.ID, xpath)))
            element.clear()
            element.send_keys(text)
            sleep(2)
            element.send_keys(Keys.ENTER)
            self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao inserir texto: {e}")
            
    def inserir_texto_enter_by_id(self, element_id, text):
        """
        Insere um texto em um elemento localizado por ID e pressiona a tecla ENTER.
        Args:
            element_id (str): ID do elemento.
            text (str): Texto a ser inserido no elemento.
        Returns:
            None
        """
        try:
            wait = WebDriverWait(self.browser, 10)
            element = wait.until(EC.presence_of_element_located((By.ID, element_id)))
            element.clear()
            element.send_keys(text)
            sleep(2)
            element.send_keys(Keys.ENTER)
            self.browser.switch_to.default_content()
        except (NoSuchElementException) as e:
            print(f"Erro ao inserir texto: {e}")

    def acessar_site_chrome(self, url, tempo_espera): 
        """
        Acessa um site usando o browser Chrome, maximiza a janela e aguarda pelo tempo especificado.
        
        Parâmetros:
        - url (str): Endereço do site que será acessado.
        - tempo_espera (int): Tempo em segundos que o método deve aguardar após acessar o site.
        
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.get(url)
            self.browser.maximize_window()
            self.contagem_regressiva(tempo_espera)
        except Exception as e:
            print(f"Erro ao acessar o site {url}: {e}")

    def contagem_regressiva(self, segundos):
        """
        Realiza uma contagem regressiva de segundos exibindo mensagens.

        Args:
            segundos (int): O número de segundos para a contagem regressiva.

        Returns:
            None
        """
        try:
            for i in range(segundos, -1, -1):
                if i == 0:
                    print("Iniciando!")
                else:
                    print(f"Iniciando em {i} segundos...")
                sleep(1)
        except Exception as e:
            print(f"Erro durante a contagem regressiva: {e}")

    def open_browser(url, browser_type="firefox"):
        try:
            if browser_type == "firefox":
                driver = webdriver.Firefox()
            elif browser_type == "chrome":
                driver = webdriver.Chrome()
            else:
                print(f"Tipo de navegador {browser_type} não suportado!")
                return None

            driver.get(url)
            return driver
        except Exception as e:
            print(f"Erro ao abrir o browser: {e}")
            return None

    def tempo_aleatorio(self, inicio, fim):
        """
        Retorna um número aleatório entre os valores 'inicio' e 'fim'.
        
        Parâmetros:
        - inicio (int): Valor mínimo para a geração do número aleatório.
        - fim (int): Valor máximo para a geração do número aleatório.
        
        Retorna:
        - int: Número aleatório entre 'inicio' e 'fim'.
        """
        try:
            tempo = random.randint(inicio, fim)
            return tempo
        except Exception as e:
            print(f"Erro ao gerar tempo aleatório: {e}")
            return (inicio + fim) // 2  # retorna um valor médio entre 'inicio' e 'fim' em caso de erro
    
    def tela_aviso(self, path): 
        """
        Fecha a janela de aviso se encontrada.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            if self.browser.find_element(By.XPATH, path):  
                self.browser.find_element(By.XPATH, path).click()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao fechar janela de aviso: {e}")

    def sel_unidade(self, path): 
        """
        Seleciona uma unidade no browser.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao selecionar unidade: {e}")

    def visua_detal(self, path): 
        """
        Seleciona a visualização detalhada no controle de processos.
        Parâmetros:
        - path (str): XPath do elemento a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            sleep(self.tempo_aleatorio())
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao selecionar visualização detalhada: {e}")

    def procura_marcador(self, path): 
        """
        Seleciona a visualização por marcadores e clica no marcador especificado.
        Parâmetros:
        - path (str): XPath do marcador a ser clicado.
        Retorna:
        - Nenhum.
        """
        try:
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            sleep(self.tempo_aleatorio())
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao procurar marcador: {e}")
            
    def procura_proc_esp(self, path, processo): 
        """
        Busca por um processo específico usando o XPath fornecido.
        Parâmetros:
        - path (str): XPath do campo de pesquisa do processo.
        - processo (str): Número ou identificação do processo a ser buscado.
        Retorna:
        - Nenhum.
        """
        try:
            tempo = random.randint(3, 7)
            self.browser.implicitly_wait(30)
            self.browser.find_element(By.XPATH, path).click()
            busca_proc = self.browser.find_element(By.XPATH, path)
            busca_proc.send_keys(processo)
            sleep(tempo)
            busca_proc.send_keys(Keys.ENTER)
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao procurar processo específico: {e}")

    def clicar_elemento_tabela_html(self, iframe, valor_alvo):
        """
        Clica em um elemento HTML se ele tiver o valor alvo e estiver na mesma linha que a palavra "Órgãos".

        Parâmetros:
        - valor_alvo (str): O valor que você está procurando, como '17000'.

        Retorna:
        - bool: True se o elemento foi clicado com sucesso, False caso contrário.
        """
        try:
            self.browser.switch_to.frame(iframe)

            # XPath para o elemento com o valor 'Órgãos'
            xpath_orgaos = '//*[@id="TGROW105_10"]/td[4]/div'
            
            # XPath para o elemento com o valor alvo
            xpath_valor_alvo = f'//*[@id="TGROW105_10"]/td[2]/div[text()="{valor_alvo}"]'

            # Verificar se ambos os elementos existem e contêm os textos desejados
            if (self.browser.find_element(By.XPATH, xpath_orgaos).text == 'Órgãos' and
                self.browser.find_element(By.XPATH, xpath_valor_alvo).text == valor_alvo):
                
                # Clicar no elemento com o valor alvo
                self.browser.find_element(By.XPATH, xpath_valor_alvo).click()
                return True

            else:
                print(f"Elemento com valor {valor_alvo} não está na mesma linha que 'Órgãos'.")
                return False

        except Exception as e:
            print(f"Erro ao tentar clicar no elemento: {e}")
            return False

    def troca_habilitacao(self, iframe_elemento, path, iframe_text, element_id, text, elemento_tabela):
        self.default_clica_elemento_by_xpath(iframe_elemento, path)
        sleep(2)
        self.inserir_texto_enter_by_id_in_iframe(iframe_text, element_id, text)
        sleep(5)
        self.clica_elemento_enter_by_xpath_com_scroll('WA1', elemento_tabela)
        sleep(7)


    def encontrar_texto_marcadores(self, texto_pesquisado):
        '''
        Localiza texto na tabela de marcadores e clica na primeira coluna que é a coluna de quantidade de 
        processos para
        entrar nos marcadores do processo
        '''
        # Localiza a tabela de marcadores
        tabela_marcadores = self.browser.find_element(By.XPATH, '//*[@id="divMarcadoresAreaTabela"]')

        # Variável para rastrear se o texto foi encontrado
        texto_encontrado = False

        # Itere pelas linhas da tabela
        for linha in tabela_marcadores.find_elements(By.TAG_NAME, 'tr'):
            # Encontrar o texto nas colunas da linha
            for coluna in linha.find_elements(By.TAG_NAME, 'td'):
                if texto_pesquisado in coluna.text:
                    print(f"Texto {texto_pesquisado} encontrado na tabela")

                    # Armazena uma referência para a célula da primeira coluna
                    primeira_coluna = linha.find_elements(By.TAG_NAME, 'td')[0]
                    print(primeira_coluna)

                    # Aqui você pode realizar ações com a célula ou linha se desejar
                    # Por exemplo, clicar na célula ou linha:
                    # linha.click()
                    primeira_coluna.click()
                    texto_encontrado = True
                    break

            if texto_encontrado:
                break  # Se o texto for encontrado, saia do loop externo

        if not texto_encontrado:
            print(f"Texto {texto_pesquisado} não encontrado na tabela.")
            
    def clicar_caixaselecao_e_marcador(self):    
        '''           
        Clicar na caixa de seleção da tabela de processos recebidos e clicar no marcador da segunda coluna para
        entrar na tela de marcadores do processo. Retornar também o número do processo e o que está escrito no 
        texto do marcador CDCONVINC
        '''
        indice_da_linha = 1

        #//*[@id="P44171630"]/td[1]/div/label
        
        # Localize a tabela de processos recebidos
        tabela_processos = self.browser.find_element(By.ID, 'divRecebidosAreaTabela')

        # Encontre todas as linhas da tabela
        linhas = tabela_processos.find_elements(By.TAG_NAME, 'tr')

        time.sleep(3)        

        # Localize a célula da primeira coluna na linha desejada (por exemplo, usando <td>)
        celula_primeira_coluna = linhas[indice_da_linha].find_elements(By.TAG_NAME, 'td')[0]
        celula_primeira_coluna.click()      
                
        time.sleep(2)        
        
        # Localize a célula da segunda coluna na linha desejada (por exemplo, usando <td>)
        celula_segunda_coluna = linhas[indice_da_linha].find_elements(By.TAG_NAME, 'td')[1]
        celula_segunda_coluna.click()
        
        time.sleep(2)

        
        
    def acessar_terminal_3270(self, local):
        '''
        Entra no Terminal 3270 - Siape Hod (tela preta)
        E vai para a tela variavel local
        '''
        app = Application().connect(title_re="^Terminal 3270.*")
        dlg = app.window(title_re="^Terminal 3270.*")
        Acesso = controle_terminal3270.Janela3270()
        time.sleep(2)
        dlg.type_keys('{F3}')

        dlg.type_keys('{F2}')
        dlg.type_keys(local)
        time.sleep(2)
        kb.press("Enter")
        #dlg.type_keys('{TAB}')
        time.sleep(2)
               
    def mover_arquivos_pdf(self, pasta_origem, pasta_destino):
        """
        Move arquivos PDF da pasta de origem para a pasta de destino.
        """
        # Use a função glob para listar todos os arquivos .pdf na pasta de origem
        arquivos_pdf = glob.glob(os.path.join(pasta_origem, '*.pdf'))

        # Itere sobre os arquivos .pdf e mova cada um deles para a pasta de destino
        for caminho_arquivo_origem in arquivos_pdf:
            caminho_arquivo_destino = os.path.join(pasta_destino, os.path.basename(caminho_arquivo_origem))
            shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)

   