import base64
from PIL import Image
from io import BytesIO
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table
from reportlab.lib.units import inch
from image64 import Image64

class Manual:
    def __init__(self):
        lista_imagens = []
        lista_imagens = Image64().get_imagens()
        
        pdf_name = "Manual-CoateBotRip.pdf"
        
        self.create_pdf(pdf_name, lista_imagens)
        
        
        
    def addpage01(self, c, lista_imagens):
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[0])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image.png"
        image.save(image_path, format='PNG')
        
        # Definir o tamanho desejado para a imagem
        width = 6 * inch  # Por exemplo, 4 polegadas
        height = 5 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(0.5)
        c.drawImage(image_path, 90, 280, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)

        
        # Adicionar texto ao PDF
        texto_capa = [
            "MANUAL DO USUÁRIO",
            "COATE BOT RIP",
            "COMUNICADO DE FALECIMENTO"
        ]
        c.setFont("Helvetica", 15)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto_capa:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto_capa) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 690 - texto_height - 20  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)

        # Remover a imagem temporária
        os.remove(image_path)

    def addpage02(self, c):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF (segunda página)
        sumario_texto = "SUMÁRIO"
        c.setFont("Helvetica", 15)
        c.drawString(50, 750, sumario_texto)
        
        y_position = 740  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        
        # Adicionar linha horizontal abaixo do texto
        c.line(50, 730, 500, 730)
        
        # Definir a fonte e o tamanho do texto
        c.setFont("Helvetica", 12)        
        
        # Definir os dados da tabela
        data = [
            ["Sobre a Automação", ".....................................................pág. 02"],
            ["Iniciando o Software", ".....................................................pág. 03"],
            ["Janela Principal", ".....................................................pág. 04"],
            ["Iniciando a Automação", ".....................................................pág. 05"],
            ["Painel de Eventos", ".....................................................pág. 06"],
            ["Editar Lista de Códigos com Vínculos Decipex", ".....................................................pág. 07"],
            ["Editar Lista de Códigos Excluído Batimento", ".....................................................pág. 08"],
            ["Editar Lista de Códigos Excluído Falecimento", ".....................................................pág. 09"],
            ["Editar Lista de Marcadores Substitutos", ".....................................................pág. 10"],
            ["Editar Outros", ".....................................................pág. 11"]           
        ]         
        
        # Criar a tabela
        table = Table(data)        

        # Definir a posição da tabela na página
        table.wrapOn(c, 0, 0)
        table.drawOn(c, 50, 550)
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Sobre a Automação"
        ]
        c.setFont("Helvetica", 15)
        y_position = 400  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 400 - texto_height - 10  # Deslocamento adicional
        print("-----------------------")
        print(x_position,y_position,"500",linha_position)
        print("-----------------------")
        c.line(x_position, linha_position, 500, linha_position)      
        
        # Adicionar texto ao PDF
        texto = [
            "    A presente automação tem por objetivo substituir os marcadores 'CDCONVINC' dos processos",
            "de comunicado de falecimento no portal SEI - Sistema Eletrônico de Informações - por marcadores",
            "'CADASTRO ATIVO', 'EXCLUÍDO BATIMENTO', 'EXCLUÍDO FALECIMENTO', 'SEM VÍNCULO' e",
            "'VERIFICAÇÃO MANUAL' e anexar um documento PDF ao processo com as telas necessárias copiadas",
            "do SIAPE - Terminal 3270.",
            "    Para substituir os marcadores a automação acessa os processos no portal SEI, copia o CPF do",
            "servidor/pensionista no marcador CDCONVINC, acessa o Terminal 3270 e verifica vínculo, código de exclusão",
            "e certidão de óbito.",
            
            #os marcadores 'CDCONVINC' dos processos de comunicado de falecimento no portal SEI
        ]
        
        c.setFont("Helvetica", 11)
        y_position = 350  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
    
    def addpage03(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Iniciando o Programa"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    Ao clicar duas vezes no ícone Coate-Bot dois arquivos serão gerados: config.json e Manual.pdf (Fig. 1).",
            "    O arquivo config.json armazena as configurações que podem ser alteradas no programa e caso seja",
            "excluído o programa criará um novo arquivo com a configuração padrão. Neste caso as alterações feitas",
            "serão descartadas.",
            "    O Manual.pdf é o manual do usuário com instruções sobre o programa. Caso seja excluído um novo será",
            "gerado assim que o programa for executado."
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[1])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image1.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 4 * inch  # Por exemplo, 4 polegadas
        height = 3 * inch  # Por exemplo, 3 polegadas  
        
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 50, 350, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)

        # Remover a imagem temporária
        os.remove(image_path)
        
    def addpage04(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Janela Principal"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O programa inicia na janela Painel Home (Fig. 2) onde é possível acessar o painel de menus e iniciar",
            "a automação.",
            "    O painel de menus contém os menus Exibir, Editar e Ajuda.",
            "    O menu Exibir permite visualizar o Painel Home e o Painel de Eventos. O menu Editar permite",
            "incluir e excluir códigos de vínculos Decipex, códigos Excluído Batimento e Excluído Falecimento. O",
            "menu Ajuda permite acessar os sub-menus Sobre e Manual do Usuário.",
            "    O Painel Home contém os campos de Usuário e Senha SEI, seleção de Unidade, botão Exibir senha",
            "e o botão Inciar."
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[2])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image2.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)        
        
        # Remover a imagem temporária
        os.remove(image_path)
        
    def addpage05(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Iniciando a Automação"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    Para iniciar a automação, digite seu usuário e senha SEI. O botão Exibir ao lado do campo",
            "senha permite visualizar a senha digitada.",
            "    A seleção de Unidade, por padrão, estará selecionada MGI e poderá ser alterada",
            "conforme o cadastro do usuário no portal SEI.",
            "    Após ter preenchido os campos Usuário e Senha clique no botão Iniciar. O painel de",
            "Eventos será exibido e por questões de segurança a senha não poderá ser alterada",
            "enquanto o programa estiver em execução.",
            "    Ao clicar no botão Iniciar o programa abrirá o portal siapenet e o usuário deverá",
            "selecionar seu certificado e digitar o código PIN, assim que for solicitado. A seguir a automação",
            "realizará as tarefas automaticamente.",
            "    IMPORTANTE!",
            "    Não é possível utilizar o computador durante o uso da automação!",
            "    Ao final será exibido a mensagem 'Terminado' no Painel de Eventos e o computador",
            "estará liberado para o uso.",
            "    Caso necessite interromper a automação feche manualmente o portal SEI e o Terminal 3270, a automação",
            "será interrompida com mensagem de erro."
            
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        
    def addpage06(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Painel de Eventos"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O Painel de Eventos (Fig.3) é exibido assim que o botão Iniciar é clicado e também pode ser",
            "mostrado através do menu Exibir - Painel de Eventos.",
            "    Nele são exibidos todos os eventos que ocorrem durante a execução da automação. Possíveis erros",
            "também aparecerão no Painel de Eventos, portanto se a automação ficar longo tempo paralizada,",
            "verifique o painel e siga as instruções que aparecerão.",
            "    Sempre que a automação terminar ou em caso de erro o Painel de Eventos exibirá o número do processo",
            "e o CPF. Basta clicar sobre qualquer deles e será copiado para a área de transferência permitindo ao",
            "usuário colar utilizado as teclas Ctrl + V."          
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[3])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image3.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)
        
        
        # Remover a imagem temporária
        os.remove(image_path)
    
    def addpage07(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Editar Lista de Códigos com vínculos Decipex"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O Painel Vínculos Decipex (Fig.4) permite inserir ou Excluir um código de vínculo.",
            "    Para acessar o painel clique no menu Editar - Vínculos Decipex.",
            "    Para inserir novo vínculo digite o novo código no campo 'Digite novo vínculo Decipex' e clique",
            "no botão Inserir.",
            "    Para excluir um vínculo, selecione o código na caixa de seleção 'Selecione um vínculo para",
            "excluir' e clique no botão Excluir.",
            "    As alterações feitas serão gravadas no arquivo config.json. Caso o arquivo config.json seja excluído",
            "as alterações serão descartadas e novo arquivo será criado com a configuração padrão."         
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[4])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image4.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)        
        
        # Remover a imagem temporária
        os.remove(image_path)    
        
    
        
    def addpage08(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Editar Lista de Códigos Excluído Batimento "
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O Painel Excluído Batimento (Fig.5) permite inserir ou excluir um código de excluído batimento.",
            "    Para acessar o painel clique no menu Editar - Excluído Batimento.",
            "    Para inserir novo código excluído batimento digite o novo código no campo 'Digite novo código para'",
            "Excluído Batimento e clique no botão Inserir.",
            "    Para excluir um código excluído batimento, selecione o código na caixa de seleção 'Selecione um",
            "código para excluir' e clique no botão Excluir.",
            "    As alterações feitas serão gravadas no arquivo config.json. Caso o arquivo config.json seja excluído",
            "as alterações serão descartadas e novo arquivo será criado com a configuração padrão."         
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[5])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image5.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)
        
        
        # Remover a imagem temporária
        os.remove(image_path)    
        
    
        
    def addpage09(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Editar Lista de Códigos Excluído Falecimento"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O Painel Excluído Falecimento (Fig.6) permite inserir ou excluir um código de excluído falecimento.",
            "    Para acessar o painel clique no menu Editar - Excluído Falecimento.",
            "    Para inserir novo código excluído falecimento digite o novo código no campo 'Digite novo código para'",
            "Excluído Falecimento e clique no botão Inserir.",
            "    Para excluir um código excluído falecimento, selecione o código na caixa de seleção 'Selecione um",
            "código para excluir' e clique no botão Excluir.",
            "    As alterações feitas serão gravadas no arquivo config.json. Caso o arquivo config.json seja excluído",
            "as alterações serão descartadas e novo arquivo será criado com a configuração padrão."         
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[6])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image6.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)
        
        
        # Remover a imagem temporária
        os.remove(image_path)
        
    def addpage10(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Editar Lista de Marcadores Substitutos"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O Painel Marcadores (Fig.7) permite inserir ou excluir um marcador substituto.",
            "    Para acessar o painel clique no menu Editar - Marcadores Substitutos.",
            "    Para inserir novo marcador substituto digite o novo marcador no campo 'Digite novo marcador",
            "substituto e clique no botão Inserir.",
            "    Para excluir um marcador substituto, selecione o código na caixa de seleção 'Selecione um",
            "marcador para excluir' e clique no botão Excluir.",
            "    As alterações feitas serão gravadas no arquivo config.json. Caso o arquivo config.json seja excluído",
            "as alterações serão descartadas e novo arquivo será criado com a configuração padrão."         
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[7])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image8.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 230, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)        
        
        # Remover a imagem temporária
        os.remove(image_path)
        
    def addpage11(self, c, lista_imagens):
        # Adicionar nova página
        c.showPage()
        
        # Adicionar texto ao PDF
        subtitulo = [
            "Editar Outros"
        ]
        c.setFont("Helvetica", 15)
        y_position = 750  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in subtitulo:
            c.drawString(x_position, y_position, line)
            y_position -= 10  # Deslocamento para a próxima linha
            
        texto_height = len(subtitulo) * 10

        # Adicionar linha horizontal abaixo do texto
        linha_position = 750 - texto_height - 10  # Deslocamento adicional
        c.line(x_position, linha_position, 500, linha_position)
        
        # Adicionar texto ao PDF
        texto = [
            "    O painel Outros (Fig.8) permite alterar a URL do SEI, a URL di SiapeNet e a unidade SEI.",
            "    Para alterar qualquer campo sobrescreva o conteúdo da caixa de texto a ser alterada e ",
            "clique no botão Alterar."         
        ]
        c.setFont("Helvetica", 11)
        y_position = 700  # Posição inicial do texto
        x_position = 50   # Recuo à esquerda
        for line in texto:
            c.drawString(x_position, y_position, line)
            y_position -= 20  # Deslocamento para a próxima linha
            
        texto_height = len(texto) * 10
        
        
        # Criar os dados da imagem a partir da string de base64
        image_data = base64.b64decode(lista_imagens[8])
        image = Image.open(BytesIO(image_data))

        # Salvar a imagem temporariamente
        image_path = "temp_image9.png"
        image.save(image_path, format='PNG')

        # Definir o tamanho desejado para a imagem
        width = 5 * inch  # Por exemplo, 4 polegadas
        height = 4 * inch  # Por exemplo, 3 polegadas
        
        c.setFillAlpha(1)
        c.drawImage(image_path, 20, 330, width=width, height=height, preserveAspectRatio=True)
        c.setFillAlpha(1)
        
        
        # Remover a imagem temporária
        os.remove(image_path)    
        
    def create_pdf(self, pdf_name, lista_imagens):
        # Criar um novo documento PDF
        c = canvas.Canvas(pdf_name, pagesize=letter)

        # Adicionar a primeira página
        self.addpage01(c, lista_imagens)

        # Adicionar a segunda página
        self.addpage02(c)
        
        # Adicionar a terceira página
        self.addpage03(c, lista_imagens)
        
        # Adicionar a quarta página
        self.addpage04(c, lista_imagens)
        
        # Adicionar a quinta página
        self.addpage05(c, lista_imagens)
        
        # Adicionar a sexta página
        self.addpage06(c, lista_imagens)
        
        # Adicionar a setima página
        self.addpage07(c, lista_imagens)
        
        # Adicionar a oitava página
        self.addpage08(c, lista_imagens)
        
        # Adicionar a nona página
        self.addpage09(c, lista_imagens)
        
        # Adicionar a décima página
        self.addpage10(c, lista_imagens)
        
        # Adicionar a décima primeira página
        self.addpage11(c, lista_imagens)

        # Salvar o PDF
        c.save()

if __name__ == "__main__":
    app = Manual()
