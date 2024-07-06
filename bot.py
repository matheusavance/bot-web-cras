
import botcity.web
from shutil import move, rmtree
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.styles import PatternFill
from webdriver_auto_update.chrome_app_utils import ChromeAppUtils
from webdriver_auto_update.webdriver_manager import WebDriverManager
from botcity.web.browsers.chrome import default_options
from botcity.web import WebBot, Browser, By
from botcity.plugins.excel import BotExcelPlugin
from botcity.web.browsers.chrome import default_options

# Instância WebBot
bot = WebBot()

def autoupdate_chromedriver():
    """
    Baixa/atualiza o arquivo chromedriver baseado na versão do chrome da máquina que está rodando o código.
    """

    # Usando ChromeAppUtils para inspecionar a versão do Chrome
    chrome_app_utils = ChromeAppUtils()
    chrome_app_version = chrome_app_utils.get_chrome_version()

    # Diretório de destino para armazenar o chromedriver
    driver_directory = r"C:\Users\Usuário\Desktop\code\python"

    # Instância WebDriverManager
    driver_manager = WebDriverManager(driver_directory)

    # Executa método main para gerenciar o chromedriver
    driver_manager.main()

def pesquisa_cras():
    """
    Preenche o campo de pesquisa do maps com o texto 'CRAS + MUNICÍPIO'.
    """

    bot.browse("https://www.google.com/maps")
    bot.find_element('searchboxinput', By.ID).send_keys('CRAS VITÓRIA')
    bot.enter()

    # Scrolldown para atualizar o número de cards
    texto_resultado = bot.find_element('/html/body/div[1]/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/h1', By.XPATH)
    bot.wait_for_element_visibility(texto_resultado)
    texto_resultado.click()
    for i in range(20):
        bot.page_down()    

def extrai_dados_cras(path_planilha):
    """
    Coleta os dados de cada cras achado na pesquisa do maps.\n
    :param path_planilha
    @return id_cras, nome, linha_saida_folha_cras, linha_saida_folha_comentarios 
    """

    # Inicializa a variável id_cras
    id_cras = 1
    
    # Extrai dados do card do CRAS
    quantidade_cras = bot.execute_javascript('return document.getElementsByClassName("hfpxzc").length')
    for item_cras in range(quantidade_cras):
        # Clica no card do CRAS
        bot.execute_javascript(f'document.getElementsByClassName("hfpxzc")[{item_cras}].click()')
        bot.wait(4000)

        # Armazena o nome do CRAS
        nome = bot.execute_javascript('return document.getElementsByClassName("DUwDvf")[0].textContent')

        # Armazena o endereço do CRAS
        endereco = bot.execute_javascript('return document.getElementsByClassName("Io6YTe fontBodyMedium kR99db")[0].textContent')

        # Armazena o telefone do CRAS
        quantidade_dados_card = bot.execute_javascript('return document.getElementsByClassName("CsEnBe").length')
        for i in range(quantidade_dados_card):
            texto_botao_card = bot.execute_javascript(f'return document.getElementsByClassName("CsEnBe")[{i}].ariaLabel')
            if 'Telefone' in texto_botao_card:
                telefone = texto_botao_card
                break
            else:
                telefone = 'Sem telefone'
            
        # Armazena o link do maps do CRAS
        bot.execute_javascript('document.getElementsByClassName("m6QErb Pf6ghf XiKgde ecceSd tLjsW ")[0].children[4].children[0].click()')
        link_maps = bot.find_element('/html/body/div[1]/div[3]/div[1]/div/div[2]/div/div[2]/div/div/div/div[3]/div[2]/div[2]/input', By.XPATH)
        bot.wait_for_element_visibility(link_maps)
        link_maps = bot.execute_javascript('return document.getElementsByClassName("vrsrZe")[0].value')
        bot.find_element('/html/body/div[1]/div[3]/div[1]/div/div[2]/div/button', By.XPATH).click()

        # Clica em "Avaliações"
        existencia_avaliacao = bot.execute_javascript(f'return document.getElementsByClassName("yx21af lLU2pe XDi3Bc")[0]')
        if existencia_avaliacao:
            texto_botao_avaliacoes = bot.execute_javascript(f'return document.getElementsByClassName("LRkQ2")[1].textContent')
            if texto_botao_avaliacoes == 'Avaliações':
                bot.execute_javascript(f'document.getElementsByClassName("LRkQ2")[1].click()')
                bot.wait(4000)
            else:
                linha_saida_folha_cras, id_cras = preenche_folha_cras(id_cras, nome, endereco, telefone, link_maps, path_planilha, quantidade_avaliacoes='Não existem avaliações', nota_cras='Sem nota')
                id_cras += 1
                continue  
        else:
            linha_saida_folha_cras, id_cras = preenche_folha_cras(id_cras, nome, endereco, telefone, link_maps, path_planilha, quantidade_avaliacoes='Não existem avaliações', nota_cras='Sem nota')
            id_cras += 1
            continue
        
        # Armazena a nota e a quantidades de avaliações do CRAS
        quantidade_avaliacoes = bot.find_element('/html/body/div[1]/div[3]/div[8]/div[9]/div/div/div[1]/div[3]/div/div[1]/div/div/div[3]/div[2]/div/div[2]/div[3]', By.XPATH).text
        nota_cras = bot.execute_javascript(f'return document.getElementsByClassName("fontDisplayLarge")[0].textContent')
        
        # Preenche a folha 'CRAS' da planilha com os dados extraídos
        linha_saida_folha_cras, id_cras = preenche_folha_cras(id_cras, nome, endereco, telefone, link_maps, path_planilha, quantidade_avaliacoes, nota_cras) 

        # Extrai dados e preenche a folha 'COMENTÁRIOS' da planilha
        linha_saida_folha_comentarios, id_cras  = extrai_dados_comentarios_cras(id_cras, nome, path_planilha)

    return id_cras, nome, linha_saida_folha_cras, linha_saida_folha_comentarios 
   
def extrai_dados_comentarios_cras(id_cras, nome, path_planilha):
    """
    Coleta os dados dos comentários de cada cras achado na pesquisa do maps.\n
    :param: id_cras
    :param: nome
    :param: path_planilha\n
    @return linha_saida_folha_comentarios, id_cras
    """

    # Scrolldown para atualizar a quantidade de comentários do CRAS analisado
    bot.find_element('/html/body/div[1]/div[3]/div[8]/div[9]/div/div/div[1]/div[3]/div/div[1]/div/div/div[3]/div[2]', By.XPATH).click()
    for i in range(20):
        bot.page_down()

    # Armazena os dados sobre os comentários
    quantidade_comentarios = bot.execute_javascript('return document.getElementsByClassName("al6Kxe").length')
    contador_childElements = 0
    for comentario in range(quantidade_comentarios):
        id_comentario = bot.execute_javascript(f'return document.getElementsByClassName("d4r55")[{comentario}].textContent')
        data_comentario = bot.execute_javascript(f'return document.getElementsByClassName("rsqaWe")[{comentario}].textContent')
        nota_comentario = bot.execute_javascript(f'return document.getElementsByClassName("kvMYJc")[{comentario}].ariaLabel')
        
        quantidade_comentarios_usuario = bot.execute_javascript(f'return document.getElementsByClassName("al6Kxe")[{comentario}].childElementCount')
        if quantidade_comentarios_usuario == 1: 
            quantidade_comentarios_usuario = 'Nenhuma avaliação'
        else:
            quantidade_comentarios_usuario = bot.execute_javascript(f'return document.getElementsByClassName("al6Kxe")[{comentario}].children[1].textContent')
        # 
        try:            
            texto_comentario = bot.execute_javascript(f'return document.getElementsByClassName("wiI7pd")[{comentario}].textContent')
        except:
            texto_comentario = 'Não existem comentários'

        linha_saida_folha_comentarios, id_cras = preenche_folha_comentarios(id_cras, nome, id_comentario, data_comentario, nota_comentario, path_planilha, quantidade_comentarios_usuario, texto_comentario)
        contador_childElements += 1

    # Incrementa a variável id_cras
    id_cras += 1

    return linha_saida_folha_comentarios, id_cras

def preenche_folha_cras(id_cras, nome, endereco, telefone, link_maps, path_planilha, quantidade_avaliacoes, nota_cras):
    """
    Preenche a folha 'CRAS' da planilha.\n
    :param id_cras
    :param nome
    :param endereco
    :param telefone
    :param link_maps
    :param quantidade_avaliacoes
    :param nota_cras\n
    @return linha_saida_folha_cras, id_cras
    """
    
    # Transforma as planilha em listas
    planilha_resultado = BotExcelPlugin('CRAS').read(path_planilha)
    lista_planilha_resultado = planilha_resultado.as_list()

    # Armazena o número da linha que deve ser preenchida nas planilhas
    linha_saida_folha_cras = len(lista_planilha_resultado) + 1

    # Preenche a folha 'CRAS' com os dados extraídos
    planilha_resultado.set_active_sheet('CRAS')
    planilha_resultado.set_cell("A", linha_saida_folha_cras, id_cras,sheet="CRAS")
    planilha_resultado.set_cell("B", linha_saida_folha_cras, nome, sheet="CRAS")
    planilha_resultado.set_cell("C", linha_saida_folha_cras, endereco, sheet="CRAS")
    planilha_resultado.set_cell("D", linha_saida_folha_cras, telefone, sheet="CRAS")
    planilha_resultado.set_cell("E", linha_saida_folha_cras, link_maps, sheet="CRAS")
    planilha_resultado.set_cell("F", linha_saida_folha_cras, quantidade_avaliacoes, sheet="CRAS")
    planilha_resultado.set_cell("G", linha_saida_folha_cras, nota_cras, sheet="CRAS")
    planilha_resultado.write(path_planilha)

    # Incrementa as variáveis contadoras
    linha_saida_folha_cras += 1

    return linha_saida_folha_cras, id_cras

def preenche_folha_comentarios(id_cras, nome, id_comentario, data_comentario, nota_comentario, path_planilha, quantidade_comentarios_usuario, texto_comentario):
    """
    Preenche a folha 'COMENTÁRIOS' da planilha.\n
    :param id_cras
    :param nome
    :param id_comentario
    :param data_comentario
    :param nota_comentario
    :param quantidade_comentarios_usuario
    :param texto_comentario\n
    @return linha_saida_folha_comentarios, id_cras 
    """
    
    # Transforma as planilha em listas
    planilha_resultado = BotExcelPlugin('COMENTÁRIOS').read(path_planilha)
    lista_planilha_resultado = planilha_resultado.as_list()

    # Armazena o número da linha que deve ser preenchida nas planilhas
    linha_saida_folha_comentarios = len(lista_planilha_resultado) + 1

    # Preenche a folha 'COMENTÁRIOS' com os dados extraídos
    planilha_resultado.set_active_sheet('COMENTÁRIOS')
    planilha_resultado.set_cell("A", linha_saida_folha_comentarios, id_cras,sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("B", linha_saida_folha_comentarios, nome, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("C", linha_saida_folha_comentarios, id_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("D", linha_saida_folha_comentarios, data_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("E", linha_saida_folha_comentarios, nota_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("F", linha_saida_folha_comentarios, quantidade_comentarios_usuario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("G", linha_saida_folha_comentarios, texto_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.write(path_planilha)
    
    # Incrementa as variáveis contadoras
    linha_saida_folha_comentarios += 1

    return linha_saida_folha_comentarios, id_cras         

def estiliza_planilha(path_planilha):
    """
    Estiliza as folhas da planilha passada.
    :param path_planilha
    """
    planilha = load_workbook(filename=path_planilha)
    quantidade_folhas_planilha = len(planilha.sheetnames)

    # Define o estilo usados
    fundo_preto = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    texto_branco_negrito = Font(color="FFFFFF", bold=True)
    borda_fina = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Aplica os estilos células que possuem conteúdo
    for numero_folha in range(quantidade_folhas_planilha):
        planilha.active = numero_folha
        folha = planilha.active

        # Header
        for celula in folha[1]:
            celula.fill = fundo_preto
            celula.font = texto_branco_negrito

        # Células
        for linha in folha.rows:
            for celula in linha:
                if celula.value is not None:
                    celula.border = borda_fina

    planilha.save(path_planilha)

# Confere e atualiza/baixa o chromedriver
autoupdate_chromedriver()

def main():
    # Modo Headless
    bot.headless = False

    # Navegador usado no processo   
    bot.browser = Browser.CHROME

    # Path chromedriver
    bot.driver_path = r"C:\Users\Usuário\Desktop\code\python\chromedriver.exe"

    # Path planilha
    path_planilha = r"C:\Users\Usuário\Desktop\code\python\bots\planilhas\Resultado.xlsx"

    # Pesquisa pelos CRAS do município analisado
    pesquisa_cras()

    # Extrai os dados dos cards do maps e preenche as folhas da planilha
    extrai_dados_cras(path_planilha)

    # Estiza as folhas da planilha
    estiliza_planilha(path_planilha)

if __name__ == '__main__':
    main()