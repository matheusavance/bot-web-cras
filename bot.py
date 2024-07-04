
import botcity.web
from shutil import move, rmtree
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

    # Using ChromeAppUtils to inspect Chrome application version
    chrome_app_utils = ChromeAppUtils()
    chrome_app_version = chrome_app_utils.get_chrome_version()

    # Target directory to store chromedriver
    driver_directory = r"C:\Users\Usuário\Desktop\code\python"

    # Create an instance of WebDriverManager
    driver_manager = WebDriverManager(driver_directory)

    # Call the main method to manage chromedriver
    driver_manager.main()

def pesquisa_cras():
    """
    Preenche o campo de pesquisa do maps com o texto 'CRAS + MUNICÍPIO'.
    """

    bot.browse("https://www.google.com/maps")
    bot.find_element('searchboxinput', By.ID).send_keys('CRAS VITÓRIA')
    bot.enter()
    bot.wait(3000)

    # Scrolldown para atualizar o número de cards
    bot.find_element('/html/body/div[1]/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/h1', By.XPATH).click()
    for i in range(20):
        bot.page_down()    

def extrai_dados_cras():
    """
    Coleta os dados de cada cras achado na pesquisa do maps.
    """

    quantidade_cras = bot.execute_javascript('return document.getElementsByClassName("hfpxzc").length')
    for item_cras in range(quantidade_cras):
        # Clica no card do CRAS
        bot.execute_javascript(f'document.getElementsByClassName("hfpxzc")[{item_cras}].click()')
        bot.wait(2000)

        # Armazena o nome do CRAS
        # nome_cras = bot.execute_javascript(f'return document.getElementsByClassName("hfpxzc")[{cras}].ariaLabel')
        nome_cras = bot.execute_javascript('return document.getElementsByClassName("DUwDvf")[0].textContent')

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
                telefone = ''
            
        # Armazena o link do maps do CRAS
        bot.execute_javascript('document.getElementsByClassName("m6QErb Pf6ghf XiKgde ecceSd tLjsW ")[0].children[4].children[0].click()')
        bot.wait(4000)
        link_maps = bot.execute_javascript('return document.getElementsByClassName("vrsrZe")[0].value')
        bot.find_element('/html/body/div[1]/div[3]/div[1]/div/div[2]/div/button', By.XPATH).click()

        # Clica em "Avaliações"
        existencia_avaliacao = bot.execute_javascript(f'return document.getElementsByClassName("yx21af lLU2pe XDi3Bc")[0]')
        if existencia_avaliacao:
            texto_botao_avaliacoes = bot.execute_javascript(f'return document.getElementsByClassName("LRkQ2")[1].textContent')
            if texto_botao_avaliacoes == 'Avaliações':
                bot.execute_javascript(f'document.getElementsByClassName("LRkQ2")[1].click()')
                bot.wait(2000)
            else:
                linha_saida_folha_cras = preenche_folha_cras(nome_cras, endereco, telefone, link_maps, quantidade_avaliacoes='Não existem avaliações', nota_cras='')
                continue  
        else:
            linha_saida_folha_cras = preenche_folha_cras(nome_cras, endereco, telefone, link_maps, quantidade_avaliacoes='Não existem avaliações', nota_cras='')
            continue
        
        # Armazena a nota e a quantidades de avaliações do CRAS
        quantidade_avaliacoes = bot.find_element('/html/body/div[1]/div[3]/div[8]/div[9]/div/div/div[1]/div[3]/div/div[1]/div/div/div[3]/div[2]/div/div[2]/div[3]', By.XPATH).text
        nota_cras = bot.execute_javascript(f'return document.getElementsByClassName("fontDisplayLarge")[0].textContent')

        # Preenche a folhas 'CRAS' da planilha com os dados extraídos
        linha_saida_folha_cras = preenche_folha_cras(nome_cras, endereco, telefone, link_maps, quantidade_avaliacoes, nota_cras) 

        # Extrai dados sobre os comentários do CRAS
        linha_saida_folha_comentarios  = extrai_dados_comentarios_cras(nome_cras)

    return nome_cras, linha_saida_folha_cras, linha_saida_folha_comentarios 
   
def extrai_dados_comentarios_cras(nome_cras):
    """
    Coleta os dados dos comentários de cada cras achado na pesquisa do maps.\n
    :param: nome_cras
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

        linha_saida_folha_comentarios = preenche_folha_comentarios(nome_cras, id_comentario, data_comentario, nota_comentario, quantidade_comentarios_usuario, texto_comentario)
        contador_childElements += 1

    return linha_saida_folha_comentarios 

def preenche_folha_cras(nome_cras, endereco, telefone, link_maps, quantidade_avaliacoes, nota_cras):
    """
    Preenche a folha 'CRAS' da planilha.\n
    :param nome_cras
    :param endereco
    :param telefone
    :param link_maps
    :param quantidade_avaliacoes
    :param nota_cras\n
    """

    # Path da planilha com os resultados da extração de dados do maps
    path_planilha_resultado = r"C:\Users\Usuário\Desktop\code\python\teste_bot\Resultado.xlsx"
    
    # Transforma as planilha em listas
    planilha_resultado = BotExcelPlugin('CRAS').read(path_planilha_resultado)
    lista_planilha_resultado = planilha_resultado.as_list()

    # Armazena o número da linha que deve ser preenchida nas planilhas
    linha_saida_folha_cras = len(lista_planilha_resultado) + 1

    # Preenche a folha 'CRAS' com os dados extraídos
    planilha_resultado.set_active_sheet('CRAS')
    planilha_resultado.set_cell("A", linha_saida_folha_cras, '',sheet="CRAS")
    planilha_resultado.set_cell("B", linha_saida_folha_cras, nome_cras, sheet="CRAS")
    planilha_resultado.set_cell("C", linha_saida_folha_cras, endereco, sheet="CRAS")
    planilha_resultado.set_cell("D", linha_saida_folha_cras, telefone, sheet="CRAS")
    planilha_resultado.set_cell("E", linha_saida_folha_cras, link_maps, sheet="CRAS")
    planilha_resultado.set_cell("F", linha_saida_folha_cras, quantidade_avaliacoes, sheet="CRAS")
    planilha_resultado.set_cell("G", linha_saida_folha_cras, nota_cras, sheet="CRAS")
    planilha_resultado.write(path_planilha_resultado)

    # Incrementa as variáveis contadoras
    linha_saida_folha_cras += 1

    return linha_saida_folha_cras

def preenche_folha_comentarios(nome_cras, id_comentario, data_comentario, nota_comentario, quantidade_comentarios_usuario, texto_comentario):
    """
    Preenche a folha 'COMENTÁRIOS' da planilha.\n
    :param nome_cras
    :param id_comentario
    :param data_comentario
    :param nota_comentario
    :param quantidade_comentarios_usuario
    :param texto_comentario\n
    """
    
    # Path da planilha com os resultados da extração de dados do maps
    path_planilha_resultado = r"C:\Users\Usuário\Desktop\code\python\teste_bot\Resultado.xlsx"
    
    # Transforma as planilha em listas
    planilha_resultado = BotExcelPlugin('COMENTÁRIOS').read(path_planilha_resultado)
    lista_planilha_resultado = planilha_resultado.as_list()

    # Armazena o número da linha que deve ser preenchida nas planilhas
    linha_saida_folha_comentarios = len(lista_planilha_resultado) + 1

    # Preenche a folha 'COMENTÁRIOS' com os dados extraídos
    planilha_resultado.set_active_sheet('COMENTÁRIOS')
    planilha_resultado.set_cell("A", linha_saida_folha_comentarios, '',sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("B", linha_saida_folha_comentarios, nome_cras, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("C", linha_saida_folha_comentarios, id_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("D", linha_saida_folha_comentarios, data_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("E", linha_saida_folha_comentarios, nota_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("F", linha_saida_folha_comentarios, quantidade_comentarios_usuario, sheet="COMENTÁRIOS")
    planilha_resultado.set_cell("G", linha_saida_folha_comentarios, texto_comentario, sheet="COMENTÁRIOS")
    planilha_resultado.write(path_planilha_resultado)
    
    # Incrementa as variáveis contadoras
    linha_saida_folha_comentarios += 1

    return linha_saida_folha_comentarios           

# Confere e atualiza/baixa o chromedriver
autoupdate_chromedriver()

def main():
    # Modo Headless
    bot.headless = False

    # Navegador usado no processo   
    bot.browser = Browser.CHROME

    # Path chromedriver
    bot.driver_path = r"C:\Users\Usuário\Desktop\code\python\chromedriver.exe"

    # Pesquisa pelos CRAS do município analisado
    pesquisa_cras()

    # Extrai os dados dos cards do maps e preenche as folhas da planilha
    extrai_dados_cras()

if __name__ == '__main__':
    main()