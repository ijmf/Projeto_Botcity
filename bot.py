"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""


# Import for the Web Bot
from aifc import Error
import shutil
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from botcity.web.util import element_as_select
from botcity.web.parsers import table_to_dict
from botcity.plugins.excel import BotExcelPlugin 
from botcity.plugins.email import BotEmailPlugin
from pandas import *

# Instanciar o plug -in
email = BotEmailPlugin()

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

excel = BotExcelPlugin()
excel.add_row(["CIDADE", "POPULAÇÃO"])

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    #Se executar pelo VScode comentar o trecho abaixo, executando pelo maestro necessário descomentar.

    maestro.login(server="https://developers.botcity.dev", 
    login = "57444048-4a34-432e-985f-88d6252065f1", 
    key="574_SFXHGJ4TTVUWBDXFN6ES")

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    # Obtendo credenciais do Maestro
    # usuario = maestro.get_credential("dados-login", "usuario")
    # senha = maestro.get_credential("dados-login", "senha")

    # Enviando alerta para o Maestro
    maestro.alert(
    task_id=execution.task_id,
    title="Iniciando processo",
    message=f"O processo de consulta foi iniciado",
    alert_type=AlertType.INFO
    )

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Chrome
    bot.browser = Browser.CHROME
    
    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\Treinamento BotCity\chromedriver-win64\chromedriver.exe"

    # Abrimos o site do Busca CEP.
    bot.browse("https://buscacepinter.correios.com.br/app/faixa_cep_uf_localidade/index.php")

    # Captura e a seleção do Estado de GO
    drop_uf = element_as_select(bot.find_element("//select[@id='uf']", By.XPATH))
    drop_uf.select_by_value("GO")

    # Clicamos no Botão de Pesquisar
    btn_pesquisar = bot.find_element("//button[@id='btn_pesquisar']", By.XPATH)
    btn_pesquisar.click()

    bot.wait(3000)

    # Captura da tabela de dados com os nomes das cidades
    table_dados = bot.find_element("//table[@id='resultado-DNEC']", By.XPATH)
    table_dados = table_to_dict(table=table_dados) #Transformação da tabela em um dicionário

    # Navegação para o site do IBGE
    bot.navigate_to ("https://cidades.ibge.gov.br/brasil/sp/panorama")


    int_Contador = 1
    str_CidadeAnterior = ""

    for cidade in table_dados:

        str_Cidade = cidade["localidade"] #Definimos o nome da cidade da vez

        if str_CidadeAnterior == str_Cidade:
            continue

        if int_Contador <=3:

            campo_pesquisa = bot.find_element("//input[@placeholder='O que você procura?']", By.XPATH)
            campo_pesquisa.send_keys(str_Cidade)

            opcao_cidade = bot.find_element(f"//a[span[contains(text(), '{str_Cidade}')] and span[contains(text(), 'GO')]]", By.XPATH)
            bot.wait(1000)
            opcao_cidade.click()
            
            bot.wait(2000)

            populacao = bot.find_element("//div[@class='indicador__valor']", By.XPATH)
            str_populacao = populacao.text

            print(str_Cidade, str_populacao)
            excel.add_row([str_Cidade, str_populacao])

            maestro.new_log_entry(
                activity_label="CIDADES", 
                values={"CIDADE": f"{str_Cidade}", 
                        "POPULACAO": f"{str_populacao}"})

            int_Contador = int_Contador + 1
            str_CidadeAnterior = str_Cidade

        else:
            print("Número de cidade já alcançada")
            break

    excel.write(r"C:\Treinamento BotCity\Projetos\RelatorioCidades\Infos_Cidades.xlsx")

    # Configure IMAP com o servidor Hotmail
    try:
        email.configure_imap("outlook.office365.com", 993)

    # Configure SMTP com o servidor Hotmail
        email.configure_smtp("smtp-mail.outlook.com", 587) #smtp.office365.com ou smtp-mail.outlook.com

    # Faça login com uma conta de email válida
        email.login("junio.str@hotmail.com", "Teste123@")

    except Exception as e:
        print(f"Erro durante a configuração do e-mail: {e}")

    # Definindo os atributos que comporão a mensagem
    para = ["junio.str@hotmail.com"]
    assunto = "Planilha Info Cidades"
    corpo_email = ""
    arquivos = [r"C:\Treinamento BotCity\Projetos\RelatorioCidades\Infos_Cidades.xlsx"]

    # Enviando a mensagem de e -mail
    email.send_message(assunto, corpo_email, para, attachments=arquivos, use_html=True)

    # Feche a conexão com os servidores IMAP e SMTP
    email.disconnect()

    # Subindo arquivo de resultados
    caminho_arquivo_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioCidades\Infos_Cidades.xlsx"
    caminho_pasta_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioCidades"
    shutil.make_archive(caminho_arquivo_xlsx, 'zip', caminho_pasta_xlsx)
    maestro.post_artifact(
    task_id=execution.task_id,
    artifact_name="Infos_Cidades",
    filepath=caminho_arquivo_xlsx + ".zip"
    )

    # Alerta de email
    maestro.alert(
    task_id=execution.task_id,
    title="E-mail OK",
    message=f"E-mail enviado com sucesso",
    alert_type=AlertType.INFO
            )

    # Implement here your logic...
    ...

    # Wait 3 seconds before closing
    bot.wait(5000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Reportando erro ao Maestro
    # maestro.error(
    # task_id=execution.task_id,
    # exception=erro,
    # screenshot="erro.png"
    #)

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
         task_id=execution.task_id,
         status=AutomationTaskFinishStatus.SUCCESS,
         message="Task Finalizada"
     )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
