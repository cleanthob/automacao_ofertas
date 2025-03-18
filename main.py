from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import pandas as pd


df_produtos = pd.read_excel("buscas.xlsx")
print(df_produtos)


driver = webdriver.Chrome()
driver.maximize_window()
driver.implicitly_wait(5)


def buscar_google_shopping(
    navegador, produto, termos_banidos, preco_minimo, preco_maximo
):
    navegador.get("https://www.google.com/search?tbm=shop")

    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    lista_termos_produto = produto.split(" ")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "q")))

    navegador.find_element(By.NAME, "q").send_keys(produto)
    navegador.find_element(By.NAME, "q").send_keys(Keys.ENTER)

    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.NAME, "lower"))
    )

    campo_preco_minimo = navegador.find_element(By.NAME, "lower")
    campo_preco_minimo.send_keys(preco_minimo)

    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.NAME, "upper"))
    )

    campo_preco_maximo = navegador.find_element(By.NAME, "upper")
    campo_preco_maximo.send_keys(preco_maximo)

    botao_filtrar = navegador.find_element(By.CSS_SELECTOR, "button.sh-dr__prs")
    botao_filtrar.click()

    resultados = navegador.find_elements(By.CLASS_NAME, "sh-dgr__content")
    lista_ofertas = []
    print("Quantidade encontrada google shopping:", len(resultados))

    sleep(5)

    for resultado in resultados:
        nome = resultado.find_element(By.CLASS_NAME, "tAxDx").text
        nome = nome.lower()
        preco = resultado.find_element(By.CLASS_NAME, "a8Pemb").text
        link = resultado.find_element(By.CSS_SELECTOR, "a.shntl").get_attribute("href")

        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True

        tem_termos_produto = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_termos_produto = False

        if tem_termos_banidos == False and tem_termos_produto == True:
            try:
                preco = (
                    preco.replace("R$", "")
                    .replace(" ", "")
                    .replace(".", "")
                    .replace(",", ".")
                )
                preco = float(preco)

                lista_ofertas.append((nome, preco, link))
            except:
                continue
    print("Lista ofertas google shopping:", lista_ofertas)
    return lista_ofertas


def buscar_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    navegador.get("https://www.buscape.com.br/")

    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.CLASS_NAME, "AutoCompleteStyle_input__WAC2Y")
        )
    )

    navegador.find_element(By.CLASS_NAME, "AutoCompleteStyle_input__WAC2Y").send_keys(
        produto
    )
    navegador.find_element(By.CLASS_NAME, "AutoCompleteStyle_input__WAC2Y").send_keys(
        Keys.ENTER
    )

    WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".AllFiltersButton_AllFilters___ayQd > button")
        )
    )

    botao_filtro_preco = navegador.find_element(
        By.CSS_SELECTOR, ".AllFiltersButton_AllFilters___ayQd > button"
    )

    botao_filtro_preco.click()

    # Insere o preço mínimo no filtro
    campo_preco_minimo = navegador.find_element(
        By.CSS_SELECTOR,
        "div.Input_InputWrapper__rY6ff:nth-child(1) > div:nth-child(1) > input:nth-child(1)",
    )
    campo_preco_minimo.send_keys(Keys.CONTROL + "a")
    campo_preco_minimo.send_keys(Keys.DELETE)
    campo_preco_minimo.send_keys(str(preco_minimo))
    # Insere o preço máximo no filtro

    campo_preco_maximo = navegador.find_element(
        By.CSS_SELECTOR,
        "div.Input_InputWrapper__rY6ff:nth-child(3) > div:nth-child(1) > input:nth-child(1)",
    )
    campo_preco_maximo.send_keys(Keys.CONTROL + "a")
    campo_preco_maximo.send_keys(Keys.DELETE)
    campo_preco_maximo.send_keys(str(preco_maximo))

    # Aplica o filtro
    navegador.find_element(
        By.CSS_SELECTOR, "button.Button_Button__zn4eq:nth-child(4)"
    ).click()
    navegador.find_element(
        By.CSS_SELECTOR, "button.IconButton_IconButton__small__Lfh6U:nth-child(2)"
    ).click()

    sleep(5)

    resultados = navegador.find_elements(By.CLASS_NAME, "Hits_ProductCard__Bonl_")
    print("Quantidade encontrada buscape:", len(resultados))
    lista_ofertas = []

    for resultado in resultados:
        try:
            nome = resultado.find_element(
                By.CSS_SELECTOR, "div > a > div > div > div > div > h2"
            ).text
            nome = nome.lower()
            preco = resultado.find_element(
                By.CSS_SELECTOR, "div > a > div > div > div > p"
            ).text
            link = resultado.find_element(By.CSS_SELECTOR, "div > a").get_attribute(
                "href"
            )

            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True

            tem_termos_produto = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_termos_produto = False

            if tem_termos_banidos == False and tem_termos_produto == True:
                preco = (
                    preco.replace("R$", "")
                    .replace(" ", "")
                    .replace(".", "")
                    .replace(",", ".")
                )
                preco = float(preco)

                lista_ofertas.append((nome, preco, link))
        except:
            pass
    print("Lista ofertas buscape:", lista_ofertas)
    return lista_ofertas


tabela_ofertas_lista = []

for linha in df_produtos.index:
    nome = df_produtos.loc[linha, "Nome"]
    termos_banidos = df_produtos.loc[linha, "Termos banidos"]
    preco_minimo = df_produtos.loc[linha, "Preço mínimo"]
    preco_maximo = df_produtos.loc[linha, "Preço máximo"]

    lista_ofertas_google_shopping = buscar_google_shopping(
        driver, nome, termos_banidos, preco_minimo, preco_maximo
    )
    if lista_ofertas_google_shopping:
        tabela_ofertas_lista.extend(lista_ofertas_google_shopping)
    else:
        tabela_google_shoppping = None

    lista_ofertas_buscape = buscar_buscape(
        driver, nome, termos_banidos, preco_minimo, preco_maximo
    )
    if lista_ofertas_buscape:
        tabela_ofertas_lista.extend(lista_ofertas_buscape)
    else:
        tabela_buscape = None
    tabela_ofertas = pd.DataFrame(
        tabela_ofertas_lista, columns=["produto", "preco", "link"]
    )
print(tabela_ofertas)


tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("ofertas.xlsx", index=False)


driver.quit()
