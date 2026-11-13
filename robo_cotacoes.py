"""
Robô de cotações (Python + RPA + Excel)

O que ele faz:
  1. Abre o Google Chrome sozinho (sem janela, modo "headless")
  2. Entra no Google Finanças e lê a cotação de várias moedas
  3. Salva tudo numa planilha Excel com data e hora
  4. Monta gráficos mostrando a variação ao longo do tempo
  5. Confere os alertas (ex.: "dólar passou de R$ 6,00")
  6. Manda a planilha por e-mail (se estiver ligado no config.ini)

Como rodar:
  python robo_cotacoes.py            # roda sem abrir janela
  python robo_cotacoes.py --ver      # abre o Chrome na tela pra você ver o robô trabalhando
  python robo_cotacoes.py --sem-email   # roda sem mandar e-mail
"""

import configparser
import re
import smtplib
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------

PASTA = Path(__file__).parent
ARQUIVO_EXCEL = PASTA / "cotacoes.xlsx"
ARQUIVO_CONFIG = PASTA / "config.ini"

# Moedas que o robô vai buscar. Pra adicionar outra, é só incluir aqui.
# "par" é o código usado na URL do Google Finanças.
# "nome" é o texto que aparece logo acima do preço na página.
MOEDAS = {
    "Dólar":   {"par": "USD-BRL", "nome": "Dólar americano / Real brasileiro"},
    "Euro":    {"par": "EUR-BRL", "nome": "Euro / Real brasileiro"},
    "Libra":   {"par": "GBP-BRL", "nome": "Libra esterlina / Real brasileiro"},
    "Bitcoin": {"par": "BTC-BRL", "nome": "Bitcoin / Real brasileiro"},
    "Ether":   {"par": "ETH-BRL", "nome": "Ether / Real brasileiro"},
}

# Um número no formato brasileiro, ex.: 5,6595 ou 512.345,67
NUMERO_BR = re.compile(r"^\d{1,3}(\.\d{3})*(,\d+)?$|^\d+(,\d+)?$")


def texto_para_numero(texto: str) -> float:
    """Converte '512.345,67' em 512345.67"""
    return float(texto.strip().replace(".", "").replace(",", "."))


def formatar_reais(valor: float, casas: int = 4) -> str:
    """Converte 512345.67 em 'R$ 512.345,6700'"""
    texto = f"{valor:,.{casas}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {texto}"


def carregar_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser(inline_comment_prefixes=("#",))
    config.optionxform = str  # mantém maiúsculas/acentos nos nomes das moedas
    if ARQUIVO_CONFIG.exists():
        config.read(ARQUIVO_CONFIG, encoding="utf-8")
    return config


# ---------------------------------------------------------------------------
# Parte 1: o navegador (RPA)
# ---------------------------------------------------------------------------

def abrir_navegador(mostrar_janela: bool) -> webdriver.Chrome:
    """Abre o Chrome. O Selenium baixa o chromedriver sozinho se precisar."""
    opcoes = webdriver.ChromeOptions()
    if not mostrar_janela:
        opcoes.add_argument("--headless=new")
    opcoes.add_argument("--lang=pt-BR")
    opcoes.add_argument("--window-size=1280,900")
    return webdriver.Chrome(options=opcoes)


def buscar_cotacao(navegador: webdriver.Chrome, par: str, nome: str) -> float:
    """
    Entra na página da moeda e devolve o preço.

    A página mostra algo assim:
        Dólar americano / Real brasileiro
        5,6595
    Então o robô procura a linha com o nome da moeda e pega o
    primeiro número que aparece logo depois.
    """
    navegador.get(f"https://www.google.com/finance/quote/{par}?hl=pt-BR")

    def achar_preco(_):
        linhas = navegador.find_element(By.TAG_NAME, "body").text.splitlines()
        linhas = [l.strip().replace("R$", "").strip() for l in linhas]
        for i, linha in enumerate(linhas):
            if linha.lower() == nome.lower():
                for seguinte in linhas[i + 1 : i + 6]:
                    if NUMERO_BR.match(seguinte):
                        return seguinte
        return False  # ainda não carregou; o WebDriverWait tenta de novo

    # Espera até 15 segundos a página carregar o preço
    texto = WebDriverWait(navegador, 15, poll_frequency=0.5).until(achar_preco)
    return texto_para_numero(texto)


def coletar_cotacoes(mostrar_janela: bool = False) -> dict:
    navegador = abrir_navegador(mostrar_janela)
    cotacoes = {}
    try:
        for moeda, info in MOEDAS.items():
            try:
                valor = buscar_cotacao(navegador, info["par"], info["nome"])
                cotacoes[moeda] = valor
                print(f"  [OK] {moeda}: {formatar_reais(valor)}")
            except Exception as erro:
                # Se uma moeda falhar, o robô não trava: anota vazio e segue
                cotacoes[moeda] = None
                print(f"  [ERRO] {moeda}: não consegui ler ({type(erro).__name__})")
            if mostrar_janela:
                time.sleep(1)  # só pra dar tempo de você ver
    finally:
        navegador.quit()  # fecha o Chrome mesmo se der erro
    return cotacoes


# ---------------------------------------------------------------------------
# Parte 2: a planilha (Excel)
# ---------------------------------------------------------------------------

def estilizar_cabecalho(aba) -> None:
    for celula in aba[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="2F5597")
        celula.alignment = Alignment(horizontal="center")
        largura = 16 if celula.column > 2 else 12
        aba.column_dimensions[celula.column_letter].width = largura


def abrir_ou_criar_planilha():
    if ARQUIVO_EXCEL.exists():
        wb = load_workbook(ARQUIVO_EXCEL)
    else:
        wb = Workbook()
        wb.active.title = "Cotações"
        wb.active.append(["Data", "Hora"])
        wb.active.freeze_panes = "A2"  # cabeçalho fica fixo ao rolar

    # Se você adicionou uma moeda nova no MOEDAS, cria a coluna dela
    aba = wb["Cotações"]
    colunas = [c.value for c in aba[1]]
    for moeda in MOEDAS:
        if moeda not in colunas:
            aba.cell(row=1, column=len(colunas) + 1, value=moeda)
            colunas.append(moeda)
    estilizar_cabecalho(aba)
    return wb


def salvar_cotacoes(cotacoes: dict) -> int:
    wb = abrir_ou_criar_planilha()
    aba = wb["Cotações"]
    colunas = [c.value for c in aba[1]]

    agora = datetime.now()
    linha = aba.max_row + 1
    aba.cell(row=linha, column=1, value=agora.strftime("%d/%m/%Y"))
    aba.cell(row=linha, column=2, value=agora.strftime("%H:%M:%S"))
    for moeda, valor in cotacoes.items():
        celula = aba.cell(row=linha, column=colunas.index(moeda) + 1, value=valor)
        celula.number_format = '"R$" #,##0.0000'

    atualizar_graficos(wb)
    wb.save(ARQUIVO_EXCEL)
    print(f"\nPlanilha salva: {ARQUIVO_EXCEL.name} ({linha - 1} registros)")
    return linha - 1


def atualizar_graficos(wb) -> None:
    """Recria a aba 'Gráficos' com um gráfico por moeda."""
    if "Gráficos" in wb.sheetnames:
        del wb["Gráficos"]
    aba_graf = wb.create_sheet("Gráficos")
    dados = wb["Cotações"]
    ultima = dados.max_row

    # Um gráfico pra cada moeda (o bitcoin vale muito mais que o dólar,
    # então no mesmo gráfico as outras linhas ficariam achatadas)
    for i, celula in enumerate(dados[1][2:]):
        grafico = LineChart()
        grafico.title = f"{celula.value} (R$)"
        grafico.y_axis.title = "R$"
        grafico.x_axis.title = "Data"
        grafico.height, grafico.width = 7, 18
        grafico.legend = None

        valores = Reference(dados, min_col=celula.column, min_row=1, max_row=ultima)
        datas = Reference(dados, min_col=1, min_row=2, max_row=ultima)
        grafico.add_data(valores, titles_from_data=True)
        grafico.set_categories(datas)

        aba_graf.add_chart(grafico, f"A{1 + i * 15}")


# ---------------------------------------------------------------------------
# Parte 3: alertas
# ---------------------------------------------------------------------------

def verificar_alertas(cotacoes: dict, config) -> list[str]:
    """Compara as cotações com os limites do config.ini."""
    alertas = []
    regras = [("alerta_acima", "subiu acima de", lambda v, l: v > l),
              ("alerta_abaixo", "caiu abaixo de", lambda v, l: v < l)]

    for secao, texto, passou in regras:
        if not config.has_section(secao):
            continue
        for moeda, limite_txt in config[secao].items():
            valor = cotacoes.get(moeda)
            if valor is None or not limite_txt.strip():
                continue
            limite = texto_para_numero(limite_txt)
            if passou(valor, limite):
                alertas.append(f"{moeda} {texto} {formatar_reais(limite, 2)}: "
                               f"está em {formatar_reais(valor)}")

    if alertas:
        print("\nALERTAS:")
        for a in alertas:
            print(f"  • {a}")
    return alertas


# ---------------------------------------------------------------------------
# Parte 4: e-mail
# ---------------------------------------------------------------------------

def enviar_email(cotacoes: dict, alertas: list[str], config) -> None:
    if not config.has_section("email") or config["email"].get("ativo", "nao") != "sim":
        return
    cfg = config["email"]
    if cfg.get("somente_com_alerta", "nao") == "sim" and not alertas:
        print("Nenhum alerta, e-mail não enviado.")
        return

    hoje = datetime.now().strftime("%d/%m/%Y %H:%M")
    msg = EmailMessage()
    msg["From"] = cfg["usuario"]
    msg["To"] = cfg["destinatario"]
    msg["Subject"] = (f"[ALERTA] {len(alertas)} alerta(s) de cotação - {hoje}" if alertas
                      else f"Cotações do dia - {hoje}")

    corpo = [f"Cotações de {hoje}:", ""]
    for moeda, valor in cotacoes.items():
        corpo.append(f"  {moeda}: {formatar_reais(valor) if valor else 'não lida'}")
    if alertas:
        corpo += ["", "ALERTAS:"] + [f"  • {a}" for a in alertas]
    corpo += ["", "Planilha completa em anexo.", "-- Robô de Cotações"]
    msg.set_content("\n".join(corpo))

    msg.add_attachment(
        ARQUIVO_EXCEL.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=ARQUIVO_EXCEL.name,
    )

    try:
        with smtplib.SMTP(cfg["servidor"], cfg.getint("porta"), timeout=30) as smtp:
            if cfg.get("tls", "sim") == "sim":
                smtp.starttls()
            if cfg.get("senha"):
                smtp.login(cfg["usuario"], cfg["senha"])
            smtp.send_message(msg)
        print(f"E-mail enviado para {cfg['destinatario']}")
    except Exception as erro:
        print(f"[ERRO] Não consegui mandar o e-mail: {erro}")


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def main():
    mostrar = "--ver" in sys.argv
    print(f"Robô de cotações iniciado ({datetime.now():%d/%m/%Y %H:%M})\n")
    config = carregar_config()

    cotacoes = coletar_cotacoes(mostrar_janela=mostrar)
    if all(v is None for v in cotacoes.values()):
        print("\nNenhuma cotação foi lida. Confere sua internet e tenta de novo.")
        sys.exit(1)

    salvar_cotacoes(cotacoes)
    alertas = verificar_alertas(cotacoes, config)
    if "--sem-email" not in sys.argv:
        enviar_email(cotacoes, alertas, config)


if __name__ == "__main__":
    main()
