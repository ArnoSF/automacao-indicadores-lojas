"""Gerenciamento de ambiente, arquivos de configuração, logs e variáveis .env."""

import logging
import sys
import tomllib
from pathlib import Path

import pandas as pd
from dotenv import dotenv_values

from modelos import Configuracoes, CredenciaisSMTP


def identificar_ambiente() -> tuple[bool | None, Path | None]:
    """Identifica o ambiente de execução e mapeia o caminho raiz do projeto."""
    try:
        eh_notebook = "ipykernel" in sys.modules
        caminho_raiz = Path.cwd().parent if eh_notebook else Path(__file__).resolve().parent.parent
        return eh_notebook, caminho_raiz
    except Exception as erro:
        print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - Erro na identificação do ambiente: {erro}!")
        return None, None


def carregar_configuracoes(eh_notebook: bool, caminho_raiz: Path) -> Configuracoes:
    """Carrega as configurações do arquivo config.toml e aplica flags de terminal, se houver."""
    try:
        with open(caminho_raiz / "config.toml", "rb") as arquivo_config:
            config = tomllib.load(arquivo_config)
    except Exception as erro:
        config = {
            "outputs": {
                "exibir_prints_acompanhamento": True,
                "manter_planilhas": True,
                "exportar_graficos": True,
                "gerar_pdfs_email": True,
            },
            "caminhos": {
                "saida_planilhas": "data/processed/backup_vendas_lojas/",
                "saida_pdfs": "reports/pdfs/",
                "saida_graficos": "reports/graficos/",
            },
        }
        print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - Erro no carregamento do config.toml: {erro}!")
        print("A execução ocorrerá com os valores padrão.")

    if not eh_notebook:
        import argparse

        parser = argparse.ArgumentParser(description="Automação de Indicadores de Lojas")
        parser.add_argument("-s", "--silencioso", action="store_true", help="Não exibe prints de acompanhamento")
        parser.add_argument("-m", "--nao-manter-planilhas", action="store_true", help="Remove planilhas geradas")
        parser.add_argument("-g", "--sem-graficos", action="store_true", help="Não exporta imagens dos gráficos")
        parser.add_argument("-p", "--sem-pdfs", action="store_true", help="Não gera PDFs de simulação dos e-mails")
        args = parser.parse_args()

        if args.silencioso:
            config["outputs"]["exibir_prints_acompanhamento"] = False
        if args.nao_manter_planilhas:
            config["outputs"]["manter_planilhas"] = False
        if args.sem_pdfs:
            config["outputs"]["gerar_pdfs_email"] = False
        if args.sem_graficos:
            config["outputs"]["exportar_graficos"] = False

    return config


def configurar_log(caminho_raiz: Path, data_hoje_str_amd: str) -> bool:
    """Configura o arquivo de registro de logs da execução diária."""
    try:
        caminho_logs = caminho_raiz / "logs"
        caminho_logs.mkdir(exist_ok=True)
        logging.basicConfig(
            filename=caminho_logs / f"{data_hoje_str_amd}-log.txt",
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            encoding="utf-8",
        )
        return True
    except Exception as erro:
        print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - Erro na configuração do log: {erro}!")
        return False


def registrar_log(mensagem: str, tipo: str = "INFO", exibir_prints: bool = True) -> bool:
    """Registra um evento no arquivo de log e, opcionalmente, o exibe no console."""
    try:
        horario = pd.Timestamp.now().strftime("%H:%M:%S")
        if tipo == "CRITICAL" or exibir_prints:
            print(f"{horario} - {mensagem}")

        funcao_log = getattr(logging, tipo.lower(), logging.info)
        funcao_log(mensagem)
        return True
    except Exception as erro:
        print(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - Erro no registro de log: {erro}!")
        return False


def buscar_variaveis_ambiente(caminho_env: Path) -> CredenciaisSMTP | None:
    """Lê e valida as variáveis de ambiente diretamente do arquivo .env."""
    try:
        if not caminho_env.is_file():
            registrar_log(f"Arquivo de variáveis de ambiente ({caminho_env}) não encontrado!", "CRITICAL")
            return None

        env_vars = dotenv_values(caminho_env)
        var_obrigatorias = ("REMETENTE_EMAIL", "REMETENTE_SENHA", "PORTA_SMTP", "SERVIDOR_SMTP", "REMETENTE_NOME")

        var_validadas = {}
        for var in var_obrigatorias:
            valor = (env_vars.get(var) or "").strip()
            if not valor:
                registrar_log(f'Variável de ambiente "{var}" ausente ou vazia!', "CRITICAL")
                return None
            var_validadas[var] = valor

        destinatario_teste = (env_vars.get("DESTINATARIO_EMAIL_TESTE") or "").strip() or None

        return CredenciaisSMTP(
            remetente=var_validadas["REMETENTE_EMAIL"],
            senha=var_validadas["REMETENTE_SENHA"],
            porta=var_validadas["PORTA_SMTP"],
            servidor=var_validadas["SERVIDOR_SMTP"],
            destinatario_teste=destinatario_teste,
            nome_remetente=var_validadas["REMETENTE_NOME"],
        )
    except Exception as erro:
        registrar_log(f"Erro no acesso às variáveis de ambiente: {erro}!", "CRITICAL")
        return None

def encerrar_log():
    """Fecha e remove todos os manipuladores de arquivo do sistema de logs."""
    logger = logging.getLogger()
    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)