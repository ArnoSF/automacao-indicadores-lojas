"""
# Automação de Indicadores de Desempenho (OnePage)

Rotina automatizada para consolidação de vendas, apuração de metas,
geração de relatórios visuais e envio de e-mails para lojas e diretoria.

## Configuração e Execução via Terminal

Os parâmetros padrão de saída e diretórios ficam definidos no arquivo `config.toml`.
Você pode sobrescrever as opções de execução diretamente pelo terminal usando as flags:

* `-s`, `--silencioso`: Oculta os prints de acompanhamento no console.
* `-m`, `--nao-manter-planilhas`: Remove as planilhas individuais das lojas após o envio.
* `-g`, `--sem-graficos`: Não salva as imagens PNG dos gráficos em disco.
* `-p`, `--sem-pdfs`: Não gera os arquivos PDF de simulação dos e-mails.

**Exemplo de uso:**
```bash
python src/main.py --silencioso --sem-pdfs
```
"""

import pandas as pd

from configuracao import (
    buscar_variaveis_ambiente,
    carregar_configuracoes,
    configurar_log,
    identificar_ambiente,
    registrar_log,
    encerrar_log,
)
from mensageria import (
    conectar_servidor_smtp,
    desconectar_servidor_smtp,
    enviar_email_diretoria,
    enviar_emails_lojas,
)
from processamento import (
    gerar_valores_consolidados,
    preencher_atingimento_metas,
    preparar_bases_processadas,
    unificar_dados_lojas,
)
from relatorios import configurar_ambiente_jinja


def main() -> None:
    """Coordena todo o fluxo de processamento de dados, geração de relatórios e envio de e-mails."""
    agora = pd.Timestamp.now()
    data_hoje = agora.normalize()
    data_ontem = data_hoje - pd.Timedelta(days=1)

    data_hoje_str_amd = data_hoje.strftime("%Y-%m-%d")
    data_ontem_str_amd = data_ontem.strftime("%Y-%m-%d")
    data_ontem_str_dma = data_ontem.strftime("%d/%m/%Y")

    print(f"{agora.strftime('%H:%M:%S')} - Identificação do ambiente de execução")
    eh_notebook, caminho_raiz = identificar_ambiente()
    if not caminho_raiz:
        return

    config = carregar_configuracoes(eh_notebook, caminho_raiz)
    exibir_prints = config["outputs"]["exibir_prints_acompanhamento"]
    manter_planilhas = config["outputs"]["manter_planilhas"]
    exportar_graficos = config["outputs"]["exportar_graficos"]
    gerar_pdfs = config["outputs"]["gerar_pdfs_email"]

    caminho_planilhas = caminho_raiz / config["caminhos"]["saida_planilhas"]
    caminho_graficos = caminho_raiz / config["caminhos"]["saida_graficos"] if exportar_graficos else None
    caminho_pdfs = caminho_raiz / config["caminhos"]["saida_pdfs"] if gerar_pdfs else None

    if exibir_prints:
        print(f"{agora.strftime('%H:%M:%S')} - Configuração do log")
    if not configurar_log(caminho_raiz, data_hoje_str_amd):
        return

    registrar_log("Início geral da execução", "INFO", exibir_prints)

    registrar_log("Busca das variáveis de ambiente", "INFO", exibir_prints)
    credenciais = buscar_variaveis_ambiente(caminho_raiz / ".env")
    if not credenciais:
        return

    registrar_log("Preparação das bases de dados", "INFO", exibir_prints)
    bases = preparar_bases_processadas(caminho_raiz, caminho_planilhas, credenciais.destinatario_teste, data_ontem)
    if not bases:
        return

    registrar_log("Unificação dos dados gerais das lojas", "INFO", exibir_prints)
    df_lojas, dic_nomes_lojas, email_diretoria = unificar_dados_lojas(bases.caminho_raw / "Lojas.csv", bases.df_emails)
    if df_lojas is None:
        return

    registrar_log("Geração dos valores consolidados", "INFO", exibir_prints)
    df_vendas_lojas, df_vendas_rede = gerar_valores_consolidados(bases.df_vendas, data_ontem)
    if df_vendas_lojas is None:
        return

    registrar_log("Cálculo de atingimento de metas", "INFO", exibir_prints)
    metas = preencher_atingimento_metas(bases.df_metas_lojas, df_vendas_lojas, df_vendas_rede)
    if not metas:
        return

    registrar_log("Configuração do ambiente Jinja", "INFO", exibir_prints)
    env_jinja, nome_template_lojas, nome_template_diretoria = configurar_ambiente_jinja(caminho_raiz)
    if not env_jinja:
        return

    registrar_log("Conexão com o servidor de e-mail", "INFO", exibir_prints)
    servidor_email = conectar_servidor_smtp(
        credenciais.remetente, credenciais.senha, credenciais.porta, credenciais.servidor
    )
    if not servidor_email:
        return

    if caminho_graficos:
        caminho_graficos.mkdir(parents=True, exist_ok=True)
    if caminho_pdfs:
        caminho_pdfs.mkdir(parents=True, exist_ok=True)

    try:
        registrar_log("Envio de e-mails às lojas", "INFO", exibir_prints)
        sucesso_lojas = enviar_emails_lojas(
            servidor_email=servidor_email,
            nome_remetente=credenciais.nome_remetente,
            df_lojas=df_lojas,
            dic_metas=metas.dic_metas_lojas,
            df_vendas_consolidadas_lojas=metas.df_vendas_lojas,
            data_ontem_str_dma=data_ontem_str_dma,
            data_ontem_str_amd=data_ontem_str_amd,
            template_compilado=env_jinja.get_template(nome_template_lojas),
            dic_planilhas_lojas=bases.dic_planilhas_lojas,
            caminho_graficos=caminho_graficos,
            caminho_pdfs=caminho_pdfs,
            eh_notebook=eh_notebook,
            exibir_prints=exibir_prints,
            manter_planilhas=manter_planilhas,
        )
        if not sucesso_lojas:
            return

        registrar_log("Envio de e-mail à diretoria com rankings das lojas", "INFO", exibir_prints)
        enviar_email_diretoria(
            servidor_email=servidor_email,
            nome_remetente=credenciais.nome_remetente,
            df_vendas=bases.df_vendas,
            dic_metas_lojas=metas.dic_metas_lojas,
            dic_metas_rede=metas.dic_metas_rede,
            df_vendas_consolidadas_lojas=metas.df_vendas_lojas,
            df_vendas_consolidadas_rede=metas.df_vendas_rede,
            email_diretoria=email_diretoria,
            data_ontem_str_dma=data_ontem_str_dma,
            data_ontem_str_amd=data_ontem_str_amd,
            template_compilado=env_jinja.get_template(nome_template_diretoria),
            caminho_graficos=caminho_graficos,
            caminho_pdfs=caminho_pdfs,
            eh_notebook=eh_notebook,
            dic_nomes_lojas=dic_nomes_lojas,
            exibir_prints=exibir_prints,
        )
    finally:
        registrar_log("Encerramento da conexão com o servidor de e-mail", "INFO", exibir_prints)
        desconectar_servidor_smtp(servidor_email, exibir_prints)
        registrar_log("Finalização da execução", "INFO", exibir_prints)
        encerrar_log()


if __name__ == "__main__":
    main()
