"""Módulo de tratamento de dados, formatação de planilhas Excel e apuração de metas."""

import gc
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

from configuracao import registrar_log
from modelos import BasesProcessadas, MetasConsolidadas


# ==============================================================================
# CONFIGURAÇÃO DO PDOC (EXIBIÇÃO DE FUNÇÕES INTERNAS)
# ==============================================================================
__all__ = [
    "moeda_para_decimal",
    "converter_notacao_europeia",
    "_formatar_aba_excel",
    "preparar_bases_processadas",
    "_calcular_metricas_periodo",
    "gerar_valores_consolidados",
    "unificar_dados_lojas",
    "_aplicar_metas_e_ordenar",
    "preencher_atingimento_metas",
]


def moeda_para_decimal(valor: Any) -> float | Any | None:
    """Converte strings monetárias no padrão brasileiro em números decimais (float).

    Args:
        valor (Any): Valor a ser avaliado e convertido.

    Returns:
        float | Any | None: Número decimal convertido, o próprio valor original se não
            for string, ou None em caso de erro de conversão.
    """
    try:
        if isinstance(valor, str):
            return float(valor.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return valor
    except Exception as erro:
        registrar_log(f"Erro na conversão de moeda para decimal: {erro}!", "CRITICAL")
        return None


def converter_notacao_europeia(nro_notacao_anglicana: float) -> str:
    """Converte o fracionário informado em notação anglo-americana para notação europeia e 2 decimais.

    Args:
        nro_notacao_anglicana (float): Fracionário em notação anglo-americana.

    Returns:
        str: Texto com número em notação europeia e 2 decimais.
    """
    return f"{nro_notacao_anglicana:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def _formatar_aba_excel(
    ws: Worksheet,
    df: pd.DataFrame,
    formatos_colunas: dict[str, str] | None = None,
    colunas_semaforo: list[str] | None = None,
) -> None:
    """Aplica máscaras de formatação numérica/data e semáforos coloridos em uma aba do Excel.

    Args:
        ws (Worksheet): Aba do openpyxl a ser formatada.
        df (pd.DataFrame): DataFrame correspondente aos dados da aba.
        formatos_colunas (dict[str, str] | None): Dicionário mapeando {nome_coluna: mascara_excel}.
        colunas_semaforo (list[str] | None): Lista de colunas booleanas a converter em círculos coloridos.
    """
    formatos_colunas = formatos_colunas or {}
    colunas_semaforo = colunas_semaforo or []

    fonte_verde = Font(name="Segoe UI Symbol", color="188038", size=14)
    fonte_vermelha = Font(name="Segoe UI Symbol", color="D93025", size=14)
    alinhamento_centro = Alignment(horizontal="center", vertical="center")

    for col_idx, col_nome in enumerate(df.columns, start=1):
        if col_nome in formatos_colunas:
            mascara = formatos_colunas[col_nome]
            for row in range(2, len(df) + 2):
                ws.cell(row=row, column=col_idx).number_format = mascara
        elif col_nome in colunas_semaforo:
            for row in range(2, len(df) + 2):
                celula = ws.cell(row=row, column=col_idx)
                atingiu = bool(celula.value)
                celula.value = "⬤"
                celula.font = fonte_verde if atingiu else fonte_vermelha
                celula.alignment = alinhamento_centro


def preparar_bases_processadas(
    caminho_raiz: Path,
    caminho_planilhas: Path,
    destinatario_teste: str | None,
    data_ontem: pd.Timestamp,
) -> BasesProcessadas | None:
    """Valida os arquivos de entrada, ajusta datas e gera as planilhas individuais das lojas.

    Args:
        caminho_raiz (Path): Caminho raiz do projeto.
        caminho_planilhas (Path): Diretório de saída para as planilhas processadas das lojas.
        destinatario_teste (str | None): E-mail destinatário para preenchimento de linhas vazias em testes.
        data_ontem (pd.Timestamp): Data de referência do dia anterior ao atual.

    Returns:
        BasesProcessadas | None: Objeto nomeado contendo os atributos:
            - caminho_raw (Path): Caminho do diretório de arquivos originais.
            - df_emails (pd.DataFrame): Tabela de e-mails com destinatários vazios preenchidos.
            - df_metas_lojas (pd.DataFrame): Tabela de metas das lojas com valores monetários convertidos.
            - df_vendas (pd.DataFrame): Tabela de vendas filtrada para os últimos 365 dias.
            - dic_planilhas_lojas (dict[int, Path]): Dicionário mapeando o ID de cada loja ao caminho da planilha.
            Retorna None em caso de erro ou ausência de arquivos obrigatórios.
    """
    try:
        caminho_raw = caminho_raiz / "data/raw"
        if not caminho_raw.is_dir():
            registrar_log(f'Diretório "data/raw" ("{caminho_raw}") não encontrado!', "CRITICAL")
            return None

        caminho_planilhas.mkdir(parents=True, exist_ok=True)

        caminho_emails = caminho_raw / "Emails.xlsx"
        if not caminho_emails.is_file():
            registrar_log(f'Arquivo "{caminho_emails.name}" não encontrado!', "CRITICAL")
            return None

        df_emails = pd.read_excel(caminho_emails, index_col="Loja", dtype={"E-mail": "string"})
        ser_sem_email = df_emails["E-mail"].isna() | (df_emails["E-mail"].str.strip() == "")
        if ser_sem_email.any():
            if destinatario_teste:
                df_emails.loc[ser_sem_email, "E-mail"] = destinatario_teste
            else:
                registrar_log(
                    'Planilha de e-mails possui linhas vazias e "DESTINATARIO_EMAIL_TESTE" não foi configurado!',
                    "CRITICAL",
                )
                return None

        caminho_metas = caminho_raw / "Metas.csv"
        if not caminho_metas.is_file():
            registrar_log(f'Arquivo "{caminho_metas.name}" não encontrado!', "CRITICAL")
            return None

        df_metas_lojas = pd.read_csv(caminho_metas, sep=";", encoding="cp1252", index_col=0)
        df_metas_lojas["Faturamento_Num"] = df_metas_lojas["Faturamento"].apply(moeda_para_decimal)
        df_metas_lojas["Ticket Médio_Num"] = df_metas_lojas["Ticket Médio"].apply(moeda_para_decimal)

        caminho_vendas = caminho_raw / "Vendas.xlsx"
        if not caminho_vendas.is_file():
            registrar_log(f'Arquivo "{caminho_vendas.name}" não encontrado!', "CRITICAL")
            return None

        df_vendas = pd.read_excel(caminho_vendas)
        df_vendas["Data"] = pd.to_datetime(df_vendas["Data"]).dt.normalize()

        qtd_dias_ajuste = (data_ontem - df_vendas["Data"].max()).days
        df_vendas["Data"] = df_vendas["Data"] + pd.to_timedelta(qtd_dias_ajuste, unit="D")

        data_corte_anual = data_ontem - pd.DateOffset(years=1) + pd.Timedelta(days=1)
        df_vendas = df_vendas[df_vendas["Data"] >= data_corte_anual]

        data_ontem_str_amd = data_ontem.strftime("%Y-%m-%d")
        dic_planilhas_lojas: dict[int, Path] = {}

        formatos_diario = {
            "Data": "dd/mm/yyyy",
            "Quantidade": "#,##0",
            "Valor Unitário": "#,##0.00",
            "Valor Final": "#,##0.00",
        }
        formatos_mensal = {
            "Mês": "mm/yyyy",
            "Quantidade": "#,##0",
            "Valor Final": "#,##0.00",
        }

        for id_loja, df_loja_dia in df_vendas.groupby("ID Loja"):
            caminho_proc_loja = caminho_planilhas / f"Loja_{id_loja}"
            caminho_proc_loja.mkdir(exist_ok=True)
            caminho_arq_loja = caminho_proc_loja / f"{data_ontem_str_amd}_Vendas_Loja_{id_loja}.xlsx"

            df_loja_dia = df_loja_dia.drop(columns=["ID Loja"])
            dic_planilhas_lojas[id_loja] = caminho_arq_loja

            df_loja_mes = (
                df_loja_dia.groupby(df_loja_dia["Data"].dt.to_period("M"))[["Quantidade", "Valor Final"]]
                .sum()
                .reset_index()
                .rename(columns={"Data": "Mês"})
            )
            df_loja_mes["Mês"] = df_loja_mes["Mês"].dt.to_timestamp()

            with pd.ExcelWriter(caminho_arq_loja, engine="openpyxl", datetime_format="mm/yyyy") as writer:
                df_loja_dia.to_excel(writer, sheet_name="Diário", index=False)
                df_loja_mes.to_excel(writer, sheet_name="Mensal", index=False)
                _formatar_aba_excel(writer.sheets["Diário"], df_loja_dia, formatos_colunas=formatos_diario)
                _formatar_aba_excel(writer.sheets["Mensal"], df_loja_mes, formatos_colunas=formatos_mensal)

        return BasesProcessadas(
            caminho_raw=caminho_raw,
            df_emails=df_emails,
            df_metas_lojas=df_metas_lojas,
            df_vendas=df_vendas.drop(columns=["Código Venda", "Valor Unitário"]),
            dic_planilhas_lojas=dic_planilhas_lojas,
        )
    except Exception as erro:
        registrar_log(f"Erro na preparação das bases processadas: {erro}!", "CRITICAL")
        return None


def _calcular_metricas_periodo(df: pd.DataFrame, agrupador: str | Callable, sufixo: str) -> pd.DataFrame:
    """Calcula soma do faturamento, quantidade de produtos distintos e ticket médio da tabela informada.

    Args:
        df (pd.DataFrame): Tabela a ter as métricas calculadas.
        agrupador (str | Callable): Coluna ou função de agrupamento.
        sufixo (str): Sufixo da coluna a ser criada na tabela.

    Returns:
        pd.DataFrame: Tabela com as métricas calculadas.
    """
    df_resumo = df.groupby(agrupador).agg(
        **{
            f"Fat_{sufixo}": pd.NamedAgg(column="Valor Final", aggfunc="sum"),
            f"Prod_Div_{sufixo}": pd.NamedAgg(column="Produto", aggfunc="nunique"),
            f"Qtd_{sufixo}": pd.NamedAgg(column="Quantidade", aggfunc="sum"),
        }
    )
    df_resumo[f"Tic_Medio_{sufixo}"] = (df_resumo[f"Fat_{sufixo}"] / df_resumo[f"Qtd_{sufixo}"]).fillna(0)
    return df_resumo.drop(columns=[f"Qtd_{sufixo}"])


def gerar_valores_consolidados(
    df_vendas: pd.DataFrame, data_ontem: pd.Timestamp
) -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    """Gera tabelas anual e do dia anterior de valores acumulados de faturamento, diversidade e ticket médio.

    Args:
        df_vendas (pd.DataFrame): Tabela geral de vendas.
        data_ontem (pd.Timestamp): Data do dia anterior.

    Returns:
        tuple: Tabela de valores consolidados por loja (pd.DataFrame) e tabela de valores consolidados
            de toda a rede (pd.DataFrame). Havendo erro, todos os itens retornam None.
    """
    try:
        df_filtrado_ontem = df_vendas[df_vendas["Data"] == data_ontem]

        def consolidar_por_nivel(agrupador: str | Callable) -> pd.DataFrame:
            df_ano = _calcular_metricas_periodo(df_vendas, agrupador, "Ano")
            df_ontem = _calcular_metricas_periodo(df_filtrado_ontem, agrupador, "Ontem")

            df_consolidado = df_ano.join(df_ontem).fillna(0)
            colunas_inteiras = ["Prod_Div_Ano", "Prod_Div_Ontem"]
            df_consolidado[colunas_inteiras] = df_consolidado[colunas_inteiras].astype(int)
            return df_consolidado

        df_vendas_consolidadas_lojas = consolidar_por_nivel("ID Loja")
        df_vendas_consolidadas_rede = consolidar_por_nivel(lambda _: "Total Rede")
        return df_vendas_consolidadas_lojas, df_vendas_consolidadas_rede
    except Exception as erro:
        registrar_log(f"Erro na geração de valores acumulados: {erro}!", "CRITICAL")
        return None, None


def unificar_dados_lojas(
    caminho_arquivo_lojas: Path, df_emails: pd.DataFrame
) -> tuple[pd.DataFrame | None, dict[int, str] | None, str | None]:
    """Unifica os dados gerais das lojas em uma única tabela.

    Args:
        caminho_arquivo_lojas (Path): Caminho do arquivo geral de lojas.
        df_emails (pd.DataFrame): Tabela de e-mails.

    Returns:
        tuple[pd.DataFrame | None, dict[int, str] | None, str | None]: Tabela contendo os dados gerais das lojas,
            dicionário mapeando {id_loja: nome_loja} e o e-mail da diretoria. Havendo erro, todos retornam None.
    """
    try:
        if not caminho_arquivo_lojas.is_file():
            registrar_log(f'Arquivo "{caminho_arquivo_lojas.name}" ("{caminho_arquivo_lojas}") não encontrado!', "CRITICAL")
            return None, None, None

        df_lojas = pd.read_csv(caminho_arquivo_lojas, sep=";", encoding="cp1252", index_col="Loja")
        df_lojas = df_lojas.join(df_emails[["Gerente", "E-mail"]]).reset_index().set_index("ID Loja")
        email_diretoria = df_emails.loc["Diretoria", "E-mail"]

        del df_emails
        gc.collect()

        dic_nomes_lojas = dict(zip(df_lojas.index, df_lojas["Loja"]))
        return df_lojas, dic_nomes_lojas, email_diretoria
    except Exception as erro:
        registrar_log(f"Erro na unificação de dados gerais das lojas: {erro}!", "CRITICAL")
        return None, None, None


def _aplicar_metas_e_ordenar(df: pd.DataFrame, dic_metas: dict) -> pd.DataFrame:
    """Cria colunas booleanas de atingimento de metas e organiza a ordem das colunas.

    Args:
        df (pd.DataFrame): Tabela consolidada de indicadores realizados.
        dic_metas (dict): Dicionário contendo as metas correspondentes ao nível da tabela.

    Returns:
        pd.DataFrame: Tabela atualizada com indicadores intercalados por seus status de meta.
    """
    regras_metas = [
        ("Meta_Fat_Ontem", "Fat_Ontem", "fat_dia"),
        ("Meta_Div_Ontem", "Prod_Div_Ontem", "div_dia"),
        ("Meta_Tic_Ontem", "Tic_Medio_Ontem", "tic_dia"),
        ("Meta_Fat_Ano", "Fat_Ano", "fat_ano"),
        ("Meta_Div_Ano", "Prod_Div_Ano", "div_ano"),
        ("Meta_Tic_Ano", "Tic_Medio_Ano", "tic_ano"),
    ]

    novas_colunas = {
        col_meta: df[col_valor] >= dic_metas[chave_meta]
        for col_meta, col_valor, chave_meta in regras_metas
    }

    ordem_colunas = []
    for col_meta, col_valor, _ in regras_metas:
        ordem_colunas.extend([col_valor, col_meta])

    return df.assign(**novas_colunas)[ordem_colunas]


def preencher_atingimento_metas(
    df_metas_lojas: pd.DataFrame,
    df_vendas_consolidadas_lojas: pd.DataFrame,
    df_vendas_consolidadas_rede: pd.DataFrame,
) -> MetasConsolidadas | None:
    """Calcula metas da rede e avalia o atingimento das metas para lojas e rede.

    Args:
        df_metas_lojas (pd.DataFrame): Tabela de metas individuais das lojas.
        df_vendas_consolidadas_lojas (pd.DataFrame): Tabela de vendas consolidadas por loja.
        df_vendas_consolidadas_rede (pd.DataFrame): Tabela de vendas consolidadas de toda a rede.

    Returns:
        MetasConsolidadas | None: Objeto nomeado contendo os atributos:
            - dic_metas_lojas (dict): Dicionário com as metas diárias e anuais das lojas.
            - df_vendas_lojas (pd.DataFrame): Tabela consolidada das lojas com status de metas.
            - dic_metas_rede (dict): Dicionário com as metas diárias e anuais da rede.
            - df_vendas_rede (pd.DataFrame): Tabela consolidada da rede com status de metas.
            Retorna None em caso de erro.
    """
    try:
        dic_metas_lojas = {
            "fat_dia": df_metas_lojas.loc["Diário", "Faturamento_Num"],
            "div_dia": df_metas_lojas.loc["Diário", "Diversidade de Produtos"],
            "tic_dia": df_metas_lojas.loc["Diário", "Ticket Médio_Num"],
            "fat_ano": df_metas_lojas.loc["Anual", "Faturamento_Num"],
            "div_ano": df_metas_lojas.loc["Anual", "Diversidade de Produtos"],
            "tic_ano": df_metas_lojas.loc["Anual", "Ticket Médio_Num"],
        }

        qtd_lojas = len(df_vendas_consolidadas_lojas)
        dic_metas_rede = {
            chave: valor if chave in {"tic_dia", "tic_ano"} else valor * qtd_lojas
            for chave, valor in dic_metas_lojas.items()
        }

        return MetasConsolidadas(
            dic_metas_lojas=dic_metas_lojas,
            df_vendas_lojas=_aplicar_metas_e_ordenar(df_vendas_consolidadas_lojas, dic_metas_lojas),
            dic_metas_rede=dic_metas_rede,
            df_vendas_rede=_aplicar_metas_e_ordenar(df_vendas_consolidadas_rede, dic_metas_rede),
        )
    except Exception as erro:
        registrar_log(f"Erro no preenchimento do atingimento de metas: {erro}!", "CRITICAL")
        return None