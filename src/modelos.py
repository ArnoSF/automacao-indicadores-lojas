"""Estruturas de dados e tipagem estática do projeto."""

from pathlib import Path
from typing import NamedTuple, TypedDict

import pandas as pd


class ConfigOutputs(TypedDict):
    exibir_prints_acompanhamento: bool
    manter_planilhas: bool
    exportar_graficos: bool
    gerar_pdfs_email: bool


class ConfigCaminhos(TypedDict):
    saida_planilhas: str
    saida_pdfs: str
    saida_graficos: str


class Configuracoes(TypedDict):
    outputs: ConfigOutputs
    caminhos: ConfigCaminhos


class CredenciaisSMTP(NamedTuple):
    remetente: str
    senha: str
    porta: str
    servidor: str
    destinatario_teste: str | None
    nome_remetente: str


class BasesProcessadas(NamedTuple):
    caminho_raw: Path
    df_emails: pd.DataFrame
    df_metas_lojas: pd.DataFrame
    df_vendas: pd.DataFrame
    dic_planilhas_lojas: dict[int, Path]


class MetasConsolidadas(NamedTuple):
    dic_metas_lojas: dict
    df_vendas_lojas: pd.DataFrame
    dic_metas_rede: dict
    df_vendas_rede: pd.DataFrame