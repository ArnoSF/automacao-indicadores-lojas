
"""Módulo de geração de mensagens HTML, gráficos estatísticos e simulações em PDF."""

import base64
from concurrent.futures import ThreadPoolExecutor
from email.utils import make_msgid
from pathlib import Path

import pandas as pd
import plotly.colors as pc
import plotly.graph_objects as go
from jinja2 import Environment, FileSystemLoader, Template
from plotly.subplots import make_subplots
from playwright.sync_api import sync_playwright
from premailer import transform

from configuracao import registrar_log
from processamento import converter_notacao_europeia


# ==============================================================================
# CONFIGURAÇÃO DO PDOC (EXIBIÇÃO DE FUNÇÕES INTERNAS)
# ==============================================================================
__all__ = [
    "configurar_ambiente_jinja",
    "_gerar_tabelas_indicadores_html",
    "construir_html_email",
    "construir_grafico_evol_vendas",
    "construir_grafico_anual_lojas",
    "_renderizar_pdf_playwright",
    "gerar_pdf_do_email",
]


def configurar_ambiente_jinja(caminho_raiz: Path) -> tuple[Environment | None, str | None, str | None]:
    """Valida a existência dos templates HTML e configura o ambiente Jinja2.

    Args:
        caminho_raiz (Path): Caminho raiz do projeto.

    Returns:
        tuple[Environment | None, str | None, str | None]: Ambiente Jinja2 configurado,
            nome do template das lojas e nome do template da diretoria.
            Retorna (None, None, None) em caso de erro.
    """
    try:
        caminho_templates = caminho_raiz / "assets" / "templates"
        if not caminho_templates.is_dir():
            registrar_log(f'Diretório de templates ("{caminho_templates}") não encontrado!', "CRITICAL")
            return None, None, None

        nome_template_lojas = "template_email_lojas.html"
        nome_template_diretoria = "template_email_diretoria.html"

        for nome_tpl in (nome_template_lojas, nome_template_diretoria):
            if not (caminho_templates / nome_tpl).is_file():
                registrar_log(f'Template "{nome_tpl}" não encontrado em "{caminho_templates}"!', "CRITICAL")
                return None, None, None

        return Environment(loader=FileSystemLoader(caminho_templates)), nome_template_lojas, nome_template_diretoria
    except Exception as erro:
        registrar_log(f"Erro ao configurar o ambiente Jinja2: {erro}!", "CRITICAL")
        return None, None, None


def _gerar_tabelas_indicadores_html(colunas_dados: pd.Series, dic_metas: dict) -> tuple[str, str]:
    """Constrói as tabelas HTML diária e anual com indicadores, metas e semáforos.

    Args:
        colunas_dados (pd.Series): Série contendo os indicadores realizados da loja ou rede.
        dic_metas (dict): Dicionário com as metas correspondentes.

    Returns:
        tuple[str, str]: Strings HTML das tabelas de desempenho diário e anual.
    """
    def montar_tabela(sufixo_col: str, sufixo_meta: str, titulo_periodo: str) -> str:
        df_visual = pd.DataFrame(
            {
                f"Valor {titulo_periodo} (R$)": [
                    int(colunas_dados[f"Prod_Div_{sufixo_col}"]),
                    converter_notacao_europeia(colunas_dados[f"Fat_{sufixo_col}"]),
                    converter_notacao_europeia(colunas_dados[f"Tic_Medio_{sufixo_col}"]),
                ],
                f"Meta {titulo_periodo} (R$)": [
                    int(dic_metas[f"div_{sufixo_meta}"]),
                    converter_notacao_europeia(dic_metas[f"fat_{sufixo_meta}"]),
                    converter_notacao_europeia(dic_metas[f"tic_{sufixo_meta}"]),
                ],
                f"Cenário {titulo_periodo}": [
                    "🟢" if colunas_dados[f"Meta_Div_{sufixo_col}"] else "🔴",
                    "🟢" if colunas_dados[f"Meta_Fat_{sufixo_col}"] else "🔴",
                    "🟢" if colunas_dados[f"Meta_Tic_{sufixo_col}"] else "🔴",
                ],
            },
            index=["Diversidade de Produtos", "Faturamento", "Ticket Médio"],
        )
        return df_visual.to_html(border=0)

    return montar_tabela("Ontem", "dia", "Dia"), montar_tabela("Ano", "ano", "Ano")


def construir_html_email(
    nome_remetente: str,
    data_ontem_str_dma: str,
    colunas_dados: pd.Series,
    dic_metas: dict,
    template: Template,
    lst_cids_imagens: list[str],
    id_log: str,
    **dados_extras,
) -> tuple[str | None, dict[str, str] | None]:
    """Renderiza o template HTML com os dados, converte estilos para inline e mapeia CIDs.

    Args:
        nome_remetente (str): Nome do remetente da mensagem.
        data_ontem_str_dma (str): Data do dia anterior no formato dd/mm/aaaa.
        colunas_dados (pd.Series): Linha de dados consolidados da loja ou rede.
        dic_metas (dict): Dicionário de metas da loja ou rede.
        template (Template): Objeto Template compilado do Jinja2.
        lst_cids_imagens (list[str]): Lista de identificadores 'cid:' presentes no template.
        id_log (str): Identificador para registro de log em caso de falha.
        **dados_extras: Variáveis adicionais específicas de cada template.

    Returns:
        tuple[str | None, dict[str, str] | None]: HTML final com CSS inline e dicionário
            mapeando {cid_original: cid_gerado_sem_colchetes}. Retorna (None, None) em caso de erro.
    """
    try:
        html_dia, html_ano = _gerar_tabelas_indicadores_html(colunas_dados, dic_metas)
        html_corpo = template.render(
            data_ontem=data_ontem_str_dma,
            indicadores_dia=html_dia,
            indicadores_ano=html_ano,
            nome_remetente=nome_remetente,
            **dados_extras,
        )
        html_corpo = transform(html_corpo)

        dic_mapa_cids = {}
        for cid_original in (lst_cids_imagens or []):
            cid_sem_colch = make_msgid(domain="controladoria")[1:-1]
            html_corpo = html_corpo.replace(f"cid:{cid_original}", f"cid:{cid_sem_colch}")
            dic_mapa_cids[cid_original] = cid_sem_colch

        return html_corpo, dic_mapa_cids
    except Exception as erro:
        registrar_log(f"Erro ao renderizar o template HTML para {id_log}: {erro}!", "CRITICAL")
        return None, None


def construir_grafico_evol_vendas(
    loja_id_nome: str | None,
    meses: pd.Series,
    quantidades: pd.Series,
    faturamentos: pd.Series,
    caminho_grafico: Path | None,
    id_log: str,
) -> bytes | None:
    """Gera o gráfico de evolução mensal (faturamento e quantidade) em bytes PNG.

    Args:
        loja_id_nome (str | None): Identificação 'ID - Nome' da loja, ou None para a rede.
        meses (pd.Series): Série temporal dos meses.
        quantidades (pd.Series): Série de quantidades vendidas por mês.
        faturamentos (pd.Series): Série de faturamento por mês.
        caminho_grafico (Path | None): Caminho para salvar a imagem em disco, se configurado.
        id_log (str): Identificador para registro de log em caso de erro.

    Returns:
        bytes | None: Imagem do gráfico em bytes PNG, ou None em caso de erro.
    """
    try:
        grafico_mensal = make_subplots(specs=[[{"secondary_y": True}]])
        grafico_mensal.add_trace(
            go.Bar(x=meses, y=faturamentos, name="Faturamento (R$)", offsetgroup=1),
            secondary_y=False,
        )
        grafico_mensal.add_trace(
            go.Bar(x=meses, y=quantidades, name="Quantidade de Produtos", offsetgroup=2),
            secondary_y=True,
        )

        titulo = (
            f"Evolução das Vendas Mensais da Loja {loja_id_nome}"
            if loja_id_nome is not None
            else "Evolução das Vendas Mensais da Rede de Lojas"
        )
        grafico_mensal.update_layout(
            template="plotly_dark",
            barmode="group",
            title=dict(text=titulo, x=0.5, xanchor="center"),
            margin=dict(t=85),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        )
        grafico_mensal.update_xaxes(ticklabelstandoff=15, tickangle=-45, automargin=True)
        grafico_mensal.update_yaxes(title_text="Faturamento (R$)", secondary_y=False)
        grafico_mensal.update_yaxes(title_text="Quantidade de Produtos", secondary_y=True)

        img_bytes = grafico_mensal.to_image(format="png", width=650, height=380, scale=2)
        if caminho_grafico:
            caminho_grafico.write_bytes(img_bytes)

        return img_bytes
    except Exception as erro:
        registrar_log(f"Erro ao gerar o gráfico de evolução mensal para {id_log}: {erro}!", "CRITICAL")
        return None

def construir_grafico_anual_lojas(df_lojas_ano: pd.DataFrame, caminho_grafico: Path | None) -> bytes | None:
    """Gera o gráfico de barras horizontais do faturamento anual por loja em bytes PNG.

    Args:
        df_lojas_ano (pd.DataFrame): Tabela ordenada contendo as colunas 'Loja' e 'Fat_Ano'.
        caminho_grafico (Path | None): Caminho para salvar a imagem em disco, se configurado.

    Returns:
        bytes | None: Imagem do gráfico em bytes PNG, ou None em caso de erro.
    """
    try:
        paleta_base = pc.qualitative.Pastel + pc.qualitative.Set3 + pc.qualitative.Bold
        cores_barras = [paleta_base[i % len(paleta_base)] for i in range(len(df_lojas_ano))]
        nomes_dentro_barras = [f"<i>{loja}</i>" for loja in df_lojas_ano["Loja"]]
        rotulos_valores_esquerda = [f"R$ {converter_notacao_europeia(valor)} " for valor in df_lojas_ano["Fat_Ano"]]

        grafico_anual = go.Figure(
            data=[
                go.Bar(
                    x=df_lojas_ano["Fat_Ano"],
                    y=df_lojas_ano["Loja"],
                    orientation="h",
                    marker_color=cores_barras,
                    text=nomes_dentro_barras,
                    textposition="inside",
                    insidetextanchor="start",
                    textfont=dict(color="black", size=12),
                    constraintext="none",
                )
            ]
        )

        maior_valor = df_lojas_ano["Fat_Ano"].max()
        grafico_anual.update_layout(
            template="plotly_dark",
            title=dict(text="Faturamento Acumulado Anual por Loja", x=0.5, xanchor="center"),
            margin=dict(l=20, r=30, t=60, b=30),
        )
        grafico_anual.update_xaxes(
            showticklabels=False,
            showgrid=False,
            zeroline=False,
            range=[0, maior_valor * 1.05 if maior_valor > 0 else 1],
        )
        grafico_anual.update_yaxes(
            tickmode="array",
            tickvals=df_lojas_ano["Loja"],
            ticktext=rotulos_valores_esquerda,
            tickfont=dict(family="Consolas, monospace", size=12, color="white"),
            automargin=True,
            showgrid=False,
        )

        altura_grafico = max(380, len(df_lojas_ano) * 28)
        img_bytes = grafico_anual.to_image(format="png", width=650, height=altura_grafico, scale=2)
        if caminho_grafico:
            caminho_grafico.write_bytes(img_bytes)

        return img_bytes
    except Exception as erro:
        registrar_log(f"Erro ao gerar o gráfico de faturamento anual por loja: {erro}!", "CRITICAL")
        return None


def _renderizar_pdf_playwright(msg_loja: bool, html_para_pdf: str, caminho_pdf: Path) -> None:
    """Converte o HTML recebido em PDF usando Chromium headless e o grava localmente.

    Args:
        msg_loja (bool): Sinalizador de mensagem enviada às lojas (ajusta para página única).
        html_para_pdf (str): HTML completo do corpo do e-mail com imagens em Base64.
        caminho_pdf (Path): Caminho completo do arquivo PDF a ser gerado.
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 750, "height": 800})
        page.set_content(html_para_pdf, wait_until="networkidle")

        altura_conteudo = page.evaluate("document.documentElement.scrollHeight")
        altura_maxima_a4 = 1060

        if msg_loja:
            escala = max(0.1, round(altura_maxima_a4 / altura_conteudo, 2)) if altura_conteudo > altura_maxima_a4 else 1.0
            margem = {"top": "4mm", "bottom": "4mm", "left": "10mm", "right": "10mm"}
        else:
            escala = 1.0
            margem = {"top": "12mm", "bottom": "12mm", "left": "12mm", "right": "12mm"}

        page.pdf(
            path=str(caminho_pdf),
            format="A4",
            print_background=True,
            scale=escala,
            margin=margem,
        )
        browser.close()


def gerar_pdf_do_email(
    msg_loja: bool,
    html_email: str,
    dic_imagens_embutidas: dict[str, bytes],
    eh_notebook: bool,
    caminho_pdf: Path,
    exibir_prints: bool,
    nome_anexo: str,
    bytes_anexo: bytes,
) -> None:
    """Simula o HTML do e-mail em arquivo PDF, suportando gráficos embutidos e cartão visual de anexo.

    Args:
        msg_loja (bool): Sinalizador de mensagem enviada às lojas.
        html_email (str): HTML do corpo do e-mail.
        dic_imagens_embutidas (dict[str, bytes]): Dicionário mapeando {cid_sem_colchetes: img_bytes}.
        eh_notebook (bool): Sinalizador de execução em ambiente Jupyter Notebook.
        caminho_pdf (Path): Caminho completo do arquivo PDF a ser gerado.
        exibir_prints (bool): Sinalizador para exibição ou não das mensagens informativas.
        nome_anexo (str): Nome do arquivo anexado para exibição visual no PDF.
        bytes_anexo (bytes): Conteúdo em bytes do anexo para cálculo do tamanho.
    """
    try:
        html_para_pdf = html_email
        for cid_s_colch, img_bytes in dic_imagens_embutidas.items():
            img_base64 = base64.b64encode(img_bytes).decode("utf-8")
            html_para_pdf = html_para_pdf.replace(f"cid:{cid_s_colch}", f"data:image/png;base64,{img_base64}")

        if nome_anexo and bytes_anexo:
            tamanho_kb = max(1, round(len(bytes_anexo) / 1024))
            bloco_anexo_html = (
                '<div style="display: inline-flex; align-items: center; border: 1px solid #dadce0; '
                'border-radius: 4px; padding: 8px 12px; margin-top: 24px; background-color: #f8f9fa; '
                'font-family: Arial, sans-serif; min-width: 240px; page-break-inside: avoid;">'
                '<div style="width: 24px; height: 24px; background-color: #188038; color: #ffffff; '
                'font-weight: bold; font-size: 13px; display: flex; align-items: center; '
                'justify-content: center; border-radius: 3px; margin-right: 12px;">X</div>'
                '<div style="line-height: 1.3;">'
                f'<div style="font-size: 12px; font-weight: bold; color: #202124;">{nome_anexo}</div>'
                f'<div style="font-size: 11px; color: #5f6368;">{tamanho_kb}K</div>'
                "</div>"
                "</div>"
            )
            html_para_pdf = html_para_pdf.replace("</body>", f"{bloco_anexo_html}\n</body>")

        if eh_notebook:
            with ThreadPoolExecutor(max_workers=1) as executor:
                futuro = executor.submit(_renderizar_pdf_playwright, msg_loja, html_para_pdf, caminho_pdf)
                futuro.result()
        else:
            _renderizar_pdf_playwright(msg_loja, html_para_pdf, caminho_pdf)
    except Exception as erro:
        registrar_log(f"Falha ao gerar PDF {caminho_pdf.name}: {erro}", "ERROR", exibir_prints)