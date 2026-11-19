"""Módulo de conexão SMTP, formatação de anexos e envio de e-mails."""

import io
import smtplib
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from jinja2 import Template

from configuracao import registrar_log
from processamento import _formatar_aba_excel
from relatorios import (
    construir_grafico_anual_lojas,
    construir_grafico_evol_vendas,
    construir_html_email,
    converter_notacao_europeia,
    gerar_pdf_do_email,
)


# ==============================================================================
# CONFIGURAÇÃO DO PDOC (EXIBIÇÃO DE FUNÇÕES INTERNAS)
# ==============================================================================
__all__ = [
    "conectar_servidor_smtp",
    "desconectar_servidor_smtp",
    "_converter_df_bytes",
    "enviar_email",
    "_processar_pdf_e_enviar_email",
    "_preparar_ranking_lojas",
    "enviar_emails_lojas",
    "enviar_email_diretoria",
]


def conectar_servidor_smtp(
    email_remetente: str, senha_email: str, porta_smtp: str, servidor_smtp: str
) -> smtplib.SMTP | None:
    """Abre conexão segura (TLS) e realiza autenticação no servidor SMTP.

    Args:
        email_remetente (str): Endereço de e-mail da conta remetente.
        senha_email (str): Senha ou senha de aplicativo da conta remetente.
        porta_smtp (str): Porta de comunicação do servidor SMTP.
        servidor_smtp (str): Endereço do host SMTP.

    Returns:
        smtplib.SMTP | None: Objeto de conexão SMTP autenticado, ou None em caso de falha.
    """
    try:
        servidor_email = smtplib.SMTP(servidor_smtp, int(porta_smtp))
        servidor_email.starttls()
        servidor_email.login(email_remetente, senha_email)
        return servidor_email
    except Exception as erro:
        registrar_log(f"Erro na conexão com o servidor SMTP: {erro}!", "CRITICAL")
        return None


def desconectar_servidor_smtp(servidor_email: smtplib.SMTP, exibir_prints: bool) -> None:
    """Encerra a conexão ativa com o servidor SMTP.

    Args:
        servidor_email (smtplib.SMTP): Instância conectada do servidor SMTP.
        exibir_prints (bool): Indica se eventuais avisos devem ser exibidos no console.
    """
    try:
        servidor_email.quit()
    except Exception as erro:
        txt_erro = f"Erro na desconexão do servidor SMTP: {erro}!"
        registrar_log(txt_erro, "WARNING", exibir_prints)


def _converter_df_bytes(dic_abas: dict[str, pd.DataFrame]) -> bytes:
    """Converte um dicionário de DataFrames em bytes de abas formatadas no formato Excel (.xlsx).

    Args:
        dic_abas (dict[str, pd.DataFrame]): Dicionário mapeando {nome_aba: dataframe}.

    Returns:
        bytes: Conteúdo binário do arquivo Excel gerado.
    """
    buffer = io.BytesIO()

    formatos_diretoria = {
        "Faturamento (R$)": "#,##0.00",
        "Meta Faturamento (R$)": "#,##0.00",
        "Diversidade de Produtos": "#,##0",
        "Meta Diversidade": "#,##0",
        "Ticket Médio (R$)": "#,##0.00",
        "Meta Ticket (R$)": "#,##0.00",
    }
    semaforos_diretoria = ["Faturamento Ok", "Diversidade Ok", "Ticket Ok"]

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for aba, df in dic_abas.items():
            df.to_excel(writer, index=False, sheet_name=aba)
            _formatar_aba_excel(
                writer.sheets[aba],
                df,
                formatos_colunas=formatos_diretoria,
                colunas_semaforo=semaforos_diretoria,
            )

    return buffer.getvalue()


def enviar_email(
    servidor_email: smtplib.SMTP,
    assunto: str,
    remetente: str,
    destinatario: str,
    html_corpo: str,
    dic_imagens_embutidas: dict[str, bytes],
    caminho_planilha: Path | None,
    bytes_anexo: bytes,
    nome_anexo: str,
    manter_planilhas: bool,
    exibir_prints: bool,
    id_log: str,
) -> bool:
    """Envia e-mails com corpo HTML, imagens embutidas e planilha Excel em anexo.

    Args:
        servidor_email (smtplib.SMTP): Servidor de e-mail a ser utilizado para o envio da mensagem.
        assunto (str): Assunto da mensagem.
        remetente (str): Endereço de e-mail do remetente da mensagem.
        destinatario (str): Endereço de e-mail do destinatário da mensagem.
        html_corpo (str): HTML do corpo da mensagem.
        dic_imagens_embutidas (dict[str, bytes]): Dicionário mapeando {cid_com_colchetes: img_bytes}.
        caminho_planilha (Path | None): Caminho da planilha de loja a ser removida, se configurado.
        bytes_anexo (bytes): Arquivo em bytes para ser anexado na mensagem.
        nome_anexo (str): Nome do arquivo que será anexado na mensagem.
        manter_planilhas (bool): Indicador de manutenção da planilha em armazenamento local.
        exibir_prints (bool): Indicador para exibição de mensagens informativas durante a execução.
        id_log (str): Texto identificador a ser registrado no arquivo de log em caso de erro.

    Returns:
        bool: Sucesso ou não da operação.
    """
    try:
        msg = EmailMessage()
        msg["Subject"] = assunto
        msg["From"] = remetente
        msg["To"] = destinatario

        msg.set_content("Seu leitor de e-mail não suporta visualização em HTML.")
        msg.add_alternative(html_corpo, subtype="html")

        for cid_com_colch, img_bytes in dic_imagens_embutidas.items():
            msg.get_payload()[1].add_related(img_bytes, maintype="image", subtype="png", cid=cid_com_colch)

        msg.add_attachment(
            bytes_anexo,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nome_anexo,
        )

        if caminho_planilha and not manter_planilhas:
            caminho_planilha.unlink(missing_ok=True)

        servidor_email.send_message(msg)
        registrar_log(f"E-mail para {id_log} enviado com sucesso!", "INFO", exibir_prints)

        return True
    except Exception as erro:
        registrar_log(f"Falha ao enviar e-mail para {id_log}: {erro}", "CRITICAL")
        return False


def _processar_pdf_e_enviar_email(
    servidor_email: smtplib.SMTP,
    assunto: str,
    destinatario: str,
    html_corpo: str,
    dic_mapa_cids: dict[str, str],
    dic_imagens_originais: dict[str, bytes],
    caminho_pdf: Path | None,
    eh_notebook: bool,
    msg_loja: bool,
    caminho_planilha: Path | None,
    bytes_anexo: bytes,
    nome_anexo_email: str,
    nome_anexo_pdf: str,
    manter_planilhas: bool,
    exibir_prints: bool,
    id_log: str,
) -> bool:
    """Mapeia os CIDs das imagens, gera o PDF opcional de simulação e envia o e-mail.

    Args:
        servidor_email (smtplib.SMTP): Servidor SMTP autenticado.
        assunto (str): Assunto do e-mail.
        destinatario (str): Endereço de e-mail do destinatário.
        html_corpo (str): Corpo HTML renderizado.
        dic_mapa_cids (dict[str, str]): Mapeamento {cid_original: cid_sem_colchetes}.
        dic_imagens_originais (dict[str, bytes]): Mapeamento {cid_original: img_bytes}.
        caminho_pdf (Path | None): Caminho para salvar o PDF de simulação, ou None.
        eh_notebook (bool): Indicador de ambiente Jupyter Notebook.
        msg_loja (bool): Indica se o layout do PDF deve ser ajustado para loja.
        caminho_planilha (Path | None): Caminho físico da planilha da loja, se houver.
        bytes_anexo (bytes): Conteúdo binário do anexo Excel.
        nome_anexo_email (str): Nome do anexo enviado no e-mail.
        nome_anexo_pdf (str): Nome do anexo exibido no cartão visual do PDF.
        manter_planilhas (bool): Indica se a planilha física deve ser mantida após o envio.
        exibir_prints (bool): Indica se mensagens informativas devem ser exibidas.
        id_log (str): Identificador para registro de log.

    Returns:
        bool: True se o envio foi bem-sucedido, False caso contrário.
    """
    dic_cids_sem_colch = {dic_mapa_cids[chave]: img_bytes for chave, img_bytes in dic_imagens_originais.items()}
    dic_cids_com_colch = {f"<{dic_mapa_cids[chave]}>": img_bytes for chave, img_bytes in dic_imagens_originais.items()}

    if caminho_pdf:
        gerar_pdf_do_email(
            msg_loja=msg_loja,
            html_email=html_corpo,
            dic_imagens_embutidas=dic_cids_sem_colch,
            eh_notebook=eh_notebook,
            caminho_pdf=caminho_pdf,
            exibir_prints=exibir_prints,
            nome_anexo=nome_anexo_pdf,
            bytes_anexo=bytes_anexo,
        )

    return enviar_email(
        servidor_email=servidor_email,
        assunto=assunto,
        remetente=servidor_email.user,
        destinatario=destinatario,
        html_corpo=html_corpo,
        dic_imagens_embutidas=dic_cids_com_colch,
        caminho_planilha=caminho_planilha,
        bytes_anexo=bytes_anexo,
        nome_anexo=nome_anexo_email,
        manter_planilhas=manter_planilhas,
        exibir_prints=exibir_prints,
        id_log=id_log,
    )


def _preparar_ranking_lojas(
    df_consolidado: pd.DataFrame,
    dic_metas_lojas: dict,
    dic_nomes_lojas: dict[int, str],
    sufixo_col: str,
    sufixo_meta: str,
) -> pd.DataFrame:
    """Ordena, renomeia, adiciona metas e estrutura a tabela de ranking para um período específico.

    Args:
        df_consolidado (pd.DataFrame): Tabela consolidada das lojas.
        dic_metas_lojas (dict): Dicionário com as metas das lojas.
        dic_nomes_lojas (dict[int, str]): Dicionário mapeando {id_loja: nome_loja}.
        sufixo_col (str): Sufixo das colunas do DataFrame ('Ano' ou 'Ontem').
        sufixo_meta (str): Sufixo das chaves no dicionário de metas ('ano' ou 'dia').

    Returns:
        pd.DataFrame: Tabela de ranking pronta para exportação.
    """
    col_fat = f"Fat_{sufixo_col}"
    df_rank = df_consolidado.sort_index(ascending=True).sort_values(by=col_fat, ascending=False).copy()

    df_rank = df_rank.rename(
        columns={
            f"Fat_{sufixo_col}": "Faturamento (R$)",
            f"Meta_Fat_{sufixo_col}": "Faturamento Ok",
            f"Prod_Div_{sufixo_col}": "Diversidade de Produtos",
            f"Meta_Div_{sufixo_col}": "Diversidade Ok",
            f"Tic_Medio_{sufixo_col}": "Ticket Médio (R$)",
            f"Meta_Tic_{sufixo_col}": "Ticket Ok",
        }
    )

    df_rank["Loja"] = df_rank.index.astype(str) + " - " + df_rank.index.map(dic_nomes_lojas)
    df_rank["Meta Faturamento (R$)"] = dic_metas_lojas[f"fat_{sufixo_meta}"]
    df_rank["Meta Diversidade"] = dic_metas_lojas[f"div_{sufixo_meta}"]
    df_rank["Meta Ticket (R$)"] = dic_metas_lojas[f"tic_{sufixo_meta}"]

    return df_rank[
        [
            "Loja",
            "Faturamento (R$)",
            "Meta Faturamento (R$)",
            "Faturamento Ok",
            "Diversidade de Produtos",
            "Meta Diversidade",
            "Diversidade Ok",
            "Ticket Médio (R$)",
            "Meta Ticket (R$)",
            "Ticket Ok",
        ]
    ]


def enviar_emails_lojas(
    servidor_email: smtplib.SMTP,
    nome_remetente: str,
    df_lojas: pd.DataFrame,
    dic_metas: dict,
    df_vendas_consolidadas_lojas: pd.DataFrame,
    data_ontem_str_dma: str,
    data_ontem_str_amd: str,
    template_compilado: Template,
    dic_planilhas_lojas: dict[int, Path],
    caminho_graficos: Path | None,
    caminho_pdfs: Path | None,
    eh_notebook: bool,
    exibir_prints: bool,
    manter_planilhas: bool,
) -> bool:
    """Envia e-mails individuais às lojas com desempenho diário/anual, gráfico mensal e planilha de vendas.

    Args:
        servidor_email (smtplib.SMTP): Servidor de e-mail ativo.
        nome_remetente (str): Nome do remetente.
        df_lojas (pd.DataFrame): Tabela de dados cadastrais das lojas.
        dic_metas (dict): Dicionário com as metas diárias e anuais das lojas.
        df_vendas_consolidadas_lojas (pd.DataFrame): Tabela de vendas consolidadas das lojas.
        data_ontem_str_dma (str): Data do dia anterior em formato dd/mm/aaaa.
        data_ontem_str_amd (str): Data do dia anterior em formato aaaa-mm-dd.
        template_compilado (Template): Template HTML compilado das lojas.
        dic_planilhas_lojas (dict[int, Path]): Dicionário mapeando ID da loja ao caminho da sua planilha.
        caminho_graficos (Path | None): Diretório para salvar imagens dos gráficos ou None.
        caminho_pdfs (Path | None): Diretório para salvar PDFs de simulação ou None.
        eh_notebook (bool): Indicador de ambiente Jupyter Notebook.
        exibir_prints (bool): Indicador de exibição de mensagens informativas.
        manter_planilhas (bool): Indicador de manutenção das planilhas em disco após o envio.

    Returns:
        bool: Sucesso ou falha no envio para todas as lojas.
    """
    try:
        df_lojas_geral = df_vendas_consolidadas_lojas.join(df_lojas[["Loja", "Gerente", "E-mail"]])

        for id_loja, colunas in df_lojas_geral.iterrows():
            id_log = f"loja {id_loja}"
            loja_id_nome = f"{id_loja} - {colunas['Loja']}"
            registrar_log(f"Processando e-mail da loja: {loja_id_nome}", "INFO", exibir_prints)

            caminho_planilha_loja = dic_planilhas_lojas[id_loja]
            df_mensal_loja = pd.read_excel(caminho_planilha_loja, sheet_name="Mensal")

            html_loja, dic_cids_loja = construir_html_email(
                nome_remetente=nome_remetente,
                data_ontem_str_dma=data_ontem_str_dma,
                colunas_dados=colunas,
                dic_metas=dic_metas,
                template=template_compilado,
                lst_cids_imagens=["grafico_vendas"],
                id_log=id_log,
                gerente=colunas["Gerente"],
                nome_loja=colunas["Loja"],
            )
            if html_loja is None:
                return False

            caminho_img_loja = (
                caminho_graficos / f"{data_ontem_str_amd}_Vendas_Mensais_Loja_{id_loja}.png"
                if caminho_graficos
                else None
            )
            img_bytes = construir_grafico_evol_vendas(
                loja_id_nome=loja_id_nome,
                meses=df_mensal_loja["Mês"],
                quantidades=df_mensal_loja["Quantidade"],
                faturamentos=df_mensal_loja["Valor Final"],
                caminho_grafico=caminho_img_loja,
                id_log=id_log,
            )
            if not img_bytes:
                return False

            bytes_planilha = caminho_planilha_loja.read_bytes()
            caminho_pdf_loja = (
                caminho_pdfs / f"{data_ontem_str_amd} - OnePage - Loja {loja_id_nome}.pdf"
                if caminho_pdfs
                else None
            )

            if not _processar_pdf_e_enviar_email(
                servidor_email=servidor_email,
                assunto=f"OnePage Dia {data_ontem_str_dma} - Loja {loja_id_nome}",
                destinatario=colunas["E-mail"],
                html_corpo=html_loja,
                dic_mapa_cids=dic_cids_loja,
                dic_imagens_originais={"grafico_vendas": img_bytes},
                caminho_pdf=caminho_pdf_loja,
                eh_notebook=eh_notebook,
                msg_loja=True,
                caminho_planilha=caminho_planilha_loja,
                bytes_anexo=bytes_planilha,
                nome_anexo_email=caminho_planilha_loja.name,
                nome_anexo_pdf=caminho_planilha_loja.name,
                manter_planilhas=manter_planilhas,
                exibir_prints=exibir_prints,
                id_log=id_log,
            ):
                return False

        return True
    except Exception as erro:
        registrar_log(f"Erro no envio de e-mails às lojas: {erro}!", "CRITICAL")
        return False


def enviar_email_diretoria(
    servidor_email: smtplib.SMTP,
    nome_remetente: str,
    df_vendas: pd.DataFrame,
    dic_metas_lojas: dict,
    dic_metas_rede: dict,
    df_vendas_consolidadas_lojas: pd.DataFrame,
    df_vendas_consolidadas_rede: pd.DataFrame,
    email_diretoria: str,
    data_ontem_str_dma: str,
    data_ontem_str_amd: str,
    template_compilado: Template,
    caminho_graficos: Path | None,
    caminho_pdfs: Path | None,
    eh_notebook: bool,
    dic_nomes_lojas: dict[int, str],
    exibir_prints: bool,
) -> bool:
    """Envia e-mail consolidado à diretoria com indicadores da rede, gráficos e planilha de rankings.

    Args:
        servidor_email (smtplib.SMTP): Servidor de e-mail ativo.
        nome_remetente (str): Nome do remetente.
        df_vendas (pd.DataFrame): Tabela geral de vendas.
        dic_metas_lojas (dict): Dicionário com as metas anuais e diárias das lojas.
        dic_metas_rede (dict): Dicionário com as metas anuais e diárias da rede.
        df_vendas_consolidadas_lojas (pd.DataFrame): Tabela de vendas consolidadas das lojas.
        df_vendas_consolidadas_rede (pd.DataFrame): Tabela de vendas consolidadas de toda a rede.
        email_diretoria (str): Endereço de e-mail da diretoria.
        data_ontem_str_dma (str): Data do dia anterior em formato dd/mm/aaaa.
        data_ontem_str_amd (str): Data do dia anterior em formato aaaa-mm-dd.
        template_compilado (Template): Template HTML compilado da diretoria.
        caminho_graficos (Path | None): Diretório para salvar imagens dos gráficos ou None.
        caminho_pdfs (Path | None): Diretório para salvar PDFs de simulação ou None.
        eh_notebook (bool): Indicador de ambiente Jupyter Notebook.
        dic_nomes_lojas (dict[int, str]): Dicionário de nomes das lojas.
        exibir_prints (bool): Indicador de exibição de mensagens informativas.

    Returns:
        bool: Sucesso ou falha no envio à diretoria.
    """
    try:
        df_ranking_ano = _preparar_ranking_lojas(
            df_vendas_consolidadas_lojas, dic_metas_lojas, dic_nomes_lojas, "Ano", "ano"
        )
        df_ranking_ontem = _preparar_ranking_lojas(
            df_vendas_consolidadas_lojas, dic_metas_lojas, dic_nomes_lojas, "Ontem", "dia"
        )

        fat_ano_max = df_ranking_ano["Faturamento (R$)"].iloc[0]
        lojas_max_ano = df_ranking_ano[df_ranking_ano["Faturamento (R$)"] == fat_ano_max]["Loja"].tolist()
        fat_ano_min = df_ranking_ano["Faturamento (R$)"].iloc[-1]
        lojas_min_ano = df_ranking_ano[df_ranking_ano["Faturamento (R$)"] == fat_ano_min]["Loja"].tolist()

        fat_ontem_max = df_ranking_ontem["Faturamento (R$)"].iloc[0]
        lojas_max_ontem = df_ranking_ontem[df_ranking_ontem["Faturamento (R$)"] == fat_ontem_max]["Loja"].tolist()
        fat_ontem_min = df_ranking_ontem["Faturamento (R$)"].iloc[-1]
        lojas_min_ontem = df_ranking_ontem[df_ranking_ontem["Faturamento (R$)"] == fat_ontem_min]["Loja"].tolist()

        id_log = "diretoria"

        html_diretoria, dic_cids_diretoria = construir_html_email(
            nome_remetente=nome_remetente,
            data_ontem_str_dma=data_ontem_str_dma,
            colunas_dados=df_vendas_consolidadas_rede.iloc[0],
            dic_metas=dic_metas_rede,
            template=template_compilado,
            lst_cids_imagens=["grafico_mensal_rede", "grafico_anual_lojas"],
            id_log=id_log,
            qtd_lojas=len(df_vendas_consolidadas_lojas),
            fat_ano_max=converter_notacao_europeia(fat_ano_max),
            lojas_max_ano=lojas_max_ano,
            fat_ano_min=converter_notacao_europeia(fat_ano_min),
            lojas_min_ano=lojas_min_ano,
            fat_ontem_max=converter_notacao_europeia(fat_ontem_max),
            lojas_max_ontem=lojas_max_ontem,
            fat_ontem_min=converter_notacao_europeia(fat_ontem_min),
            lojas_min_ontem=lojas_min_ontem,
        )
        if html_diretoria is None:
            return False

        df_rede_mes = (
            df_vendas.groupby(df_vendas["Data"].dt.to_period("M"))[["Quantidade", "Valor Final"]]
            .sum()
            .reset_index()
            .rename(columns={"Data": "Mês"})
        )
        df_rede_mes["Mês"] = df_rede_mes["Mês"].astype(str)

        caminho_img_rede = (
            caminho_graficos / f"{data_ontem_str_amd}_Vendas_Mensais_Rede.png" if caminho_graficos else None
        )
        img_bytes_mensal_rede = construir_grafico_evol_vendas(
            loja_id_nome=None,
            meses=df_rede_mes["Mês"],
            quantidades=df_rede_mes["Quantidade"],
            faturamentos=df_rede_mes["Valor Final"],
            caminho_grafico=caminho_img_rede,
            id_log=id_log,
        )
        if not img_bytes_mensal_rede:
            return False

        df_grafico_anual = (
            df_ranking_ano[["Loja", "Faturamento (R$)"]]
            .rename(columns={"Faturamento (R$)": "Fat_Ano"})
            .sort_values(by="Fat_Ano", ascending=True)
            .reset_index(drop=True)
        )
        caminho_img_anual = (
            caminho_graficos / f"{data_ontem_str_amd}_Vendas_Anuais_Lojas.png" if caminho_graficos else None
        )
        img_bytes_anual_lojas = construir_grafico_anual_lojas(df_grafico_anual, caminho_img_anual)
        if not img_bytes_anual_lojas:
            return False

        bytes_df_rede = _converter_df_bytes({"Ranking Anual": df_ranking_ano, "Ranking Diário": df_ranking_ontem})
        caminho_pdf_diretoria = (
            caminho_pdfs / f"{data_ontem_str_amd} - OnePage - Rede de Lojas.pdf" if caminho_pdfs else None
        )

        return _processar_pdf_e_enviar_email(
            servidor_email=servidor_email,
            assunto=f"OnePage Dia {data_ontem_str_dma} - Rede de Lojas",
            destinatario=email_diretoria,
            html_corpo=html_diretoria,
            dic_mapa_cids=dic_cids_diretoria,
            dic_imagens_originais={
                "grafico_mensal_rede": img_bytes_mensal_rede,
                "grafico_anual_lojas": img_bytes_anual_lojas,
            },
            caminho_pdf=caminho_pdf_diretoria,
            eh_notebook=eh_notebook,
            msg_loja=False,
            caminho_planilha=None,
            bytes_anexo=bytes_df_rede,
            nome_anexo_email=f"{data_ontem_str_amd} - Ranking_Vendas_Lojas.xlsx",
            nome_anexo_pdf="Ranking_Vendas_Lojas.xlsx",
            manter_planilhas=True,
            exibir_prints=exibir_prints,
            id_log=id_log,
        )
    except Exception as erro:
        registrar_log(f"Erro no envio de e-mail à diretoria: {erro}!", "CRITICAL")
        return False
