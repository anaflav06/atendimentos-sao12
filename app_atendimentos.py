import streamlit as st
import json
import re
import base64
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO

import pandas as pd
import requests

st.set_page_config(
    page_title="Atendimentos",
    page_icon="📋",
    layout="centered"
)

FUNCIONARIOS = ["PAULO HENRIQUE", "PAULO SERGIO", "CLEBER", "RENATA"]

# GitHub configurado pelo Streamlit Secrets.
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
GITHUB_DATABASE_PATH = st.secrets.get("GITHUB_DATABASE_PATH", "database.json")


def github_headers():
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def validar_configuracao_github():
    faltando = []
    if not GITHUB_TOKEN:
        faltando.append("GITHUB_TOKEN")
    if not GITHUB_REPO:
        faltando.append("GITHUB_REPO")
    return faltando


def ler_database_github():
    """Retorna (dados, sha_atual)."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATABASE_PATH}"
    resp = requests.get(
        url,
        headers=github_headers(),
        params={"ref": GITHUB_BRANCH},
        timeout=20,
    )

    if resp.status_code == 404:
        return [], None

    resp.raise_for_status()
    payload = resp.json()

    conteudo = base64.b64decode(payload["content"]).decode("utf-8")
    dados = json.loads(conteudo) if conteudo.strip() else []

    if not isinstance(dados, list):
        raise ValueError("O database.json precisa conter uma lista JSON.")

    return dados, payload["sha"]


def salvar_database_github(dados, sha_atual=None):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATABASE_PATH}"

    conteudo_json = json.dumps(
        dados,
        ensure_ascii=False,
        indent=2
    )

    payload = {
        "message": f"Atualiza atendimentos - {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')}",
        "content": base64.b64encode(
            conteudo_json.encode("utf-8")
        ).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }

    if sha_atual:
        payload["sha"] = sha_atual

    resp = requests.put(
        url,
        headers=github_headers(),
        json=payload,
        timeout=20,
    )
    resp.raise_for_status()
    return resp.json()


def carregar_dados():
    dados, _ = ler_database_github()
    return dados


def adicionar_registro_github(registro, tentativas=3):
    """
    Lê a versão mais recente antes de salvar.
    Se dois usuários salvarem quase ao mesmo tempo, tenta novamente.
    """
    ultimo_erro = None

    for _ in range(tentativas):
        try:
            dados, sha = ler_database_github()
            dados.append(registro)
            salvar_database_github(dados, sha)
            return
        except requests.HTTPError as exc:
            ultimo_erro = exc
            # Conflito por SHA desatualizado: tenta reler e salvar novamente.
            if exc.response is not None and exc.response.status_code in (409, 422):
                continue
            raise

    if ultimo_erro:
        raise ultimo_erro


def somente_numeros(valor):
    return re.sub(r"\D", "", valor or "")


def formatar_telefone(valor):
    numeros = somente_numeros(valor)[:11]

    if len(numeros) == 10:
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    if len(numeros) == 11:
        return f"({numeros[:2]}) {numeros[2:7]}-{numeros[7:]}"
    return valor


def telefone_valido(valor):
    numeros = somente_numeros(valor)
    return len(numeros) in (10, 11)


def email_valido(email):
    email = (email or "").strip()
    padrao = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    return bool(re.fullmatch(padrao, email))


def cotacao_valida(valor):
    return bool(re.fullmatch(r"\d{8}", (valor or "").strip()))


def gerar_excel(dados):
    df = pd.DataFrame(dados)

    colunas = [
        "data",
        "hora",
        "funcionario",
        "nome_cliente",
        "telefone",
        "email",
        "numero_cotacao",
    ]

    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = ""

    df = df[colunas]

    df = df.rename(columns={
        "data": "DATA",
        "hora": "HORÁRIO",
        "funcionario": "FUNCIONÁRIO",
        "nome_cliente": "NOME DO CLIENTE",
        "telefone": "TELEFONE",
        "email": "E-MAIL",
        "numero_cotacao": "NÚMERO COTAÇÃO",
    })

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Atendimentos")
        ws = writer.book["Atendimentos"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        larguras = {
            "A": 13,
            "B": 11,
            "C": 22,
            "D": 28,
            "E": 20,
            "F": 36,
            "G": 18,
        }

        for coluna, largura in larguras.items():
            ws.column_dimensions[coluna].width = largura

        for celula in ws["E"][1:]:
            celula.number_format = "@"

        for celula in ws["G"][1:]:
            celula.number_format = "@"

    buffer.seek(0)
    return buffer.getvalue()


faltando = validar_configuracao_github()

if faltando:
    st.error(
        "Configuração do GitHub incompleta. "
        f"Adicione nos Secrets do Streamlit: {', '.join(faltando)}."
    )
    st.stop()


st.title("📋 Registro de Atendimento")
st.caption(
    "Cadastre a cotação de forma rápida. "
    "Data e horário são registrados automaticamente."
)

aba_cadastro, aba_relatorio = st.tabs(
    ["➕ Novo atendimento", "📊 Relatório"]
)

with aba_cadastro:
    with st.form("form_atendimento", clear_on_submit=True):

        funcionario = st.selectbox(
            "Funcionário *",
            ["SELECIONE..."] + FUNCIONARIOS
        )

        nome_cliente = st.text_input(
            "Nome do cliente *",
            placeholder="Ex.: MARIA SILVA"
        )

        telefone = st.text_input(
            "Telefone *",
            placeholder="Ex.: (11) 99999-9999"
        )

        email = st.text_input(
            "E-mail do cliente *",
            placeholder="Ex.: cliente@empresa.com.br"
        )

        numero_cotacao = st.text_input(
            "Número da cotação *",
            placeholder="Exatamente 8 dígitos",
            max_chars=8
        )

        salvar = st.form_submit_button(
            "💾 SALVAR ATENDIMENTO",
            use_container_width=True,
            type="primary"
        )

    if salvar:
        erros = []

        nome_cliente = nome_cliente.strip().upper()
        telefone_numeros = somente_numeros(telefone)
        email = email.strip().lower()
        numero_cotacao = somente_numeros(numero_cotacao)

        if funcionario == "SELECIONE...":
            erros.append("Selecione o funcionário.")

        if not nome_cliente:
            erros.append("Informe o nome do cliente.")

        if not telefone_valido(telefone_numeros):
            erros.append(
                "O telefone deve possuir 10 ou 11 dígitos, incluindo o DDD."
            )

        if not email_valido(email):
            erros.append("Informe um e-mail válido.")

        if not cotacao_valida(numero_cotacao):
            erros.append(
                "O número da cotação deve possuir exatamente 8 dígitos."
            )

        if erros:
            for erro in erros:
                st.error(erro)

        else:
            agora = datetime.now(
                ZoneInfo("America/Sao_Paulo")
            )

            registro = {
                "data": agora.strftime("%d/%m/%Y"),
                "hora": agora.strftime("%H:%M:%S"),
                "funcionario": funcionario,
                "nome_cliente": nome_cliente,
                "telefone": formatar_telefone(
                    telefone_numeros
                ),
                "email": email,
                "numero_cotacao": numero_cotacao,
            }

            try:
                adicionar_registro_github(registro)

                st.success(
                    "✅ Atendimento registrado com sucesso — "
                    f"Cotação {numero_cotacao}"
                )

            except Exception as exc:
                st.error(
                    "Não foi possível salvar o atendimento "
                    "na base do GitHub."
                )
                st.exception(exc)


with aba_relatorio:
    try:
        dados = carregar_dados()
    except Exception as exc:
        st.error("Não foi possível carregar a base do GitHub.")
        st.exception(exc)
        dados = []

    if not dados:
        st.info(
            "Nenhum atendimento registrado até o momento."
        )

    else:
        df = pd.DataFrame(dados)

        st.subheader("Atendimentos registrados")

        col1, col2 = st.columns(2)

        with col1:
            filtro_funcionario = st.multiselect(
                "Filtrar por funcionário",
                FUNCIONARIOS
            )

        datas_validas = pd.to_datetime(
            df.get("data", pd.Series(dtype=str)),
            format="%d/%m/%Y",
            errors="coerce"
        )

        data_min = datas_validas.min()
        data_max = datas_validas.max()

        with col2:
            if pd.notna(data_min) and pd.notna(data_max):
                periodo = st.date_input(
                    "Período",
                    value=(
                        data_min.date(),
                        data_max.date()
                    ),
                    format="DD/MM/YYYY"
                )
            else:
                periodo = ()

        df_filtrado = df.copy()

        if filtro_funcionario:
            df_filtrado = df_filtrado[
                df_filtrado["funcionario"].isin(
                    filtro_funcionario
                )
            ]

        if (
            isinstance(periodo, (tuple, list))
            and len(periodo) == 2
        ):
            datas_df = pd.to_datetime(
                df_filtrado["data"],
                format="%d/%m/%Y",
                errors="coerce"
            ).dt.date

            df_filtrado = df_filtrado[
                (datas_df >= periodo[0])
                & (datas_df <= periodo[1])
            ]

        exibicao = df_filtrado.rename(columns={
            "data": "DATA",
            "hora": "HORÁRIO",
            "funcionario": "FUNCIONÁRIO",
            "nome_cliente": "NOME DO CLIENTE",
            "telefone": "TELEFONE",
            "email": "E-MAIL",
            "numero_cotacao": "NÚMERO COTAÇÃO",
        })

        ordem = [
            "DATA",
            "HORÁRIO",
            "FUNCIONÁRIO",
            "NOME DO CLIENTE",
            "TELEFONE",
            "E-MAIL",
            "NÚMERO COTAÇÃO",
        ]

        exibicao = exibicao[
            [c for c in ordem if c in exibicao.columns]
        ]

        st.dataframe(
            exibicao,
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            f"{len(df_filtrado)} atendimento(s) encontrado(s)."
        )

        nome_arquivo = (
            "relatorio_atendimentos_"
            + datetime.now(
                ZoneInfo("America/Sao_Paulo")
            ).strftime("%Y%m%d_%H%M")
            + ".xlsx"
        )

        st.download_button(
            "📥 Baixar relatório em Excel",
            data=gerar_excel(
                df_filtrado.to_dict("records")
            ),
            file_name=nome_arquivo,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

st.divider()
st.caption("Base de dados persistente: database.json no GitHub.")
