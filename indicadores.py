"""
LOCALIZADOR INTELIGENTE DE INDICADORES (v4.0)
=============================================
Aplicação Streamlit de arquivo único.

Requisitos: streamlit, pandas>=2.0, numpy, plotly, openpyxl, python-pptx, python-docx, PyMuPDF (leitura de PDF)
Execução:   streamlit run localizador_indicadores.py

Novidades da v4.0 (tudo local; sem banco, API, IA, nuvem ou envio externo)
- Documentos de texto (aba própria, isolada das planilhas): base documental com identidade por conteúdo (SHA-256), duplicidade e
  versões explícitas; busca (BM25 + sinônimos) com filtros; anotações e etiquetas separadas do texto original; extração segura de
  números; pontos de atenção e perguntas; comparação de documentos; vínculos manuais com indicadores; reuniões e atas (só com o
  que o analista confirmou); produção de documentos a partir de 11 modelos (MD/TXT/DOCX); pacote de trabalho JSON (exportar/importar).
- Design: tokens visuais, cartão de KPI com especificação única (tela, relatório HTML e PowerPoint) e estados padronizados.
- Produtividade: dicionário de dados (CSV/Excel) e comparação de dois arquivos por coluna-chave, sem alterar as bases.

Novidades da v3.6 (múltiplas fontes e relacionamentos)
- Importação de VÁRIOS CSV (cada um com a mesma validação de separador/integridade) e aba "Modelo de dados":
  tabela fato, relacionamentos muitos-para-um validados (chave única, nulos, órfãos, cardinalidade), sugestões
  automáticas, tabelas filhas (1:N) apenas AGREGADAS, e base analítica que NUNCA altera o nº de linhas da fato.
- O motor de análise (classificação, indicadores, filtros, painel, relatório) continua igual: passa a receber a base analítica.
- Órfãos, duplicidades e relacionamentos ignorados entram no diagnóstico; a linhagem das fontes entra no relatório.

Novidades da v3.4 (terceira auditoria)
- Taxas: a ESCALA que reproduziu a coluna é preservada (× 100 para "25"; × 1 para "0.25"). Dados validados antes do
  cálculo (denominador ≤ 0, numerador negativo, vazios, exclusões e denominador total zero); mostra numerador,
  denominador e escala. Sem validação, nenhuma taxa oficial é apresentada.
- CSV: nº de campos de CADA registro conferido no arquivo inteiro; nada é descartado nem preenchido em silêncio;
  contagens (válidos/curtos/longos/aproveitados/descartados) e confirmação do tratamento pelo usuário.
- Média ponderada: peso escolhido explicitamente (só quantidade/valor monetário), com tratamento de pesos
  nulos, zero e negativos informado.
- Consolidação por categoria: informa total da base, registros com categoria válida e excluídos.
- Meses parciais marcados como ESTIMATIVA; período esperado opcional (aba Indicadores).
- Auditoria de formatos: datas MM/DD vs DD/MM e números com ponto (milhar x decimal).
- Fórmulas exibidas = fórmulas calculadas; relatório com validações e notas metodológicas.
- Identidade visual com as cores da logo da Santa Casa de Barra Mansa (logo opcional ao lado do script).

Novidades da v3.3 (segunda revisão técnica)
- Unidades: 'total'/'valor' são nomes ambíguos (não atribuem R$ nem soma). Naturezas numéricas explícitas:
  quantidade | monetario | percentual | indefinido (editáveis na Visão geral). Identificadores nunca são somados.
- Motor de agregações (soma, contagem, contagem distinta, média, média ponderada, mediana, taxa) com validação
  e consolidação correta de grupos (nunca média simples de médias).
- Taxa consolidada = Σnumerador ÷ Σdenominador (detectada só quando reproduz o percentual linha a linha, ou manual);
  a média de percentuais é rotulada como DESCRITIVA.
- Séries mensais com calendário completo: meses ausentes = lacuna (NaN), nunca zero; comparação mês a mês
  (calendário) diferenciada da comparação entre meses observados; comparações inválidas viram indicadores indisponíveis.
- Indicadores temporais por medida (sem depender de uma única coluna somável).
- Dashboard inclui avaliações (escala/unidade/média/distribuição) e média ponderada opcional.
- Importação: separador validado por consistência de cabeçalho e linhas; ambiguidade pede confirmação manual.
- Relatório/CSV: unidades, regra de agregação, filtros e indicadores indisponíveis com motivo.

Novidades da v3.2 (auditoria)
- Agregação por tipo: tempo e percentuais NÃO são somados (média/mediana); unidade (R$, %, min) em InfoColuna.
- Variação entre meses consecutivos/completos e detecção de mês parcial nas pontas.
- Importação: avisos de cabeçalho duplicado/vazio, linhas descartadas (modo tolerante) e encoding não-UTF-8.
- Zeros à esquerda (CEP, matrícula) classificam a coluna como identificador.
- Limiares do diagnóstico nomeados (LIMIAR_*) e exibidos na UI.
- Widgets com keys estáveis; uploader limpo ao limpar a sessão; relatório informa filtros e escopo.

Principais melhorias da v3.0
- Correção de bugs:
  * `resumo_importacao` inexistente (removido hack com globals()).
  * Números no formato brasileiro ("1.234,56") eram convertidos para NaN.
  * Detecção de colunas por substring errada ("id" em "cidade", "dia" em "media");
    agora é feita por palavras inteiras (sem acento / camelCase).
  * Colunas contínuas únicas (ex.: valores float) eram classificadas como identificador.
  * Alerta de concentração por categoria nunca disparava (o indicador não gerava o dict esperado).
  * Variação percentual com base zero retornava 0% (agora fica indisponível).
  * Slider quebrava com bases de poucas linhas; filtro de data ignorava horas do último dia.
  * Filtro de categoria truncava valores em ordem alfabética (agora: mais frequentes).
- Performance:
  * Conversões numéricas/datas feitas UMA vez (DataFrame tipado em cache de sessão).
  * Cache invalidado por assinatura dos tipos (sem st.rerun em loop).
  * Leitura do CSV: 1 decodificação + detecção de separador (antes: até 20 parses).
- Funcionalidades:
  * Seleção de coluna de data/categoria/medida no dashboard.
  * Indicadores para todas as colunas numéricas, distribuição por categoria.
  * Diagnóstico de valores não convertidos e colunas com muitos ausentes.
  * Relatório HTML mantido em sessão (o botão de download não some).

TODO (v5): avaliar divisão em módulos (core.py, importers.py, ui_planilhas.py, ui_documentos.py, docs.py, pptx_render.py,
app.py) para permitir testes unitários e reduzir custo de import. (Fix: registrado no docstring, fora do escopo — bug #21)
"""
from __future__ import annotations

import base64
import csv
import hashlib
import html
import importlib.util
import inspect
import io
import json
import math
import os
import re
import sys
import threading
import uuid
import unicodedata
import warnings
from collections import Counter, OrderedDict
from contextlib import contextmanager
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta
from datetime import time as hora_do_dia
from functools import lru_cache
from pathlib import Path
from typing import Any, Literal, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# CONFIGURAÇÃO
# =============================================================================

def configurar_pagina() -> None:
    st.set_page_config(
        page_title="Localizador Inteligente de Indicadores",
        page_icon="🔎",
        layout="wide",
        initial_sidebar_state="expanded",
    )


NOME_SISTEMA = "LOCALIZADOR INTELIGENTE DE INDICADORES"
VERSAO = "4.0"
SUBTITULO = "Transforme dados em informações para a tomada de decisão."

SEPARADORES_CANDIDATOS = [",", ";", "\t", "|"]
# latin-1 nunca falha, por isso fica por último.
ENCODINGS_CANDIDATOS = ["utf-8-sig", "cp1252", "latin-1"]

TipoSemantico = Literal[
    "data", "numero", "categoria", "texto", "identificador", "tempo", "avaliacao"
]
TIPOS_DISPONIVEIS: list[str] = [
    "data", "numero", "categoria", "texto", "identificador", "tempo", "avaliacao",
]
TIPOS_NUMERICOS = ("numero", "tempo", "avaliacao")

PALAVRAS_DATA = ("data", "date", "dt", "dia", "mes", "mês", "ano", "periodo", "período", "timestamp")
PALAVRAS_TEMPO = ("tempo", "duracao", "duração", "minutos", "segundos", "horas", "prazo", "sla", "atraso", "delay")
PALAVRAS_AVALIACAO = ("nota", "score", "rating", "avaliacao", "avaliação", "satisfacao", "satisfação", "nps", "estrela")
PALAVRAS_ID = ("id", "cd", "codigo", "código", "cod", "matricula", "matrícula", "cpf", "cnpj", "cep", "telefone", "protocolo", "registro")
PALAVRAS_CATEGORIA = ("categoria", "setor", "grupo", "tipo", "regiao", "região", "estado", "cidade", "produto", "canal", "segmento", "status", "classe")

# Identidade visual — Santa Casa de Misericórdia de Barra Mansa
# Fonte das cores: logomarca (azul-petróleo #286770 e verde-lima #B0C458 medidos no arquivo da logo). Verde-petróleo,
# cinza-claro e amarelo suave são cores de APOIO derivadas/escolhidas para harmonizar com a logo (a apresentação
# institucional não foi fornecida): ajuste aqui, em um único lugar, se a identidade oficial tiver outros códigos.
NOME_INSTITUICAO = "Santa Casa de Misericórdia de Barra Mansa"
COR_AZUL_PETROLEO = "#286770"     # logo: texto "SantaCasa" e mãos
COR_LIMA = "#B0C458"              # logo: coração (verde-lima)
COR_VERDE_PETROLEO = "#2F7D6B"    # apoio (derivada): verde-petróleo
COR_BRANCO = "#FFFFFF"
COR_CINZA_CLARO = "#F1F4F4"       # apoio: fundos de página/cartões
COR_CINZA_BORDA = "#D5E1E1"       # apoio: bordas e grades
COR_AMARELO_SUAVE = "#FFF3C4"     # apoio: destaques e avisos suaves
COR_AMARELO_DESTAQUE = "#E6C84F"  # apoio: série de destaque nos gráficos
COR_ESCURA = "#1B3F45"            # títulos e texto (alto contraste sobre branco)
COR_TEXTO_SUAVE = "#52696D"       # texto secundário (contraste ≥ 4,5:1 sobre branco)
COR_AVISO_MESES = "#8A5A00"       # Fix: aviso de meses parciais nos gráficos, antes um literal fora da paleta (bug #10)
# Aliases mantidos por compatibilidade com versões anteriores
COR_PRIMARIA = COR_AZUL_PETROLEO
COR_SECUNDARIA = COR_LIMA
COR_SUAVE = "#EAF1F1"
COR_APOIO = "#4E8289"
PALETA = [COR_AZUL_PETROLEO, COR_LIMA, COR_VERDE_PETROLEO, COR_AMARELO_DESTAQUE, "#8FB3B8", COR_ESCURA, "#D5E1A0", "#5B8F6A"]
ARQUIVO_LOGO = "logo_santa_casa.png"   # opcional: na mesma pasta do script; embutido no cabeçalho e no relatório

# Palavras que indicam a unidade/natureza da medida (comparadas por palavra inteira)
PALAVRAS_PERCENTUAL = ("percentual", "pct", "taxa", "rate", "share", "proporcao", "proporção")
# Somente nomes INEQUÍVOCOS de valor monetário. "valor", "total", "saldo" etc. são ambíguos
# (podem ser contagens, pontos, quilos...) e NÃO atribuem R$ nem permitem soma automática.
PALAVRAS_MOEDA = ("preco", "preço", "receita", "custo", "fatura", "salario", "salário",   # "fatura" cobre faturado/faturamento
                  "montante", "vlr", "reais", "brl", "glosa", "repasse")
# Substantivos que indicam QUANTIDADE (contagem de ocorrências, aditiva). "total_atendimentos" -> quantidade.
PALAVRAS_QUANTIDADE = ("qtd", "qtde", "qte", "quantidade", "atendimento", "chamado", "unidade", "item", "itens",
                       "pedido", "ocorrencia", "caso", "visita", "lead", "conversao", "conversoes",
                       "cliente", "acesso", "clique", "ticket")
# Nomes genéricos de medida: não bastam para decidir unidade nem somabilidade.
PALAVRAS_MEDIDA_AMBIGUA = ("valor", "total", "soma", "saldo", "media", "mediana", "indice", "numero", "num", "nr")

# Naturezas possíveis de uma coluna numérica (subtipo) e suas regras
NATUREZAS_NUMERICAS = ("quantidade", "monetario", "percentual", "indefinido")
MAX_MEDIDAS_TEMPORAIS = 8         # indicadores temporais automáticos por medida (as demais: aba Dashboard)
MODOS_AGREGACAO = ("soma", "contagem", "contagem_distinta", "media", "media_ponderada", "mediana", "minimo", "maximo",
                   "taxa", "nps")

# Limiares do diagnóstico (documentados na aba Diagnóstico)
LIMIAR_AUSENTE_ALERTA = 20.0      # % de células vazias -> alerta
LIMIAR_AUSENTE_ATENCAO = 5.0      # % -> atenção
LIMIAR_COLUNA_VAZIA = 50.0        # % de ausentes em uma coluna -> alerta
LIMIAR_CONCENTRACAO = 50.0        # % da categoria principal -> atenção
LIMIAR_VARIACAO_ATENCAO = 20.0    # % de variação -> atenção
LIMIAR_VARIACAO_ALERTA = 50.0     # % -> alerta
LIMIAR_AMOSTRA_PEQUENA = 30       # nº de linhas -> informação
LIMIAR_OUTLIERS_PCT = 5.0         # % de outliers (IQR) -> atenção
LIMIAR_CONVERSAO_PCT = 5.0        # % de valores não convertidos -> atenção
LIMIAR_EXCLUSAO_TAXA_PCT = 5.0    # % de registros (com dados) excluídos por inconsistência -> taxa REPROVADA
LIMIAR_EXCLUSAO_NPS_PCT = 5.0     # % de respostas preenchidas fora da escala 0–10 -> NPS REPROVADO

# Privacidade — textos conforme o ambiente real de execução (nenhuma declaração de conformidade com a LGPD)
MENSAGEM_PRIVACIDADE_CURTA = ("🔒 Os dados são processados no equipamento/servidor onde este app Streamlit está em execução; "
                              "o app não envia dados a serviços externos. Sem validação de conformidade com a LGPD.")
MENSAGEM_PRIVACIDADE = (
    "Os dados importados são processados no equipamento ou servidor em que este aplicativo Streamlit está em execução; "
    "o aplicativo não envia dados a serviços externos nem registra o conteúdo dos arquivos em logs. Se o aplicativo estiver "
    "hospedado em servidor ou nuvem, os dados trafegam até esse ambiente. Relatórios e exportações podem conter dados "
    "da instituição: compartilhe-os apenas conforme a política institucional. Este aplicativo não foi validado quanto à "
    "LGPD; essa validação cabe à instituição.")
# Colunas tratadas como identificadoras/sensíveis (ocultas por padrão em prévias, filtros, rankings e exportações)
PALAVRAS_SENSIVEIS = ("cpf", "cnpj", "rg", "cns", "nome", "paciente", "telefone", "celular", "email", "endereco", "rua",
                      "cep", "prontuario", "nascimento", "mae", "pai", "protocolo", "matricula")
# Dado de saúde (LGPD, art. 5º, II): pode aparecer AGREGADO (contagens, rankings), nunca linha a linha em exportações.
PALAVRAS_SAUDE = ("cid", "diagnostico", "diagnóstico", "doenca", "doença", "patologia", "comorbidade", "hipotese_diagnostica")
# Política centralizada de privacidade (uma categoria por coluna):
#   identificavel -> pode identificar uma pessoa: fora de filtros, rankings, prévias e exportações (por padrão)
#   revisar       -> texto livre/alta variedade: pode conter dado pessoal; tratado como identificável até revisão
#   saude         -> dado de saúde: só agregado; fora das exportações linha a linha
#   liberada      -> sem indício de dado pessoal
CATEGORIAS_PRIVACIDADE = ("identificavel", "revisar", "saude", "liberada")
DIAS_SEMANA = ("Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo")
MESES_ABREV = ("jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez")
LIMIAR_SUPRESSAO = 5   # contagens abaixo deste valor aparecem como "<5" em relatórios e apresentações exportados

# Padrões de dados pessoais e de saúde. Definidos aqui, antes de exemplo_seguro/_classificar_tipo/politica_privacidade,
# que os usam (A1: sem referências adiantadas).
_RE_CPF = r"^\d{3}\.?\d{3}\.?\d{3}-?\d{2}$"
_RE_CNPJ = r"^\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}$"
_RE_TELEFONE = r"^(\+?55\s?)?(\(?\d{2}\)?\s?)?9?\d{4}[-\s]?\d{4}$"
_RE_EMAIL = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
_RE_CNS = r"^\d{15}$"
_RE_NOME_PESSOA = r"^[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+(?:d[aeo]s?|[A-ZÀ-Ý][a-zà-ÿ]+))+$|^[A-ZÀ-Ý]{2,}(?:\s+[A-ZÀ-Ý]{2,}){1,5}$"
_RE_CID = r"^[A-Z]\d{2}(\.?\d)?$"   # CID-10: letra + 2 dígitos + subcategoria opcional


# =============================================================================
# UTILITÁRIOS
# =============================================================================


def formatar_numero(valor: Any, casas: int = 2) -> str:
    if valor is None:
        return "-"
    if isinstance(valor, (dict, list, tuple, set)):
        return str(valor)
    if isinstance(valor, float) and pd.isna(valor):
        return "-"
    try:
        if isinstance(valor, (int, np.integer)) and not isinstance(valor, bool):
            return f"{valor:,}".replace(",", ".")
        return f"{float(valor):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(valor)


def formatar_percentual(valor: float, casas: int = 2) -> str:
    if valor is None or pd.isna(valor):
        return "-"
    try:
        return f"{float(valor):.{casas}f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "-"


def resultado_exportavel(resultado: Any) -> str:
    """Igual a `formatar_resultado`, mas contagens inteiras de 1 a 4 (inclusive dentro de listas como 'top: ...')
    aparecem como "<5". Usado em tudo que sai do aplicativo."""
    def proteger(v: Any) -> Any:
        if isinstance(v, dict):
            return {k: proteger(x) for k, x in v.items()}
        if isinstance(v, (int, np.integer)) and not isinstance(v, bool) and 0 < int(v) < limiar_ativo():
            return suprimir_contagem(int(v))
        return v
    return formatar_resultado(proteger(resultado))


def formatar_resultado(resultado: Any, aninhado: bool = False) -> str:
    """Formata resultados de indicadores (números, dicts) para texto legível."""
    if isinstance(resultado, dict):
        texto = ", ".join(f"{k}: {formatar_resultado(v, True)}" for k, v in resultado.items())
        return f"({texto})" if aninhado else texto
    if isinstance(resultado, (int, float, np.integer, np.floating)) and not isinstance(resultado, bool):
        return formatar_numero(resultado)
    return str(resultado)


def safe_divide(numerador: float, denominador: float) -> Optional[float]:
    """Divide com segurança; retorna None quando a divisão não é definida."""
    try:
        if denominador is None or pd.isna(denominador) or denominador == 0:
            return None
        return float(numerador) / float(denominador)
    except (TypeError, ValueError, ZeroDivisionError):
        return None


def _normalizar(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return sem_acento.lower()


def _tokens(nome: str) -> list[str]:
    separado = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", str(nome))  # camelCase -> camel Case
    return [t for t in re.split(r"[\W_]+", _normalizar(separado)) if t]


def _nome_contem(nome: str, palavras: tuple[str, ...]) -> bool:
    """Compara por palavra inteira (ou prefixo, para palavras com 4+ letras).

    Evita falsos positivos como 'id' em 'cidade' ou 'dia' em 'media'.
    """
    tokens = _tokens(nome)
    alvo = {_normalizar(p) for p in palavras}
    return any(t == p or (len(p) >= 4 and t.startswith(p)) for t in tokens for p in alvo)


def limpar_e_converter_numerico(serie: pd.Series) -> pd.Series:
    """Converte texto para número, aceitando formato BR e US.

    Exemplos: "R$ 1.234,56" -> 1234.56 | "1,234.56" -> 1234.56 | "(50)" -> -50 | "3,5" -> 3.5
    Regra para ponto isolado: "1.234" / "1.234.567" são tratados como milhar (padrão BR);
    "0.5" e "12.5" são tratados como decimal.
    """
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    s = serie.astype(str).str.strip()
    preenchido = serie.notna()
    if preenchido.any() and s[preenchido].str.fullmatch(r"[+-]?\d+").all():   # BLOCO 7: só inteiros simples → atalho
        return pd.to_numeric(s.where(preenchido), errors="coerce")
    negativo = s.str.match(r"^\(.*\)$")
    s = s.str.replace(r"[R$€£%\s()]", "", regex=True)

    tem_virgula = s.str.contains(",", regex=False)
    tem_ponto = s.str.contains(".", regex=False)
    virgula_decimal = tem_virgula & (~tem_ponto | (s.str.rfind(",") > s.str.rfind(".")))
    ponto_decimal_us = tem_virgula & tem_ponto & ~virgula_decimal
    ponto_milhar = ~tem_virgula & s.str.match(r"^[+-]?[1-9]\d{0,2}(\.\d{3})+$")

    s = s.mask(virgula_decimal, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    s = s.mask(ponto_decimal_us, s.str.replace(",", "", regex=False))
    s = s.mask(ponto_milhar, s.str.replace(".", "", regex=False))

    numeros = pd.to_numeric(s, errors="coerce")
    return numeros.where(~negativo, -numeros.abs())


def _para_datetime(texto: pd.Series, dayfirst: bool) -> pd.Series:
    """pd.to_datetime tolerante (formatos mistos), sempre sem fuso horário."""
    def sem_fuso(s: pd.Series) -> pd.Series:
        return s.dt.tz_convert(None) if getattr(s.dt, "tz", None) is not None else s

    try:
        res = sem_fuso(pd.to_datetime(texto, errors="coerce", dayfirst=dayfirst, format="mixed"))
        # Fix: resgata só as linhas preenchidas que o format="mixed" deixou em NaT (bugs de dayfirst no pandas
        # 2.0.0–2.0.3); nunca sobrescreve uma data já interpretada (bug #7)
        faltam = res.isna() & texto.notna()
        if faltam.any():
            try:
                alt = sem_fuso(pd.to_datetime(texto[faltam], errors="coerce", dayfirst=dayfirst))
                res = res.copy()
                res.loc[faltam] = alt
            except (ValueError, TypeError):
                pass
    except (ValueError, TypeError):
        try:
            res = sem_fuso(pd.to_datetime(texto, errors="coerce", dayfirst=dayfirst))
        except (ValueError, TypeError):
            # Fix: fusos mistos que o regex não removeu: lê valor a valor (hora do arquivo, sem fuso); o que não der
            # vira NaT, auditado em auditar_conversao, em vez de derrubar a coluna inteira (bug #16)
            def um(x: Any) -> Any:
                v = pd.to_datetime(x, errors="coerce", dayfirst=dayfirst) if isinstance(x, str) else pd.NaT
                return v.tz_localize(None) if getattr(v, "tzinfo", None) is not None else v
            res = pd.Series([um(x) for x in texto], index=texto.index, dtype="datetime64[ns]")
    return res


def converter_datas(serie: pd.Series, dayfirst: bool = True) -> pd.Series:
    """Converte para datetime.

    Datas ISO (AAAA-MM-DD, AAAA/MM/DD, com ou sem hora) são lidas SEMPRE como ano-mês-dia. Antes, `dayfirst=True` era
    aplicado também a elas e "2024-02-03" virava 3 de março (dia e mês trocados quando ambos ≤ 12). Os demais formatos
    (DD/MM/AAAA...) são lidos com o dia primeiro. Valores inválidos viram NaT (e são auditados em `auditar_conversao`)."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    texto = serie.astype("string").str.strip()
    # Fuso horário (só aparece depois de uma hora, ex.: 10:00:00Z ou 10:00-03:00): ignorado; vale a hora do arquivo.
    # Fix: também remove "GMT-3", "UTC+10", "UTC"/"GMT" sozinhos; antes, fusos GMT mistos chegavam ao pandas e a
    # conversão da coluna inteira falhava com "Mixed timezones" (bug #16)
    texto = texto.str.replace(
        r"(\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)\s*(?:Z|[+-]\d{2}:?\d{2}|(?:GMT|UTC)(?:\s*[+-]\d{1,2}(?::?\d{2})?)?)$",
        r"\1", regex=True)
    iso = texto.str.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:[ T].*)?$").fillna(False).astype(bool)
    resto = (~iso) & texto.notna()
    partes = []
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        if iso.any():
            partes.append(_para_datetime(texto[iso], dayfirst=False))
        if resto.any():
            partes.append(_para_datetime(texto[resto], dayfirst=dayfirst))
    if not partes:
        return pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    return pd.concat(partes).reindex(serie.index)


def colunas_sensiveis(tipos: dict[str, "InfoColuna"]) -> list[str]:
    """Colunas que podem identificar uma pessoa ('identificavel') ou ainda não revisadas ('revisar'). Ficam fora de
    filtros, rankings, prévias e exportações (por padrão). Ver `politica_privacidade`."""
    return [c for c, i in tipos.items() if i.tipo == "identificador" or _nome_contem(c, PALAVRAS_SENSIVEIS)
            or getattr(i, "privacidade", "liberada") in ("identificavel", "revisar")]


def colunas_saude(tipos: dict[str, "InfoColuna"]) -> list[str]:
    """Colunas com dado de saúde (ex.: CID): permitidas só de forma agregada; fora das exportações linha a linha."""
    return [c for c, i in tipos.items() if getattr(i, "privacidade", "liberada") == "saude"]


def colunas_nao_exportaveis(tipos: dict[str, "InfoColuna"]) -> list[str]:
    """Colunas retiradas por padrão de qualquer exportação linha a linha (identificáveis, a revisar e de saúde)."""
    return list(dict.fromkeys(colunas_sensiveis(tipos) + colunas_saude(tipos)))


def exemplo_seguro(valor: Any, coluna: str = "") -> str:
    """Exemplo de valor que pode aparecer em mensagens: códigos curtos (ex.: '099') sim; qualquer coisa que pareça
    documento, telefone, e-mail ou nome — ou que venha de coluna com nome de dado pessoal — vira '(valor oculto)'."""
    txt = str(valor).strip()
    if coluna and _nome_contem(coluna, PALAVRAS_SENSIVEIS):
        return "(valor oculto)"
    if (len(txt) > 12 or len(re.sub(r"\D", "", txt)) >= 8 or "@" in txt or re.match(_RE_NOME_PESSOA, txt)):
        return "(valor oculto)"
    return txt


# v3.9-B — PROTEÇÃO DE GRUPOS PEQUENOS CONFIGURÁVEL ---------------------------------------------------------------
# Por padrão o aplicativo protege (limiar 5). A equipe pode escolher: "automatico" (só quando a base tem colunas
# sensíveis), "sempre" ou "nunca" (a planilha já vem tratada; exige confirmação quando há colunas sensíveis).
# O limiar vigente fica numa variável da thread (várias sessões não interferem) e vale só dentro de `politica_protecao`.
MODOS_PROTECAO = ("automatico", "sempre", "nunca")
_POLITICA = threading.local()


def limiar_ativo() -> int:
    """Limiar de supressão em vigor: 5 por padrão; 0 quando a proteção foi desligada pela decisão da equipe."""
    return int(getattr(_POLITICA, "limiar", LIMIAR_SUPRESSAO))


@contextmanager
def politica_protecao(limiar: int):
    anterior = getattr(_POLITICA, "limiar", None)
    _POLITICA.limiar = int(limiar)
    try:
        yield
    finally:
        if anterior is None:
            if hasattr(_POLITICA, "limiar"):
                del _POLITICA.limiar
        else:
            _POLITICA.limiar = anterior


@dataclass
class DecisaoProtecao:
    modo: str
    limiar: int
    colunas: list[str]            # colunas sensíveis detectadas (identificáveis, a revisar ou de saúde)
    confirmou: bool = False       # confirmação de que a planilha já vem tratada (só vale para modo "nunca")
    pendente: bool = False        # "nunca" escolhido com colunas sensíveis e sem confirmação -> segue protegendo

    @property
    def aplicada(self) -> bool:
        return self.limiar > 0

    def registro(self) -> str:
        """Frase que fica no rodapé do relatório e da apresentação: quem ler depois sabe o que foi feito."""
        if self.aplicada:
            motivo = ("a base tem colunas que podem identificar pessoas ou dados de saúde" if self.colunas
                      else "escolha do usuário")
            return (f'Proteção de grupos pequenos: aplicada ({motivo}); contagens de 1 a {self.limiar - 1} aparecem '
                    f'como "<{self.limiar}".')
        if self.modo == "nunca":
            extra = (f" Colunas sensíveis detectadas: {', '.join(self.colunas)}." if self.colunas else "")
            return ("Proteção de grupos pequenos: desligada por decisão do usuário, que informou que a planilha já vem "
                    "tratada. O aplicativo não verifica isso." + extra)
        return ("Proteção de grupos pequenos: não aplicada, pois nenhuma coluna que identifique pessoas ou dado de "
                "saúde foi detectada na base.")


def decidir_protecao(modo: str, tipos: dict[str, "InfoColuna"], confirmou: bool = False) -> DecisaoProtecao:
    """Traduz a escolha da equipe em limiar. Na dúvida, protege: modo desconhecido ou "nunca" sem confirmação (havendo
    colunas sensíveis) mantêm o limiar padrão."""
    colunas = list(dict.fromkeys(colunas_sensiveis(tipos) + colunas_saude(tipos)))
    if modo == "nunca":
        if colunas and not confirmou:
            return DecisaoProtecao("nunca", LIMIAR_SUPRESSAO, colunas, False, True)
        return DecisaoProtecao("nunca", 0, colunas, bool(confirmou))
    if modo == "automatico":
        return DecisaoProtecao("automatico", LIMIAR_SUPRESSAO if colunas else 0, colunas)
    return DecisaoProtecao("sempre", LIMIAR_SUPRESSAO, colunas)


def suprimir_contagem(n: Any, limite: Optional[int] = None) -> str:
    """Contagem para saídas compartilháveis: valores de 1 a (limite − 1) aparecem como '<limite' (ex.: '<5')."""
    lim = limiar_ativo() if limite is None else limite
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "-"
    n = int(n)
    return f"<{lim}" if 0 < n < lim else formatar_numero(n)


def estilizar_figura(fig: go.Figure) -> go.Figure:
    """Aplica a identidade visual a um gráfico Plotly (cores, fontes, grades). Não altera os dados."""
    fig.update_layout(
        template="plotly_white", colorway=PALETA, paper_bgcolor=COR_BRANCO, plot_bgcolor=COR_BRANCO,
        font=dict(family="Segoe UI, Roboto, Arial, sans-serif", color=COR_ESCURA, size=13),
        title=dict(font=dict(color=COR_AZUL_PETROLEO, size=15)),
        hoverlabel=dict(bgcolor=COR_BRANCO, bordercolor=COR_AZUL_PETROLEO, font=dict(color=COR_ESCURA)),
        legend=dict(font=dict(color=COR_ESCURA)))
    fig.update_xaxes(gridcolor=COR_CINZA_BORDA, linecolor=COR_CINZA_BORDA, zerolinecolor=COR_CINZA_BORDA)
    fig.update_yaxes(gridcolor=COR_CINZA_BORDA, linecolor=COR_CINZA_BORDA, zerolinecolor=COR_CINZA_BORDA)
    return fig


# =============================================================================
# IMPORTAÇÃO
# =============================================================================


@dataclass
class ResultadoImportacao:
    sucesso: bool
    dataframe: Optional[pd.DataFrame] = None
    separador: Optional[str] = None
    encoding: Optional[str] = None
    mensagem: str = ""
    erro: Optional[str] = None
    avisos: list[str] = field(default_factory=list)
    ambiguo: bool = False                      # True: o separador exige confirmação manual
    candidatos: list[dict] = field(default_factory=list)
    requer_confirmacao: bool = False           # True: há registros inconsistentes; o usuário escolhe o tratamento
    aba: str = ""                               # Excel: nome da aba que originou esta tabela
    origem: str = "csv"                         # csv | excel
    integridade: Optional[dict] = None         # contagens de registros (válidos/curtos/longos/aproveitados...)


def _decodificar(conteudo: bytes) -> tuple[Optional[str], Optional[str]]:
    for enc in ENCODINGS_CANDIDATOS:
        try:
            return conteudo.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return None, None


def _detectar_separador(texto: str) -> str:
    amostra = "\n".join(texto.splitlines()[:50])
    try:
        return csv.Sniffer().sniff(amostra, delimiters="".join(SEPARADORES_CANDIDATOS)).delimiter
    except csv.Error:
        contagens = {s: amostra.count(s) for s in SEPARADORES_CANDIDATOS}
        if not contagens or max(contagens.values()) == 0:
            return ","  # fallback: arquivo de coluna única ou sem delimitadores
        return max(contagens, key=contagens.get)  # type: ignore[arg-type]


TRATAMENTOS_CSV = ("descartar", "completar", "truncar")
LIMITE_LINHAS_INCONSISTENTES = 50_000   # teto de linhas inconsistentes guardadas para download (sem conteúdo da linha)
LIMITE_EVENTOS_REGISTRO = 300           # teto do registro desta análise
DESCRICAO_TRATAMENTOS = {
    "descartar": "Descartar as linhas inconsistentes (curtas e longas)",
    "completar": "Completar as linhas curtas com valores vazios (as linhas longas continuam descartadas)",
    "truncar": "Truncar as linhas longas (descarta apenas os campos excedentes)",
}


@dataclass
class IntegridadeCSV:
    """Contagem de registros do CSV por nº de campos, feita sobre TODO o arquivo (não só uma amostra)."""
    cabecalho: Optional[list[str]]
    n_campos: int
    lidos: int                # registros de dados com algum conteúdo
    validos: int              # nº de campos == cabeçalho
    curtos: int               # menos campos que o cabeçalho (o pandas preencheria com vazio, em silêncio)
    longos: int               # mais campos que o cabeçalho (excedente seria perdido ou deslocaria colunas)
    vazios: int               # registros só com campos vazios (ignorados)
    linhas_em_branco: int
    exemplos: list[dict]      # até 10 registros inconsistentes: {"linha", "campos", "tipo"} — sem o conteúdo da linha

    @property
    def inconsistentes(self) -> int:
        return self.curtos + self.longos


def analisar_integridade(texto: str, sep: str) -> tuple[IntegridadeCSV, list[tuple[list[str], bool]]]:
    """Lê todos os registros com `csv.reader` (aspas e quebras de linha entre aspas respeitadas) e confere o nº
    de campos de CADA um, mesmo quando o pandas leria o arquivo sem erro. Retorna (integridade, guardados), onde
    `guardados` = [(campos, é_curto)] dos registros válidos, curtos e longos, na ordem do arquivo. Um registro longo
    tem é_curto=False e len(campos) > nº de campos do cabeçalho; só é aproveitado (truncado) com o tratamento "truncar".
    Levanta ValueError se as aspas forem malformadas (ex.: aspas não fechadas, que engoliriam o restante do arquivo
    dentro de um único campo)."""
    try:
        csv.field_size_limit(10_000_000)
    except OverflowError:
        pass
    cabecalho: Optional[list[str]] = None
    n = 0
    guardados: list[tuple[list[str], bool]] = []
    exemplos: list[dict] = []
    validos = curtos = longos = vazios = brancas = 0
    try:
        leitor = csv.reader(io.StringIO(texto), delimiter=sep, strict=True)
        for reg in leitor:
            if not reg or not any(c.strip() for c in reg):
                if cabecalho is None or not reg:
                    brancas += 1
                else:
                    vazios += 1
                continue
            if cabecalho is None:
                cabecalho = [c.strip() for c in reg]
                n = len(cabecalho)
                continue
            k = len(reg)
            if k == n:
                validos += 1
                guardados.append((reg, False))
                continue
            tipo = "curta" if k < n else "longa"
            if k < n:
                curtos += 1
                guardados.append((reg, True))
            else:
                longos += 1
                guardados.append((reg, False))   # C4: guardado para o tratamento "truncar"
            if len(exemplos) < LIMITE_LINHAS_INCONSISTENTES:
                # Guarda nº da linha, quantidade de campos e motivo; nunca o conteúdo da linha.
                exemplos.append({
                    "linha": leitor.line_num, "campos": k, "esperados": n, "tipo": tipo,
                    "motivo": ("menos campos que o cabeçalho (faltam valores)" if tipo == "curta"
                               else "mais campos que o cabeçalho (valores excedentes)"),
                })
    except csv.Error as e:
        raise ValueError(f"CSV malformado próximo à linha {leitor.line_num}: {e}. Verifique aspas não fechadas ou texto "
                         "após as aspas de fechamento.") from e
    integ = IntegridadeCSV(cabecalho, n, validos + curtos + longos, validos, curtos, longos, vazios, brancas, exemplos)
    return integ, guardados


def csv_trecho(reg: list[str], limite: int = 80) -> str:
    texto = " | ".join(reg)
    return texto if len(texto) <= limite else texto[:limite - 1] + "…"


def _nomes_unicos(cabecalho: list[str]) -> list[str]:
    """Mesma convenção do pandas: cabeçalho vazio -> 'Unnamed: i'; duplicado -> 'nome.1', 'nome.2'..."""
    nomes, vistos = [], {}
    for i, c in enumerate(cabecalho):
        nome = c.strip() or f"Unnamed: {i}"
        if nome in vistos:
            vistos[nome] += 1
            nome = f"{nome}.{vistos[nome]}"
        else:
            vistos[nome] = 0
        nomes.append(nome)
    return nomes


def _dict_integridade(i: IntegridadeCSV, tratamento: Optional[str], aproveitados: Optional[int] = None) -> dict:
    completados = i.curtos if tratamento == "completar" else 0
    truncados = i.longos if tratamento == "truncar" else 0   # C4: nº de LINHAS longas truncadas
    return {
        "campos_cabecalho": i.n_campos, "registros_lidos": i.lidos, "validos": i.validos, "curtos": i.curtos,
        "longos": i.longos, "vazios_ignorados": i.vazios, "linhas_em_branco": i.linhas_em_branco,
        "tratamento": tratamento or "pendente", "aproveitados": aproveitados,
        "completados": completados if aproveitados is not None else None,
        "truncados": truncados if aproveitados is not None else None,
        "descartados": (i.lidos - aproveitados) if aproveitados is not None else None,
        "exemplos": i.exemplos,
    }


SEP_UNICA = "\x1f"  # separador "inexistente": força a leitura como uma única coluna
NOMES_SEPARADORES = {",": "vírgula (,)", ";": "ponto e vírgula (;)", "\t": "tabulação", "|": "barra vertical (|)",
                     SEP_UNICA: "uma única coluna"}


@dataclass
class AvaliacaoSeparador:
    """Resultado da validação de um separador candidato sobre uma amostra do arquivo."""
    sep: str
    colunas: int
    consistencia: float          # fração das linhas de amostra com o mesmo nº de campos do cabeçalho
    cabecalho_ok: bool
    cabecalho: list[str]
    amostra: list[list[str]]

    @property
    def plausivel(self) -> bool:
        return self.colunas >= 2 and self.consistencia >= 0.9 and self.cabecalho_ok

    @property
    def toleravel(self) -> bool:
        return self.colunas >= 2 and self.consistencia >= 0.6 and self.cabecalho_ok


def _amostrar_registros(texto: str, sep: str, limite: int = 200) -> list[list[str]]:
    try:
        saida: list[list[str]] = []
        for linha in csv.reader(io.StringIO(texto), delimiter=sep):
            if any(c.strip() for c in linha):
                saida.append(linha)
            if len(saida) > limite:
                break
        return saida
    except csv.Error:
        return []


def _cabecalho_plausivel(cabecalho: list[str]) -> bool:
    """Cabeçalho plausível: 2+ campos, no máximo 50% vazios, sem campos enormes e não todo numérico."""
    if len(cabecalho) < 2:
        return False
    preenchidos = [c.strip() for c in cabecalho if c.strip()]
    if not preenchidos or len(preenchidos) < len(cabecalho) * 0.5:
        return False
    if any(len(c) > 80 for c in preenchidos):
        return False
    if all(pd.notna(limpar_e_converter_numerico(pd.Series([c])).iloc[0]) for c in preenchidos):
        return False  # cabeçalho inteiro numérico: provavelmente é uma linha de dados
    return True


def avaliar_separador(texto: str, sep: str) -> AvaliacaoSeparador:
    """Valida o separador pela consistência do cabeçalho e das linhas de amostra.

    Produzir 2+ colunas NÃO basta: as linhas de dados precisam ter o mesmo nº de campos do cabeçalho.
    """
    registros = _amostrar_registros(texto, sep)
    if not registros:
        return AvaliacaoSeparador(sep, 0, 0.0, False, [], [])
    cabecalho = [c.strip() for c in registros[0]]
    corpo = registros[1:]
    n = len(cabecalho)
    consistencia = 1.0 if not corpo else sum(1 for l in corpo if len(l) == n) / len(corpo)
    return AvaliacaoSeparador(sep, n, consistencia, _cabecalho_plausivel(cabecalho), cabecalho, corpo[:3])


def _candidato_para_ui(a: AvaliacaoSeparador) -> dict:
    nomes = [c or f"coluna_{i + 1}" for i, c in enumerate(a.cabecalho)] or ["coluna_1"]
    linhas = [(l + [""] * len(nomes))[:len(nomes)] for l in a.amostra]
    vistos: dict[str, int] = {}
    unicos = []
    for nome in nomes:  # nomes únicos só para a pré-visualização
        vistos[nome] = vistos.get(nome, -1) + 1
        unicos.append(nome if vistos[nome] == 0 else f"{nome}.{vistos[nome]}")
    return {"sep": a.sep, "colunas": a.colunas, "consistencia": round(a.consistencia, 3),
            "cabecalho": a.cabecalho, "amostra": pd.DataFrame(linhas, columns=unicos)}


def _ambiguo(avals: list[AvaliacaoSeparador], mensagem: str) -> ResultadoImportacao:
    avals = sorted(avals, key=lambda a: (-a.consistencia, -a.colunas))
    candidatos = [_candidato_para_ui(a) for a in avals]
    return ResultadoImportacao(False, mensagem=mensagem, erro="separador_ambiguo", ambiguo=True, candidatos=candidatos)


def _strip_textos(s: pd.Series) -> pd.Series:
    """Remove espaços nas pontas dos valores de TEXTO; o resto da coluna fica como estava (bug #9)."""
    if not (pd.api.types.is_string_dtype(s) or s.dtype == object):
        return s
    limpo = s.str.strip()
    return limpo.where(limpo.notna() | s.isna(), s)


def _limpar_tabela(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Limpeza comum a CSV e Excel: nomes e valores sem espaços nas pontas, vazios viram NaN, colunas sem cabeçalho e
    sem dados são removidas e as demais sem cabeçalho viram 'coluna_N'. Retorna (tabela, nº de colunas renomeadas)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # Fix: strip só em texto; colunas numéricas/datas ficam intactas e, numa coluna mista (object), valores que não são
    # texto são preservados — .str.strip() os transformaria em NaN em silêncio (bug #9)
    df = df.apply(_strip_textos)
    df = df.mask(df == "")
    vazias = [c for c in df.columns if c.startswith("Unnamed") and df[c].isna().all()]
    df = df.drop(columns=vazias).reset_index(drop=True)
    sem_nome = [c for c in df.columns if str(c).startswith("Unnamed")]
    df.columns = [f"coluna_{i + 1}" if str(c).startswith("Unnamed") else c for i, c in enumerate(df.columns)]
    return df, len(sem_nome)


def ler_csv(conteudo: bytes, nome_arquivo: str = "arquivo.csv", separador: Optional[str] = None,
            tratamento: Optional[str] = None) -> ResultadoImportacao:
    """Importa o CSV.

    `separador`: escolha manual do usuário (quando a detecção for ambígua).
    `tratamento`: como tratar registros com nº de campos diferente do cabeçalho — "descartar" (curtas e longas),
    "completar" (curtas completadas com vazio; longas descartadas) ou "truncar" (longas perdem só os campos
    excedentes; curtas descartadas, como em "descartar"). Sem `tratamento`, se houver registros
    inconsistentes a leitura NÃO prossegue: retorna `requer_confirmacao=True` com as contagens (nada é
    descartado nem preenchido em silêncio).
    """
    if not conteudo or len(conteudo.strip()) == 0:
        return ResultadoImportacao(False, mensagem="O arquivo está vazio.", erro="arquivo_vazio")
    if not nome_arquivo.lower().endswith(".csv"):
        return ResultadoImportacao(False, mensagem="Formato não suportado. Envie um .csv.", erro="formato_invalido")
    if tratamento is not None and tratamento not in TRATAMENTOS_CSV:
        return ResultadoImportacao(False, mensagem=f"Não reconheci a opção '{tratamento}'. Escolha descartar, completar ou truncar as linhas.", erro="tratamento_invalido")

    texto, enc = _decodificar(conteudo)
    if texto is None:
        return ResultadoImportacao(False, mensagem="Não foi possível decodificar o arquivo.", erro="encoding")

    avisos: list[str] = []
    if separador is not None:
        sep = separador
    else:
        avals = [avaliar_separador(texto, s) for s in SEPARADORES_CANDIDATOS]
        multi = [a for a in avals if a.colunas >= 2]
        if not multi:  # nenhum separador aparece no cabeçalho: arquivo de uma única coluna
            sep = SEP_UNICA
            avisos.append("Nenhum separador foi encontrado no cabeçalho; o arquivo foi lido como uma única coluna.")
        else:
            plausiveis = [a for a in multi if a.plausivel] or [a for a in multi if a.toleravel]
            if len(plausiveis) == 1:
                sep = plausiveis[0].sep
            elif len(plausiveis) >= 2:
                nomes = ", ".join(NOMES_SEPARADORES[a.sep] for a in plausiveis)
                return _ambiguo(plausiveis, f"Mais de uma leitura é plausível ({nomes}). Confirme o separador.")
            else:
                cands = [a for a in multi if a.consistencia > 0]
                if not cands:
                    return ResultadoImportacao(False, mensagem="Não foi possível identificar o separador do arquivo.", erro="parse_falhou")
                if len(cands) > 1:
                    return _ambiguo(cands, "Nenhum separador gerou uma leitura consistente (linhas com nº de campos diferente "
                                           "do cabeçalho). Escolha o separador correto.")
                sep = cands[0].sep   # única interpretação possível: a conferência de integridade abaixo exige confirmação

    # Integridade: nº de campos de CADA registro do arquivo inteiro (independe de o pandas gerar exceção).
    try:
        integ, guardados = analisar_integridade(texto, sep)
    except ValueError as e:
        return ResultadoImportacao(False, mensagem=str(e), erro="csv_malformado")
    if integ.cabecalho is None or integ.lidos == 0:
        return ResultadoImportacao(False, mensagem="O arquivo não contém dados válidos.", erro="sem_dados")

    if integ.inconsistentes and tratamento is None:
        return ResultadoImportacao(
            False, separador=sep, encoding=enc, erro="integridade_csv", requer_confirmacao=True,
            integridade=_dict_integridade(integ, None),
            mensagem=(f"{integ.inconsistentes} de {integ.lidos} registros têm nº de campos diferente do cabeçalho "
                      f"({integ.n_campos} campos): {integ.curtos} curtos e {integ.longos} longos. "
                      "Nada foi descartado nem preenchido. Escolha o tratamento."))

    n = integ.n_campos
    linhas: list[list[str]] = []
    for campos, curto in guardados:
        if len(campos) > n:                      # longa: só entra (truncada) com "truncar"
            if tratamento == "truncar":
                linhas.append(campos[:n])
            continue
        if curto and tratamento != "completar":
            continue
        linhas.append(campos + [""] * (n - len(campos)) if curto else campos)
    df = pd.DataFrame(linhas, columns=_nomes_unicos(integ.cabecalho), dtype=str)
    duplicados = sorted({c for c in integ.cabecalho if c and integ.cabecalho.count(c) > 1}) if sep != SEP_UNICA else []

    df, n_sem_nome = _limpar_tabela(df)

    if df.shape[1] == 0 or df.empty:
        return ResultadoImportacao(False, mensagem="O arquivo não contém dados válidos.", erro="sem_dados")

    integridade = _dict_integridade(integ, tratamento or "nenhum", aproveitados=len(df))
    if integ.inconsistentes:
        exemplos = "; ".join(f"linha {e['linha']} ({e['campos']} campos)" for e in integ.exemplos[:5])
        avisos.append(
            f"Integridade do CSV — {integ.lidos} registros lidos: {integ.validos} válidos, {integ.curtos} curtos, "
            f"{integ.longos} longos. Tratamento '{tratamento}': {len(df)} aproveitados, "
            f"{integridade['completados']} completados com valores vazios, "
            + (f"{integridade['truncados']} truncados (campos excedentes descartados), " if tratamento == "truncar" else "")
            + f"{integridade['descartados']} descartados. Exemplos: {exemplos}.")
    if integ.vazios:
        avisos.append(f"{integ.vazios} registro(s) totalmente vazios foram ignorados.")
    if duplicados:
        lista = ", ".join(duplicados[:5]) + ("…" if len(duplicados) > 5 else "")
        avisos.append(f"Cabeçalhos duplicados no arquivo ({lista}). Foram renomeados (ex.: 'col.1'); confira o significado de cada coluna.")
    if n_sem_nome:
        avisos.append(f"{n_sem_nome} coluna(s) tinham cabeçalho vazio e dados; foram nomeadas 'coluna_N'.")
    if separador is not None and sep != SEP_UNICA:
        a = avaliar_separador(texto, sep)
        outros_geram_mais = any(avaliar_separador(texto, o).colunas >= 2 for o in SEPARADORES_CANDIDATOS if o != sep)
        if a.colunas == 1 and outros_geram_mais:
            avisos.append(f"O separador escolhido ({NOMES_SEPARADORES.get(sep, sep)}) resultou em uma única coluna; "
                          "outro separador pode ser mais adequado.")
    if enc != "utf-8-sig":
        avisos.append(f"Arquivo não está em UTF-8 (lido como {enc}). Confira se os acentos aparecem corretamente.")

    rotulo_sep = "(única coluna)" if sep == SEP_UNICA else f"'{sep}'"
    mensagem = f"Arquivo importado. Separador: {rotulo_sep} | Encoding: {enc}."
    if avisos:
        mensagem += " Avisos: " + " ".join(avisos)
    return ResultadoImportacao(True, df, sep, enc, mensagem=mensagem, avisos=avisos, integridade=integridade)


# =============================================================================
# IMPORTAÇÃO DE EXCEL (.xlsx / .xlsm): cada aba com dados vira uma tabela, com a mesma limpeza do CSV
# =============================================================================
EXTENSOES_EXCEL = (".xlsx", ".xlsm")
CHAVE_MANTER_TOTAIS = "excel_manter_totais"   # chave do widget "Manter linhas de total/subtotal do Excel" (C3)
TRECHO_AVISO_TOTAIS = "linha(s) de total ou subtotal foram"   # presente nos avisos de totais retirados ou mantidos


def remover_totais_excel() -> bool:
    """Escolha atual do usuário para linhas de total/subtotal do Excel. Padrão: retirar."""
    return not st.session_state.get(CHAVE_MANTER_TOTAIS, False)
ERROS_EXCEL = {"#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#VALOR!", "#NOME?", "#REF!"}
_RE_TOTAL = re.compile(r"^\s*(sub\s*)?total(\s+geral)?\b", re.I)


def excel_disponivel() -> bool:
    """A leitura de Excel usa a biblioteca `openpyxl` (pip install openpyxl)."""
    return importlib.util.find_spec("openpyxl") is not None


def _vazia(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def _celula_para_texto(v: Any) -> str:
    """Valor da célula -> texto que o restante do app lê SEM ambiguidade: datas em ISO (AAAA-MM-DD [hh:mm:ss]) e números
    sem separador de milhar, com vírgula decimal. Textos (inclusive datas digitadas como texto) ficam como estão."""
    if _vazia(v):
        return ""
    if isinstance(v, bool):
        return "sim" if v else "não"
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, hora_do_dia):
        return v.strftime("%H:%M:%S")
    if isinstance(v, timedelta):
        seg = int(v.total_seconds())
        return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        if float(v).is_integer():
            return str(int(v))
        return format(float(v), ".10g").replace(".", ",")
    return str(v).strip()


def _detectar_cabecalho(linhas: list[list], limite: int = 30) -> int:
    """Linha do cabeçalho: a primeira (entre as 30 primeiras) com ≥ 60% da largura útil preenchida e ≥ 80% de textos.
    Planilhas costumam ter título, data de emissão ou filtros acima do cabeçalho."""
    amostra = linhas[:limite]
    largura = max((sum(not _vazia(v) for v in l) for l in amostra), default=0)
    for i, l in enumerate(amostra):
        pre = [v for v in l if not _vazia(v)]
        if len(pre) >= max(2, int(0.6 * largura)) and sum(isinstance(v, str) for v in pre) / len(pre) >= 0.8:
            return i
    return 0


def _linha_total(l: list) -> bool:
    return any(isinstance(v, str) and _RE_TOTAL.match(v) for v in l[:3] if not _vazia(v))


def _ler_aba(ws: Any, remover_totais: bool = True) -> ResultadoImportacao:
    """Lê uma aba do Excel como tabela.

    Linhas de total/subtotal (primeira célula preenchida começando por "Total", "Subtotal" ou "Total geral"):
    - `remover_totais=True` (padrão): são retiradas, para não contar os valores duas vezes, e um aviso informa quantas.
    - `remover_totais=False`: são mantidas por escolha do usuário, e um aviso alerta que os valores podem ser contados
      duas vezes. A escolha vem da caixa "Manter linhas de total/subtotal do Excel" (aba Modelo de dados) e é
      propagada por `ler_excel` e `ler_arquivo`.
    """
    titulo = str(ws.title)
    linhas = [list(r) for r in ws.iter_rows(values_only=True)]
    acima = 0
    while linhas and all(_vazia(v) for v in linhas[0]):
        linhas.pop(0)
        acima += 1
    if not linhas:
        return ResultadoImportacao(False, mensagem=f"A aba '{titulo}' está vazia e foi ignorada.", erro="aba_vazia",
                                   aba=titulo, origem="excel")
    ih = _detectar_cabecalho(linhas)
    cab, corpo = linhas[ih], linhas[ih + 1:]
    ncol = max((i + 1 for l in [cab] + corpo for i, v in enumerate(l) if not _vazia(v)), default=0)
    if ncol == 0 or not corpo:
        return ResultadoImportacao(False, mensagem=f"A aba '{titulo}' não tem dados abaixo do cabeçalho.", erro="sem_dados",
                                   aba=titulo, origem="excel")
    cab_txt = [("" if _vazia(v) else str(v).strip()) for v in (cab + [None] * ncol)[:ncol]]
    registros = [(l + [None] * ncol)[:ncol] for l in corpo]
    em_branco = sum(all(_vazia(v) for v in l) for l in registros)
    registros = [l for l in registros if any(not _vazia(v) for v in l)]
    totais = [i for i, l in enumerate(registros) if _linha_total(l)]
    if totais and remover_totais:
        marcados = set(totais)
        registros = [l for i, l in enumerate(registros) if i not in marcados]
    erros_celulas = 0
    texto: list[list[str]] = []
    for l in registros:
        nova = []
        for v in l:
            if isinstance(v, str) and v.strip().upper() in ERROS_EXCEL:
                erros_celulas += 1
                nova.append("")
            else:
                nova.append(_celula_para_texto(v))
        texto.append(nova)
    if not texto:
        return ResultadoImportacao(False, mensagem=f"A aba '{titulo}' não tem registros.", erro="sem_dados", aba=titulo,
                                   origem="excel")
    duplicados = sorted({c for c in cab_txt if c and cab_txt.count(c) > 1})
    df = pd.DataFrame(texto, columns=_nomes_unicos(cab_txt), dtype=str)
    df, n_sem_nome = _limpar_tabela(df)
    if df.shape[1] == 0 or df.empty:
        return ResultadoImportacao(False, mensagem=f"A aba '{titulo}' não contém dados.", erro="sem_dados", aba=titulo,
                                   origem="excel")
    linha_cab = acima + ih + 1
    avisos: list[str] = []
    if linha_cab > 1:
        avisos.append(f"Aba '{titulo}': o cabeçalho está na linha {linha_cab}. As {linha_cab - 1} linha(s) de cima "
                      "(títulos ou filtros) foram ignoradas.")
    if totais and remover_totais:
        avisos.append(f"Aba '{titulo}': {len(totais)} linha(s) de total ou subtotal foram retiradas para não contar os "
                      "valores duas vezes.")
    elif totais:
        avisos.append(f"Aba '{titulo}': {len(totais)} linha(s) de total ou subtotal foram mantidas, por sua escolha. "
                      "Atenção: somas e contagens podem contar esses valores duas vezes.")
    if em_branco:
        avisos.append(f"Aba '{titulo}': {em_branco} linha(s) em branco ignoradas.")
    if erros_celulas:
        avisos.append(f"Aba '{titulo}': {erros_celulas} célula(s) com erro do Excel (#DIV/0!, #N/A…) ficaram vazias.")
    if duplicados:
        avisos.append(f"Aba '{titulo}': colunas com o mesmo nome ({', '.join(duplicados[:5])}) foram renomeadas (ex.: 'col.1').")
    if n_sem_nome:
        avisos.append(f"Aba '{titulo}': {n_sem_nome} coluna(s) sem cabeçalho foram chamadas de 'coluna_N'.")
    retiradas = len(totais) if remover_totais else 0
    lidos = len(texto) + retiradas
    # Fix: "validos" = só os registros aproveitados (como no CSV); totais retirados ficam em campo próprio (bug #3)
    integridade = {"campos_cabecalho": df.shape[1], "registros_lidos": lidos, "validos": len(df),
                   "totais_retirados": retiradas, "curtos": 0, "longos": 0,
                   "vazios_ignorados": em_branco, "linhas_em_branco": em_branco, "tratamento": "nenhum",
                   "aproveitados": len(df), "completados": 0, "descartados": retiradas, "exemplos": []}
    return ResultadoImportacao(True, df, "(planilha)", "xlsx",
                               f"Planilha importada (aba '{titulo}'): {formatar_numero(len(df))} registros e {df.shape[1]} colunas.",
                               None, avisos, integridade=integridade, aba=titulo, origem="excel")


def ler_excel(conteudo: bytes, nome_arquivo: str = "planilha.xlsx",
              remover_totais: bool = True) -> list[ResultadoImportacao]:
    """Lê todas as abas VISÍVEIS de um .xlsx/.xlsm (usa os valores já calculados das fórmulas). Uma tabela por aba."""
    if not nome_arquivo.lower().endswith(EXTENSOES_EXCEL):
        return [ResultadoImportacao(False, mensagem="Este formato não é aceito. Envie .xlsx ou .xlsm (um .xls antigo pode "
                                                    "ser salvo como .xlsx no próprio Excel).", erro="formato_invalido")]
    if not conteudo:
        return [ResultadoImportacao(False, mensagem="O arquivo está vazio.", erro="arquivo_vazio")]
    if not excel_disponivel():
        return [ResultadoImportacao(False, mensagem="Para abrir planilhas Excel, instale a biblioteca openpyxl "
                                                    "(no terminal: pip install openpyxl).", erro="dependencia_ausente")]
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    except Exception:   # mensagem genérica: nunca devolver trechos do conteúdo
        return [ResultadoImportacao(False, mensagem="Não consegui abrir a planilha. Ela pode estar corrompida, protegida por "
                                                    "senha ou não ser um arquivo Excel de verdade.", erro="excel_ilegivel",
                                    origem="excel")]
    resultados: list[ResultadoImportacao] = []
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                resultados.append(ResultadoImportacao(False, mensagem=f"A aba '{ws.title}' está oculta e foi ignorada.",
                                                      erro="aba_vazia", aba=str(ws.title), origem="excel"))
                continue
            resultados.append(_ler_aba(ws, remover_totais=remover_totais))
    finally:
        wb.close()
    return resultados or [ResultadoImportacao(False, mensagem="A planilha não tem abas visíveis.", erro="sem_dados",
                                              origem="excel")]


def ler_arquivo(conteudo: bytes, nome_arquivo: str, remover_totais: bool = True) -> list[ResultadoImportacao]:
    """Porta de entrada única para arquivos: CSV (uma tabela) ou Excel (uma tabela por aba com dados).
    `remover_totais` só vale para Excel (ver `_ler_aba`)."""
    if nome_arquivo.lower().endswith(EXTENSOES_EXCEL):
        return ler_excel(conteudo, nome_arquivo, remover_totais=remover_totais)
    return [ler_csv(conteudo, nome_arquivo)]


# =============================================================================
# QUALIDADE
# =============================================================================


def resumo_qualidade(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"total_registros": 0, "total_colunas": 0, "valores_ausentes": 0,
                "percentual_ausentes": 0.0, "duplicidades": 0, "percentual_duplicidades": 0.0,
                "colunas_constantes": []}

    total_celulas = df.shape[0] * df.shape[1]
    ausentes = int(df.isna().sum().sum())
    duplicidades = int(df.duplicated().sum())
    nunicos = df.nunique(dropna=True)

    return {
        "total_registros": int(df.shape[0]),
        "total_colunas": int(df.shape[1]),
        "valores_ausentes": ausentes,
        "percentual_ausentes": round(100 * ausentes / total_celulas, 2) if total_celulas else 0.0,
        "duplicidades": duplicidades,
        "percentual_duplicidades": round(100 * duplicidades / df.shape[0], 2),
        "colunas_constantes": nunicos[nunicos <= 1].index.tolist(),
    }


def qualidade_por_coluna(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    total = len(df)
    ausentes = df.isna().sum()
    unicos = df.nunique(dropna=True)
    return pd.DataFrame({
        "Coluna": df.columns,
        "Valores_ausentes": ausentes.values,
        "%_Ausentes": (100 * ausentes / total).round(2).values,
        "Valores_únicos": unicos.values,
        "Cardinalidade_%": (100 * unicos / total).round(2).values,
    })


def detectar_outliers_iqr(serie: pd.Series) -> dict:
    valores = limpar_e_converter_numerico(serie).dropna()
    if len(valores) < 4:
        return {"qtd": 0, "limite_inferior": None, "limite_superior": None, "percentual": 0.0}
    q1, q3 = valores.quantile(0.25), valores.quantile(0.75)
    iqr = q3 - q1
    li, ls = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    qtd = int(((valores < li) | (valores > ls)).sum())
    return {
        "qtd": qtd,
        "limite_inferior": float(li),
        "limite_superior": float(ls),
        "percentual": round(100 * qtd / len(valores), 2),
    }


def auditar_formato_data(serie: pd.Series) -> dict:
    """Conta datas com barra (DD/MM/AAAA ou MM/DD/AAAA) que SÓ fazem sentido em um dos dois formatos.
    O app lê datas como DD/MM; um 2º campo > 12 indica formato americano (MM/DD)."""
    s = serie.dropna().astype(str).str.strip()
    m = s.str.extract(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})").dropna()
    if m.empty:
        return {"total": 0, "primeiro_maior_12": 0, "segundo_maior_12": 0, "ambiguas": 0, "nivel": "ok"}
    m = m.astype(int)
    segundo = int((m[1] > 12).sum())
    return {"total": int(len(m)), "primeiro_maior_12": int((m[0] > 12).sum()),
            "segundo_maior_12": segundo,
            "ambiguas": int(((m[0] <= 12) & (m[1] <= 12)).sum()),
            "nivel": "alerta" if segundo else "ok"}   # BLOCO 10c: com 2º campo > 12, o usuário escolhe o formato


def auditar_formato_numerico(serie: pd.Series) -> dict:
    """Conta valores só com ponto (sem vírgula) lidos como MILHAR ('1.234' -> 1234; padrão brasileiro) e como
    DECIMAL ('12.5' -> 12,5). Mistura dos dois indica arquivo com formatos diferentes na mesma coluna."""
    s = serie.dropna().astype(str).str.strip().str.replace(r"[R$€£%\s]", "", regex=True)
    so_ponto = s[~s.str.contains(",", regex=False)]
    milhar = so_ponto.str.match(r"^[+-]?[1-9]\d{0,2}(\.\d{3})+$")
    decimal = so_ponto.str.match(r"^[+-]?\d+\.\d+$") & ~milhar
    pontos = so_ponto.str.count(r"\.")
    return {"milhar_ambiguo": int((milhar & (pontos == 1)).sum()), "milhar_inequivoco": int((milhar & (pontos >= 2)).sum()),
            "ponto_decimal": int(decimal.sum())}


# =============================================================================
# MODELO DE DADOS: MÚLTIPLAS FONTES E RELACIONAMENTOS
# =============================================================================
# Ideia central: o motor de análise continua trabalhando sobre UMA tabela. O modelo (tabelas + relacionamentos
# validados) produz essa tabela — a "base analítica" — ligando a tabela FATO às dimensões. Regras que protegem os
# números: relacionamentos são muitos-para-um (a chave do lado "um" deve ser única), a junção é à esquerda e o nº de
# linhas da fato nunca pode mudar, e tabelas filhas (1:N) só entram AGREGADAS (nunca linha a linha).

LIMIAR_ORFAOS_ALERTA_PCT = 20.0           # % de registros da origem sem correspondência -> alerta
MODOS_NORMALIZACAO = {
    "exata": "Comparar as chaves como estão (apenas sem espaços nas pontas)",
    "texto": "Ignorar maiúsculas, acentos e espaços repetidos",
    "numerica": "Ignorar zeros à esquerda e '.0' final (ex.: 007 = 7; 12.0 = 12)",
}
OPCOES_DUPLICIDADE = {
    "bloquear": "Bloquear se a chave do lado 'um' tiver duplicatas (recomendado)",
    "primeira": "Manter a 1ª ocorrência das chaves duplicadas (descarta as demais — informado)",
}
AGREGACOES_FILHA = ("soma", "media", "contagem", "contagem_distinta", "minimo", "maximo")
_PALAVRAS_CHAVE_GENERICAS = {"id", "cod", "codigo", "chave", "key", "cd", "num"}


MAX_TABELAS_MODELO = 3   # tabela principal + até 2 ligações (o uso real da equipe)


class ModeloInvalido(ValueError):
    """O modelo ou um relacionamento violaria a integridade da análise (ex.: duplicaria linhas da fato)."""


@dataclass
class TabelaFonte:
    nome: str                          # identificador da tabela no modelo
    arquivo: str                       # nome do arquivo de origem
    df: pd.DataFrame                   # dados brutos (texto), já limpos pela importação
    separador: Optional[str] = None
    encoding: Optional[str] = None
    integridade: Optional[dict] = None
    avisos: list[str] = field(default_factory=list)
    resultado: Any = None              # ResultadoImportacao original (usado no caminho de uma única tabela)
    aba: str = ""                       # Excel: aba de origem
    arquivos_empilhados: list[str] = field(default_factory=list)   # quando a tabela é a união de vários arquivos
    # Fix: identidade do conteúdo para a fila de importação não confundir arquivos de mesmo nome (bug #4)
    conteudo_id: str = ""              # SHA-256 dos bytes do arquivo de origem ("" = tabela derivada/empilhada)


@dataclass
class Relacionamento:
    """`origem` é o lado 'muitos' (ex.: atendimentos) e `destino` o lado 'um' (ex.: setores, com chave única).
    Com `agregacoes` preenchidas a origem é uma tabela FILHA (1:N) do destino: ela é agregada por chave antes de entrar."""
    origem: str
    coluna_origem: str
    destino: str
    coluna_destino: str
    normalizacao: str = "exata"
    duplicidade: str = "bloquear"
    agregacoes: list[tuple[str, str]] = field(default_factory=list)   # [(coluna, agregação)]
    ativo: bool = True

    def descricao(self) -> str:
        return f"{self.origem}.{self.coluna_origem} → {self.destino}.{self.coluna_destino}"

    def assinatura(self) -> tuple:
        return (self.origem, self.coluna_origem, self.destino, self.coluna_destino, self.normalizacao,
                self.duplicidade, tuple(map(tuple, self.agregacoes)), self.ativo)


@dataclass
class ValidacaoRelacionamento:
    descricao: str
    status: str                        # valida | valida_com_ressalvas | invalida
    motivo: str
    cardinalidade: str                 # observada: um-para-um | um-para-muitos | muitos-para-um | muitos-para-muitos | indefinida
    linhas_origem: int = 0
    linhas_destino: int = 0
    chaves_nulas_origem: int = 0
    chaves_nulas_destino: int = 0
    chaves_duplicadas_destino: int = 0
    linhas_em_chaves_duplicadas: int = 0          # linhas A MAIS por código repetido (3 ocorrências = 2) — bug #6
    exemplos_duplicadas: list[str] = field(default_factory=list)
    com_correspondencia: int = 0
    sem_correspondencia: int = 0
    pct_sem_correspondencia: float = 0.0
    chaves_orfas_distintas: int = 0
    exemplos_orfas: list[str] = field(default_factory=list)
    origem_unica: bool = False                 # cada chave aparece uma vez na origem?
    destino_unico: bool = False                # cada chave aparece uma vez no destino?
    cardinalidade_permitida: str = "muitos-para-um ou um-para-um"   # regra para incorporar o destino à base
    linhas_estimadas_apos_juncao: int = 0      # linhas da origem depois de uma junção linha a linha
    risco_multiplicacao: bool = False          # a junção linha a linha aumentaria o nº de linhas?


@dataclass
class JoinInfo:
    descricao: str
    tipo: str                          # dimensão | filha agregada
    linhas_base: int
    com_correspondencia: int
    sem_correspondencia: int
    pct_sem_correspondencia: float
    colunas_adicionadas: list[str]
    duplicidades_descartadas: int
    validacao: ValidacaoRelacionamento
    linhas_antes: int = 0
    linhas_depois: int = 0


@dataclass
class ResultadoBase:
    df: pd.DataFrame
    fato: str
    joins: list[JoinInfo]
    avisos: list[str]
    tabelas_nao_conectadas: list[str]
    linhas_fato: int
    naturezas: dict[str, str] = field(default_factory=dict)   # natureza sugerida das colunas agregadas
    mapa_colunas: list[dict] = field(default_factory=list)    # origem de CADA coluna da base (rastreabilidade)
    transformacoes: list[str] = field(default_factory=list)   # registro das etapas aplicadas, em ordem
    granularidade: str = ""                                   # o que é uma linha da base (ex.: "1 linha = 1 registro de atendimentos")

    def origem_da_coluna(self, coluna: str) -> Optional[dict]:
        return next((m for m in self.mapa_colunas if m["coluna_base"] == coluna), None)


@dataclass
class ModeloDados:
    tabelas: dict[str, TabelaFonte] = field(default_factory=dict)
    relacionamentos: list[Relacionamento] = field(default_factory=list)
    fato: Optional[str] = None
    colunas: dict[str, list[str]] = field(default_factory=dict)   # tabela -> colunas trazidas (ausente = todas, exceto as de nome sensível)
    resultado: Optional[ResultadoBase] = None
    assinatura_resultado: Optional[tuple] = None   # assinatura do modelo no momento em que a base foi construída

    def desatualizado(self) -> bool:
        """True se a base construída não corresponde mais ao modelo atual (tabelas, relacionamentos, colunas ou fato mudaram)."""
        return self.resultado is not None and self.assinatura_resultado != self.assinatura()

    def assinatura(self) -> tuple:
        return (self.fato, tuple(sorted(self.tabelas)), tuple(r.assinatura() for r in self.relacionamentos),
                tuple(sorted((t, tuple(c)) for t, c in self.colunas.items())))


def nome_tabela(nome_arquivo: str, existentes: set[str]) -> str:
    """Nome seguro e único da tabela a partir do nome do arquivo."""
    base = re.sub(r"\.(csv|xlsx|xlsm)$", "", nome_arquivo.strip(), flags=re.I)
    base = re.sub(r"[^\w]+", "_", base, flags=re.UNICODE).strip("_") or "tabela"
    nome, i = base, 2
    while nome in existentes:
        nome, i = f"{base}_{i}", i + 1
    return nome


def escolher_fato_padrao(modelo: ModeloDados) -> Optional[str]:
    """Fato padrão: a tabela com mais linhas (empate: a primeira importada)."""
    if not modelo.tabelas:
        return None
    return max(modelo.tabelas, key=lambda n: len(modelo.tabelas[n].df))


def normalizar_chave(serie: pd.Series, modo: str = "exata") -> pd.Series:
    """Chave comparável (texto). Vazios viram NA. `texto`: sem acentos/caixa; `numerica`: sem zeros à esquerda."""
    s = serie.astype("string").str.strip()
    s = s.mask(s == "")
    if modo == "texto":
        s = (s.str.normalize("NFKD").str.replace("[\u0300-\u036f]", "", regex=True).str.lower()
             .str.replace(r"\s+", " ", regex=True))
    elif modo == "numerica":
        s = s.str.replace(r"\.0+$", "", regex=True).str.replace(r"^0+(?=\d)", "", regex=True)
    return s


def validar_relacionamento(modelo: ModeloDados, rel: Relacionamento) -> ValidacaoRelacionamento:
    """Valida um relacionamento SEM alterar dados: unicidade da chave do lado 'um', chaves nulas, órfãos e cardinalidade."""
    desc = rel.descricao()
    to, td = modelo.tabelas.get(rel.origem), modelo.tabelas.get(rel.destino)
    if to is None or td is None:
        return ValidacaoRelacionamento(desc, "invalida", "Tabela inexistente no modelo.", "indefinida")
    if rel.coluna_origem not in to.df.columns or rel.coluna_destino not in td.df.columns:
        return ValidacaoRelacionamento(desc, "invalida", "Coluna inexistente na tabela.", "indefinida")
    ko_all = normalizar_chave(to.df[rel.coluna_origem], rel.normalizacao)
    kd_all = normalizar_chave(td.df[rel.coluna_destino], rel.normalizacao)
    ko, kd = ko_all.dropna(), kd_all.dropna()
    dup = kd.duplicated(keep=False)
    # Fix: contar só as linhas a mais (a 1ª ocorrência de cada código não é duplicata) (bug #6)
    chaves_dup = int(kd[dup].nunique())
    linhas_dup = int(dup.sum()) - chaves_dup
    conjunto = set(kd.unique())
    pertence = ko.isin(conjunto)
    com, sem = int(pertence.sum()), int((~pertence).sum())
    pct = 100.0 * sem / len(ko) if len(ko) else 0.0
    origem_unica, destino_unico = bool(ko.is_unique), bool(kd.is_unique)
    # Cardinalidade OBSERVADA: unicidade avaliada separadamente em cada lado ("um" = chave única naquele lado).
    card = {(True, True): "um-para-um", (True, False): "um-para-muitos",
            (False, True): "muitos-para-um", (False, False): "muitos-para-muitos"}[(origem_unica, destino_unico)]
    # Linhas da origem depois de uma junção linha a linha: cada chave casada vira tantas linhas quantas houver no destino.
    contagem_destino = kd.value_counts()
    estimadas = int(len(to.df) - len(ko) + ko.map(contagem_destino).fillna(1).sum())
    oculta = lambda vals, col: [exemplo_seguro(x, col) for x in vals]
    v = ValidacaoRelacionamento(
        desc, "valida", "", card, len(to.df), len(td.df), int(ko_all.isna().sum()), int(kd_all.isna().sum()),
        chaves_dup, linhas_dup, oculta(kd[dup].drop_duplicates().head(5).tolist(), rel.coluna_destino), com, sem,
        round(pct, 2), int(ko[~pertence].nunique()), oculta(ko[~pertence].drop_duplicates().head(5).tolist(), rel.coluna_origem),
        origem_unica, destino_unico, "muitos-para-um ou um-para-um", estimadas, estimadas > len(to.df))
    exemplos_orfas = ", ".join(v.exemplos_orfas[:3])
    if chaves_dup and rel.duplicidade == "bloquear":
        v.status = "invalida"
        # Fix: a mensagem informa as linhas a mais, não o total de ocorrências (bug #6)
        v.motivo = (f"Em '{rel.destino}', a coluna '{rel.coluna_destino}' tem códigos repetidos ({chaves_dup} código(s), "
                    f"{linhas_dup} linha(s) a mais, ex.: {', '.join(v.exemplos_duplicadas[:3])}). Ligando assim, cada registro de "
                    f"'{rel.origem}' poderia virar vários: {formatar_numero(len(to.df))} linhas passariam a "
                    f"{formatar_numero(estimadas)}, e somas e contagens sairiam infladas. Escolha outra coluna, corrija os "
                    "códigos repetidos ou use a opção de manter só a primeira ocorrência.")
    elif len(ko) == 0:
        v.status, v.motivo = "invalida", f"A coluna '{rel.coluna_origem}' de '{rel.origem}' está vazia."
    elif com == 0:
        v.status = "invalida"
        v.motivo = (f"Nenhum código de '{rel.origem}' foi encontrado em '{rel.destino}'"
                    + (f" (ex.: {exemplos_orfas})" if exemplos_orfas else "")
                    + ". Confira se escolheu as colunas certas ou mude a forma de comparar os códigos "
                      "(por exemplo, ignorando zeros à esquerda).")
    else:
        partes = []
        if sem:
            partes.append(f"{sem} de {len(ko)} registros de '{rel.origem}' ({pct:.1f}%) não encontraram o código em "
                          f"'{rel.destino}'" + (f" (ex.: {exemplos_orfas})" if exemplos_orfas else ""))
        if v.chaves_nulas_origem:
            partes.append(f"{v.chaves_nulas_origem} registro(s) estão sem código para ligar")
        if chaves_dup:
            partes.append(f"{chaves_dup} código(s) repetido(s) em '{rel.destino}': foi mantida a primeira ocorrência")
        if partes:
            v.status, v.motivo = "valida_com_ressalvas", "; ".join(partes) + "."
    return v


def _cria_ciclo(modelo: ModeloDados, rel: Relacionamento) -> bool:
    """True se `origem` e `destino` já estiverem ligados por outro caminho (caminhos ambíguos duplicariam/confundiriam)."""
    pai = {n: n for n in modelo.tabelas}

    def raiz(x: str) -> str:
        while pai[x] != x:
            pai[x] = pai[pai[x]]
            x = pai[x]
        return x
    for r in modelo.relacionamentos:
        if r.ativo and r.origem in pai and r.destino in pai:
            pai[raiz(r.origem)] = raiz(r.destino)
    return raiz(rel.origem) == raiz(rel.destino)


def adicionar_relacionamento(modelo: ModeloDados, rel: Relacionamento) -> ValidacaoRelacionamento:
    """Valida e adiciona. Relacionamentos inválidos, repetidos, entre a mesma tabela ou que fechem ciclo NÃO são adicionados."""
    v = validar_relacionamento(modelo, rel)
    envolvidas = {t for r in modelo.relacionamentos if r.ativo for t in (r.origem, r.destino)} | {rel.origem, rel.destino}
    if len(envolvidas) > MAX_TABELAS_MODELO:
        v.status = "invalida"
        v.motivo = (f"O aplicativo trabalha com até {MAX_TABELAS_MODELO} tabelas ligadas por vez (a principal e mais duas). "
                    "Remova uma ligação antes de adicionar esta.")
        return v
    if rel.origem == rel.destino:
        v.status, v.motivo = "invalida", "Origem e destino são a mesma tabela."
    elif any(r.assinatura()[:4] == rel.assinatura()[:4] for r in modelo.relacionamentos):
        v.status, v.motivo = "invalida", "Este relacionamento já existe."
    elif v.status != "invalida" and _cria_ciclo(modelo, rel):
        v.status = "invalida"
        v.motivo = ("Já existe um caminho entre estas tabelas: um segundo caminho tornaria a análise ambígua. "
                    "Remova o outro relacionamento antes de adicionar este.")
    if v.status != "invalida":
        modelo.relacionamentos.append(rel)
    return v


def _tokens_nome(nome: str) -> set[str]:
    return set(_tokens(nome))


def _prefixo_comum(a: set[str], b: set[str]) -> bool:
    return any(min(len(x), len(y)) >= 4 and (x.startswith(y) or y.startswith(x)) for x in a for y in b)


def _nomes_compativeis(col_a: str, tab_a: str, col_b: str, tab_b: str) -> bool:
    """Colunas com cara de chave do mesmo conceito: mesmo nome, ou diferença só em palavras de chave (id/cod/codigo)
    ('setor' × 'id_setor'), ou chave genérica ('id') de uma tabela cujo nome combina com a outra coluna."""
    ta, tb = _tokens_nome(col_a), _tokens_nome(col_b)
    if not ta or not tb:
        return False
    esp_a, esp_b = ta - _PALAVRAS_CHAVE_GENERICAS, tb - _PALAVRAS_CHAVE_GENERICAS
    if esp_a and esp_a == esp_b:
        return True
    if not esp_a and not esp_b:
        return False                         # 'id' × 'id': genérico demais para sugerir
    if not esp_a:                            # a é genérica (ex.: 'id' da tabela setores)
        return _prefixo_comum(esp_b, _tokens_nome(tab_a))
    if not esp_b:
        return _prefixo_comum(esp_a, _tokens_nome(tab_b))
    return False


@dataclass
class SugestaoRelacionamento:
    relacionamento: Relacionamento
    cobertura: float                         # fração das chaves da origem encontrada no destino
    forca: str                               # forte | moderada
    validacao: ValidacaoRelacionamento


def sugerir_relacionamentos(modelo: ModeloDados, cobertura_minima: float = 0.5, limite: int = 20) -> list[SugestaoRelacionamento]:
    """Sugere relacionamentos muitos-para-um: colunas com nomes compatíveis, chave única em um dos lados e boa
    sobreposição de valores. Nunca sugere muitos-para-muitos. São SUGESTÕES: nada é aplicado sem confirmação."""
    existentes = {r.assinatura()[:4] for r in modelo.relacionamentos}
    saida: list[SugestaoRelacionamento] = []
    nomes = list(modelo.tabelas)
    for i, a in enumerate(nomes):
        for b in nomes[i + 1:]:
            da, db = modelo.tabelas[a].df, modelo.tabelas[b].df
            for ca in list(da.columns)[:40]:
                for cb in list(db.columns)[:40]:
                    if not _nomes_compativeis(ca, a, cb, b):
                        continue
                    melhor = None
                    for modo in ("exata", "numerica"):           # 'numerica' resolve 001 × 1 (zeros à esquerda)
                        ka, kb = normalizar_chave(da[ca], modo).dropna(), normalizar_chave(db[cb], modo).dropna()
                        if ka.empty or kb.empty or not (ka.is_unique or kb.is_unique):
                            continue
                        a_e_destino = (ka.is_unique and not kb.is_unique) or (ka.is_unique and kb.is_unique and len(da) <= len(db))
                        (o, co, ko), (d, cd, kd) = (((b, cb, kb), (a, ca, ka)) if a_e_destino else ((a, ca, ka), (b, cb, kb)))
                        cob = float(ko.isin(set(kd)).mean())
                        if melhor is None or cob > melhor[0] + 1e-9:
                            melhor = (cob, modo, o, co, d, cd)
                    if melhor is None or melhor[0] < cobertura_minima:
                        continue
                    cob, modo, o, co, d, cd = melhor
                    rel = Relacionamento(o, co, d, cd, modo)
                    if rel.assinatura()[:4] in existentes:
                        continue
                    forca = "forte" if cob >= 0.95 else "moderada"
                    saida.append(SugestaoRelacionamento(rel, round(cob, 4), forca, validar_relacionamento(modelo, rel)))
    saida.sort(key=lambda s: (s.forca != "forte", -s.cobertura))
    return saida[:limite]


def _fmt_br(x: Any) -> Any:
    """Número -> texto no padrão brasileiro SEM milhar (ex.: 1234,5), que o importador lê sem ambiguidade."""
    if x is None or pd.isna(x):
        return pd.NA
    x = float(x)
    return str(int(x)) if x.is_integer() else format(x, ".10g").replace(".", ",")


def agregar_tabela_filha(df: pd.DataFrame, coluna_chave: str, modo: str, agregacoes: list[tuple[str, str]],
                         prefixo: str) -> tuple[pd.DataFrame, dict[str, str]]:
    """Agrega a tabela filha por chave (uma linha por chave). Sempre inclui `<tabela>.qtd_linhas`.
    Retorna (tabela agregada indexada por chave normalizada, naturezas sugeridas das colunas geradas)."""
    d = df.copy()
    d["__k__"] = normalizar_chave(d[coluna_chave], modo)
    d = d[d["__k__"].notna()]
    g = d.groupby("__k__")
    saida = pd.DataFrame({f"{prefixo}.qtd_linhas": g.size()})
    naturezas = {f"{prefixo}.qtd_linhas": "quantidade"}
    for col, agg in agregacoes:
        if col not in d.columns or agg not in AGREGACOES_FILHA:
            continue
        nome = f"{prefixo}.{col} ({agg})"
        if agg == "contagem":
            saida[nome], naturezas[nome] = g[col].count(), "quantidade"
        elif agg == "contagem_distinta":
            saida[nome], naturezas[nome] = g[col].nunique(), "quantidade"
        else:
            num = limpar_e_converter_numerico(d[col]).groupby(d["__k__"])
            saida[nome] = {"soma": num.sum(min_count=1), "media": num.mean(), "minimo": num.min(), "maximo": num.max()}[agg]
            nat = _detectar_natureza(col, d[col].dropna().head(200)) if agg == "soma" else "indefinido"
            naturezas[nome] = nat if nat in ("quantidade", "monetario") else "indefinido"
    return saida.apply(lambda s: s.map(_fmt_br)).astype("string"), naturezas


def construir_base_analitica(modelo: ModeloDados, fato: Optional[str] = None) -> ResultadoBase:
    """Constrói a base analítica a partir da FATO.

    - Dimensões (muitos-para-um): junção à esquerda; colunas entram como `tabela.coluna`; chain (snowflake) suportado.
    - Tabelas filhas (1:N) com `agregacoes`: agregadas por chave e então juntadas (nunca linha a linha).
    - O nº de linhas da fato NÃO pode mudar (ModeloInvalido se mudar). Registros sem correspondência ficam com
      valores vazios nas colunas da dimensão e são contados (órfãos).
    - Relacionamentos que não puderam ser aplicados (tabela inalcançável, ciclo, filha sem agregação) são avisados.
    """
    fato = fato or modelo.fato
    if fato not in modelo.tabelas:
        raise ModeloInvalido("Escolha a tabela fato.")
    base = modelo.tabelas[fato].df.copy().reset_index(drop=True)
    n0 = len(base)
    col_id = "id_registro" if "id_registro" not in base.columns else "id_registro_base"
    base.insert(0, col_id, np.arange(1, n0 + 1))          # identificador estável: nº da linha na tabela principal
    incorporadas = {fato}
    mapa: dict[str, dict[str, str]] = {fato: {c: c for c in base.columns if c != col_id}}   # tabela -> coluna -> nome na base
    mapa_colunas: list[dict] = [{"coluna_base": col_id, "tabela": fato, "coluna_original": "(nº da linha)",
                                 "relacionamento": "—", "chave_base": "—", "chave_destino": "—", "derivada": True}]
    mapa_colunas += [{"coluna_base": c, "tabela": fato, "coluna_original": c, "relacionamento": "—", "chave_base": "—",
                      "chave_destino": "—", "derivada": False} for c in mapa[fato]]
    transformacoes = [f"Tabela principal '{fato}': {formatar_numero(n0)} linhas; coluna '{col_id}' criada com o nº de cada linha."]
    pendentes = [r for r in modelo.relacionamentos if r.ativo]
    joins: list[JoinInfo] = []
    avisos: list[str] = []
    naturezas: dict[str, str] = {}

    def _nome_livre(nome: str) -> str:
        """Nome de coluna sem conflito com a base; conflitos são avisados (nunca sobrescritos)."""
        if nome not in base.columns:
            return nome
        k = 2
        while f"{nome} ({k})" in base.columns:
            k += 1
        avisos.append(f"A coluna '{nome}' já existia na base; a nova foi chamada '{nome} ({k})'.")
        return f"{nome} ({k})"
    progresso = True
    while progresso and pendentes:
        progresso = False
        for rel in list(pendentes):
            val = None
            if rel.origem in incorporadas and rel.destino not in incorporadas and not rel.agregacoes:
                col_base = mapa[rel.origem].get(rel.coluna_origem)
                if col_base is None:
                    raise ModeloInvalido(f"A coluna '{rel.origem}.{rel.coluna_origem}' não está na base: inclua-a nas colunas trazidas.")
                val = validar_relacionamento(modelo, rel)
                if val.status == "invalida":
                    raise ModeloInvalido(f"Relacionamento inválido ({rel.descricao()}): {val.motivo}")
                dim = modelo.tabelas[rel.destino].df
                exigidas = {r.coluna_origem for r in modelo.relacionamentos if r.ativo and r.origem == rel.destino}
                escolhidas = modelo.colunas.get(rel.destino)
                if escolhidas is None:     # padrão: todas, exceto colunas com nome de dado pessoal (privacidade)
                    escolhidas = [c for c in dim.columns if not _nome_contem(c, PALAVRAS_SENSIVEIS)]
                cols = [c for c in dict.fromkeys(list(escolhidas) + sorted(exigidas)) if c in dim.columns and c != rel.coluna_destino]
                # A chave do DESTINO entra como coluna própria ('destino.chave'): em relacionamentos encadeados, a próxima
                # junção usa o valor que veio do destino — vazio quando não houve correspondência —, e nunca o código
                # original da tabela de origem (que poderia casar, por engano, com outra tabela).
                cols_dest = [rel.coluna_destino] + cols
                dd = dim[cols_dest].copy()
                dd["__k__"] = normalizar_chave(dd[rel.coluna_destino], rel.normalizacao)
                dd = dd[dd["__k__"].notna()]
                descartadas = int(dd["__k__"].duplicated().sum())
                if descartadas:
                    dd = dd.drop_duplicates("__k__", keep="first")      # só chega aqui com duplicidade == "primeira"
                renome = {c: _nome_livre(f"{rel.destino}.{c}") for c in cols_dest}
                dd = dd.rename(columns=renome)[["__k__"] + list(renome.values())]
                base["__k__"] = normalizar_chave(base[col_base], rel.normalizacao)
                casou = base["__k__"].notna() & base["__k__"].isin(set(dd["__k__"]))
                chaves_preenchidas = int(base["__k__"].notna().sum())      # chave nula não é "órfã": é contada à parte
                linhas_antes = len(base)
                base = base.merge(dd, on="__k__", how="left", validate="many_to_one")
                if len(base) != n0:
                    raise ModeloInvalido(f"A junção {rel.descricao()} alterou o nº de linhas da tabela principal "
                                         f"({n0} → {len(base)}).")
                # Conferência dos VALORES (não só do nº de linhas): onde houve correspondência, a chave trazida do destino
                # é igual à chave da base; onde não houve, as colunas do destino ficam vazias.
                chave_trazida = normalizar_chave(base[renome[rel.coluna_destino]], rel.normalizacao)
                if not (chave_trazida[casou.to_numpy()] == base.loc[casou.to_numpy(), "__k__"]).all() \
                        or chave_trazida[~casou.to_numpy()].notna().any():
                    raise ModeloInvalido(f"A conferência da junção {rel.descricao()} falhou: chaves trazidas não correspondem.")
                base = base.drop(columns="__k__")
                mapa[rel.destino] = {c: renome[c] for c in cols_dest}
                for c in cols_dest:
                    mapa_colunas.append({"coluna_base": renome[c], "tabela": rel.destino, "coluna_original": c,
                                         "relacionamento": rel.descricao(), "chave_base": col_base,
                                         "chave_destino": rel.coluna_destino, "derivada": False})
                novas, tipo = [renome[c] for c in cols_dest], "dimensão"
                transformacoes.append(f"Ligação {rel.descricao()} (comparação '{rel.normalizacao}'): {linhas_antes} linhas antes, "
                                      f"{len(base)} depois; {int(casou.sum())} com correspondência; chave usada na base: "
                                      f"'{col_base}'.")
                com = int(casou.sum())
                sem = chaves_preenchidas - com
            elif rel.destino in incorporadas and rel.origem not in incorporadas and rel.agregacoes:
                col_base = mapa[rel.destino].get(rel.coluna_destino)
                if col_base is None:
                    raise ModeloInvalido(f"A coluna '{rel.destino}.{rel.coluna_destino}' não está na base.")
                val = validar_relacionamento(modelo, rel)
                if val.status == "invalida":
                    raise ModeloInvalido(f"Relacionamento inválido ({rel.descricao()}): {val.motivo}")
                chave_base = normalizar_chave(base[col_base], rel.normalizacao)
                if chave_base.dropna().duplicated().any():
                    raise ModeloInvalido(
                        f"A chave '{rel.destino}.{rel.coluna_destino}' repete-se na base: agregar '{rel.origem}' duplicaria "
                        "as medidas agregadas (somas e contagens infladas). Use uma chave única por linha da base.")
                ag, nats = agregar_tabela_filha(modelo.tabelas[rel.origem].df, rel.coluna_origem, rel.normalizacao,
                                                rel.agregacoes, rel.origem)
                base["__k__"] = chave_base
                casou = base["__k__"].notna() & base["__k__"].isin(set(ag.index))
                chaves_preenchidas = int(base["__k__"].notna().sum())
                linhas_antes = len(base)
                renome_ag = {c: _nome_livre(c) for c in ag.columns}
                ag = ag.rename(columns=renome_ag)
                naturezas.update({renome_ag[c]: n for c, n in nats.items()})
                base = base.merge(ag.reset_index().rename(columns={"index": "__k__"}), on="__k__", how="left", validate="many_to_one")
                if len(base) != n0:
                    raise ModeloInvalido(f"A junção {rel.descricao()} alterou o nº de linhas da tabela principal "
                                         f"({n0} → {len(base)}).")
                base = base.drop(columns="__k__")
                mapa[rel.origem] = {c: c for c in ag.columns}
                for c in ag.columns:
                    mapa_colunas.append({"coluna_base": c, "tabela": rel.origem, "coluna_original": c.split(".", 1)[-1],
                                         "relacionamento": rel.descricao(), "chave_base": col_base,
                                         "chave_destino": rel.coluna_origem, "derivada": True})
                novas, tipo, descartadas = list(ag.columns), "filha agregada", 0
                transformacoes.append(f"Tabela '{rel.origem}' agregada por '{rel.coluna_origem}' (1 linha por código) e ligada "
                                      f"pela chave '{col_base}': {linhas_antes} linhas antes, {len(base)} depois.")
                com = int(casou.sum())
                sem = chaves_preenchidas - com
            else:
                continue
            incorporadas.add(rel.destino if tipo == "dimensão" else rel.origem)
            pendentes.remove(rel)
            progresso = True
            total_chaves = com + sem
            joins.append(JoinInfo(rel.descricao(), tipo, n0, com, sem, round(100 * sem / total_chaves, 2) if total_chaves else 0.0,
                                  novas, descartadas, val, linhas_antes, len(base)))
    for rel in pendentes:
        if rel.agregacoes and rel.origem in incorporadas:
            avisos.append(f"Relacionamento ignorado ({rel.descricao()}): está marcado como agregação de tabela filha, mas "
                          f"'{rel.origem}' já está na base (é a fato ou foi incorporada). Remova as agregações para usá-lo como dimensão.")
        elif rel.origem not in incorporadas and rel.destino not in incorporadas:
            avisos.append(f"Relacionamento ignorado ({rel.descricao()}): nenhuma das tabelas está ligada à fato '{fato}'.")
        elif rel.destino in incorporadas and rel.origem not in incorporadas and not rel.agregacoes:
            avisos.append(f"Relacionamento ignorado ({rel.descricao()}): '{rel.origem}' é filha (1:N) de '{rel.destino}'. "
                          "Incorporá-la linha a linha duplicaria registros; defina agregações para usá-la.")
        else:
            avisos.append(f"Relacionamento ignorado ({rel.descricao()}): caminho já coberto por outro relacionamento.")
    nao = [t for t in modelo.tabelas if t not in incorporadas]
    if nao:
        avisos.append("Tabelas sem ligação com a fato (não incluídas na base): " + ", ".join(nao) + ".")
    granularidade = f"1 linha = 1 registro de '{fato}' ({formatar_numero(n0)} linhas, preservadas em todas as ligações)"
    return ResultadoBase(base, fato, joins, avisos, nao, n0, naturezas, mapa_colunas, transformacoes, granularidade)


@dataclass
class ResultadoEmpilhamento:
    df: pd.DataFrame
    arquivos: list[str]
    linhas_por_arquivo: dict[str, int]
    linhas_repetidas: int                     # linhas idênticas que aparecem em mais de um arquivo
    removidas: int
    chave: Optional[str]                      # coluna identificadora usada para conferir repetições
    chaves_repetidas: int
    periodos: dict[str, tuple]                # arquivo -> (data mínima, data máxima), quando há coluna de data
    sobreposicoes: list[str]
    avisos: list[str]


def grupos_mesmo_layout(modelo: "ModeloDados") -> list[list[str]]:
    """Tabelas com exatamente as mesmas colunas (em qualquer ordem): candidatas a empilhar (ex.: extratos mensais)."""
    grupos: dict[tuple, list[str]] = {}
    for nome, t in modelo.tabelas.items():
        grupos.setdefault(tuple(sorted(map(str, t.df.columns))), []).append(nome)
    return [g for g in grupos.values() if len(g) > 1]


def empilhar_tabelas(tabelas: list["TabelaFonte"], remover_repetidas: bool = False) -> ResultadoEmpilhamento:
    """Une arquivos do MESMO layout (um embaixo do outro), com a coluna 'arquivo_origem'. Nada é descartado em silêncio:
    linhas idênticas entre arquivos e períodos sobrepostos são contados e avisados; só são removidos se o usuário pedir."""
    colunas = list(tabelas[0].df.columns)
    if any(sorted(map(str, t.df.columns)) != sorted(map(str, colunas)) for t in tabelas[1:]):
        raise ModeloInvalido("Só dá para empilhar arquivos com as mesmas colunas.")
    partes = []
    for t in tabelas:
        d = t.df[colunas].copy()
        d.insert(0, "arquivo_origem", t.arquivo + (f" [{t.aba}]" if t.aba else ""))
        partes.append(d)
    df = pd.concat(partes, ignore_index=True)
    tipos = classificar_dataframe(df[colunas])
    # linha "repetida entre arquivos": uma linha idêntica já apareceu antes em OUTRO arquivo
    primeiro_arquivo = df.fillna("\x00").groupby(colunas, sort=False)["arquivo_origem"].transform("first")
    entre_arquivos = df["arquivo_origem"] != primeiro_arquivo
    n_repetidas = int(entre_arquivos.sum())
    chave = next((c for c, i in tipos.items() if i.tipo == "identificador"
                  and all(t.df[c].dropna().is_unique for t in tabelas)), None)
    chaves_rep = 0
    if chave:
        por_arquivo = df.dropna(subset=[chave]).drop_duplicates(["arquivo_origem", chave])
        chaves_rep = int(por_arquivo[chave].duplicated().sum())
    periodos: dict[str, tuple] = {}
    col_data = next((c for c, i in tipos.items() if i.tipo == "data"), None)
    if col_data:
        datas = converter_datas(df[col_data])
        for arq, g in datas.groupby(df["arquivo_origem"]):
            g = g.dropna()
            if len(g):
                periodos[arq] = (g.min(), g.max())
    sobre = []
    itens = sorted(periodos.items(), key=lambda kv: kv[1][0])
    for (a1, (i1, f1)), (a2, (i2, f2)) in zip(itens, itens[1:]):
        if i2 <= f1:
            sobre.append(f"'{a1}' ({i1:%d/%m/%Y}–{f1:%d/%m/%Y}) e '{a2}' ({i2:%d/%m/%Y}–{f2:%d/%m/%Y}) cobrem datas em comum")
    avisos = []
    if n_repetidas:
        avisos.append(f"{formatar_numero(n_repetidas)} linha(s) aparecem iguais em mais de um arquivo"
                      + (" e foram retiradas, como você pediu." if remover_repetidas else
                         ". Elas foram mantidas; se forem o mesmo registro exportado duas vezes, os totais ficarão inflados."))
    if chaves_rep:
        avisos.append(f"{formatar_numero(chaves_rep)} código(s) de '{chave}' aparecem em mais de um arquivo. Confira se os "
                      "extratos não se sobrepõem.")
    avisos += [f"Períodos sobrepostos: {x}." for x in sobre]
    removidas = 0
    if remover_repetidas and n_repetidas:
        df = df[~entre_arquivos].reset_index(drop=True)
        removidas = n_repetidas
    return ResultadoEmpilhamento(df, [t.arquivo for t in tabelas], {t.arquivo: len(t.df) for t in tabelas}, n_repetidas,
                                 removidas, chave, chaves_rep, periodos, sobre, avisos)


def aplicar_empilhamento(modelo: "ModeloDados", nomes: list[str], remover_repetidas: bool = False) -> ResultadoEmpilhamento:
    """Substitui as tabelas `nomes` do modelo por UMA tabela empilhada (relacionamentos delas são removidos)."""
    tabs = [modelo.tabelas[n] for n in nomes]
    r = empilhar_tabelas(tabs, remover_repetidas)
    novo = nome_tabela(f"{nomes[0]}_empilhada", set(modelo.tabelas) - set(nomes))
    for n in nomes:
        modelo.tabelas.pop(n)
        modelo.colunas.pop(n, None)
    modelo.relacionamentos = [x for x in modelo.relacionamentos if x.origem not in nomes and x.destino not in nomes]
    modelo.tabelas[novo] = TabelaFonte(novo, " + ".join(r.arquivos), r.df, "(empilhado)", "—", None, list(r.avisos), None,
                                       "", list(r.arquivos))
    if modelo.fato in nomes or modelo.fato is None:
        modelo.fato = novo
    modelo.resultado = None
    return r


def dot_modelo(modelo: ModeloDados) -> str:
    """Diagrama do modelo (Graphviz DOT): tabelas como caixas (fato em verde-lima) e relacionamentos N:1 como setas."""
    lim = lambda s: str(s).replace('"', "'")
    linhas = ["digraph modelo {", "rankdir=LR;",
              f'node [shape=box, style="rounded,filled", fontname="Segoe UI", color="{COR_AZUL_PETROLEO}", fillcolor="{COR_BRANCO}", fontcolor="{COR_ESCURA}"];',
              f'edge [color="{COR_AZUL_PETROLEO}", fontname="Segoe UI", fontsize=10, fontcolor="{COR_TEXTO_SUAVE}"];']
    for nome, t in modelo.tabelas.items():
        cor = COR_LIMA if nome == modelo.fato else COR