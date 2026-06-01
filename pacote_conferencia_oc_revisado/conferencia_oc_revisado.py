# -*- coding: utf-8 -*-
"""
Módulo: Conferência de Ordem de Compra x Consumo Necessário
Projeto: EXPRESS - Bot Telegram de Cálculo de Proteínas / Insumos
Versão revisada: robusta, auditável e pronta para integração
By Maicon

Dependências:
    pdfplumber==0.11.4
    reportlab==4.2.2
    pandas==2.2.2
    openpyxl==3.1.5

Função principal:
    executar_conferencia_oc(relatorio_pdf, ordem_compra_pdf, pasta_saida, logo_path=None)

Saídas:
    - relatorio_itens_faltantes.pdf
    - relatorio_itens_sobras.pdf
    - conferencia_oc.xlsx
"""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, asdict
from datetime import datetime
from difflib import SequenceMatcher
from typing import Dict, Iterable, List, Optional, Tuple
from copy import copy

import pandas as pd
import pdfplumber
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

UNIDADES = {"KG", "LT", "UN", "PC", "CX"}
TOLERANCIA = 0.001
SCORE_MINIMO_SEGURO = 0.78

TOKENS_GENERICOS = {
    "CARNE", "BOVINA", "BOVINO", "SUINA", "SUINO", "FRANGO", "DE", "DA", "DO", "DAS", "DOS",
    "BIFE", "FILE", "FATIADO", "CONGELADO", "SECA", "SECO", "BRANCO", "BRANCA", "MEDIA", "MEDIO",
    "KG", "UN", "LT", "PC", "CX", "X", "COM", "PARA", "EM", "AO", "A", "O", "ALIMENTICIA",
    "MISTURA", "PO", "PLASTICO", "DESCARTAVEL", "MASSA", "SUCO", "BRASSUCO", "MOLHO", "TEMPERO",
}

# Sinônimos curtos e seguros usados nos PDFs da Teknisa/relatório.
SINONIMOS = {
    "DIAFRAG": "DIAFRAGMA",
    "TORTILHO": "TORTILHONE",
    "CANJIQUINHA": "CANJICA",
    "ALVEJADO": "ALVEJADO",
    "DESCARTAVEL": "DESCARTAVEL",
    "DESIDRATADA": "DESIDRATADO",
    "PEROLA": "PEROLA",
}


class ErroConferenciaOC(Exception):
    """Erro controlado da conferência de Ordem de Compra."""


@dataclass
class ItemNecessario:
    item: str
    un: str
    necessario: float
    categoria: str = ""


@dataclass
class ItemOC:
    codigo: str
    produto_original: str
    produto_base: str
    un_original: str
    quantidade_original: float
    un_convertida: str
    quantidade_convertida: float
    entrega: str
    utilizacao: str
    observacao_conversao: str = ""


@dataclass
class LinhaConferencia:
    produto: str
    unidade: str
    necessario: float
    comprado: float
    diferenca: float
    data_utilizacao: str
    observacao: str
    produto_oc: str = ""
    score: float = 0.0
    status: str = ""
    entrega: str = ""
    categoria: str = ""


def br_float(valor: object) -> float:
    """Converte números brasileiros: 4.873,400 -> 4873.4 / 166,860 -> 166.86."""
    if valor is None:
        return 0.0
    s = str(valor).strip().replace(" ", "")
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_qtd(valor: float) -> str:
    return f"{float(valor):,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")


def sem_acento(txt: object) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", str(txt or ""))
        if unicodedata.category(c) != "Mn"
    )


def compactar_porcionamento(s: str) -> str:
    """Transforma '90 G' em '90G' para não confundir hambúrguer 90G com 56G."""
    return re.sub(r"\b(\d+[,.]?\d*)\s*G\b", lambda m: m.group(1).replace(",", ".") + "G", s)


def aplicar_sinonimos(s: str) -> str:
    for origem, destino in SINONIMOS.items():
        s = re.sub(rf"\b{re.escape(origem)}\b", destino, s)
    return s


def limpar_texto_produto(txt: object, remover_embalagem: bool = True) -> str:
    s = sem_acento(txt).upper()
    s = compactar_porcionamento(s)
    if remover_embalagem:
        # Remove somente descrições de embalagem, preservando porcionamentos do produto como 90G/56G.
        s = re.sub(r"\b\d+[,.]?\d*\s*KG\s*X\s*(CX|PC)\b", " ", s)
        s = re.sub(r"\b\d+\s*UN\s*X\s*\d+\s*BJ\s*X\s*CX\b", " ", s)
        s = re.sub(r"\b\d+\s*UN\s*X\s*CX\b", " ", s)
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = aplicar_sinonimos(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalizar_nome(txt: object) -> str:
    return limpar_texto_produto(txt, remover_embalagem=True)


def tokens_relevantes(txt: object) -> List[str]:
    s = normalizar_nome(txt)
    return [t for t in s.split() if t not in TOKENS_GENERICOS and len(t) > 1]


def porcionamentos(txt: object) -> List[str]:
    return re.findall(r"\b\d+(?:\.\d+)?G\b", normalizar_nome(txt))


def extrair_texto_pdf(caminho_pdf: str) -> str:
    if not caminho_pdf or not os.path.exists(caminho_pdf):
        raise ErroConferenciaOC(f"Arquivo PDF não encontrado: {caminho_pdf}")
    partes: List[str] = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                partes.append(page.extract_text(x_tolerance=1, y_tolerance=3) or "")
    except Exception as exc:
        raise ErroConferenciaOC(f"Não foi possível ler o PDF: {os.path.basename(caminho_pdf)}. Erro: {exc}") from exc
    texto = "\n".join(partes).strip()
    if not texto:
        raise ErroConferenciaOC(f"O PDF não retornou texto legível: {os.path.basename(caminho_pdf)}")
    return texto


def extrair_periodo_relatorio(texto: str) -> str:
    m = re.search(r"Per[ií]odo conferido\s+(\d{2}/\d{2}/\d{4}\s+a\s+\d{2}/\d{2}/\d{4})", texto, flags=re.I)
    if m:
        return m.group(1)
    m = re.search(r"RELAT[ÓO]RIO.*?(\d{2}/\d{2}/\d{4})\s+A\s+(\d{2}/\d{2}/\d{4})", texto, flags=re.I | re.S)
    if m:
        return f"{m.group(1)} a {m.group(2)}"
    return "Período não identificado"


def extrair_periodo_oc(texto: str) -> str:
    m = re.search(r"Per[ií]odo de Entrega\s+(\d{2}/\d{2}/\d{4}\s+a\s+\d{2}/\d{2}/\d{4})", texto, flags=re.I)
    return m.group(1) if m else "Período da OC não identificado"


def extrair_itens_necessarios(relatorio_pdf: str) -> Tuple[List[ItemNecessario], str]:
    """Extrai a lista consolidada por categoria do relatório gerado pelo bot."""
    texto = extrair_texto_pdf(relatorio_pdf)
    periodo = extrair_periodo_relatorio(texto)
    marcador = "3. LISTA COMPLETA CONSOLIDADA POR CATEGORIA"
    if marcador not in texto:
        raise ErroConferenciaOC("O relatório não possui a seção '3. LISTA COMPLETA CONSOLIDADA POR CATEGORIA'.")

    bloco = texto.split(marcador, 1)[1]
    itens: List[ItemNecessario] = []
    categoria = ""

    for linha in bloco.splitlines():
        linha = linha.strip()
        if not linha:
            continue
        cat = re.match(r"^(PROTE[IÍ]NAS|N[ÃA]O PEREC[IÍ]VEIS|HORTIFRUTI)\b", linha, flags=re.I)
        if cat:
            categoria = sem_acento(cat.group(1)).upper().replace("NAO", "NÃO")
            continue
        if re.match(r"^(ITEM\s+UN\s+TOTAL|OBSERVA|CONFER[ÊE]NCIA|P[ÁA]GINA|RELAT[ÓO]RIO)", linha, flags=re.I):
            continue
        # Nome + Unidade + Total no final. Evita capturar linhas quebradas ou rodapé.
        m = re.match(r"^(.+?)\s+(KG|LT|UN|PC|CX)\s+([\d\.]+,\d{1,3}|\d+,\d{1,3}|\d+)$", linha, flags=re.I)
        if not m:
            continue
        nome, un, qtd = m.groups()
        nome = re.sub(r"\s+", " ", nome.strip().upper())
        if len(nome) >= 3 and categoria:
            itens.append(ItemNecessario(nome, un.upper(), br_float(qtd), categoria))

    if not itens:
        raise ErroConferenciaOC("Nenhum item necessário foi extraído do relatório. Confira se o PDF é o relatório correto.")

    consol: Dict[Tuple[str, str, str], float] = {}
    for it in itens:
        chave = (it.item, it.un, it.categoria)
        consol[chave] = consol.get(chave, 0.0) + it.necessario

    return [ItemNecessario(k[0], k[1], round(v, 3), k[2]) for k, v in consol.items()], periodo


def converter_pacote(produto: str, un: str, qtd: float) -> Tuple[str, float, str]:
    """
    Converte compra em CX/PC para unidade real quando o nome informa embalagem.
    Exemplos:
    - 8 KG X CX, CX 6 -> KG 48
    - 2 KG X PC, PC 9 -> KG 18
    - 30 UN X 12 BJ X CX, CX 14 -> UN 5040
    """
    p = sem_acento(produto).upper()
    un = un.upper()

    m = re.search(r"(\d+[,.]?\d*)\s*KG\s*X\s*(CX|PC)", p)
    if un in {"CX", "PC"} and m:
        fator = br_float(m.group(1))
        return "KG", round(qtd * fator, 3), f"Convertido de {fmt_qtd(qtd)} {un} x {fmt_qtd(fator)} KG."

    m = re.search(r"(\d+)\s*UN\s*X\s*(\d+)\s*BJ\s*X\s*CX", p)
    if un == "CX" and m:
        fator = int(m.group(1)) * int(m.group(2))
        return "UN", round(qtd * fator, 3), f"Convertido de {fmt_qtd(qtd)} CX x {fator} UN."

    m = re.search(r"(\d+[,.]?\d*)\s*UN\s*X\s*CX", p)
    if un == "CX" and m:
        fator = br_float(m.group(1))
        return "UN", round(qtd * fator, 3), f"Convertido de {fmt_qtd(qtd)} CX x {fmt_qtd(fator)} UN."

    return un, round(qtd, 3), ""


def limpar_nome_oc(produto: str) -> str:
    return limpar_texto_produto(produto, remover_embalagem=True)


def extrair_itens_oc(oc_pdf: str) -> Tuple[List[ItemOC], str]:
    texto = extrair_texto_pdf(oc_pdf)
    periodo_oc = extrair_periodo_oc(texto)
    itens: List[ItemOC] = []

    # Regex tolerante a espaçamentos e linhas iniciando por código Teknisa.
    padrao = re.compile(
        r"^(\d(?:\.\d+)+)\s+(.+?)\s+(KG|LT|UN|PC|CX)\s+([\d\.]+,\d{1,3}|\d+,\d{1,3}|\d+)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})$",
        flags=re.I,
    )

    for linha in texto.splitlines():
        linha = re.sub(r"\s+", " ", linha.strip())
        if not linha or not re.match(r"^\d(?:\.\d+)+\s+", linha):
            continue
        m = padrao.match(linha)
        if not m:
            continue
        codigo, produto, un, qtd, entrega, utilizacao = m.groups()
        qtd_float = br_float(qtd)
        un_conv, qtd_conv, obs_conv = converter_pacote(produto, un, qtd_float)
        itens.append(ItemOC(
            codigo=codigo,
            produto_original=produto.strip().upper(),
            produto_base=limpar_nome_oc(produto),
            un_original=un.upper(),
            quantidade_original=qtd_float,
            un_convertida=un_conv,
            quantidade_convertida=qtd_conv,
            entrega=entrega,
            utilizacao=utilizacao,
            observacao_conversao=obs_conv,
        ))

    if not itens:
        raise ErroConferenciaOC("Nenhum item foi extraído da Ordem de Compra. Confira se o PDF está no modelo da matriz.")
    return itens, periodo_oc


def produto_compatível(a: str, b: str) -> Tuple[bool, float, str]:
    """Valida se dois nomes parecem ser o mesmo produto, reduzindo falsos positivos."""
    na, nb = normalizar_nome(a), normalizar_nome(b)
    if not na or not nb:
        return False, 0.0, "nome vazio"

    pa, pb = set(porcionamentos(na)), set(porcionamentos(nb))
    if pa and pb and pa != pb:
        return False, 0.0, f"porcionamento diferente ({'/'.join(sorted(pa))} x {'/'.join(sorted(pb))})"

    if na == nb:
        return True, 1.0, "match exato"

    ta, tb = set(tokens_relevantes(na)), set(tokens_relevantes(nb))
    if ta and tb:
        inter = ta & tb
        cobertura_menor = len(inter) / max(1, min(len(ta), len(tb)))
        cobertura_maior = len(inter) / max(1, max(len(ta), len(tb)))
        # Rejeita produtos que só compartilham termos genéricos removidos, ex.: BISTECA x ALCATRA.
        if cobertura_menor < 0.55 or cobertura_maior < 0.35:
            return False, 0.0, "tokens principais diferentes"
    else:
        cobertura_menor = 0.0

    seq = SequenceMatcher(None, na, nb).ratio()
    overlap = len(set(na.split()) & set(nb.split())) / max(1, min(len(na.split()), len(nb.split())))
    score = (seq * 0.50) + (overlap * 0.25) + (cobertura_menor * 0.25)
    return score >= SCORE_MINIMO_SEGURO, round(score, 4), "match por similaridade" if score >= SCORE_MINIMO_SEGURO else "similaridade baixa"


def agrupar_oc(ocs: Iterable[ItemOC]) -> List[ItemOC]:
    """Soma itens repetidos da OC após conversão de unidade."""
    grupos: Dict[Tuple[str, str], ItemOC] = {}
    for oc in ocs:
        chave = (normalizar_nome(oc.produto_base), oc.un_convertida)
        if chave not in grupos:
            grupos[chave] = oc
            continue
        base = grupos[chave]
        base.quantidade_convertida = round(base.quantidade_convertida + oc.quantidade_convertida, 3)
        base.quantidade_original = round(base.quantidade_original + oc.quantidade_original, 3)
        base.entrega = ", ".join(sorted(set((base.entrega + "," + oc.entrega).split(","))))
        base.utilizacao = ", ".join(sorted(set((base.utilizacao + "," + oc.utilizacao).split(","))))
        if oc.observacao_conversao and oc.observacao_conversao not in base.observacao_conversao:
            base.observacao_conversao = (base.observacao_conversao + " " + oc.observacao_conversao).strip()
    return list(grupos.values())


def encontrar_oc(item: ItemNecessario, ocs: List[ItemOC], usados: set) -> Tuple[Optional[int], float, str]:
    melhor_i: Optional[int] = None
    melhor_score = 0.0
    melhor_motivo = "não localizado"
    for i, oc in enumerate(ocs):
        if i in usados:
            continue
        if item.un != oc.un_convertida:
            continue
        ok, score, motivo = produto_compatível(item.item, oc.produto_base)
        if ok and score > melhor_score:
            melhor_score = score
            melhor_i = i
            melhor_motivo = motivo
    return melhor_i, melhor_score, melhor_motivo


def montar_observacao_base(oc: Optional[ItemOC], score: float, motivo: str) -> str:
    obs = []
    if motivo:
        obs.append(f"Critério: {motivo}.")
    if score and score < 0.90:
        obs.append(f"Conferir nome/modelo manualmente; similaridade {int(score * 100)}%.")
    if oc and oc.observacao_conversao:
        obs.append(oc.observacao_conversao)
    return " ".join(obs).strip()


def conferir_necessario_x_oc(relatorio_pdf: str, oc_pdf: str) -> Tuple[List[LinhaConferencia], List[LinhaConferencia], List[dict], str, str]:
    necessarios, periodo_relatorio = extrair_itens_necessarios(relatorio_pdf)
    ocs_raw, periodo_oc = extrair_itens_oc(oc_pdf)
    ocs = agrupar_oc(ocs_raw)
    usados: set = set()
    faltantes: List[LinhaConferencia] = []
    sobras: List[LinhaConferencia] = []
    geral: List[dict] = []

    for nec in necessarios:
        idx, score, motivo = encontrar_oc(nec, ocs, usados)
        if idx is None:
            linha = LinhaConferencia(
                produto=nec.item, unidade=nec.un, necessario=nec.necessario, comprado=0.0,
                diferenca=round(nec.necessario, 3), data_utilizacao="-", entrega="-", categoria=nec.categoria,
                observacao="Produto necessário no relatório e não localizado de forma segura na Ordem de Compra.",
                score=score, status="FALTANTE",
            )
            faltantes.append(linha)
            geral.append(asdict(linha))
            continue

        oc = ocs[idx]
        usados.add(idx)
        diff = round(nec.necessario - oc.quantidade_convertida, 3)
        obs_base = montar_observacao_base(oc, score, motivo)

        if diff > TOLERANCIA:
            status = "FALTANTE"
            obs = f"FALTA: quantidade comprada/solicitada não atende ao consumo necessário. {obs_base}".strip()
            linha = LinhaConferencia(nec.item, nec.un, nec.necessario, oc.quantidade_convertida, diff, oc.utilizacao, obs,
                                      produto_oc=oc.produto_original, score=score, status=status, entrega=oc.entrega, categoria=nec.categoria)
            faltantes.append(linha)
        elif diff < -TOLERANCIA:
            status = "SOBRA"
            obs = f"SOBRA/EXCEDENTE: quantidade comprada/solicitada acima do consumo necessário. {obs_base}".strip()
            linha = LinhaConferencia(nec.item, nec.un, nec.necessario, oc.quantidade_convertida, abs(diff), oc.utilizacao, obs,
                                      produto_oc=oc.produto_original, score=score, status=status, entrega=oc.entrega, categoria=nec.categoria)
            sobras.append(linha)
        else:
            status = "OK"
            obs = f"OK: quantidade comprada/solicitada atende exatamente ao consumo necessário. {obs_base}".strip()
            linha = LinhaConferencia(nec.item, nec.un, nec.necessario, oc.quantidade_convertida, 0.0, oc.utilizacao, obs,
                                      produto_oc=oc.produto_original, score=score, status=status, entrega=oc.entrega, categoria=nec.categoria)
        geral.append(asdict(linha))

    for i, oc in enumerate(ocs):
        if i not in usados:
            linha = LinhaConferencia(
                produto=oc.produto_base, unidade=oc.un_convertida, necessario=0.0,
                comprado=oc.quantidade_convertida, diferenca=oc.quantidade_convertida,
                data_utilizacao=oc.utilizacao, entrega=oc.entrega,
                observacao="Produto consta na Ordem de Compra, mas não foi localizado no consumo necessário do relatório.",
                produto_oc=oc.produto_original, score=0.0, status="SOBRA/NAO PREVISTO", categoria="OC"
            )
            sobras.append(linha)
            geral.append(asdict(linha))

    faltantes.sort(key=lambda x: (x.categoria, -x.diferenca, x.produto))
    sobras.sort(key=lambda x: (x.status, -x.diferenca, x.produto))
    return faltantes, sobras, geral, periodo_relatorio, periodo_oc


def safe_paragraph(txt: object, style: ParagraphStyle) -> Paragraph:
    s = str(txt or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(s, style)


def gerar_pdf_conferencia(
    caminho_pdf: str,
    titulo: str,
    subtitulo: str,
    linhas: List[LinhaConferencia],
    tipo: str,
    logo_path: Optional[str] = None,
) -> None:
    doc = SimpleDocTemplate(
        caminho_pdf,
        pagesize=landscape(A4),
        leftMargin=0.7 * cm,
        rightMargin=0.7 * cm,
        topMargin=0.6 * cm,
        bottomMargin=0.7 * cm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Titulo", parent=styles["Title"], alignment=TA_CENTER, fontSize=16, leading=19, spaceAfter=5)
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"], alignment=TA_CENTER, fontSize=8.5, leading=11, spaceAfter=8)
    cell_style = ParagraphStyle("Cel", parent=styles["Normal"], fontSize=6.5, leading=7.8, alignment=TA_LEFT)
    head_style = ParagraphStyle("Cab", parent=styles["Normal"], fontSize=6.8, leading=8, textColor=colors.white, alignment=TA_CENTER)

    story: List[object] = []
    if logo_path and os.path.exists(logo_path):
        try:
            story.append(Image(logo_path, width=2.2 * cm, height=2.2 * cm))
        except Exception:
            pass
    story.append(Paragraph(titulo, title_style))
    story.append(Paragraph(subtitulo, sub_style))
    story.append(Spacer(1, 0.08 * cm))

    if not linhas:
        story.append(Paragraph("Nenhuma divergência encontrada.", styles["Normal"]))
    else:
        dados = [[
            safe_paragraph("Nº", head_style), safe_paragraph("PRODUTO", head_style), safe_paragraph("UN", head_style),
            safe_paragraph("NECESSÁRIO", head_style), safe_paragraph("OC/COMPRADO", head_style),
            safe_paragraph("DIFERENÇA", head_style), safe_paragraph("ENTREGA", head_style),
            safe_paragraph("UTILIZAÇÃO", head_style), safe_paragraph("OBSERVAÇÃO", head_style),
        ]]
        for n, l in enumerate(linhas, 1):
            obs = l.observacao
            if l.produto_oc and normalizar_nome(l.produto) != normalizar_nome(l.produto_oc):
                obs += f" Produto na OC: {l.produto_oc}."
            dados.append([
                str(n), safe_paragraph(l.produto, cell_style), l.unidade,
                fmt_qtd(l.necessario), fmt_qtd(l.comprado), fmt_qtd(l.diferenca),
                l.entrega or "-", l.data_utilizacao or "-", safe_paragraph(obs, cell_style),
            ])

        cor_cab = colors.HexColor("#8B0000") if tipo == "faltantes" else colors.HexColor("#1F4E79")
        tabela = Table(
            dados,
            colWidths=[0.65 * cm, 6.2 * cm, 0.8 * cm, 1.8 * cm, 1.9 * cm, 1.8 * cm, 1.9 * cm, 2.0 * cm, 10.2 * cm],
            repeatRows=1,
        )
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), cor_cab),
            ("GRID", (0, 0), (-1, -1), 0.30, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 1), (-1, -1), 6.5),
            ("ALIGN", (0, 1), (0, -1), "CENTER"),
            ("ALIGN", (2, 1), (7, -1), "CENTER"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6F7")]),
        ]))
        story.append(tabela)

    def rodape(canvas, doc_obj):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        data = datetime.now().strftime("%d/%m/%Y")
        canvas.drawRightString(28.8 * cm, 0.42 * cm, f"{data} - By Maicon")
        canvas.drawString(0.7 * cm, 0.42 * cm, f"{titulo} - Página {doc_obj.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=rodape, onLaterPages=rodape)


def gerar_excel_conferencia(caminho_xlsx: str, geral: List[dict]) -> None:
    colunas = [
        "status", "categoria", "produto", "unidade", "necessario", "comprado", "diferenca",
        "entrega", "data_utilizacao", "produto_oc", "score", "observacao",
    ]
    df = pd.DataFrame(geral)
    for c in colunas:
        if c not in df.columns:
            df[c] = ""
    df = df[colunas]
    with pd.ExcelWriter(caminho_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Conferencia", index=False)
        resumo = df.groupby("status", dropna=False).agg(
            itens=("produto", "count"),
            total_diferenca=("diferenca", "sum"),
        ).reset_index()
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                fonte = copy(cell.font)
                fonte.bold = True
                cell.font = fonte
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col[:100])
                width = min(max(max_len + 2, 10), 45)
                ws.column_dimensions[col[0].column_letter].width = width


def executar_conferencia_oc(
    relatorio_pdf: str,
    ordem_compra_pdf: str,
    pasta_saida: str = ".",
    logo_path: Optional[str] = None,
) -> Dict[str, str]:
    os.makedirs(pasta_saida, exist_ok=True)
    faltantes, sobras, geral, periodo_relatorio, periodo_oc = conferir_necessario_x_oc(relatorio_pdf, ordem_compra_pdf)

    pdf_faltantes = os.path.join(pasta_saida, "relatorio_itens_faltantes.pdf")
    pdf_sobras = os.path.join(pasta_saida, "relatorio_itens_sobras.pdf")
    xlsx = os.path.join(pasta_saida, "conferencia_oc.xlsx")

    subtitulo = (
        "Conferência automática: Consumo necessário x Ordem de Compra | "
        f"Período do relatório: {periodo_relatorio} | Período de entrega da OC: {periodo_oc}"
    )
    gerar_pdf_conferencia(pdf_faltantes, "RELATÓRIO DE ITENS FALTANTES", subtitulo, faltantes, "faltantes", logo_path)
    gerar_pdf_conferencia(pdf_sobras, "RELATÓRIO DE ITENS COM SOBRA / EXCEDENTE", subtitulo, sobras, "sobras", logo_path)
    gerar_excel_conferencia(xlsx, geral)

    return {
        "pdf_faltantes": pdf_faltantes,
        "pdf_sobras": pdf_sobras,
        "xlsx": xlsx,
        "faltantes": str(len(faltantes)),
        "sobras": str(len(sobras)),
        "total_linhas": str(len(geral)),
        "periodo_relatorio": periodo_relatorio,
        "periodo_oc": periodo_oc,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Conferir relatório de consumo necessário x Ordem de Compra")
    parser.add_argument("relatorio_pdf")
    parser.add_argument("ordem_compra_pdf")
    parser.add_argument("--saida", default="saida_conferencia")
    parser.add_argument("--logo", default=None)
    args = parser.parse_args()
    print(executar_conferencia_oc(args.relatorio_pdf, args.ordem_compra_pdf, args.saida, args.logo))
