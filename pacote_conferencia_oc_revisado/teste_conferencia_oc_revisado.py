# -*- coding: utf-8 -*-
"""Teste local rápido da função Conferência OC."""

import os
import sys
import pandas as pd

from conferencia_oc_revisado import executar_conferencia_oc, extrair_itens_necessarios, extrair_itens_oc


def main():
    if len(sys.argv) < 3:
        print("Uso: python teste_conferencia_oc_revisado.py RELATORIO.pdf ORDEM_COMPRA.pdf")
        raise SystemExit(1)

    relatorio_pdf = sys.argv[1]
    oc_pdf = sys.argv[2]
    saida = sys.argv[3] if len(sys.argv) >= 4 else "saida_teste_conferencia"

    necessarios, periodo = extrair_itens_necessarios(relatorio_pdf)
    itens_oc, periodo_oc = extrair_itens_oc(oc_pdf)
    print(f"Itens extraídos do relatório: {len(necessarios)} | Período: {periodo}")
    print(f"Itens extraídos da OC: {len(itens_oc)} | Período OC: {periodo_oc}")

    resultado = executar_conferencia_oc(relatorio_pdf, oc_pdf, saida)
    print(resultado)

    df = pd.read_excel(resultado["xlsx"], sheet_name="Conferencia")
    print(df["status"].value_counts(dropna=False))

    obrigatorios = [resultado["pdf_faltantes"], resultado["pdf_sobras"], resultado["xlsx"]]
    for caminho in obrigatorios:
        if not os.path.exists(caminho) or os.path.getsize(caminho) <= 0:
            raise RuntimeError(f"Arquivo de saída inválido: {caminho}")

    print("TESTE OK: arquivos gerados com sucesso.")


if __name__ == "__main__":
    main()
