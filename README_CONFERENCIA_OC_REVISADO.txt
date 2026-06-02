FUNÇÃO REVISADA - CONFERÊNCIA DE ORDEM DE COMPRA
Projeto: EXPRESS - Bot Telegram de Cálculo de Proteínas / Insumos
By Maicon

ARQUIVOS PRINCIPAIS
1) conferencia_oc_revisado.py
   - Módulo principal revisado.
   - Lê o relatório de consumo necessário gerado pelo bot.
   - Lê a Ordem de Compra no modelo da matriz.
   - Confere quantidade necessária x quantidade comprada/solicitada.
   - Gera PDF de itens faltantes, PDF de sobras/excedentes e Excel de conferência.

2) trecho_integracao_bot_revisado.py
   - Trecho pronto para integrar no bot.py.
   - Fluxo Telegram:
     📦 Conferir OC -> enviar relatório necessário -> enviar Ordem de Compra -> receber relatórios.
   - Possui cancelamento, tratamento de erro e processamento em thread para não travar o bot.

3) teste_conferencia_oc_revisado.py
   - Teste local rápido para validar extração, conferência e geração dos arquivos.

CORREÇÕES FEITAS NESTA REVISÃO
- Evita falso positivo entre produtos parecidos, exemplo:
  * CARNE BOVINA BIFE COXÃO DURO 90G não é mais confundida com CARNE BOVINA BIFE DE FÍGADO 90G.
  * CARNE SUÍNA BISTECA 100G não é mais confundida com CARNE SUÍNA ALCATRA.
  * HAMBÚRGUER 90G não é mais confundido com HAMBÚRGUER 56G.
- Preserva gramatura/porcionamento no nome do produto, como 90G, 56G e 100G.
- Soma itens repetidos na OC após conversão.
- Converte embalagens automaticamente quando o PDF informa fator:
  * 8 KG X CX
  * 2 KG X PC
  * 30 UN X 12 BJ X CX
- Valida arquivos de entrada e informa erro claro se o PDF não estiver legível.
- Melhora o Excel com aba Conferencia e aba Resumo.
- Melhora os PDFs com coluna Entrega, Utilização, Observação, rodapé e paginação.
- Se o produto da OC não tiver correspondente seguro no relatório, ele entra como SOBRA/NAO PREVISTO.

DEPENDÊNCIAS PARA requirements.txt
pdfplumber==0.11.4
reportlab==4.2.2
pandas==2.2.2
openpyxl==3.1.5
python-telegram-bot==20.7

COMANDO DE TESTE LOCAL
python teste_conferencia_oc_revisado.py "Relatorio .pdf" "rancho 1905_260506_233805(1).pdf" saida_teste_conferencia

INTEGRAÇÃO NO BOT.PY
1) Copie conferencia_oc_revisado.py para a raiz do projeto.
2) Copie o conteúdo de trecho_integracao_bot_revisado.py para o bot.py ou importe o handler.
3) Adicione no application:
   application.add_handler(conferencia_oc_handler)
4) Adicione no menu principal:
   InlineKeyboardButton("📦 Conferir OC", callback_data="conferir_oc")

OBSERVAÇÃO OPERACIONAL IMPORTANTE
O exemplo enviado compara um relatório de 07/05/2026 a 13/05/2026 com uma OC de utilização 21/05/2026 a 27/05/2026. O sistema confere por produto e quantidade, mas a conferência operacional perfeita exige que o relatório e a OC sejam da mesma semana/período.
