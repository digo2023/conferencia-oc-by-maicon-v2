# -*- coding: utf-8 -*-
"""
Trecho de integração revisado para adicionar no bot.py existente.
Compatível com python-telegram-bot==20.7.
Usa o módulo conferencia_oc_revisado.py.
"""

import asyncio
import os
import traceback
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    CallbackQueryHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from conferencia_oc_revisado import ErroConferenciaOC, executar_conferencia_oc

AGUARDANDO_RELATORIO_NECESSARIO, AGUARDANDO_ORDEM_COMPRA = range(900, 902)


def menu_cancelar_oc():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_conferencia_oc")]
    ])


async def iniciar_conferencia_oc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Chamar pelo botão 📦 Conferir OC com callback_data='conferir_oc'."""
    context.user_data.pop("relatorio_necessario_pdf", None)
    context.user_data.pop("ordem_compra_pdf", None)

    texto = (
        "📦 *Conferência de Ordem de Compra*\n\n"
        "1️⃣ Envie agora o PDF do *relatório de consumo necessário* gerado pelo bot.\n"
        "2️⃣ Depois vou pedir o PDF da *Ordem de Compra*.\n"
        "3️⃣ No final vou gerar PDF de faltantes, PDF de sobras e Excel de conferência."
    )
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(texto, parse_mode="Markdown", reply_markup=menu_cancelar_oc())
    else:
        await update.message.reply_text(texto, parse_mode="Markdown", reply_markup=menu_cancelar_oc())
    return AGUARDANDO_RELATORIO_NECESSARIO


async def receber_relatorio_necessario(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name or not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ Envie um arquivo PDF válido do relatório de consumo necessário.")
        return AGUARDANDO_RELATORIO_NECESSARIO

    pasta = os.path.join("downloads_conferencia", str(update.effective_user.id))
    os.makedirs(pasta, exist_ok=True)
    arquivo = await doc.get_file()
    caminho = os.path.join(pasta, f"relatorio_necessario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    await arquivo.download_to_drive(caminho)
    context.user_data["relatorio_necessario_pdf"] = caminho

    await update.message.reply_text(
        "✅ Relatório de consumo necessário recebido.\n\n"
        "Agora envie o PDF da *Ordem de Compra* no modelo da matriz.",
        parse_mode="Markdown",
        reply_markup=menu_cancelar_oc(),
    )
    return AGUARDANDO_ORDEM_COMPRA


async def receber_ordem_compra(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name or not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ Envie um arquivo PDF válido da Ordem de Compra.")
        return AGUARDANDO_ORDEM_COMPRA

    relatorio_pdf = context.user_data.get("relatorio_necessario_pdf")
    if not relatorio_pdf or not os.path.exists(relatorio_pdf):
        await update.message.reply_text("⚠️ Não encontrei o relatório necessário. Comece novamente pela opção 📦 Conferir OC.")
        return ConversationHandler.END

    pasta = os.path.join("downloads_conferencia", str(update.effective_user.id))
    os.makedirs(pasta, exist_ok=True)
    arquivo = await doc.get_file()
    oc_pdf = os.path.join(pasta, f"ordem_compra_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    await arquivo.download_to_drive(oc_pdf)
    context.user_data["ordem_compra_pdf"] = oc_pdf

    msg = await update.message.reply_text("⏳ Conferindo consumo necessário x Ordem de Compra. Aguarde...")

    saida = os.path.join("relatorios_conferencia", str(update.effective_user.id), datetime.now().strftime("%Y%m%d_%H%M%S"))
    try:
        # Evita travar o loop assíncrono do Telegram enquanto processa PDF/Excel.
        resultado = await asyncio.to_thread(executar_conferencia_oc, relatorio_pdf, oc_pdf, saida)
    except ErroConferenciaOC as exc:
        await msg.edit_text(f"❌ Não foi possível concluir a conferência:\n{exc}")
        return ConversationHandler.END
    except Exception:
        traceback.print_exc()
        await msg.edit_text("❌ Erro inesperado durante a conferência. Verifique os PDFs e tente novamente.")
        return ConversationHandler.END

    await msg.edit_text(
        "✅ Conferência concluída.\n\n"
        f"🔴 Itens faltantes: {resultado['faltantes']}\n"
        f"🔵 Itens com sobra/excedente: {resultado['sobras']}\n"
        f"📄 Linhas analisadas: {resultado['total_linhas']}"
    )

    for caminho in [resultado["pdf_faltantes"], resultado["pdf_sobras"], resultado["xlsx"]]:
        if os.path.exists(caminho):
            with open(caminho, "rb") as f:
                await update.message.reply_document(document=f, filename=os.path.basename(caminho))

    context.user_data.pop("relatorio_necessario_pdf", None)
    context.user_data.pop("ordem_compra_pdf", None)
    return ConversationHandler.END


async def cancelar_conferencia_oc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("relatorio_necessario_pdf", None)
    context.user_data.pop("ordem_compra_pdf", None)
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text("Conferência de OC cancelada.")
    else:
        await update.message.reply_text("Conferência de OC cancelada.")
    return ConversationHandler.END


conferencia_oc_handler = ConversationHandler(
    entry_points=[CallbackQueryHandler(iniciar_conferencia_oc, pattern="^conferir_oc$")],
    states={
        AGUARDANDO_RELATORIO_NECESSARIO: [MessageHandler(filters.Document.PDF, receber_relatorio_necessario)],
        AGUARDANDO_ORDEM_COMPRA: [MessageHandler(filters.Document.PDF, receber_ordem_compra)],
    },
    fallbacks=[
        CallbackQueryHandler(cancelar_conferencia_oc, pattern="^cancelar_conferencia_oc$"),
        MessageHandler(filters.Regex("^(Cancelar|❌ Cancelar)$"), cancelar_conferencia_oc),
    ],
    allow_reentry=True,
)

# No main/application do bot.py, adicione:
# application.add_handler(conferencia_oc_handler)
#
# Botão recomendado no menu principal:
# InlineKeyboardButton("📦 Conferir OC", callback_data="conferir_oc")
