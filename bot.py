# -*- coding: utf-8 -*-
import os
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, ContextTypes

from trecho_integracao_bot_revisado import conferencia_oc_handler, iniciar_conferencia_oc

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TOKEN")

ADMIN_IDS = {
    int(x.strip())
    for x in os.getenv("ADMIN_IDS", "").split(",")
    if x.strip().isdigit()
}


def usuario_autorizado(user_id: int) -> bool:
    if not ADMIN_IDS:
        return True
    return user_id in ADMIN_IDS


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if not usuario_autorizado(user.id):
        await update.message.reply_text(
            f"⛔ Usuário não autorizado.\n\nSeu ID é: {user.id}"
        )
        return

    teclado = InlineKeyboardMarkup([
        [InlineKeyboardButton("📦 Conferir Ordem de Compra", callback_data="conferir_oc")],
    ])

    await update.message.reply_text(
        "🤖 *Cálculo de Proteínas - Express*\n\n"
        "✅ Bot online e funcionando.\n\n"
        "Escolha uma opção abaixo:",
        parse_mode="Markdown",
        reply_markup=teclado
    )


async def meu_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Seu ID Telegram é: {update.effective_user.id}")


def main():
    if not TOKEN:
        raise RuntimeError("Variável BOT_TOKEN ou TOKEN não encontrada no Railway.")

    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("meuid", meu_id))
    app.add_handler(CommandHandler("conferir", iniciar_conferencia_oc))

    app.add_handler(conferencia_oc_handler)

    print("Bot Conferência OC iniciado com sucesso.")
    app.run_polling()


if __name__ == "__main__":
    main()
