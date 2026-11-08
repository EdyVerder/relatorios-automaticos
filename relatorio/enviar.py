"""Monta a mensagem e envia por SMTP. Em modo simulação só grava o .eml (abre no Outlook) e a prévia em HTML."""
import os
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

XLSX = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def mensagem(assunto, html, para, copia, anexo, nome_anexo, remetente):
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = ", ".join(para)
    if copia:
        msg["Cc"] = ", ".join(copia)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    msg.set_content("Este e-mail precisa de um leitor com suporte a HTML. A planilha completa está em anexo.")
    msg.add_alternative(html, subtype="html")
    msg.add_attachment(anexo, maintype=XLSX[0], subtype=XLSX[1], filename=nome_anexo)
    return msg


def enviar(msg):
    """Credenciais só por variável de ambiente (ou .env carregado antes); nunca no código nem no log."""
    host, porta = os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORTA", "587"))
    with smtplib.SMTP(host, porta, timeout=30) as smtp:
        smtp.starttls(context=ssl.create_default_context())
        smtp.login(os.environ["SMTP_USUARIO"], os.environ["SMTP_SENHA"])
        smtp.send_message(msg)
