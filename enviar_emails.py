import os
import smtplib
import time
import mimetypes
from datetime import datetime
import tkinter as tk
from email.message import EmailMessage

from openpyxl import load_workbook

from dotenv import load_dotenv
load_dotenv()

#Configurações do programa
REMETENTE = os.environ["GMAIL_USER"]
SENHA = os.environ["GMAIL_APP_PASSWORD"]
PLANILHA_FUNCIONARIOS = "funcionarios.xlsx"
ANEXO = "painel_vendas.xlsx"
ASSUNTO = "Dashboard de vendas da papelaria - 2024"
PAUSA_ENTRE_ENVIOS = 2

MENSAGEM = """Olá, {nome}!

Segue o dashboard de vendas da papelaria referente a 2024, para análise
de insights do setor de {setor}.

Atenciosamente,
Antonio Carlos
"""

#Função que lê todos os contatos da planilha e salva em um dicionário
def ler_contatos(caminho):
    planilha = load_workbook(caminho)
    aba = planilha.active
    contatos = []
    for nome, email, setor in aba.iter_rows(min_row=2, values_only=True):
        if email:
            contatos.append({"nome": nome, "email": email, "setor": setor})
    return contatos

#Monta o email com as informações que configuramos antes e as novas
def montar_email(contato):
    msg = EmailMessage()
    msg["From"] = REMETENTE
    msg["To"] = contato["email"]
    msg["Subject"] = ASSUNTO
    msg.set_content(MENSAGEM.format(nome=contato["nome"], setor=contato["setor"]))

    if ANEXO:
        tipo, _ = mimetypes.guess_type(ANEXO)
        tipo = tipo or "application/octet-stream"
        maintype, subtype = tipo.split("/")
        with open(ANEXO, "rb") as arquivo:
            msg.add_attachment(
                arquivo.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(ANEXO),
            )
    return msg

#Escolher para qual(is) setor(es) o email será enviado
def escolher_setor(funcionarios):
    setores = sorted({f["setor"] for f in funcionarios if f["setor"]})
    resultado = {"setor": None, "confirmado": False}

    janela = tk.Tk()
    janela.title("Enviar e-mails")

    tk.Label(janela, text="Para qual setor deseja enviar?").pack(padx=20, pady=10)

    escolha = tk.StringVar(value="Todos")
    for opcao in ["Todos"] + setores:
        tk.Radiobutton(janela, text=opcao, variable=escolha, value=opcao).pack(anchor="w", padx=30)

    def confirmar():
        resultado["setor"] = None if escolha.get() == "Todos" else escolha.get()
        resultado["confirmado"] = True
        janela.destroy()

    tk.Button(janela, text="Enviar", command=confirmar).pack(pady=15)

    janela.eval('tk::PlaceWindow . center')
    janela.mainloop()

    return resultado

def main():
    funcionarios = ler_contatos(PLANILHA_FUNCIONARIOS)
    print(f"{len(funcionarios)} contatos encontrados.")

    #Escolher para qual setor vai ser enviado o e-mail
    resultado = escolher_setor(funcionarios)
    if not resultado["confirmado"]:
        print("Envio cancelado.")
        return

    #Filtrar a lista
    if resultado["setor"]:
        funcionarios = [f for f in funcionarios if f["setor"] == resultado["setor"]]

    tipo_envio = resultado["setor"] or "Todos os setores"

    enviados, falhas = 0, []

    planilha = load_workbook(PLANILHA_FUNCIONARIOS)

    #Verifica se a aba 'Registros' já existe
    if "Registros" in planilha.sheetnames:
        aba_registro = planilha["Registros"]
    else:
        aba_registro = planilha.create_sheet("Registros")
        aba_registro.append(["Nome", "Email", "Tipo de envio", "Data", "Status"])

    #Conexão única para todos os envios
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
        servidor.login(REMETENTE, SENHA)

        #Tenta enviar o email para todos os funcionários
        for funcionario in funcionarios:
            try:
                servidor.send_message(montar_email(funcionario))
                enviados += 1
                print(f"Enviado para {funcionario['nome']} <{funcionario['email']}>")
                status = 'Enviado com sucesso'
            except Exception as erro:
                falhas.append(funcionario["email"])
                print(f"✘ Falha com {funcionario['email']}: {erro}")
                status = 'Falha no envio'

            #Salva o resultado na aba Registros
            data = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            aba_registro.append([funcionario['nome'], funcionario['email'], tipo_envio, data, status])
            time.sleep(PAUSA_ENTRE_ENVIOS)

    planilha.save(PLANILHA_FUNCIONARIOS)

    print(f"\nConcluído: {enviados} enviados, {len(falhas)} falhas.")
    if falhas:
        print("Falharam:", ", ".join(falhas))


if __name__ == "__main__":
    main()