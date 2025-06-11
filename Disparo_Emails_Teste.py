# INFORMAÇÕES:
# Esse código considera uma planilha com uma aba para Dados (NOME | EMAIL | FUNCAO | LOCAL | Enviar) e uma para a mensagem em HTML (FUNCAO | TEXTO).
# Também utiliza uma pasta denominada "anexos" com subpastas com nome da função e os anexos desejados.

import pandas as pd
import os
import win32com.client as win32
from datetime import datetime

# === ⚙️ CONFIGURAÇÕES ===
ARQUIVO_PLANILHA = r"C:\PPT\EnvioEmails\BASE_GE_24_(Envio_Emails).xlsx"
ABA_DADOS = "Base_GE_2024"
ABA_HTML = "Corpo_HTML"
PASTA_ANEXOS = r"C:\PPT\EnvioEmails\anexos"  # <- caminho base dos anexos

# === 🔍 MODO DE EXECUÇÃO ===
ENVIAR = False

print("\n📤 Qual modo deseja executar?")
print("[1] Exibir os e-mails (modo seguro)")
print("[2] Enviar os e-mails (modo real)")

while True:
    modo = input("Digite 1 ou 2: ").strip()
    if modo == "1":
        ENVIAR = False
        print("🔒 Modo SEGURO selecionado: e-mails serão apenas exibidos.\n")
        break
    elif modo == "2":
        confirm = input("🚨 CONFIRMA ENVIO REAL DOS E-MAILS? (s para sim): ").strip().lower()
        if confirm == 's':
            ENVIAR = True
            print("📤 Modo ENVIO REAL confirmado!\n")
        else:
            print("❌ Envio cancelado. Continuando em modo Display.")
        break
    else:
        print("[X] Entrada inválida. Digite 1 ou 2.")

print(f"💡 EXECUTANDO EM: {'ENVIO REAL (Send)' if ENVIAR else 'DISPLAY (modo teste)'}\n")

# === 📊 LER PLANILHAS ===
df_dados = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_DADOS, engine='openpyxl')
df_html = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_HTML, engine='openpyxl')

# === LIMPEZA ===
df_dados['LOCAL'] = df_dados['LOCAL'].astype(str).str.strip()
df_dados['FUNCAO'] = df_dados['FUNCAO'].astype(str).str.strip().str.upper()
df_dados['Enviar'] = df_dados['Enviar'].fillna(False).astype(bool)

df_html['FUNCAO'] = df_html['FUNCAO'].astype(str).str.strip().str.upper()
df_html.set_index('FUNCAO', inplace=True)

# === LISTA DE LOCAIS E FUNÇÕES ===
locais_validos = sorted(df_dados['LOCAL'].unique())
funcoes_validas = sorted(df_dados['FUNCAO'].unique())

# === SELEÇÃO INTERATIVA ===
print("\n📍 Locais disponíveis:")
for i, local in enumerate(locais_validos, 1):
    print(f"[{i}] {local}")
while True:
    entrada = input("\nDigite o número do LOCAL desejado: ").strip()
    if entrada.isdigit() and 1 <= int(entrada) <= len(locais_validos):
        LOCAL_FILTRADO = locais_validos[int(entrada) - 1]
        break
    print("[X] Entrada inválida.")

print("\n👤 Funções disponíveis:")
for i, funcao in enumerate(funcoes_validas, 1):
    print(f"[{i}] {funcao}")
while True:
    entrada_funcao = input("\nDigite o número da FUNÇÃO desejada: ").strip()
    if entrada_funcao.isdigit() and 1 <= int(entrada_funcao) <= len(funcoes_validas):
        FUNCAO_FILTRADA = funcoes_validas[int(entrada_funcao) - 1]
        break
    print("[X] Entrada inválida.")

# === VALIDAR HTML PARA FUNÇÃO ===
if FUNCAO_FILTRADA not in df_html.index:
    print(f"\n[X] Nenhum corpo HTML encontrado para a função: {FUNCAO_FILTRADA}")
    exit()
TEMPLATE_HTML = df_html.loc[FUNCAO_FILTRADA]['TEXTO']

# === ASSUNTO PADRÃO ===
ASSUNTO_EMAIL = "Avaliação Periódica de Desempenho e Competências para Gestores das Unidades Escolares da SME - Ciclo 2024"

# === FILTRO DOS DADOS ===
df_filtrado = df_dados[
    (df_dados['LOCAL'] == LOCAL_FILTRADO) &
    (df_dados['FUNCAO'] == FUNCAO_FILTRADA) &
    (df_dados['Enviar'] == True)
]

if df_filtrado.empty:
    print(f"\n[X] Nenhum registro encontrado para LOCAL '{LOCAL_FILTRADO}' e FUNÇÃO '{FUNCAO_FILTRADA}'")
    exit()

print(f"\n🔎 {len(df_filtrado)} e-mail(s) prontos para envio ou exibição.\n")

# === CLIENTE OUTLOOK ===
outlook = win32.Dispatch("Outlook.Application")

# === LOG DE ENVIO ===
logs = []

for _, row in df_filtrado.iterrows():
    nome = row['NOME']
    email = row['EMAIL']
    local = row['LOCAL']
    funcao = row['FUNCAO']
    status = ""
    erro = ""

    try:
        corpo_email = TEMPLATE_HTML.replace("{nome}", nome)

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = ASSUNTO_EMAIL
        mail.HTMLBody = corpo_email

        # Anexar assinatura visual com Content-ID
        imagem_assinatura = r"C:\Projetos_Python\Disparo_Emails\Assinatura.png"
        if os.path.exists(imagem_assinatura):
            attachment = mail.Attachments.Add(imagem_assinatura)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "assinatura_img"
        )

        # === BUSCAR ANEXOS POR FUNÇÃO ===
        pasta_funcao = os.path.join(PASTA_ANEXOS, funcao)
        if os.path.exists(pasta_funcao):
            anexos = [os.path.join(pasta_funcao, f) for f in os.listdir(pasta_funcao) if os.path.isfile(os.path.join(pasta_funcao, f))]
            for anexo in anexos:
                mail.Attachments.Add(anexo)
        else:
            print(f"[!] Pasta de anexos não encontrada para: {funcao}")

        # === ENVIAR OU EXIBIR ===
        if ENVIAR:
            mail.Send()
            status = "enviado"
            print(f"[📤] E-mail ENVIADO para: {email}")
        else:
            mail.Display()
            status = "exibido"
            print(f"[👁️] E-mail EXIBIDO para: {email}")

    except Exception as e:
        erro = str(e)
        status = "erro"
        print(f"[X] Erro com {email}: {erro}")

    logs.append({
        "NOME": nome,
        "EMAIL": email,
        "FUNCAO": funcao,
        "LOCAL": local,
        "STATUS": status,
        "ERRO": erro
    })

# === SALVAR LOG ===
nome_log = f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
pd.DataFrame(logs).to_csv(nome_log, index=False, encoding='utf-8-sig')
print(f"\n📁 LOG salvo em: {nome_log}")
