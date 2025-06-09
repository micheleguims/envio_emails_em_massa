import pandas as pd
import os
import win32com.client as win32
from datetime import datetime

# === ⚙️ CONFIGURAÇÕES GERAIS ===
ARQUIVO_PLANILHA = r"C:\PPT\EvioEmails\BASE_GE_24_(Envio_Emails).xlsx"
ABA = "Base_GE_2024"
CAMINHO_ANEXO = r"C:\PPT\EvioEmails\Exemplo de Anexo.pdf"

# === 🔍 MODO DE EXECUÇÃO (DISPLAY / SEND) ===
ENVIAR = False  # Inicializa como falso

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

# === 💌 TEMPLATE DE E-MAIL ===
TEMPLATE_HTML = """
<h2> Prezado(a) {funcao}: <em><span style="color:pink;">{nome}</span></em>, </h2>
<br>
<p>Informamos que a avaliação irá iniciar em 11/06/2025.</p>
<br>
<p>Atenciosamente,<br>
<strong>Equipe Avaliação e Desempenho.</strong></p>
"""

ASSUNTO_EMAIL = "Avaliação Diretores Ciclo 2024"

# === 📊 LER PLANILHA ===
df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA, engine='openpyxl')
df['LOCAL'] = df['LOCAL'].astype(str).str.strip()
df['FUNCAO'] = df['FUNCAO'].astype(str).str.strip().str.upper()
df['Enviar'] = df['Enviar'].fillna(False).astype(bool)

# === VALORES VÁLIDOS ===
locais_validos = [str(i) for i in range(1, 13)] + ["NC"]
funcoes_validas = ["DIRETOR ADJUNTO", "DIRETOR IV", "COORDENADOR", "GERENTE", "SECRETARIA"]

# === SELECIONAR LOCAL ===
print("\n📍 Lista de LOCALS disponíveis:")
for i, local in enumerate(locais_validos, start=1):
    print(f"[{i}] {local}")

while True:
    entrada = input("\nDigite o número do LOCAL desejado: ").strip()
    if entrada.isdigit() and 1 <= int(entrada) <= len(locais_validos):
        LOCAL_FILTRADO = locais_validos[int(entrada) - 1]
        break
    else:
        print("[X] Entrada inválida.")

# === SELECIONAR FUNÇÃO ===
print(f"\n👤 Lista de FUNÇÕES permitidas:")
for i, funcao in enumerate(funcoes_validas, start=1):
    print(f"[{i}] {funcao}")

while True:
    entrada_funcao = input("\nDigite o número da FUNÇÃO desejada: ").strip()
    if entrada_funcao.isdigit() and 1 <= int(entrada_funcao) <= len(funcoes_validas):
        FUNCAO_FILTRADA = funcoes_validas[int(entrada_funcao) - 1]
        break
    else:
        print("[X] Entrada inválida.")

# === FILTRO ===
df_filtrado = df[
    (df['LOCAL'] == LOCAL_FILTRADO) &
    (df['FUNCAO'] == FUNCAO_FILTRADA) &
    (df['Enviar'] == True)
]

if df_filtrado.empty:
    print(f"\n[X] Nenhum registro encontrado para LOCAL '{LOCAL_FILTRADO}' e FUNÇÃO '{FUNCAO_FILTRADA}'")
    exit()

print(f"\n🔎 {len(df_filtrado)} e-mail(s) encontrados para envio ou exibição.\n")

# === OUTLOOK ===
outlook = win32.Dispatch("Outlook.Application")

# === LOG ===
logs = []

for _, row in df_filtrado.iterrows():
    nome = row['NOME']
    email = row['EMAIL']
    funcao = row['FUNCAO']
    local = row['LOCAL']
    status = ""
    erro = ""

    try:
        corpo_email = TEMPLATE_HTML.format(nome=nome, funcao=funcao)

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = ASSUNTO_EMAIL
        mail.HTMLBody = corpo_email

        if os.path.exists(CAMINHO_ANEXO):
            mail.Attachments.Add(CAMINHO_ANEXO)
        else:
            erro = f"Anexo não encontrado: {CAMINHO_ANEXO}"
            status = "erro"
            print(f"[!] {erro}")
            continue

        if ENVIAR:
            mail.Send()
            status = "enviado"
            print(f"[📤] E-mail ENVIADO para: {email}")
        else:
            mail.Display()
            status = "exibido"
            print(f"[👁️] E-mail EXIBIDO para: {email}")

    except Exception as e:
        status = "erro"
        erro = str(e)
        print(f"[X] Erro com {email}: {erro}")

    logs.append({
        "NOME": nome,
        "EMAIL": email,
        "FUNCAO": funcao,
        "LOCAL": local,
        "STATUS": status,
        "ERRO": erro
    })

# === SALVA LOG ===
nome_arquivo_log = f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
pd.DataFrame(logs).to_csv(nome_arquivo_log, index=False, encoding='utf-8-sig')
print(f"\n📁 LOG salvo com sucesso: {nome_arquivo_log}")
