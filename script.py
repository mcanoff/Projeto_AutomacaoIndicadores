import pandas as pd
import smtplib
import email.message
import os
from senha_enviar_email import EMAIL_PASSWORD
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# importar base de dados
emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

# colocar a coluna de nome das lojas na base de dados 'vendas'
vendas = vendas.merge(lojas, on='ID Loja')
dia_indicador = vendas['Data'].max()

# criar uma tabela para cada loja (colunas: data, loja, produto, quantidade, valor un, valor final, loja)
dic_lojas = {}
for loja in vendas['Loja'].unique():
    dic_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]

# criar pasta de backups
root_dir = r'C:\Users\miria\OneDrive\Área de Trabalho\Programacao\learning-python\Projeto AutomacaoIndicadores'
backup_dir = os.path.join(root_dir, "Backup Arquivos Lojas")

if not os.path.exists(backup_dir):
    os.makedirs(backup_dir)

for loja in dic_lojas:
    # criar uma pasta com o nome da loja
    loja_dir = os.path.join(backup_dir, loja)
    os.makedirs(loja_dir, exist_ok=True)
    # cria um arquivo em excel com o nome da loja e armazenar o informação do dicionario no arquivo
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    caminho_arquivo = os.path.join(loja_dir, nome_arquivo)
    excel_loja = dic_lojas[loja]
    excel_loja.to_excel(caminho_arquivo, index=False)

# criar função que calcula o indicador para cada loja

def calcular_indicador(loja):

    # definir um df para as vendas anuais e outro para as vendas do dia
    vendas_ano = dic_lojas[loja]
    vendas_dia = vendas_ano.loc[vendas_ano['Data'] == dia_indicador, :]

    # calcular o faturamento usando os dfs acima
    faturamento_ano = vendas_ano['Valor Final'].sum()
    faturamento_dia = vendas_dia['Valor Final'].sum()

    # calcular a diversidade de produtos
    diversidade_ano = len(vendas_ano['Produto'].unique())
    diversidade_dia = len(vendas_dia['Produto'].unique())

    # ticket médio
    valores_vendas_ano = vendas_ano.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valores_vendas_ano['Valor Final'].mean()

    valores_vendas_dia = vendas_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valores_vendas_dia['Valor Final'].mean()

    indicadores = (
    faturamento_ano,
    faturamento_dia,
    diversidade_ano,
    diversidade_dia,
    ticket_medio_ano,
    ticket_medio_dia
)
    return indicadores

def enviar_email(email_gerente, loja, dia_indicador, indicadores):

    meta_faturamento_dia = 1000
    meta_faturamento_ano = 1650000
    meta_qtdeprodutos_dia = 4
    meta_qtdeprodutos_ano = 120
    meta_ticketmedio_dia = 500
    meta_ticketmedio_ano = 500

    faturamento_ano, faturamento_dia, diversidade_ano, diversidade_dia, ticket_medio_ano, ticket_medio_dia = indicadores


    msg = MIMEMultipart()
    msg['Subject'] = f'OnePage {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    msg['From'] = 'mcanoff16@gmail.com'
    # Separar com ; os e-mails em cópia
    msg['To'] = email_gerente
    msg['Cc'] = 'mcanoff16+copia@gmail.com'
    msg['Bcc'] = ''
    nome_gerente = emails.loc[emails['E-mail'] == email_gerente, 'Gerente'].values[0]

   # Corpo do e-mail com a tabela embutida
    corpo_email = f"""
    <html>
    <head>
        <style>
            table {{
                border-collapse: collapse;
                width: 100%;
            }}
            th, td {{
                border: 1px solid black;
                padding: 8px;
                text-align: center;
            }}
            th {{
                background-color: #f2f2f2;
            }}
            caption {{
                font-size: 18px;
                font-weight: bold;
                color: #333;
                text-align: center;
                padding: 10px;
                background-color: #f8f8f8;
                margin-bottom: 10px;
                }}
        </style>
    </head>
    <body>
        <p>Bom dia, {nome_gerente}</p>
        <p>O resultado de ontem <strong>(dia {dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong> foi:</p>
        <table>
            <caption>Indicadores Diários</caption>
            <thead>
                <tr>
                    <th></th>
                    <th>Valor Dia</th>
                    <th>Meta Dia</th>
                    <th>Cenário Dia</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <th>Faturamento</th>
                    <td>R$ {faturamento_dia:.2f}</td>
                    <td>R$ {meta_faturamento_dia:.2f}</td>
                    <td>{'✔️' if faturamento_dia >= meta_faturamento_dia else '❌'}</td>
                </tr>
                <tr>
                    <th>Diversidade de Produtos</th>
                    <td>{diversidade_dia}</td>
                    <td>{meta_qtdeprodutos_dia}</td>
                    <td>{'✔️' if diversidade_dia >= meta_qtdeprodutos_dia else '❌'}</td>
                </tr>
                <tr>
                    <th>Ticket Médio</th>
                    <td>R$ {ticket_medio_dia:.2f}</td>
                    <td>R$ {meta_ticketmedio_dia:.2f}</td>
                    <td>{'✔️' if ticket_medio_dia >= meta_ticketmedio_dia else '❌'}</td>
                </tr>
            </tbody>
        </table>

        <table>
            <caption>Indicadores Anuais</caption>
            <thead>
                <tr>
                    <th></th>
                    <th>Valor Ano</th>
                    <th>Meta Ano</th>
                    <th>Cenário Ano</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <th>Faturamento</th>
                    <td>R$ {faturamento_ano:.2f}</td>
                    <td>R$ {meta_faturamento_ano:.2f}</td>
                    <td>{'✔️' if faturamento_ano >= meta_faturamento_ano else '❌'}</td>
                </tr>
                <tr>
                    <th>Diversidade de Produtos</th>
                    <td>{diversidade_ano}</td>
                    <td>{meta_qtdeprodutos_ano}</td>
                    <td>{'✔️' if diversidade_ano >= meta_qtdeprodutos_ano else '❌'}</td>
                </tr>
                <tr>
                    <th>Ticket Médio</th>
                    <td>R$ {ticket_medio_ano:.2f}</td>
                    <td>R$ {meta_ticketmedio_ano:.2f}</td>
                    <td>{'✔️' if ticket_medio_ano >= meta_ticketmedio_ano else '❌'}</td>
                </tr>
            </tbody>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att.,<br>Mirian</p>
    </body>
    </html>
    """

    msg.attach(MIMEText(corpo_email, "html"))

    # anexar o corpo de texto e os anexos
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    with open(caminho_arquivo, "rb") as arquivo:
        msg.attach(MIMEApplication(arquivo.read(), Name=nome_arquivo))

    # conectar no servidor de email - cada servidor vai ter parâmetros diferentes (configurações smtp)
    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    # determinar o formato de criptografia - padão TLS
    servidor.starttls()

    # fazer login no e-mail - para criar a senha -> conf. conta do Google -> senhas de app
    servidor.login(msg['From'], EMAIL_PASSWORD)
    # enviar o e-mail
    servidor.send_message(msg)
    # fechar a conexão com o servidor
    servidor.quit()
    print("Email enviado.")


# Automatizar todas as lojas
for loja in lojas['Lojas']:
    
    indicadores_loja = calcular_indicador(loja)

    try:
        email_gerente = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]  # Pegando o primeiro e-mail encontrado
        print(email_gerente)
        enviar_email(email_gerente, loja, dia_indicador, indicadores_loja)
    except IndexError:
        print(f"Erro: Não foi encontrado um e-mail para a loja {loja}.")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")