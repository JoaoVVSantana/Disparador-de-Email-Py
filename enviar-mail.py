import smtplib
import os
import pandas as pd
import re

# Multipurpose Internet Mail Extensions #
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


#--------------------------------------------------------------------------------------------------------#
#                                                                                                        #
#                                   ** É necessário atualizar: **                                        #
#                                                                                                        #         
#  - O host SMTP (atualmente usando o gmail)                                                     #
#  - O diretório dos arquivos                                                                            #
#  - O arquivo de texto, que contém o corpo do Email                                                     #
#                                                                                                        #
#--------------------------------------------------------------------------------------------------------#





def main():
    conectar_smtp()
    print(emails_invalidos)

#Inserir a assinatura pessoal ou institucional aqui
assinatura_html = """

"""
#Digitar o e-mail e senha, pensando em qual servidor SMTP vai ser usado#
meu_email=''
email_login=''
minha_senha=''

emails_invalidos = []




def conectar_smtp(): #Cria uma sessão SMTP#
     
    try:
        with smtplib.SMTP('smtp.gmail.com',587) as server:
            
            server.starttls() #Criptografia#
            
            print('Realizando login... ')
            server.login(email_login,minha_senha) #Realiza login com o e-mail e senha de quem vai enviar#
            print('Login Concluido')

            ler_planilha(server)  #Começa o processo de leitura da planilha#      
    except Exception as e:
        print(f'Falha no servidor SMTP: {str(e)}') #Verificar qual o erro#


#------------------------------------------------------------------------------------------------------#
#                                                                                                      #
#                                   ** Padrões da planilha **                                          #
#                                                                                                      #
#  - [1,A] = Data/Hora                                                                                 #
#  - [1,B] = Nome Completo                                                                             #
#  - [1,C] = Unidade                                                                                   #
#  - [1,D] = Cargo                                                                                     #
#  - [1,E] = E-mail                                                                                     #
#                                                                                                      #
#------------------------------------------------------------------------------------------------------#

def ler_planilha(server):

    
    lista=pd.read_excel(r'Lista_de_presenca_respostas.xlsx')  #Leitura da planilha do excel#
    
    lista_filtrada=lista.dropna(subset=['Nome Completo', 'E-mail']) #Filtra a planilha lida pra exibir só as linhas totalmente preenchidas#

    email_regex = r'^[\w\.-]+@[\w\.-]+\.\w+$' #verificar se o email está no padrão correto

    for i, row in lista_filtrada.iterrows(): #Itera a planilha filtrada, salvando nome, email e nome do certificado#
        nome_destinatario=row['Nome Completo']
        email_destinatario=row['E-mail']
        #Tratando os inválidos
        if re.match(email_regex, email_destinatario):
            print(f'Escrevendo para {nome_destinatario}')        
            escrever_email(nome_destinatario,email_destinatario,server)
        else:
            emails_invalidos.append((nome_destinatario, email_destinatario))
            print(f"Email inválido encontrado: {email_destinatario}. Pulando este destinatário.")

        




def escrever_email(nome_destinatario,email_destinatario,server):

    
    conteudo=escrever_mensagem(nome_destinatario) #Salva o nome do destinatário e o conteúdo do arquivo txt#

    corpo_html = f"""
    <html>
    <body>
    <p>{conteudo.replace('\n', '<br>')}</p>
    {assinatura_html}
    </body>
    </html>
    """

    
    arq_excel=r"Relatorios.xlsx"
    
    

    #Configurando o MIME#
    mensagem=MIMEMultipart()
    mensagem['From']=meu_email
    mensagem['To']=email_destinatario
    mensagem['Subject']='- Reunião dia XX/XX'

    

    mensagem.attach(MIMEText(corpo_html,"html"))
    
    
    with open(arq_excel,"rb") as anexo: # Abrindo o arquivo em modo de leitura binária (rb) - "as anexo" é o objeto de arquivo usado para ler os dados#

        
        base = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')#Instanciando o MIMEBase, com application/octet-stream para anexar arquivos binários#

        
        base.set_payload(anexo.read()) #Lê todo conteúdo do arquivo binário e define como o payload (corpo de dados) da parte MIME#

        
    encoders.encode_base64(base) #Convertendo os dados binários em caractéres ASCII#

    
    base.add_header('Content-Disposition ',f"attachment; filename={os.path.basename(arq_excel)}") #Adiciona cabeçalho, especifica como o conteúdo deve ser exibido, indica que o conteúdo deve ser tratado como anexo#

    
    mensagem.attach(base) #Anexa o Excel ao e-mail#

    enviar_email(mensagem, email_destinatario,server) #Pega a mensagem MIME, quem vai receber o email e de quem tá mandando#

    
    
def escrever_mensagem(nome):

    
    caminho_texto=(r'mensagem.txt') #Pega o arquivo txt, lê e salva o conteúdo numa string#

    
    conteudo=f'Prezado(a) {nome}, bom dia!\n' #A string recebe o nome da pessoa que vai receber o Email#

    with open(caminho_texto, 'r', encoding='utf-8') as arquivo:
        conteudo_arquivo=arquivo.read()
        conteudo+=conteudo_arquivo
    return conteudo

def enviar_email(mensagem, email_destinatario, server):

    #Criação de uma sessão de SMTP para enviar o e-mail#
    try:
        
            #Converte a mensagem construída no MIME em uma string a ser enviada#
            texto = mensagem.as_string()
            server.sendmail(meu_email, email_destinatario, texto)
            
            print(f'E-mail enviado para {email_destinatario}')
    except Exception as e:
        print(f'Falha ao enviar e-mail para {email_destinatario}: {str(e)}')


if (__name__=='__main__') : 
    main()