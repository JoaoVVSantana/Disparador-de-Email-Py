import smtplib
import os
import pandas as pd

# Multipurpose Internet Mail Extensions #
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


#--------------------------------------------------------------------------------------------------------#
#                                                                                                        #
#                                   ** É necessário atualizar: **                                        #
#                                                                                                        #         
#  - O host SMTP (atualmente usando office 365)                                                          #
#  - O diretório dos arquivos                                                                            #
#  - O arquivo de texto, que contém o corpo do Email                                                     #
#                                                                                                        #
#--------------------------------------------------------------------------------------------------------#





def main():
    conectar_smtp()

#Digitar o e-mail e senha, pensando em qual servidor SMTP vai ser usado#
meu_email=input('Digite seu Email: ')
minha_senha=input('Digite a Senha: ')






def conectar_smtp(): #Cria uma sessão SMTP#
     
    try:
        with smtplib.SMTP('smtp.office365.com',587) as server:
            
            server.starttls() #Criptografia#
            
            print('Realizando login... ')
            server.login(meu_email,minha_senha) #Realiza login com o e-mail e senha de quem vai enviar#
            print('Login Concluido')

            ler_planilha(server)  #Começa o processo de leitura dos arquivos#      
    except Exception as e:
        print(f'Falha ao conectar no servidor SMTP: {str(e)}') #Verificar qual o erro, servidor gmail apresenta problemas de autenticação#


#--------------------------------------------------------------------------------------------------------#
#                                                                                                        #
#                                   ** Padrões da planilha **                                           #
#                                                                                                        #
#  - [1,A] = Nome                                                                                        #
#  - [1,B] = Unidade                                                                                     #
#  - [1,C] = Setor                                                                                       #
#  - [1,D] = Email                                                                                       #
#  - [1,E] = Arquivo                                                                                     #
#  - [1,F] = Setor                                                                                       #
#                                                                                                        #
#--------------------------------------------------------------------------------------------------------#

def ler_planilha(server):

    
    lista=pd.read_excel(r'C:\Users\jvict\OneDrive\Desktop\Python\Script_Emails\teste Lista .xlsx')  #Leitura da planilha do excel#
    
    lista_filtrada=lista.dropna(subset=['Nome', 'Email', 'Arquivo']) #Filtra a planilha lida pra exibir só as linhas totalmente preenchidas#

    
    for index, row in lista_filtrada.iterrows(): #Itera a planilha filtrada, salvando nome, email e nome do certificado#

        
        nome_destinatario=row['Nome']
        
        email_destinatario=row['Email']

        nome_arquivo=row['Arquivo']

        string_Arq=f'{nome_arquivo}.pdf' #Transforma o nome do arquivo com .pdf pra colocar no caminho relativo
        print(f'Escrevendo para {nome_destinatario}')        
        escrever_email(string_Arq,nome_destinatario,email_destinatario,server)




def escrever_email(nome_arq,nome_destinatario,email_destinatario,server):

    
    conteudo=escrever_mensagem(nome_destinatario) #Salva o nome do destinatário e o conteúdo do arquivo txt#

    
    arq_pdf=f'C:/Users/jvict/OneDrive/Desktop/Python/Script_Emails/{nome_arq}' #Gambiarra que muda o nome do relative path pro certificado da pessoa da iteração#
    
    

    #Configurando o MIME#
    mensagem=MIMEMultipart()
    mensagem['From']=meu_email
    mensagem['To']=email_destinatario
    mensagem['Subject']=conteudo


    mensagem.attach(MIMEText(conteudo,"plain"))

    
    with open(arq_pdf,"rb") as anexo: # Abrindo o arquivo em modo de leitura binária (rb) - "as anexo" é o objeto de arquivo usado para ler os dados#

        
        base=MIMEBase('application','octet-stream') #Instanciando o MIMEBase, com application/octet-stream para anexar arquivos binários#

        
        base.set_payload(anexo.read()) #Lê todo conteúdo do arquivo binário e define como o payload (corpo de dados) da parte MIME#

        
    encoders.encode_base64(base) #Convertendo os dados binários em caractéres ASCII#

    
    base.add_header('Content-Disposition ',f"attachment; filename={os.path.basename(arq_pdf)}" ) #Adiciona cabeçalho, especifica como o conteúdo deve ser exibido, indica que o conteúdo deve ser tratado como anexo#

    
    mensagem.attach(base) #Anexa o PDF ao e-mail#

    enviar_email(mensagem, email_destinatario,server) #Pega a mensagem MIME, quem vai receber o email e de quem tá mandando#

    
    
def escrever_mensagem(nome):

    
    caminho_texto=(r'C:\Users\jvict\OneDrive\Desktop\Python\Script_Emails\mensagem.txt') #Pega o arquivo txt, lê e salva o conteúdo numa string#

    
    conteudo=f'Prezado(a) {nome},\n' #A string recebe o nome da pessoa que vai receber o Email#

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