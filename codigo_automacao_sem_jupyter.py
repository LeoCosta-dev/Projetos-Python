#Importação de bibliotecas.

import pyautogui
import time
import pyperclip
import pandas as pd

#Atribuição de variável para o link onde se encontra hospedada a tabela.

link = 'https://drive.google.com/drive/folders/1Fb3RRj9rzhw_Zh3rrp-8PEWZQYTOjbBl?usp=sharing'

#Baixar base de dados (excel) do servidor (google drive).

pyautogui.PAUSE = 3

#Aviso para o usuário do ínicio da aplicação.
pyautogui.alert('A automação irá iniciar, favor não usar a máquina até o final do processo.')

pyautogui.hotkey('ctrl', 't') #Atalhos usados em virtude da aplicação iniciar no navegador.
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(16)
pyautogui.click(315, 348, clicks=2)

time.sleep(20)
pyautogui.click(991, 185)

time.sleep(20)

#Atribuíção de variável para manipulação da tabela com o pandas.
tabela = pd.read_excel('/home/familia/Downloads/Vendas-Dez.xlsx')

display(tabela)

#Atribuição de variáveis para acomodar os valores extraídos dos dados.
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
ticketMedio = faturamento/quantidade

#acessar o servidor de e-mail (oulook) e enviar para o destinatário (e-mail fictício)
pyautogui.hotkey('ctrl', 't')
linkServidorDeEmail = 'https://outlook.live.com/mail/0/inbox'
pyperclip.copy(linkServidorDeEmail)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#Automação do click para iniciar o envio do e-mail.
time.sleep(26)
pyautogui.click(171, 237)

#Período de aguardo para o carregamento da página e definição dos destinatários.
time.sleep(26)
pyautogui.write('e-mailexemplo+@gmail.com')
pyautogui.press('enter')

#Escrita do assunto ao qual se trata o e-mail.
pyautogui.click(410, 347)
assunto = "Relatório geral de vendas do mês de dezembro de 2019"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

#Escrita e formatação do corpo do e-mail.
texto = f'''Caríssimos,

segue relatório de vendas referente ao mês de Dezembro de 2019:

Faturamento: R$ {faturamento:,.2f}
Quantidade de produtos vendidos:  {quantidade:,}
Ticket Médio de venda: R$ {ticketMedio:,.2f}

Desejo a todos boas festas, atenciosamente...

Léo Costa.

'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

#Envio
time.sleep(10)
pyautogui.hotkey('ctrl', 'enter')

#Aviso para o usuário que a aplicação terminou.
time.sleep(10)
pyautogui.alert('Final do processamento, automação concluída com sucesso!')

#Observação final, todo texto com caractéres especiais foram abordados como variáveis para evitar conflito no momento da escrita.

#código para descobrir a posição onde o mouse deve clicar de acordo com o seu monitor.

#time.sleep(15)
#pyautogui.position()