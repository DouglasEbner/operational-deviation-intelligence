import win32com.client as win32
import os

# Pasta dos arquivos
pasta = "Pasta"

grafico1 = os.path.join(pasta, "grafico_dia_atual.png")
grafico2 = os.path.join(pasta, "grafico_evolucao_uf.png")

# Inicializa Outlook
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

email.To = "mail"
email.Subject = "subject"

# Adiciona anexos (para gráficos embutidos, também precisa como anexo)
att1 = email.Attachments.Add(grafico1)
att1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "sample01.png")

att2 = email.Attachments.Add(grafico2)
att2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "sample02.png")

# Corpo HTML com gráficos embutidos
email.HTMLBody = f"""
<p>Olá,</p>
<p>Aqui vai o descritivo do material</p>
<p><b>Direcionamento do grafico#1</b><br>
<img src="cid:GraficoDiaAtual"></p>
<p><b>desvcritivo do dado #2</b><br>
<img src="cid:GraficoEvolucaoUF"></p>
<p>Anexo com todos os dados</p>
<p>Att,<br>assinatura</p>
"""

# Exibir e-mail antes de enviar
email.Display()  # usa email.Send() para enviar direto
