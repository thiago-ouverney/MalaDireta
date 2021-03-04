from functions import MalaDireta

email_solicitante = "Servicos.Integrados.de.RH@Vale.com"
email_destino = "thiago.ouverney@vale.com"
Outlook = MalaDireta()

#Outlook.mail.Bcc = email_copia_oculta
#Outlook.mail.Cc = email_copia
Outlook.mail.To =  email_destino
Outlook.mail.SendUsingAccount = email_solicitante
Outlook.mail.SentOnBehalfOfName = email_solicitante

Outlook.mail.Subject = "Envio Mala Direta"
Outlook.add_html_body("""<p><strong>Envio Teste</strong></p>
<p><br></p>
<p><strong>Att,</strong></p>
<p><strong><u>Thiago Ouverney</u></strong></p>""")
# Outlook.add_attachment(arquivo_anexo)
# https://wordtohtml.net/

Outlook.mail.Send()