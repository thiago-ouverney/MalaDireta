import win32com.client as win32
import mammoth

class MalaDireta:
    """
    Agrupa todas as principais funcionalidades para envio de Mala Direta
    mail.To : Destinatário do email ; para múltiplos (n) destinatários "To1;To2;To3;..;Ton"
    mail.Bcc : Cópia Oculta
    mail.Cc : Cópia
    mail.SendUsingAccount : Conta que o outlook irá usar para enviar o email
    mail.SentOnBehalfOfName : Em qual nome será enviado (Importante alinhar com SendUsingAccount)
    mail.Subject : Assunto do e-mail
    """
    def __init__(self):
        outlook = win32.Dispatch('outlook.application')
        self.mail = outlook.CreateItem(0)
        self.mail.HTMLBody = ''
        self.myNamespace = outlook.GetNamespace("MAPI")
        self.myFolder = self.myNamespace.GetDefaultFolder(6)

    def add_doc_body(self,docx_file):
        """
        mais informações: https://pypi.org/project/mammoth/#:~:text=Mammoth%20is%20designed%20to%20convert,document%2C%20and%20ignoring%20other%20details.
        :param docx_file: Path do arquivo .doc que será convertido em html para adicionar como corpo da mesangem
        :return:
        """
        result = mammoth.convert_to_html(docx_file)
        html = result.value
        self.mail.HTMLBody += html
        self.mail.BodyFormat = 2

    def add_html_body(self,html):
        """
        Adiciona diretamente no corpo do email já no formato html
        Convertor doc to html
        https://wordtohtml.net/
        :param html:
        :return:
        """
        self.mail.HTMLBody += html
        self.mail.BodyFormat = 2

    def add_attachment(self,arquivo):
        self.mail.Attachments.Add(arquivo)
