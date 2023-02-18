import csv
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    # cria o servidor SMTP
    context = ssl.create_default_context()
    server = smtplib.SMTP('smtp.kinghost.net', 587)
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()

    # dados do remetente
    remetente = 'seuemail'
    senha = 'suasenha' # caso queira deixar o código mais seguro, pode dar um input para o usuário colocar a senha

    server.login(remetente, senha)

    for email in csv.DictReader(open('emails.csv')):
        nome = email['nome']
        email = email['email']

        # dados do e-mail
        body = """Lorem ipsum dolor sit amet. Et quia dolorem est maiores explicabo id sunt rerum ut quia cumque et eligendi laboriosam At nemo nulla. Ab itaque neque qui fuga tempora sit nobis atque. Non incidunt exercitationem qui nulla dolore sit porro omnis eum illum expedita est perferendis sint eum facilis blanditiis et exercitationem tempora? Qui corporis accusamus At reiciendis omnis quo amet itaque et labore consequatur eos ratione delectus.

Ut rerum blanditiis ea laudantium nihil et atque ipsa id labore aliquid sed esse dolorem. Sit pariatur numquam id itaque minus hic porro incidunt!

A deleniti error sit accusantium sapiente 33 neque optio et enim doloribus qui enim rerum. Aut voluptatum enim ut eligendi ullam non nobis odio eum quas animi eum amet sunt! Ad cupiditate consectetur vel accusantium reprehenderit aut voluptatem harum est provident iusto qui deleniti voluptas et voluptas ipsa.
    """
        message = MIMEMultipart()
        message['Subject'] = f"ASSUNTO - {nome.upper()}"
        message['From'] = remetente
        message['To'] = email
        message.attach(
        MIMEText(body, 'plain'))

        # define os atributos do anexo
        filename = f"C://users/vitor/desktop/email/{nome.upper()}.pdf" # para cada pessoa enviar um pdf que estará escrito o nome dela
        attachment = MIMEBase('application', 'octet-stream')

        with open(filename, 'rb') as f:
            attachment.set_payload(f.read())

            encoders.encode_base64(attachment)

        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename={nome}',
        )

        # anexa o arquivo no e-mail
        message.attach(attachment)

        # realiza login no servidor
        server.login(remetente, senha)

        # envia o email
        server.sendmail(remetente, email, message.as_string())

except Exception as e:
  print(f'Erros: {e}')
finally:

  server.quit()
