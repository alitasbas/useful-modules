import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Read data from Excel
data = pd.read_excel(".xlsx")

# Email details
sender_email = 'sender@gmail.com'  # Update with your email address
password = 'code'  # Update with your email password
smtp_server = 'smtp.gmail.com'  # Update with your SMTP server details
smtp_port = 587  # Update with your SMTP server port number

# Email template
email_template = '''
Merhabalar {} {},
Ben ODTÜ İstatistik ve Veri Bilimi Topluluğu aktif üyesi .
ODTÜ bünyesinde 1991 yılında hayat bulan ODTÜ İstatistik Topluluğu olarak üniversite öğrencilerine, istatistiğin değerini ve farklı alanlardaki uygulamalarını göstermeyi, bilime ve teknolojiye değer katmayı, insanları bilinçlendirmeyi ve onların çevrelerine karşı farkındalık kazanmalarını amaçlamaktayız.
Bu düşüncelerle başlatılan ve alanımızda çalışan okulumuz mezunlarını ağırladığımız geleneksel Speed Networking etkinliğimize 17 Haziran 2023 tarihinde Teknokent ZGarage’a sizin de katılımınızı bekliyoruz.
Daha ayrıntılı bilgileri size sunmak adına etkinliğimizin tanıtım dosyası ektedir.
Geri dönüşlerinizi bekliyor olacağım. 

İyi çalışmalar dilerim.
Saygılarımla,
'''

# CC address
cc_email = 'CC@gmail.com'  # Update with the CC email address

# File attachment
attachment_file = '.pptx'  # Update with your PowerPoint file path and name

# SMTP Connection
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, password)

# Batch size for sending emails
batch_size = 10  # Update with your preferred batch size

# Iterate over recipients and send emails in batches
for batch_start in range(0, len(data), batch_size):
    batch_end = min(batch_start + batch_size, len(data))
    batch_data = data.iloc[batch_start:batch_end]

    for index, row in batch_data.iterrows():
        recipient_name = row['Name']
        recipient_email = row['Email']

        # Modify email layout
        personalized_email = email_template.format(name=recipient_name, title="Hocam")

        # Create email message
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Cc'] = cc_email
        message['Subject'] = 'subject'

        # Attach email content
        message.attach(MIMEText(personalized_email, 'plain'))

        # Attach PowerPoint file
        with open(attachment_file, 'rb') as file:
            attachment = MIMEBase('application', 'vnd.openxmlformats-officedocument.presentationml.presentation')
            attachment.set_payload(file.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename='.pptx')
        message.attach(attachment)

        # Add email to SMTP transaction
        server.sendmail(sender_email, [recipient_email, cc_email], message.as_string())

# Close SMTP connection
server.quit()
