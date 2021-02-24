# TRY PYINSTALLER
import os
import json
import shutil
import sqlite3
import os.path
import smtplib
import pathlib
from time import time
from time import ctime
from datetime import date
from email import encoders
from datetime import datetime
from datetime import timedelta
from os.path import getmtime as hora
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


def copy_xml_zip(original_path, copy_path, log_file, report_error):

    file_name = time_conversor()
    temp_path = str(copy_path)+'\\'+file_name
    full_name = temp_path+".zip"
    files_cp = 0

    try:
        if not os.path.exists(copy_path):
            os.makedirs(copy_path)
            log_file.write(today() + " --- Pasta Copia Criada\n")

        if not os.path.exists(temp_path):
            os.makedirs(temp_path)
            log_file.write(today() + " --- Pasta Temporária Criada\n")

        for root, folder, files in os.walk(original_path):
            if (len(files)) > 0:
                for file_ in files:
                    if file_.endswith(".xml"):
                        full_path_file = pathlib.Path(original_path, file_)
                        file_date = datetime.strptime(
                            ctime(hora(full_path_file)), "%a %b %d %H:%M:%S %Y")
                        if last_month() in str(file_date):
                            shutil.copy2(full_path_file, temp_path)
                            files_cp += 1

                if files_cp > 0:
                    log_file.write(today() + " --- " + str(files_cp) +
                                   " arquivos copiados com sucesso!\n")
                    os.chdir(copy_path)
                    shutil.make_archive(
                        file_name, 'zip', base_dir=time_conversor())
                    log_file.write(
                        today() + " --- Arquivos compactados com sucesso!\n")

                    for root, folder, files in os.walk(temp_path, topdown=False):
                        for file_ in files:
                            os.remove(os.path.join(root, file_))
                        os.rmdir(temp_path)
                        log_file.write(
                            today() + " --- Pastas/arquivos deletados com sucesso!\n")

                    return(True, full_name, report_error)

                else:
                    log_file.write(
                        today() + " --- Nenhum arquivo copiado de Original_path\n")
                    report_error = 1
                    return(False, '', report_error)
            else:
                log_file.write(
                    today() + " --- Nenhum arquivo encontrado em Original_path\n")
                report_error = 1
                return(False, '', report_error)

    except Exception:
        return(False, '', report_error)


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location):

    email_login = 'enviarxml@gmail.com'
    email_pass = '123456abc'

    msg = MIMEMultipart()
    msg['From'] = email_login
    msg['Subject'] = email_subject

    if type(email_recipient) == list and len(email_recipient) > 1:
        msg['To'] = ",".join(email_recipient)
    else:
        msg['To'] = email_recipient

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(email_login, email_pass)
        text = msg.as_string()
        server.sendmail(email_login, email_recipient, text)
        server.quit()
        return True
    except Exception:
        return False


def time_conversor():
    today = date.today()
    first = today.replace(day=1)
    enterprise = config['enterprise']
    mes = ((first - timedelta(days=1)).strftime("%m-%Y"))
    return('XMLs de ' + mes + ' - ' + enterprise)


def first_day():
    today = datetime.now()
    first_day = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return(datetime.timestamp(first_day))


def last_month():
    today = date.today()
    first = today.replace(day=1)

    return((first - timedelta(days=1)).strftime("%Y-%m"))


def create_table():
    cursor.execute("""CREATE TABLE IF NOT EXISTS sent (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        unix INTEGER, sent INTEGER
    )""")


def feed_table(unix, sent):
    infos = [(float(unix), int(sent))]

    cursor.executemany("""INSERT INTO sent (unix, sent)
                    VALUES (?,?)""", infos)

    connc.commit()


def read_db(first_day):
    sql = 'SELECT * FROM sent WHERE unix > ?'
    for row in cursor.execute(sql, (first_day,)):
        x = row[2]
        if x == 1:
            return True

    else:
        return False


def today():
    return datetime.today().strftime("%d-%m-%Y %H:%M:%S:%f")


def create_json():
    with open('config.json', 'w+', encoding='utf-8') as conf_json:
        conf_json.write(r"""{
    "root_path":"C:/dev/python/copiar arquivos/xml original",
    "new_path":"C:/xmls",
    "recipient":["pedroamaral0119@outlook.com", "pedrobrito0119ti@outlook.com"],
    "enterprise":"Pedro Co.",
    "content":"Olá, segue o arquivo zip com os XMLs referente ao último mês\n\n\n\n\nTI_ Automação Comercial\nE-mail automático, não responda.\nQualquer dúvida entre em contato através do endereço: ti_consultoria@hotmail.com"
}""")


if __name__ == '__main__':
    log_file = open('log.txt', 'a+', encoding='utf-8')
    log_file.write(today() + " --- ||Abrindo Programa||\n")
    report_error = 0

    if not os.path.exists('config.json'):
        create_json()

    with open('config.json', encoding='utf-8') as confjson:
        config = json.load(confjson)

    log_file.write(today() + " --- " + config['enterprise'] + "\n")
    connc = sqlite3.connect('sent.db')
    cursor = connc.cursor()
    create_table()

    root_path = config['root_path']
    new_path = config['new_path']
    recipient = config['recipient']
    enterprise = config['enterprise']
    subject = time_conversor()
    content = config['content']

    if not os.path.exists(root_path):
        log_file.write(today() + " --- Pasta inexistente.\n")
        report_error = 1
    else:
        if not read_db(first_day()):
            boole, att, report_error = copy_xml_zip(
                root_path, new_path, log_file, report_error)
            if boole:
                if send_email(recipient, subject, content, att):
                    feed_table(int(time()), 1)
                    log_file.write(
                        today() + " --- Email enviado com sucesso!!\n")
                else:
                    feed_table(time(), 0)
                    log_file.write(
                        today() + " --- Falha ao enviar o email.\n")
                    report_error = 1
            else:
                log_file.write(
                    today() + " --- Falha ao copiar/comprimir arquivos\n")
                report_error = 1

        else:
            log_file.write(today() + " --- Email já enviado\n")
            connc.close()

    log_file.write(today() + " --- ||Fechando Programa||\n\n")
    log_file.close()

    if report_error == 1:
        text_for_fail = 'Falha no Programa ' + time_conversor()
        send_email('pedroab0119@gmail.com', text_for_fail,
                   text_for_fail, 'log.txt')
