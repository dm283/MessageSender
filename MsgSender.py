import sys, os, configparser, re, datetime, json, asyncio, email, tkinter as tk
import requests, aioodbc
from aiosmtplib import SMTP
from aioimaplib import aioimaplib
from cryptography.fernet import Fernet
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# загрузка конфигурации
CONFIG_FILE = 'config.ini'
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding='utf-8')

# загрузка ключа шифрования
with open('rec-k.txt') as f:
    rkey = f.read().encode('utf-8')
refKey = Fernet(rkey)

# расшифровка хэшированных значений конфигурации
hashed_user_credentials_password = config['user_credentials']['password'].split('\t#')[0]
user_credentials_password = (refKey.decrypt(hashed_user_credentials_password).decode('utf-8'))
hashed_common_bot_token = config['telegram_bot']['bot_token'].split('\t#')[0]
common_bot_token = (refKey.decrypt(hashed_common_bot_token).decode('utf-8'))
hashed_email_server_password = config['email']['server_password'].split('\t#')[0]
email_server_password = (refKey.decrypt(hashed_email_server_password).decode('utf-8'))

CHECK_DB_PERIOD = int(config['common']['check_db_period'].split('\t#')[0])  # период проверки новых записей в базе данных
DIR_EMAIL_ATTACHMENTS = Path(config['common']['dir_email_attachments'].split('\t#')[0])  # директория с файлами для отправки
DIR_TELEGRAM_ATTACHMENTS = Path(config['common']['dir_telegram_attachments'].split('\t#')[0])
#Path('c:/Users/dm283/Documents/Tech/ALTA/MessageSender/attachments')
#Path().absolute() / 'attachments'

USER_NAME = config['user_credentials']['name'].split('\t#')[0]
USER_PASSWORD = user_credentials_password

ADMIN_EMAIL = config['admin_credentials']['email'].split('\t#')[0]  # почта админа

BOT_NAME = config['telegram_bot']['bot_name'].split('\t#')[0]
BOT_TOKEN = common_bot_token
TELEGRAM_DB = config['telegram_bot']['db'].split('\t#')[0]  # база данных mssql/posgres
TELEGRAM_DB_TABLE_MESSAGES = config['telegram_bot']['db_table_messages'].split('\t#')[0]  # db.schema.table
TELEGRAM_DB_TABLE_CHATS = config['telegram_bot']['db_table_chats'].split('\t#')[0]  # db.schema.table  таблица с telegram-чатами
TELEGRAM_DB_CONNECTION_STRING = config['telegram_bot']['db_connection_string'].split('\t#')[0]  # odbc driver system dsn name
ADMIN_BOT_CHAT_ID = str()  # объявление глобальной константы, которая записывается в функции load_telegram_chats_from_db
MODE_EMAIL, MODE_TELEGRAM = bool(), bool()

SENDER_EMAIL, EMAIL_SERVER_PASSWORD = config['email']['sender_email'].split('\t#')[0], email_server_password
SMTP_HOST, SMTP_PORT = config['email']['smtp_host'].split('\t#')[0], config['email']['smtp_port'].split('\t#')[0]
TEST_MESSAGE = f"""To: {ADMIN_EMAIL}\nFrom: {SENDER_EMAIL}
Subject: Mailsender - тестовое сообщение\n
Это тестовое сообщение отправленное сервисом Mailsender.""".encode('utf8')
UNDELIVERED_MESSAGE = f"""To: {ADMIN_EMAIL}\nFrom: {SENDER_EMAIL}
Subject: Mailsender - недоставленное сообщение\n
Это сообщение отправленно сервисом Mailsender.\n""".encode('utf8')
IMAP_HOST, IMAP_PORT = config['email']['imap_host'].split('\t#')[0], config['email']['imap_port'].split('\t#')[0]
EMAIL_DB = config['email']['db'].split('\t#')[0]
EMAIL_DB_CONNECTION_STRING = config['email']['db_connection_string'].split('\t#')[0]
EMAIL_DB_TABLE_EMAILS = config['email']['db_table_emails'].split('\t#')[0]

REGEX_EMAIL_VALID = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' # шаблон для валидации e-mail адреса

ERROR_EMAIL_LIST = []  # список несуществующих адресов
error_email_list_path = 'error-emails-list.log'
if os.path.exists(error_email_list_path):
    with open(error_email_list_path, 'r') as f:
        for l in f.readlines():
            ERROR_EMAIL_LIST.append(l.split('\t')[-1].strip())

ROBOT_START = False
ROBOT_STOP = False
APP_EXIT = False
SIGN_IN_FLAG = False
THEME_COLOR = 'Gainsboro'
TK_FONT = 'Segoe UI'
LBL_COLOR = THEME_COLOR
LBL_ROBOT_MSG_COLOR = 'LightGray'
ENT_COLOR = 'White'
BTN_SIGN_IN_COLOR = 'Green'
BTN_START_COLOR = 'SeaGreen'
BTN_STOP_COLOR = 'SlateGray'
BTN_EXIT_COLOR = 'OrangeRed'
BTN_TEXT_COLOR = 'White'
BTN_EXIT_TEXT_COLOR = 'Black'
BTN_FONT = (TK_FONT, 12, 'bold')
RUNNER_COLOR = 'DodgerBlue'


# ПРОВЕРКА АРГУМЕНТОВ CMD ПРИ ЗАПУСКЕ ПРИЛОЖЕНИЯ - ДЛЯ ВЫБОРА РЕЖИМА РАБОТЫ (APPMODE_CONSOLE/APPMODE_INTERFACE)
# APPMODE_INTERFACE - none argv
# APPMODE_CONSOLE - MsgSender -console -username -userpassword -email/telegram/all-channels -cntrecs/all
APPMODE_INTERFACE, APPMODE_CONSOLE, IS_ALL_RECS = bool(), bool(), bool()
if len(sys.argv) == 1:
    APPMODE_INTERFACE = True
elif len(sys.argv) >= 5 and len(sys.argv) <= 6 and sys.argv[1] == '-console':
    APPMODE_CONSOLE = True
    user, password = sys.argv[2][1:], sys.argv[3][1:]
    if user != USER_NAME or password != USER_PASSWORD:
        print('Ошибка:  некорректные имя пользователя/пароль.')
        sys.exit()
    if sys.argv[4][1:] not in ('email', 'telegram', 'all-channels'):
        print('Ошибка:  в аргументах указан некорректный канал сообщений [email/telegram/all-channels].')
        print('Формат запуска:  MsgSender -console -username -userpassword -email/telegram/all-channels -cntrecs/all')
        sys.exit()
    MODE_EMAIL = True if sys.argv[4][1:] in ('email', 'all-channels') else False
    MODE_TELEGRAM = True if sys.argv[4][1:] in ('telegram', 'all-channels') else False
    if len(sys.argv) == 6:
        if sys.argv[5][1:].isdigit():
            IS_ALL_RECS = False
            CNT_RECS = int(sys.argv[5][1:])
            print('CNT_RECS', CNT_RECS)
        elif sys.argv[5][1:] == 'all':
            IS_ALL_RECS = True
        else:
            print('Ошибка аргумента количества записей  -  должен быть целым числом или all.')
            sys.exit()
    elif len(sys.argv) == 5:
        scheduler_handling_db_recs = config['common']['scheduler_handling_db_recs'].split('\t#')[0]
        if scheduler_handling_db_recs.isdigit():
            IS_ALL_RECS = False                         # флаг чтения всех записей из бд
            CNT_RECS = int(scheduler_handling_db_recs)  # кол-во записей читаемых из бд
            print('CNT_RECS', CNT_RECS)
        elif scheduler_handling_db_recs == 'all':
            IS_ALL_RECS = True
else:
    print('Ошибка запуска приложения.')
    print('Формат запуска:  MsgSender -console -username -userpassword -email/telegram/all-channels -cntrecs/all')
    sys.exit()


# === MESSENGER FUNCTIONS ===
async def robot():
    # запускает робота
    global ROBOT_START, ROBOT_STOP, ADMIN_BOT_CHAT_ID, MODE_EMAIL, MODE_TELEGRAM
    if ROBOT_START or ROBOT_STOP:
        return
    ROBOT_START = True  # флаг старта робота, предотвращает запуск нескольких экземпляров робота
    # режимы обработки сообщений: email, telegram
    if APPMODE_INTERFACE:
        MODE_EMAIL, MODE_TELEGRAM = cbt_msg_type_v1['email'].get(), cbt_msg_type_v1['telegram'].get()
    # подключение к базе данных TELEGRAM_DB
    if MODE_TELEGRAM:
        try:
            cnxn_telegram_db = await aioodbc.connect(dsn=TELEGRAM_DB_CONNECTION_STRING, loop=loop_robot)
            cursor_telegram_db = await cnxn_telegram_db.cursor()
            print(f'Создано подключение к базе данных telegram {TELEGRAM_DB}')  ###
        except Exception as e:
            print(f"Подключение к базе данных telegram {TELEGRAM_DB} -  ошибка.", e)
            return 1
    # подключение к базе данных EMAIL_DB
    if MODE_EMAIL:
        try:
            cnxn_email_db = await aioodbc.connect(dsn=EMAIL_DB_CONNECTION_STRING, loop=loop_robot)
            cursor_email_db = await cnxn_email_db.cursor()
            print(f'Создано подключение к базе данных email {EMAIL_DB}')  ###
        except Exception as e:
            print(f"Подключение к базе данных email {EMAIL_DB} -  ошибка.", e)
            return 1
    # чтение из бд данных о telegram-группах
    if MODE_TELEGRAM:
        telegram_chats, ADMIN_BOT_CHAT_ID = await load_telegram_chats_from_db(cursor_telegram_db)
        if telegram_chats == 1:
            await cursor_telegram_db.close()
            await cnxn_telegram_db.close()
            ROBOT_START, ROBOT_STOP = False, False
            if APPMODE_INTERFACE:
                lbl_msg_robot["text"] = f'Ошибка чтения из базы данных telegram {TELEGRAM_DB}'
            elif APPMODE_CONSOLE:
                print(f'Ошибка чтения из базы данных telegram {TELEGRAM_DB}')
            return 1

    if APPMODE_INTERFACE:
        lbl_msg_robot["text"] = 'Робот в рабочем режиме'
    elif APPMODE_CONSOLE:
            print('Робот в рабочем режиме')

    while not ROBOT_STOP:
        # обработка telegram-сообщений ====================================================================
        if MODE_TELEGRAM:
            telegram_msg_data_records = await load_records_from_telegram_db(cursor_telegram_db)
            if telegram_msg_data_records == 1:
                await cursor_telegram_db.close()
                await cnxn_telegram_db.close()
                ROBOT_START, ROBOT_STOP = False, False
                if APPMODE_INTERFACE:
                    lbl_msg_robot["text"] = f'Ошибка чтения из базы данных telegram {TELEGRAM_DB}'
                elif APPMODE_CONSOLE:
                    print(f'Ошибка чтения из базы данных telegram {TELEGRAM_DB}')
                return 1
            print(telegram_msg_data_records)
            print()
            if len(telegram_msg_data_records) > 0:
                await robot_send_telegram_msg(cnxn_telegram_db, cursor_telegram_db, telegram_msg_data_records, telegram_chats)
            else:
                print(f'Нет новых сообщений в базе данных telegram {TELEGRAM_DB}.')  ### test

        # обработка email-сообщений ======================================================================
        if MODE_EMAIL:
            email_msg_data_records = await load_records_from_email_db(cursor_email_db)
            if email_msg_data_records == 1:
                await cursor_email_db.close()
                await cnxn_email_db.close()
                ROBOT_START, ROBOT_STOP = False, False
                if APPMODE_INTERFACE:
                    lbl_msg_robot["text"] = f'Ошибка чтения из базы данных email {EMAIL_DB}'
                elif APPMODE_CONSOLE:
                    print(f'Ошибка чтения из базы данных email {EMAIL_DB}')
                return 1
            print(email_msg_data_records)
            print()
            if len(email_msg_data_records) > 0:
                await robot_send_email_msg(cnxn_email_db, cursor_email_db, email_msg_data_records)
            else:
                print(f'Нет новых сообщений в базе данных email {EMAIL_DB}.')  ### test

            #  email - недоставленные сообщения: проверка оповещений, запись в лог и отправка на почту админа
            undelivereds = await check_undelivered_emails(IMAP_HOST, SENDER_EMAIL, EMAIL_SERVER_PASSWORD)
            if len(undelivereds) > 0:
                smtp_client = SMTP(hostname=SMTP_HOST, port=SMTP_PORT, use_tls=True, username=SENDER_EMAIL, password=EMAIL_SERVER_PASSWORD)
                await smtp_client.connect()
                for u in undelivereds:
                    print(f'undelivered:  {u}')
                    log_rec = f'Недоставлено сообщение, отправленное {u[0]} на несуществующий адрес {u[1]}'
                    await rec_to_log(log_rec)

                    # запись несуществующего адреса в error-email-list
                    eel_rec = f'{u[0]}\t{u[1]}'
                    await rec_to_error_emails_list(eel_rec)
                    if u[1] not in ERROR_EMAIL_LIST:
                        ERROR_EMAIL_LIST.append(u[1])

                    msg = UNDELIVERED_MESSAGE + log_rec.encode('utf-8')
                    await smtp_client.sendmail(SENDER_EMAIL, ADMIN_EMAIL, msg)
                await smtp_client.quit()
            
        if APPMODE_CONSOLE:
            break

        await asyncio.sleep(CHECK_DB_PERIOD)

    #  действия после остановки робота
    if MODE_TELEGRAM:
        await cursor_telegram_db.close()
        await cnxn_telegram_db.close()
    if MODE_EMAIL:
        await cursor_email_db.close()
        await cnxn_email_db.close()
    print("Робот остановлен")
    ROBOT_START, ROBOT_STOP = False, False
    if APPMODE_INTERFACE:
        lbl_msg_robot["text"] = 'Робот остановлен'
    if APP_EXIT:
        sys.exit()


async def robot_send_email_msg(cnxn_email_db, cursor_email_db, email_msg_data_records):
    # отправляет почту
    smtp_client = SMTP(hostname=SMTP_HOST, port=SMTP_PORT, use_tls=True, username=SENDER_EMAIL, password=EMAIL_SERVER_PASSWORD)
    await smtp_client.connect() 
    for e in email_msg_data_records:
        # e =  (1, 'test1', 'This is the test message 1!', 'testbox283@yandex.ru; testbox283@mail.ru', 'at1.txt; at2.jpg')
        print("Новая запись в EMAIL_DB ", e)  ### test
        addrs = e[3].strip().split(';')
        for a in addrs:
            a = a.strip()
            if(re.fullmatch(REGEX_EMAIL_VALID, a)):
                if a not in ERROR_EMAIL_LIST:
                    if not e[4]:   #  если нет приложенных файлов формируется простое сообщение
                        msg = f'To: {a}\nFrom: {SENDER_EMAIL}\nSubject: {e[1]}\n\n{e[2]}'.encode("utf-8")
                    else:   # если есть приложенные файлы формируется составное сообщение
                        message = MIMEMultipart()
                        message['From'] = SENDER_EMAIL
                        message['To'] = a
                        message['Subject'] = e[1]
                        message_text = e[2]
                        message.attach(MIMEText(message_text, 'plain'))
                        files = e[4].strip().split(';')
                        for f in files:  # add files to the message
                            f = f.strip()
                            file_path = DIR_EMAIL_ATTACHMENTS / f
                            if not file_path.exists():
                                print(f'файл {f} не существует!')
                                continue

                            with open(file_path, 'rb') as fp:
                                # attach_file = open(file_path, 'rb')
                                attach_file = fp
                                payload = MIMEBase('application', 'octate-stream')
                                payload.set_payload((attach_file).read())
                                encoders.encode_base64(payload)
                                payload.add_header('Content-Disposition', 'attachment', filename=f)
                                message.attach(payload)

                        msg = message.as_string()
                    await smtp_client.sendmail(SENDER_EMAIL, a, msg)
                    log_rec = f'send message to {a} [ id = {e[0]} ]'
                else:
                    print(f'адрес {a} в error-list')   ###
                    log_rec = f'адрес {a} в error-list'  ###
            else:
                print(f'invalid email address {a} [ id = {e[0]} ]')  ###
                log_rec = f'invalid email address {a} [ id = {e[0]} ]'  ###
            await rec_to_log(log_rec)
        await set_record_handling_time_in_email_db(cnxn_email_db, cursor_email_db, id=e[0])
        print('Запись из EMAIL_DB обработана')  ####
    await smtp_client.quit()


async def check_undelivered_emails(host, user, password):
    # проверяет неотправленные сообщения, написано для imap@yandex, для других серверов может потребоваться корректировка функции
    imap_client = aioimaplib.IMAP4_SSL(host=host)
    await imap_client.wait_hello_from_server()
    await imap_client.login(user, password)
    await imap_client.select('INBOX')
    typ, msg_nums_unseen = await imap_client.search('UNSEEN')
    typ, msg_nums_from_subject = await imap_client.search('(FROM "mailer-daemon@yandex.ru" SUBJECT "Недоставленное сообщение")')
    msg_nums_unseen = set(msg_nums_unseen[0].decode().split())
    msg_nums_from_subject = set(msg_nums_from_subject[0].decode().split())
    msg_nums = ' '.join(list(msg_nums_unseen & msg_nums_from_subject))
    #msg_nums = '7 8'  #  для разработки
    l = len(msg_nums.split())
    if l == 0:
        print('Нет новых сообщений о недоставленной почте')    ###
        await imap_client.close()
        await imap_client.logout()
        return []
    print(f'Получено {l} оповещений о недоставленной почте')    ###
    msg_nums = msg_nums.replace(' ', ',')
    typ, data = await imap_client.fetch(msg_nums, '(UID BODY[TEXT])')

    undelivered = []
    for m in range(1, len(data), 3):
        msg = email.message_from_bytes(data[m])
        msg = msg.get_payload()
        msg_arrival_date = re.search(r'(?<=Arrival-Date: ).*', msg)[0].strip()
        msg_recipient = re.search(r'(?<=Original-Recipient: rfc822;).*', msg)[0].strip()
        print(msg_arrival_date, msg_recipient)  ###
        undelivered.append((msg_arrival_date, msg_recipient))

    await imap_client.close()
    await imap_client.logout()

    print(undelivered)  ###
    return undelivered


async def robot_send_telegram_msg(cnxn_telegram_db, cursor_telegram_db, msg_data_records, telegram_chats):
    # отправляет сообщения через telegram
    for record in msg_data_records:
        # структура данных record =  (1, 'This is the test message 1!', 'test-group-1; test-group-2', 'file1.txt; file2.txt')
        print("Новая запись в TELEGRAM_DB ", record)  ###
        record_id = record[0]
        record_msg = record[1]
        record_addresses = record[2].strip().split(';')
        record_attachments = record[3].strip().split(';') if record[3] else None

        for address in record_addresses:
            address = address.strip()
            if address not in telegram_chats:
                # print(f'Бот не является участником группы {address} или группа не добавлена в базу данных.\nСообщение не отправлено.\n')
                print(f'У бота не создан чат с {address} или чат не добавлен в базу данных.\nСообщение не отправлено.\n')
                # добавить оповещение админа, только 1 раз
                msg = (f"Получен запрос на отправку сообщения в чат с {address}, в котором бот не участвует, " +
                        "или чат не добавлен в базу данных.\n" +
                        f"Запись в таблице {TELEGRAM_DB_TABLE_MESSAGES} с id={record_id}")
                url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage?chat_id={ADMIN_BOT_CHAT_ID}&text={msg}"
                requests.get(url).json()
                continue
            chat_id = telegram_chats[address]
            print(f'Отправка сообщения {address}', 'chat_id =', chat_id)
            url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage?chat_id={chat_id}&text={record_msg}"
            try:
                requests.get(url).json()
                print(f'Сообщение {address} отправлено.\n')
            except Exception as e:
                print('Ошибка отправки:\n', e)

            # обработка поля attachments - отправка файлов
            if not record_attachments:   #  если нет приложенных файлов
                continue
            # если есть приложенные файлы
            for document in record_attachments:
                document = document.strip()
                file_path = DIR_TELEGRAM_ATTACHMENTS / document
                
                if not file_path.exists():
                    print(f'файл {document} не существует!')
                    continue

                with open(file_path, "rb") as f:
                    files = {"document": f}
                    try:
                        r = requests.post(f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument", data={"chat_id":chat_id}, files=files)
                        if r.status_code != 200:
                            raise Exception()
                        print(f'Файл {document} отправлен {address}.\n')
                    except Exception as e:
                        print('Ошибка отправки:\n', e)    
                    
        await set_record_handling_time_in_telegram_db(cnxn_telegram_db, cursor_telegram_db, record_id)
        print('Запись из TELEGRAM_DB обработана.')  ####


# === DATABASE FUNCTIONS ===
async def load_telegram_chats_from_db(cursor_telegram_db):
    # выборка из базы данных параметров telegram-чатов бота
    try:
        query = f"select entity_name, chat_id, entity_type from {TELEGRAM_DB_TABLE_CHATS} where bot_name='{BOT_NAME}' and is_active"
        await cursor_telegram_db.execute(query)
        rows = await cursor_telegram_db.fetchall()
        telegram_chats_dict = {row[0]: row[1] for row in rows}
        ADMIN_BOT_CHAT_ID = [row[1] for row in rows if row[2] == 'administrator'][0]
        return telegram_chats_dict, ADMIN_BOT_CHAT_ID
    except Exception as e:
        print('Ошибка чтения из базы данных TELEGRAM_DB.', e)
        return 1


async def load_records_from_email_db(cursor_email_db):
    # выборка из базы данных EMAIL_DB необработанных (новых) записей
    try:
        await cursor_email_db.execute(f"""select UniqueIndexField, subj, textemail, adrto, attachmentfiles from {EMAIL_DB_TABLE_EMAILS} 
                where dates is null order by datep""")
        rows = await cursor_email_db.fetchall()  # список кортежей
    except Exception as e:
        print(f'Ошибка чтения из базы данных EMAIL_DB {EMAIL_DB}.', e)
        return 1

    if APPMODE_CONSOLE:
        rows = rows[:CNT_RECS] if not IS_ALL_RECS and len(rows) > 0 else rows

    return rows

async def load_records_from_telegram_db(cursor_telegram_db):
    # выборка из базы данных TELEGRAM_DB необработанных (новых) записей
    try:
        await cursor_telegram_db.execute(f"""select UniqueIndexField, msg_text, adrto, attachmentfiles from {TELEGRAM_DB_TABLE_MESSAGES} 
            where dates is null order by datep""")
        rows = await cursor_telegram_db.fetchall()  # список кортежей
    except Exception as e:
        print(f'Ошибка чтения из базы данных TELEGRAM_DB {TELEGRAM_DB}.', e)
        return 1

    if APPMODE_CONSOLE:
        rows = rows[:CNT_RECS] if not IS_ALL_RECS and len(rows) > 0 else rows

    return rows


async def set_record_handling_time_in_email_db(cnxn_email_db, cursor_email_db, id):
    # пишет в базу EMAIL_DB дату/время отправки сообщения
    dt_string = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    await cursor_email_db.execute(f"update {EMAIL_DB_TABLE_EMAILS} set dates = '{dt_string}' where UniqueIndexField = {id}")
    await cnxn_email_db.commit()

async def set_record_handling_time_in_telegram_db(cnxn_telegram_db, cursor_telegram_db, id):
    # пишет в базу TELEGRAM_DB дату/время отправки сообщения
    dt_string = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    await cursor_telegram_db.execute(f"update {TELEGRAM_DB_TABLE_MESSAGES} set dates = '{dt_string}' where UniqueIndexField = {id}")
    await cnxn_telegram_db.commit()


async def rec_to_error_emails_list(rec):
    # добавляет несуществующий email адрес в error_emails_list
    current_time = str(datetime.datetime.now())
    with open('error-emails-list.log', 'a') as f:
        f.write(f'{current_time}\t{rec}\n')


async def rec_to_log(rec):
    # пишет в лог-файл запись об отправке сообщения
    current_time = str(datetime.datetime.now())
    with open('log-mailsender.log', 'a') as f:
        f.write(f'{current_time}\t{rec}\n')



# === INTERFACE FUNCTIONS ===
async def btn_sign_click():
    # кнопка sign-in
    global SIGN_IN_FLAG
    user = ent_user.get()
    password = ent_password.get()
    if user == USER_NAME and password == USER_PASSWORD:
        lbl_msg_sign["text"] = ''
        SIGN_IN_FLAG = True
        root.destroy()
    else:
        lbl_msg_sign["text"] = 'Incorrect username or password'

async def show_password_signin():
    # показывает/скрывает пароль в окне входа
    ent_password['show'] = '' if(cbt_sign_show_pwd_v1.get() == 1) else '*'

async def btn_exit_click():
    # кнопка Send test email
    global ROBOT_START, ROBOT_STOP, APP_EXIT
    if ROBOT_START:
        lbl_msg_robot["text"] = 'Остановка робота...\nВыход из приложения...'
        ROBOT_STOP = True
        APP_EXIT = True
    else:
        sys.exit()

async def btn_robot_run_click():
    # кнопка Start robot
    global ROBOT_START, ROBOT_STOP
    if not ROBOT_START:
        if not (cbt_msg_type_v1['email'].get() or cbt_msg_type_v1['telegram'].get()):
            lbl_msg_robot["text"] = 'Выберите сообщения для обработки'
            return 1
        lbl_msg_robot["text"] = 'Запуск робота...'
        # при запуске робота checkbuttons выбора сообщений деактивируются
        cbt_msg_type['email']['state'], cbt_msg_type['telegram']['state'] = 'disabled', 'disabled'
        await asyncio.sleep(1)
        await robot()
        # после остановки или незапуска робота checkbuttons выбора сообщений активируются
        cbt_msg_type['email']['state'], cbt_msg_type['telegram']['state'] = 'normal', 'normal'

async def btn_robot_stop_click():
    # кнопка Stop robot
    global ROBOT_START, ROBOT_STOP
    if ROBOT_START:
        lbl_msg_robot["text"] = 'Остановка робота...'
        ROBOT_STOP = True

async def window_signin():
    # рисует окно входа
    frm.pack()
    lbl_user.place(x=95, y=43)
    ent_user.place(x=95, y=86)
    lbl_password.place(x=95, y=110)
    ent_password.place(x=95, y=153)
    cbt_sign_show_pwd.place(x=95, y=180)
    btn_sign.place(x=95, y=220)
    lbl_msg_sign.place(x=95, y=270)

async def window_robot():
    # рисует окно админки
    frm.pack()
    #lbl_robot.place(x=95, y=30)
    btn_robot_run.place(x=30, y=30)
    btn_robot_stop.place(x=30, y=78)
    btn_exit.place(x=60, y=126)
    lbl_runner.place(x=200, y=187)
    lbl_msg_robot.place(x=30, y=217)

    lbl_msg_type.place(x=300, y=30)
    cbt_msg_type['email'].place(x=300, y=63)
    cbt_msg_type['telegram'].place(x=300, y=93)



# ===========  ЗАПУСК ПРИЛОЖЕНИЯ В APPMODE_CONSOLE
if APPMODE_CONSOLE:
    loop_robot = asyncio.get_event_loop()
    loop_robot.run_until_complete(robot())
    sys.exit()


# ============== window sign in
root = tk.Tk()
root.resizable(0, 0)  # делает неактивной кнопку Развернуть
root.title('MessageSender')
frm = tk.Frame(bg=THEME_COLOR, width=400, height=400)
lbl_sign = tk.Label(master=frm, text='Sign in to MessageSender', bg=LBL_COLOR, font=(TK_FONT, 15), width=21, height=2)
lbl_user = tk.Label(master=frm, text='Username', bg=LBL_COLOR, font=(TK_FONT, 12), anchor='w', width=25, height=2)
ent_user = tk.Entry(master=frm, bg=ENT_COLOR, font=(TK_FONT, 12), width=25, )
lbl_password = tk.Label(master=frm, text='Password', bg=LBL_COLOR, font=(TK_FONT, 12), anchor='w', width=25, height=2)
ent_password = tk.Entry(master=frm, show='*', bg=ENT_COLOR, font=(TK_FONT, 12), width=25, )

cbt_sign_show_pwd_v1 = tk.IntVar(value = 0)
cbt_sign_show_pwd = tk.Checkbutton(frm, bg=THEME_COLOR, text='Show password', variable=cbt_sign_show_pwd_v1, onvalue=1, offvalue=0, 
                                    command=lambda: loop.create_task(show_password_signin()))

btn_sign = tk.Button(master=frm, bg=BTN_SIGN_IN_COLOR, fg='White', text='Sign in', font=(TK_FONT, 12, "bold"), 
                    width=22, height=1, command=lambda: loop.create_task(btn_sign_click()))
lbl_msg_sign = tk.Label(master=frm, bg=LBL_COLOR, fg='PaleVioletRed', font=(TK_FONT, 12), width=25, height=2)

async def show():
    # показывает и обновляет окно входа
    await window_signin()
    while not SIGN_IN_FLAG:
        root.update()
        await asyncio.sleep(.01)

development_mode = False     # True - для разработки окна робота переход сразу на него без sign in
if development_mode:    # для разработки окна робота переход сразу на него без sign in
    SIGN_IN_FLAG = True
else:
    loop = asyncio.get_event_loop()
    loop.run_until_complete(show())

# выход из приложения если принудительно закрыто окно логина
# c asyncio не работает, надо выяснять!
if not SIGN_IN_FLAG:
    print('SIGN IN FALSE')
    #print('loop = ', loop)
    sys.exit()


# ============== window robot
root_robot = tk.Tk()
root_robot.resizable(0, 0)  # делает неактивной кнопку Развернуть
root_robot.title('MessageSender')
frm = tk.Frame(bg=THEME_COLOR, width=555, height=290)
#lbl_robot = tk.Label(master=frm, text='MessageSender', bg=LBL_COLOR, font=("Arial", 15), width=20, height=2)
btn_robot_run = tk.Button(master=frm, bg=BTN_START_COLOR, fg=BTN_TEXT_COLOR, text='Запуск робота', font=BTN_FONT, 
                    width=22, height=1, command=lambda: loop_robot.create_task(btn_robot_run_click()))
btn_robot_stop = tk.Button(master=frm, bg=BTN_STOP_COLOR, fg=BTN_TEXT_COLOR, text='Остановка робота', font=BTN_FONT, 
                    width=22, height=1, command=lambda: loop_robot.create_task(btn_robot_stop_click()))
btn_exit = tk.Button(master=frm, bg=BTN_EXIT_COLOR, fg=BTN_EXIT_TEXT_COLOR, text='Выход', font=BTN_FONT, 
                    width=16, height=1, command=lambda: loop_robot.create_task(btn_exit_click()))
animation = "░▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒"
lbl_runner = tk.Label(master=frm, fg=RUNNER_COLOR, text="", font=(TK_FONT, 4))
lbl_msg_robot = tk.Label(master=frm, bg=LBL_ROBOT_MSG_COLOR, font=(TK_FONT, 10), width=70, height=2)


lbl_msg_type = tk.Label(master=frm, text='Выбор сообщений для обработки:', bg=THEME_COLOR, font=(TK_FONT, 10, 'bold'), width=27, height=1)
cbt_msg_type_v1, cbt_msg_type = {}, {}
cbt_msg_type_v1['email'] = tk.IntVar(value=0)
cbt_msg_type_v1['telegram'] = tk.IntVar(value=0)
cbt_msg_type['email'] = tk.Checkbutton(master=frm, bg=THEME_COLOR, text = 'E-mail сообщения', font=(TK_FONT, 10),
                variable = cbt_msg_type_v1['email'], 
                onvalue = 1, offvalue = 0)
cbt_msg_type['telegram'] = tk.Checkbutton(master=frm, bg=THEME_COLOR, text = 'Telegram сообщения', font=(TK_FONT, 10),
                variable = cbt_msg_type_v1['telegram'], 
                onvalue = 1, offvalue = 0)

async def show_robot():
    # показывает и обновляет окно робота
    global animation

    await window_robot()
    while True:
        lbl_runner["text"] = animation
        if ROBOT_START:
            # animation = animation[1:] + animation[0]
            animation = animation[-1] + animation[:-1]

        root_robot.update()
        await asyncio.sleep(.1)

loop_robot = asyncio.get_event_loop()
loop_robot.run_until_complete(show_robot())
