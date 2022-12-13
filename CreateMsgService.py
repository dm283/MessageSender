import sys, configparser, re, datetime, asyncio, shutil, tkinter as tk
import aioodbc
from cryptography.fernet import Fernet
from pathlib import Path
from tkinter import ttk, filedialog

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

DIR_EMAIL_ATTACHMENTS = Path(config['common']['dir_email_attachments'].split('\t#')[0])  # директория с файлами для отправки
DIR_TELEGRAM_ATTACHMENTS = Path(config['common']['dir_telegram_attachments'].split('\t#')[0])
USER_NAME = config['user_credentials']['name'].split('\t#')[0]
USER_PASSWORD = user_credentials_password
BOT_NAME = config['telegram_bot']['bot_name'].split('\t#')[0]
TELEGRAM_DB = config['telegram_bot']['db'].split('\t#')[0]  # база данных mssql/posgres
TELEGRAM_DB_TABLE_MESSAGES = config['telegram_bot']['db_table_messages'].split('\t#')[0]  # db.schema.table
TELEGRAM_DB_TABLE_CHATS = config['telegram_bot']['db_table_chats'].split('\t#')[0]  # db.schema.table  таблица с telegram-чатами
TELEGRAM_DB_CONNECTION_STRING = config['telegram_bot']['db_connection_string'].split('\t#')[0]  # odbc driver system dsn name
MODE_EMAIL, MODE_TELEGRAM = bool(), bool()
SENDER_EMAIL, EMAIL_SERVER_PASSWORD = config['email']['sender_email'].split('\t#')[0], email_server_password
SMTP_HOST, SMTP_PORT = config['email']['smtp_host'].split('\t#')[0], config['email']['smtp_port'].split('\t#')[0]
EMAIL_DB = config['email']['db'].split('\t#')[0]
EMAIL_DB_CONNECTION_STRING = config['email']['db_connection_string'].split('\t#')[0]
EMAIL_DB_TABLE_EMAILS = config['email']['db_table_emails'].split('\t#')[0]

REGEX_EMAIL_VALID = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' # шаблон для валидации e-mail адреса

RECORDS_EMAIL, RECORDS_TELEGRAM = [], []  # записи из базы данных сообщений
RECORDS_EMAIL_POINTER, RECORDS_TELEGRAM_POINTER = 0, 0  # указатель срезов записей (по блокам 10 штук)

SIGN_IN_FLAG = False
THEME_COLOR = 'Gainsboro'
TK_FONT = 'Segoe UI'
LBL_COLOR = THEME_COLOR
BTN_INSERT_COLOR = 'SeaGreen'
ENT_COLOR = 'White'
BTN_SIGN_IN_COLOR = 'Green'


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


async def btn_email_insert_db_click():
    # Кнопка записи в бд email-сообщения
    adrto = ent['email']['to'].get().strip()
    subj = ent['email']['subj'].get().strip()
    textemail = ent['email']['msg_text'].get(1.0, "end-1c").strip()
    attachments = ent['email']['attachments'].get(1.0, "end-1c").strip()
    
    if adrto == '' or subj == '' or textemail == '':
        lbl_msg_send['email']['text'] = 'Заполните все поля e-mail сообщения.'
        return 1

    # Проверка корректности e-mail адресов
    addrs = adrto.split(';')
    for a in addrs:
        a = a.strip()
        if not (re.fullmatch(REGEX_EMAIL_VALID, a)):
            lbl_msg_send['email']['text'] = f'Некорректный e-mail адрес {a}.'
            return 1

    try:
        cnxn = await aioodbc.connect(dsn=EMAIL_DB_CONNECTION_STRING, loop=loop_msg_service)
        cursor = await cnxn.cursor()
    except:
        lbl_msg_send['email']['text'] = f"Подключение к базе данных {EMAIL_DB} -  ошибка."
        return 1

    lbl_msg_send['email']['text'] = f'Запись в базу данных .....'
    await asyncio.sleep(0.5)

    datep = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    query = f""" insert into {EMAIL_DB_TABLE_EMAILS} (subj, textemail, adrto, attachmentfiles, datep) values
                ('{subj}', '{textemail}', '{adrto}', '{attachments}', '{datep}') """
    try:
        await cursor.execute(query)
        await cnxn.commit()
        lbl_msg_send['email']['text'] = f'Записано в базу данных.'
    except:
        lbl_msg_send['email']['text'] = f"Ошибка записи в базу данных {EMAIL_DB} -  ошибка."
        await cursor.close()
        await cnxn.close()
        return 1

    await cursor.close()
    await cnxn.close()


async def btn_telegram_insert_db_click():
    # Кнопка записи в бд telegram-сообщения
    msg_text = ent['telegram']['msg_text'].get(1.0, "end-1c").strip()
    adrto = cmbx['telegram']['to'].get().strip()
    attachments = ent['telegram']['attachments'].get(1.0, "end-1c").strip()

    if adrto == '' or msg_text == '':
        lbl_msg_send['telegram']['text'] = 'Заполните все поля telegram сообщения'
        return 1

    try:
        cnxn = await aioodbc.connect(dsn=TELEGRAM_DB_CONNECTION_STRING, loop=loop_msg_service)
        cursor = await cnxn.cursor()
    except:
        lbl_msg_send['telegram']['text'] = f"Подключение к базе данных {TELEGRAM_DB} -  ошибка."
        return 1

    lbl_msg_send['telegram']['text'] = f'Запись в базу данных .....'
    await asyncio.sleep(0.5)

    datep = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    query = f""" insert into {TELEGRAM_DB_TABLE_MESSAGES} (msg_text, adrto, attachmentfiles, datep) values
                ('{msg_text}', '{adrto}', '{attachments}', '{datep}') """
    try:
        await cursor.execute(query)
        await cnxn.commit()
        lbl_msg_send['telegram']['text'] = f'Записано в базу данных.'
    except:
        lbl_msg_send['telegram']['text'] = f"Ошибка записи в базу данных {TELEGRAM_DB} -  ошибка."
        await cursor.close()
        await cnxn.close()
        return 1

    await cursor.close()
    await cnxn.close()


async def load_records_from_db(mode):
    # выборка записей из базы данных
    global RECORDS_EMAIL, RECORDS_TELEGRAM
    if mode == 'email':
        dsn = EMAIL_DB_CONNECTION_STRING
        current_db = EMAIL_DB
        query = f"""select UniqueIndexField, adrto, subj, textemail, attachmentfiles, datep, dates
            from {EMAIL_DB_TABLE_EMAILS} order by datep desc"""
    elif mode == 'telegram':
        dsn = TELEGRAM_DB_CONNECTION_STRING
        current_db = TELEGRAM_DB
        query = f"""select UniqueIndexField, adrto, msg_text, attachmentfiles, datep, dates 
            from {TELEGRAM_DB_TABLE_MESSAGES} order by datep desc"""
    try:
        cnxn = await aioodbc.connect(dsn=dsn, loop=loop_msg_service)
        cursor = await cnxn.cursor()
    except:
        print(f"Подключение к базе данных {current_db} -  ошибка.")
        await cursor.close()
        await cnxn.close()
        return 1
    try:
        await cursor.execute(query)
        if mode == 'email':
            RECORDS_EMAIL = await cursor.fetchall()  # список кортежей
        elif mode == 'telegram':
            RECORDS_TELEGRAM = await cursor.fetchall()  # список кортежей
        await cursor.close()
        await cnxn.close()
    except Exception as e:
        print(f'Ошибка чтения из базы данных {current_db}.', e)
        await cursor.close()
        await cnxn.close()
        return 1


async def btn_load_records_from_db_click(mode):
    # кнопка ЗАГРУЗИТЬ ИЗ БД
    global RECORDS_EMAIL_POINTER, RECORDS_TELEGRAM_POINTER

    await load_records_from_db(mode)

    if mode == 'email':
        records_from_db = RECORDS_EMAIL
        RECORDS_EMAIL_POINTER = 0
        records_pointer = RECORDS_EMAIL_POINTER
        labels_list = ['id', 'adrto', 'subj', 'textemail', 'attachments', 'datep', 'dates', ]
        labels_list_check_len = ['adrto', 'subj', 'textemail', 'attachments']
    elif mode == 'telegram':
        records_from_db = RECORDS_TELEGRAM
        RECORDS_TELEGRAM_POINTER = 0
        records_pointer = RECORDS_TELEGRAM_POINTER
        labels_list = ['id', 'adrto', 'msg_text', 'attachments', 'datep', 'dates']
        labels_list_check_len = ['adrto', 'msg_text', 'attachments']

    # Отрисовка шапки таблицы сообщений
    c = 0
    for l in labels_list:
        lbl_sent_messages_header[mode][l].grid(row=0, column=c, sticky='w', padx=1, pady=1)
        c += 1
    # Заполнение таблицы сообщений
    cnt_rows = len(records_from_db) if len(records_from_db) < 10 else 10
    await fill_msg_table(mode, labels_list, labels_list_check_len, records_from_db, records_pointer, cnt_rows)
    for i in range(cnt_rows):
        column = 0
        for l in labels_list:
            lbl_message[mode][l][i].grid(row=i+1, column=column, sticky='w', padx=1, pady=1)
            column += 1


async def fill_msg_table(mode, labels_list, labels_list_check_len, records_from_db, records_pointer, cnt_rows):
    # Заполнение/Обновление таблицы сообщений
    row = 0
    start = 10*records_pointer
    finish = cnt_rows+10*records_pointer
    if finish > len(records_from_db):
        finish = len(records_from_db)
    for i in range(start, finish):
        row += 1
        column = 0
        for l in labels_list:
            text_str = records_from_db[i][column] if records_from_db[i][column] else ''
            # Замена переноса строк на пробелы
            if l in labels_list_check_len and ('\n' in text_str or '\r' in text_str):
                text_str = text_str.replace('\n', ' ').replace('\r', ' ')
            # Обрезка длинных строк
            if l == 'adrto' and len(text_str) > 19:
                text_str = text_str[:19] + ' ...'
            elif l == 'subj' and len(text_str) > 17:
                text_str = text_str[:17] + ' ...'
            elif l == 'textemail' and len(text_str) > 37:
                text_str = text_str[:37] + ' ...'   
            elif l == 'msg_text' and len(text_str) > 57:
                text_str = text_str[:57] + ' ...'
            elif l == 'attachments' and text_str:
                cnt_files = len(text_str.split(';'))
                if cnt_files > 1:
                    text_str = f'{cnt_files} файлов'
                if cnt_files == 1 and len(text_str) > 19:
                    text_str = text_str[:17] + ' ...'       
            lbl_message[mode][l][row-1]['text'] = text_str
            column += 1
    # Если строк в срезе менее 10, оставшиеся заполняются пустыми значениями
    for r in range(row, 10):
        for l in labels_list:
            lbl_message[mode][l][r]['text'] = ''
    lbl_header_records_numbers[mode]['text'] = f'{start+1}-{finish} из {len(records_from_db)}'


async def btn_slice_msg_click(mode: str, direction: int):
    # Кнопки перемещения по отправленным сообщениям
    global RECORDS_EMAIL_POINTER, RECORDS_TELEGRAM_POINTER
    if mode == 'email':
        records_from_db = RECORDS_EMAIL
        records_pointer = RECORDS_EMAIL_POINTER
    elif mode == 'telegram':
        records_from_db = RECORDS_TELEGRAM
        records_pointer = RECORDS_TELEGRAM_POINTER
    if len(records_from_db) < 10:  # Если сообщений < 10 передвижения по срезам нет
        return 0
    if (records_pointer + direction) < 0:  # Ограничение снизу изменения RECORDS_EMAIL_POINTER
        return 0
    if (records_pointer + direction) >= len(records_from_db)/10:  # Ограничение сверху изменения RECORDS_EMAIL_POINTER
        return 0
    cnt_rows = len(records_from_db) if len(records_from_db) < 10 else 10
    if mode == 'email':
        RECORDS_EMAIL_POINTER += direction
        labels_list = ['id', 'adrto', 'subj', 'textemail', 'attachments', 'datep', 'dates', ]
        labels_list_check_len = ['adrto', 'subj', 'textemail', 'attachments']
        await fill_msg_table(mode, labels_list, labels_list_check_len, records_from_db, RECORDS_EMAIL_POINTER, cnt_rows)
    elif mode == 'telegram':
        RECORDS_TELEGRAM_POINTER += direction
        labels_list = ['id', 'adrto', 'msg_text', 'attachments', 'datep', 'dates']
        labels_list_check_len = ['adrto', 'msg_text', 'attachments']
        await fill_msg_table(mode, labels_list, labels_list_check_len, records_from_db, RECORDS_TELEGRAM_POINTER, cnt_rows)


async def load_from_telegram_db(event):
    # выборка записей из базы данных TELEGRAM_DB
    try:
        cnxn = await aioodbc.connect(dsn=TELEGRAM_DB_CONNECTION_STRING, loop=loop_msg_service)
        cursor = await cnxn.cursor()
    except:
        print(f"Подключение к базе данных {TELEGRAM_DB} -  ошибка.")
        await cursor.close()
        await cnxn.close()
        return 1
    entity_type = cmbx['telegram']['entity'].get()
    entity_types = "'user', 'administrator'" if entity_type == 'Telegram-пользователь' else "'group'"
    try:
        query = f"""select distinct entity_name from {TELEGRAM_DB_TABLE_CHATS} 
            where entity_type in ({entity_types}) and bot_name='{BOT_NAME}'"""
        await cursor.execute(query)
        res = await cursor.fetchall()
        await cursor.close()
        await cnxn.close()
        cmbx['telegram']['to']['values'] = [v[0] for v in res]
    except:
        print(f'Ошибка чтения из базы данных {TELEGRAM_DB}.')   ###
        await cursor.close()
        await cnxn.close()
        return 1


async def btn_attached_files_path_click(mode):
    # добавляет к записи прикрепленные файлы
    filepath = filedialog.askopenfilenames()
    if filepath != "":
        filepath_dir_attachments = list(map(lambda x: x.split('/')[-1], list(filepath)))
        postfix = '; ' if ent[mode]['attachments'].get(1.0, "end-1c").strip() else ''
        ent[mode]['attachments'].insert("1.0", '; '.join(filepath_dir_attachments) + postfix)
    # копирование выбранных файлов в папки вложений
    dst_dir = DIR_EMAIL_ATTACHMENTS if mode == 'email' else DIR_TELEGRAM_ATTACHMENTS
    for file in filepath:
        dest = shutil.copy(file, dst_dir)


async def show_signin():
    # рисует окно входа
    frm.pack()
    lbl_user.place(x=95, y=43)
    ent_user.place(x=95, y=86)
    lbl_password.place(x=95, y=110)
    ent_password.place(x=95, y=153)
    cbt_sign_show_pwd.place(x=95, y=180)
    btn_sign.place(x=95, y=220)
    lbl_msg_sign.place(x=95, y=270)

    while not SIGN_IN_FLAG:
        root.update()
        await asyncio.sleep(.01)


async def show_send_msg():
    # рисует окно записи сообщений    
    notebook.pack(padx=10, pady=10, fill='both', expand=True)
    for m in ['email', 'telegram']:
        notebook.add(frm[m], text='E-mail' if m=='email' else 'Telegram')
        frm_msg_form[m].pack(padx=5, pady=(5, 0), fill='both', expand=True)
        lbl[m]['description'].grid(row=0, columnspan=2, sticky='w', padx=5, pady=5)
        ent[m]['msg_text'].grid(row=3, columnspan=2, sticky='w', padx=5, pady=5)
        lbl[m]['attachments'].grid(row=4, column=0, sticky='w', padx=5, pady=5)
        ent[m]['attachments'].grid(row=4, column=1, sticky='w', padx=5, pady=5)
        frm_sending[m].pack(padx=5, pady=(1, 5), fill='both', expand=True)
        btn_send[m].grid(row=0, column=0, padx=5, pady=5)
        btn_attached_files_path[m].grid(row=0, column=1, padx=5, pady=5)
        lbl_msg_send[m].grid(row=0, column=2, padx=5, pady=5)
        frm_sent_msg_header[m].pack(padx=5, pady=(1, 5), fill='both', expand=True)
        lbl_header_title[m].grid(row=0, column=0, sticky='w', padx=5, pady=5)
        btn_load_msg_from_db[m].grid(row=0, column=1, sticky='w', padx=5, pady=5)
        lbl_header_records_numbers[m].grid(row=0, column=2, sticky='w', padx=5, pady=5)
        btn_prev[m].grid(row=0, column=3, sticky='w', padx=5, pady=5)
        btn_next[m].grid(row=0, column=4, sticky='w', padx=5, pady=5)
        frm_sent_messages[m].pack(padx=5, pady=(1, 5), fill='both', expand=True)
    # уникальные виджеты или их расположение
    lbl['email']['to'].grid(row=1, column=0, sticky='w', padx=5, pady=5)
    ent['email']['to'].grid(row=1, column=1, sticky='w', padx=5, pady=5)
    lbl['email']['subj'].grid(row=2, column=0, sticky='w', padx=5, pady=5)
    ent['email']['subj'].grid(row=2, column=1, sticky='w', padx=5, pady=5)
    lbl['telegram']['entity'].grid(row=1, column=0, sticky='w', padx=5, pady=5)
    cmbx['telegram']['entity'].grid(row=1, column=1, sticky='w', padx=5, pady=5)
    lbl['telegram']['to'].grid(row=2, column=0, sticky='w', padx=5, pady=5)
    cmbx['telegram']['to'].grid(row=2, column=1, sticky='w', padx=5, pady=5)

    while True:
        root_send_msg.update()
        await asyncio.sleep(.01)

# ============== window sign in
root = tk.Tk()
root.resizable(0, 0)  # делает неактивной кнопку Развернуть
root.title('CreateMsgService')
frm = tk.Frame(bg=THEME_COLOR, width=400, height=350)
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


development_mode = False     # True - для разработки окна робота переход сразу на него без sign in
if development_mode:    # для разработки окна робота переход сразу на него без sign in
    SIGN_IN_FLAG = True
else:
    loop = asyncio.get_event_loop()
    loop.run_until_complete(show_signin())

# выход из приложения если принудительно закрыто окно логина
# c asyncio не работает, надо выяснять!
if not SIGN_IN_FLAG:
    print('SIGN IN FALSE')
    #print('loop = ', loop)
    sys.exit()

# ============== window send_message
root_send_msg = tk.Tk()
root_send_msg.resizable(0, 0)  # делает неактивной кнопку Развернуть
root_send_msg.title('CreateMsgService')
notebook = ttk.Notebook(root_send_msg)

frm, frm_msg_form, frm_sending, lbl, ent = {}, {}, {}, {}, {}

# Размеры виджетов
lbl_form_width = 13  # Фрейм №1  -  labels наименований полей ввода данных
ent_form_width = 129  # Фрейм №1  -  entry полей ввода данных
ent_form_msg_width = 147  # Фрейм №1  -  text сообщения ширина
ent_form_msg_height = 6  # Фрейм №1  -  text сообщения высота
lbl_msg_send_width = 89  # Фрейм №2  -  label сервисных сообщений
lbl_header_title_width = 28  # Фрейм №3  -  заголовок База данных ... сообщений
btn_move_width = 12  #  Фрейм №3 -  кнопки перемещения по срезам сообщений
frm_sent_messages_height = 230  #  Фрейм №4  -  фрейм таблицы сообщений
telegram_combo_width = 126   #  Фрейм №1 telegram - comboboxы

# Вкладка email-сообщений =============================================================================
frm['email'] = tk.Frame(notebook, bg=THEME_COLOR, width=400, )

# === Фрейм №1 - формы сообщения ===
frm_msg_form['email'], lbl['email'], ent['email'] = tk.Frame(frm['email'], bg=THEME_COLOR, width=400, ), {}, {}
lbl['email']['description'] = tk.Label(frm_msg_form['email'], bg=THEME_COLOR, text='Запись в базу данных e-mail сообщений', 
                                        font=('Segoe UI', 10, 'bold'))
lbl['email']['to'] = tk.Label(frm_msg_form['email'], bg=THEME_COLOR,
            text = 'Адреса (через ;):', width=lbl_form_width, anchor='w', )
ent['email']['to'] = tk.Entry(frm_msg_form['email'], width=ent_form_width, highlightthickness=1, highlightcolor = "Gainsboro", )
lbl['email']['subj'] = tk.Label(frm_msg_form['email'], bg=THEME_COLOR,
            text = 'Тема:', width=lbl_form_width, anchor='w', )
ent['email']['subj'] = tk.Entry(frm_msg_form['email'], width=ent_form_width, highlightthickness=1, highlightcolor = "Gainsboro", )
ent['email']['msg_text'] = tk.Text(frm_msg_form['email'], width=ent_form_msg_width, height=ent_form_msg_height, 
    highlightthickness=1, highlightcolor = "Gainsboro", font=((TK_FONT, 9)))
lbl['email']['attachments'] = tk.Label(frm_msg_form['email'], bg=THEME_COLOR,
            text = 'Файлы:', width=lbl_form_width, anchor='w', )
ent['email']['attachments'] = tk.Text(frm_msg_form['email'], width=ent_form_width, height=1, 
    highlightthickness=1, highlightcolor = "Gainsboro", font=((TK_FONT, 9)))

# === Фрейм №2 - кнопка отправки, кнопка добавления файлов и информационные сообщения ===
frm_sending['email'] = tk.Frame(frm['email'], bg=THEME_COLOR, width=400, )
btn_send, btn_attached_files_path, lbl_msg_send = {}, {}, {}
btn_send['email'] = tk.Button(frm_sending['email'], text='Записать в БД', bg=BTN_INSERT_COLOR, width = 15, 
        command=lambda: loop_msg_service.create_task(btn_email_insert_db_click()))
btn_attached_files_path['email'] = tk.Button(frm_sending['email'], text='Прикрепить файлы', width = 16, 
        command=lambda: loop_msg_service.create_task(btn_attached_files_path_click('email')))
lbl_msg_send['email'] = tk.Label(frm_sending['email'], text='', 
        bg=THEME_COLOR, width=lbl_msg_send_width, anchor='w', )

# === Фрейм №3 - управление отправленными сообщениями ===
frm_sent_msg_header, lbl_header_title, btn_load_msg_from_db, lbl_header_records_numbers, btn_prev, btn_next = {}, {}, {}, {}, {}, {}

frm_sent_msg_header['email'] = tk.Frame(frm['email'], width=400, ) 
lbl_header_title['email'] = tk.Label(frm_sent_msg_header['email'], text='База данных e-mail сообщений', 
    width=lbl_header_title_width, font=('Segoe UI', 10, 'bold'))
btn_load_msg_from_db['email'] = tk.Button(frm_sent_msg_header['email'], text='Загрузить из БД', width = 15, 
                    command=lambda: loop_msg_service.create_task(btn_load_records_from_db_click('email'))) 
                    # btn_load_records_from_email_db_click()
lbl_header_records_numbers['email'] = tk.Label(frm_sent_msg_header['email'], bg=THEME_COLOR, width = 15, font=('Segoe UI', 10))
btn_prev['email'] = tk.Button(frm_sent_msg_header['email'], text='<', width=btn_move_width, 
                    command=lambda: loop_msg_service.create_task(btn_slice_msg_click('email', -1)))
                    # btn_slice_email_msg_click(-1)
btn_next['email'] = tk.Button(frm_sent_msg_header['email'], text='>', width=btn_move_width, 
                    command=lambda: loop_msg_service.create_task(btn_slice_msg_click('email', 1)))
            

# Вкладка telegram-сообщений =============================================================================
frm['telegram'] = tk.Frame(notebook, bg=THEME_COLOR, width=400, )
cmbx, cmbx['telegram'] = {}, {}

# === Фрейм №1 - формы сообщения ===
frm_msg_form['telegram'], lbl['telegram'], ent['telegram'] = tk.Frame(frm['telegram'], bg=THEME_COLOR, width=400, ), {}, {}
lbl['telegram']['description'] = tk.Label(frm_msg_form['telegram'], bg=THEME_COLOR, text='Запись в базу данных telegram сообщений', 
                                        font=('Segoe UI', 10, 'bold'))
lbl['telegram']['entity'] = tk.Label(frm_msg_form['telegram'], bg=THEME_COLOR,
            text = 'Тип получателя:', width=lbl_form_width, anchor='w', )
var_telegram_entity_type = tk.StringVar()
entity_type_list = ['Telegram-группа', 'Telegram-пользователь']
cmbx['telegram']['entity'] = ttk.Combobox(frm_msg_form['telegram'], textvariable=var_telegram_entity_type, 
    width=telegram_combo_width)
cmbx['telegram']['entity']['values'] = entity_type_list
cmbx['telegram']['entity']['state'] = 'readonly'
cmbx['telegram']['entity'].bind('<<ComboboxSelected>>', lambda event: asyncio.ensure_future(load_from_telegram_db(event)))
var_telegram_to = tk.StringVar()
cmbx['telegram']['to'] = ttk.Combobox(frm_msg_form['telegram'], textvariable=var_telegram_to, 
    width=telegram_combo_width, )
cmbx['telegram']['to']['state'] = 'readonly'
lbl['telegram']['to'] = tk.Label(frm_msg_form['telegram'], bg=THEME_COLOR,
            text = 'Кому:', width=lbl_form_width, anchor='w', )
ent['telegram']['msg_text'] = tk.Text(frm_msg_form['telegram'], width=ent_form_msg_width, height=ent_form_msg_height, 
    highlightthickness=1, highlightcolor = "Gainsboro", font=((TK_FONT, 9)))

lbl['telegram']['attachments'] = tk.Label(frm_msg_form['telegram'], bg=THEME_COLOR,
            text = 'Файлы:', width=lbl_form_width, anchor='w', )
ent['telegram']['attachments'] = tk.Text(frm_msg_form['telegram'], width=ent_form_width, height=1, 
    highlightthickness=1, highlightcolor = "Gainsboro", font=((TK_FONT, 9)))

# === Фрейм №2 - кнопка отправки и информационные сообщения ===
frm_sending['telegram'] = tk.Frame(frm['telegram'], bg=THEME_COLOR, width=400, )
btn_send['telegram'] = tk.Button(frm_sending['telegram'], text='Записать в БД', bg=BTN_INSERT_COLOR, width = 15, 
        command=lambda: loop_msg_service.create_task(btn_telegram_insert_db_click()))
btn_attached_files_path['telegram'] = tk.Button(frm_sending['telegram'], text='Прикрепить файлы', width = 16, 
        command=lambda: loop_msg_service.create_task(btn_attached_files_path_click('telegram')))
lbl_msg_send['telegram'] = tk.Label(frm_sending['telegram'], text='', 
        bg=THEME_COLOR, width=lbl_msg_send_width, anchor='w', )

# === Фрейм №3 - управление отправленными сообщениями ===
frm_sent_msg_header['telegram'] = tk.Frame(frm['telegram'], width=400, ) 
lbl_header_title['telegram'] = tk.Label(frm_sent_msg_header['telegram'], text='База данных telegram сообщений', 
    width=lbl_header_title_width, font=('Segoe UI', 10, 'bold'))
btn_load_msg_from_db['telegram'] = tk.Button(frm_sent_msg_header['telegram'], text='Загрузить из БД', width = 15, 
                    command=lambda: loop_msg_service.create_task(btn_load_records_from_db_click('telegram')))
lbl_header_records_numbers['telegram'] = tk.Label(frm_sent_msg_header['telegram'], bg=THEME_COLOR, width = 15, font=('Segoe UI', 10))
btn_prev['telegram'] = tk.Button(frm_sent_msg_header['telegram'], text='<', width=btn_move_width, 
                    command=lambda: loop_msg_service.create_task(btn_slice_msg_click('telegram', -1)))
btn_next['telegram'] = tk.Button(frm_sent_msg_header['telegram'], text='>', width=btn_move_width, 
                    command=lambda: loop_msg_service.create_task(btn_slice_msg_click('telegram', 1)))

# ОБЪЕДИНЕННЫЙ === Фрейм №4 - просмотр отправленных сообщений ===
frm_sent_messages, lbl_sent_messages_header, lbl_message = {}, {}, {}

common_columns_list = [('id', 5, 'id'), ('adrto', 20, 'Адреса'), ('attachments', 20, 'Файлы'),
        ('datep', 17, 'Дата записи'), ('dates', 17, 'Дата обработки')]
email_columns_list = [('subj', 20, 'Тема'), ('textemail', 40, 'Сообщение'), ]
telegram_columns_list = [('msg_text', 61, 'Сообщение')]

for m in ['email', 'telegram']:
    frm_sent_messages[m] = tk.Frame(frm[m], width=400, height=frm_sent_messages_height, )
    lbl_message[m], lbl_sent_messages_header[m] = {}, {}
    columns_list = common_columns_list+email_columns_list if m == 'email' else common_columns_list+telegram_columns_list
    for l in columns_list:
        lbl_sent_messages_header[m][l[0]] = tk.Label(frm_sent_messages[m],
            font=('Segoe UI', 8), width=l[1], text=l[2])
        lbl_message[m][l[0]] = {}
        for i in range(10):
            lbl_message[m][l[0]][i] = tk.Label(frm_sent_messages[m], bg=THEME_COLOR, font=('Segoe UI', 8), width=l[1])
            lbl_message[m][l[0]][i]['anchor'] = 'w' if l[0] != 'id' else 'c'



loop_msg_service = asyncio.get_event_loop()
loop_msg_service.run_until_complete(show_send_msg())
