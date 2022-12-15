import sys, configparser, asyncio, tkinter as tk
import requests, aioodbc
from aiosmtplib import SMTP
from aioimaplib import aioimaplib
from cryptography.fernet import Fernet
from tkinter import ttk
from pathlib import Path

# загрузка конфигурации
CONFIG_FILE = 'config.ini'
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding='utf-8')

# загрузка ключа шифрования
with open('rec-k.txt') as f:
    rkey = f.read().encode('utf-8')
refKey = Fernet(rkey)

# config {} = реальные значения; config_show {} = значения для отрисовки
config_show = {}
for s in config.sections():
    config_show[s] = {}
    for k, v in config.items(s):
        config[s][k] = v.split('\t# ')[0]
        config_show[s][k] = v.split('\t# ')

# расшифровка паролей
password_section_key_list = [ ('user_credentials', 'password'), ('admin_credentials', 'password'), ('email', 'server_password')]
hashed_section_key_list = [ ('user_credentials', 'password'), ('admin_credentials', 'password'), ('email', 'server_password'), ('telegram_bot', 'bot_token')]
for s in hashed_section_key_list:
    hashed = config[s[0]][s[1]]
    config[s[0]][s[1]] = (refKey.decrypt(hashed).decode('utf-8')) if hashed != '' else config[s[0]][s[1]]

ADMIN_BOT_CHAT_ID = str()

to_str = config['admin_credentials']['email'].split('\t#')[0]
from_str = config['email']['sender_email'].split('\t#')[0]
TEST_MESSAGE = f"""To: {to_str}\nFrom: {from_str}\nSubject: MessageSender - тестовое сообщение\n
Это тестовое сообщение отправленное сервисом MessageSender.""".encode('utf8')

SIGN_IN_FLAG = False

THEME_COLOR = 'Gainsboro'
LBL_COLOR = THEME_COLOR
ENT_COLOR = 'White'
BTN_COLOR = 'Green'


# === TEST FUNCTIONS ===
async def btn_test_telegram_click():
    # тестирует работу с telegram-сообщениями: подключение к базе данных, отправка тестового сообщения в telegram-чат с админом
    test_1_res = await test_db_connect('telegram_bot', config['telegram_bot']['db'], config['telegram_bot']['db_connection_string'], 
                        [config['telegram_bot']['db_table_messages'], config['telegram_bot']['db_table_chats'], ])
    if test_1_res == 1:
        return 1
    test_2_res = await test_telegram_send_msg_to_admin()  # тестирует отправку сообщения в чат бота с админом
    if test_2_res == 1:
        return 1
    lbl_msg_test['telegram_bot']['text'] = 'Тестирование успешно завершено'


async def test_telegram_send_msg_to_admin():
    # тестирует отправку сообщения в чат бота с админом
    if not ADMIN_BOT_CHAT_ID:
        res = await load_admin_bot_chat_id_from_db()
        if res == 1:
            return 1
    lbl_msg_test['telegram_bot']['text'] = f"Отправка telegram-сообщения администратору ....."
    await asyncio.sleep(1)
    msg = 'Тестовое сообщение'
    url = f"https://api.telegram.org/bot{config['telegram_bot']['bot_token']}/sendMessage?chat_id={ADMIN_BOT_CHAT_ID}&text={msg}"
    try:
        res = requests.get(url).json()
        if res['ok'] == False:
            lbl_msg_test['telegram_bot']['text'] = f"Отправка telegram-сообщения администратору - ошибка"
            await asyncio.sleep(1)
            return 1
        else:
            lbl_msg_test['telegram_bot']['text'] = 'Отправка telegram-сообщения администратору - OK'
            await asyncio.sleep(1)
            return 0
    except:
        lbl_msg_test['telegram_bot']['text'] = 'Отправка telegram-сообщения администратору - ошибка'
        await asyncio.sleep(1)
        return 1


async def load_admin_bot_chat_id_from_db():
    # выборка из базы данных id чата бота с админом
    global ADMIN_BOT_CHAT_ID
    try:
        cnxn = await aioodbc.connect(dsn=config['telegram_bot']['db_connection_string'], loop=loop_admin)
        cursor = await cnxn.cursor()
    except:
        lbl_msg_test['telegram_bot']['text'] = f"Подключение к базе данных {config['database']['db']}  -  ошибка"
        return 1
    try:
        query = f"""select chat_id from {config['telegram_bot']['db_table_chats']} 
            where entity_type='administrator' and bot_name='{config['telegram_bot']['bot_name']}'"""
        await cursor.execute(query)
        rows = await cursor.fetchall()
        print(rows)
        if len(rows) == 0:
            lbl_msg_test['telegram_bot']['text'] = 'Ошибка: чат с администратором отсутствует в базе данных'
            await cursor.close()
            await cnxn.close()
            return 1
        ADMIN_BOT_CHAT_ID = [row[0] for row in rows][0]
    except:
        await cursor.close()
        await cnxn.close()
        lbl_msg_test['telegram_bot']['text'] = 'Ошибка чтения id чата бота с администратором\nиз базы данных'
        return 1
    await cursor.close()
    await cnxn.close()


async def test_db_connect(label, db, connection_string, tables):
    # тестирует подключение к базе данных и обращение к таблицам бд
    lbl_msg_test[label]['text'] = f"Подключение к базе данных {db} ....."
    await asyncio.sleep(1)
    try:
        cnxn = await aioodbc.connect(dsn=connection_string, loop=loop_admin)
        cursor = await cnxn.cursor()
        lbl_msg_test[label]['text'] = f"Подключение к базе данных {db} - OK"
        await asyncio.sleep(1)
        lbl_msg_test[label]['text'] = 'Обращение к таблицам базы данных .....'
        await asyncio.sleep(1)
        for table in tables:
            try:
                await cursor.execute(f"select count(id) from {table}")
                await cursor.fetchall()
                if table == tables[-1]:
                    lbl_msg_test[label]['text'] = f"Обращение к таблицам базы данных - OK"
            except:
                lbl_msg_test[label]['text'] = f"Обращение к таблице {table.split('.')[-1]} - ошибка"
                await cursor.close()
                await cnxn.close()
                await asyncio.sleep(1)
                return 1
        await cursor.close()
        await cnxn.close()
        await asyncio.sleep(1)
        return 0
    except:
        lbl_msg_test[label]['text'] = f"Подключение к базе данных {db} - ошибка"
        await asyncio.sleep(1)
        return 1


async def btn_test_email_click():
    # тестирует работу с email-сообщениями: подключение к базе данных, smtp и imap серверам, отправка тестового сообщения на email админа
    test_1_res = await test_db_connect('email', config['email']['db'], config['email']['db_connection_string'], 
                         [config['email']['db_table_emails'], ])
    if test_1_res == 1:
        return 1
    test_2_res = await test_smtp_server(host=config['email']['smtp_host'], port=config['email']['smtp_port'], 
                                        sender=config['email']['sender_email'],
                                        pwd=config['email']['server_password'], receiver=config['admin_credentials']['email'])
    if test_2_res == 1:
        return 1
    test_3_res = await test_imap_server(host=config['email']['imap_host'], port=config['email']['imap_port'], 
                                        sender=config['email']['sender_email'], pwd=config['email']['server_password'])
    if test_3_res == 1:
        return 1
    lbl_msg_test['telegram_bot']['text'] = 'Тестирование успешно завершено'


async def test_smtp_server(host, port, sender, pwd, receiver):
    # тестирует подключение к smtp-серверу: доступность smtp-сервера и отправка тестового сообщения
    lbl_msg_test['email']['text'] = f"Подключение к {host}: {port} ....."
    try:
        smtp_client = SMTP(hostname=host, 
                    port=port, 
                    use_tls=True, 
                    username=sender, 
                    password=pwd)
        await smtp_client.connect()
        lbl_msg_test['email']['text'] = f"Подключение к {host}: {port} - OK"
        await asyncio.sleep(1)
        try:
            lbl_msg_test['email']['text'] = f"Отправка сообщения {receiver} ....."
            await smtp_client.sendmail(sender, receiver, TEST_MESSAGE)
            lbl_msg_test['email']['text'] = f'Отправка сообщения {receiver} - OK'
            await asyncio.sleep(1)
            await smtp_client.quit()
            lbl_msg_test['email']['text'] = 'Тестирование успешно завершено'
        except:
            lbl_msg_test['email']['text'] = f'Отправка сообщения {receiver} - ошибка'
    except:
        lbl_msg_test['email']['text'] = f"Подключение к {host}:{port} - ошибка"


async def test_imap_server(host, port, sender, pwd):
    # тестирует подключение к imap-сервер
    lbl_msg_test['email']['text'] = f"Подключение к {host}: {port} ....."
    try:
        imap_client = aioimaplib.IMAP4_SSL(host=host)  # не ловиться исключение здесь!
        await imap_client.wait_hello_from_server()
        lbl_msg_test['email']['text'] = f"Подключение к {host}: {port}  -  OK"
        await asyncio.sleep(1)
        try:
            await imap_client.login(sender, pwd)
            await imap_client.select('INBOX')
            lbl_msg_test['email']['text'] = f'Чтение входящих сообщений  -  OK'
            await asyncio.sleep(1)
            await imap_client.close()
            await imap_client.logout()
            lbl_msg_test['email']['text'] = 'Тестирование успешно завершено'
        except:
                lbl_msg_test['email']['text'] = f'Чтение входящих сообщений  -  ошибка'
    except:
        lbl_msg_test['email']['text'] = f"Подключение к {host}:{port}  -  ошибка"


# === SAVE CONFIG FUNCTION ===
async def btn_save_config_click():
    global config
    # сохраняет установленные значения конфига
    for s in config.sections():
        for k, v in config.items(s):
            # описания секций остаются неизменными
            if k in ['section_description', 'section_label']:
                continue
            if (s, k) in hashed_section_key_list:
                # если параметр подлежит шифрованию - хэширование перед записью в config
                scrt_value = ent[s][k].get().encode('utf-8')
                hashed_scrt_value = refKey.encrypt(scrt_value)
                config[s][k] = hashed_scrt_value.decode('utf-8')
            else:
                # обычный параметр
                config[s][k] = ent[s][k].get()
            # добавляет комментарий к параметру, если он предусмотрен (записан 2-м элементом в список value словаря config_show )
            if len(config_show[s][k]) > 1:
                config[s][k] = config[s][k] + '\t# ' + config_show[s][k][1]
    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
        config.write(configfile)
    lbl_config_msg['text'] = f'Конфигурация сохранена в файл {CONFIG_FILE}'
    # после сохранения конфига сообщения о тестах меняются на изначальные
    lbl_msg_test['telegram_bot']['text'], lbl_msg_test['email']['text'] = '', ''
    # запись обратно в переменные config значений без комментариев
    for s in config.sections():
        for k, v in config.items(s):
            if k not in ['section_description', 'section_label']:
                config[s][k] = ent[s][k].get()
    # создание папок приложений, если отсутствуют
    print('проверка наличия папок приложений')
    dir_email_attachments = Path(config['common']['dir_email_attachments'])
    dir_telegram_attachments = Path(config['common']['dir_telegram_attachments'])
    if not dir_email_attachments.exists():
        dir_email_attachments.mkdir()
    if not dir_telegram_attachments.exists():
        dir_telegram_attachments.mkdir()


# === INTERFACE FUNCTIONS ===
async def btn_sign_click():
    # кнопка sign-in
    global SIGN_IN_FLAG
    user = ent_user.get()
    password = ent_password.get()

    print(config['admin_credentials']['name'], config['admin_credentials']['password'])

    if user == config['admin_credentials']['name'] and password == config['admin_credentials']['password']:
        lbl_msg_sign["text"] = ''
        SIGN_IN_FLAG = True
        root.destroy()
    else:
        lbl_msg_sign["text"] = 'Incorrect username or password'

async def show_password_signin():
    # показывает/скрывает пароль в окне входа
    ent_password['show'] = '' if(cbt_sign_show_pwd_v1.get() == 1) else '*'
        
async def show_password(s, k):
    # показывает/скрывает пароль на вкладках окна админки
    ent[s][k]['show'] = '' if(cbt_v1[s][k].get() == 1) else '*'

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

async def show_admin():
    # рисует окно администрирования
    notebook.pack(padx=10, pady=10, fill='both', expand=True)
    for s in config.sections():
        notebook.add(frm[s], text = config[s]['section_label'])
        frm_params[s].pack(padx=5, pady=(5, 0), fill='both', expand=True)
        frm_test[s].pack(padx=5, pady=(1, 5), fill='both', expand=True)
        lbl[s]['section_description'].grid(row = 0, columnspan = 3, sticky = 'w', padx = 5, pady = 5)
        r = 1
        for k, v in config.items(s):
            if k in ['section_description', 'section_label']:
                continue
            lbl[s][k].grid(row = r, column = 0, sticky = 'w', padx = 5, pady = 5)
            ent[s][k].grid(row = r, column = 1, sticky = 'w', padx = 5, pady = 5)
            ent[s][k].insert(0, v)
            if (s, k) in password_section_key_list:
                cbt[s][k].grid(row = r, column = 2, sticky = 'w', )
            r += 1
        # положение кнопок тестирования (в окне с наибольшим кол-вом параметров - grid, в прочих - place)
        if s == 'telegram_bot':
            btn_test['telegram_bot'].place(x=5, y=34)
            lbl_msg_test['telegram_bot'].place(x=130, y=36)
        if s == 'email':
            btn_test['email'].grid(row=0, column=0, padx=5, pady=5)    
            lbl_msg_test['email'].grid(row=0, column=1, padx=5, pady=5)

    # специальный фрейм определения-добавления telegram-чатов
    notebook.add(frm_chat_detect, text='Telegram-чаты')
    lbl_chat_detect_descr.grid(row=0, columnspan=3, sticky='w', padx=5, pady=5)
    lbl_chat_detect_entity_type.grid(row=1, column=0, sticky='w', padx=5, pady=5)
    cmbx_chat_detect.grid(row=1, column=1, sticky='w', padx=5, pady=5)
    lbl_chat_detect_name.grid(row=2, column=0, sticky='w', padx=5, pady=5)
    ent_chat_detect_name.grid(row=2, column=1, sticky='w', padx=5, pady=5)
    btn_chat_detect_detect.grid(row=3, column=0, sticky='w', padx=5, pady=5)
    lbl_chat_detect_msg.grid(row=4, columnspan=3, sticky='w', padx=5, pady=5)

    # фрейм 'Сохранить'
    frm_footer.pack(padx=10, pady=(0, 10), fill='both', expand=True)  
    btn_save_config.grid(row = 0, column = 0)
    lbl_config_msg.grid(row = 0, column = 1, padx = 10)

    while True:
        root_admin.update()
        await asyncio.sleep(.01)



# === TELEGRAM CHAT DETECT FUNCTIONS ===
async def detect_telegram_chat_id():
# определяет для бота и telegram-сущности id их чата, после его создания
    tet = cmbx_chat_detect.get().strip()
    TELEGRAM_ENTITY_NAME = ent_chat_detect_name.get().strip()
    print('tet', tet, 'ten', TELEGRAM_ENTITY_NAME)

    if tet == '' or TELEGRAM_ENTITY_NAME == '':
        lbl_chat_detect_msg['text'] = 'Заполните все поля с данными контакта.'
        return 1
    else:
        lbl_chat_detect_msg['text'] = ''
    
    if tet in ('Telegram-пользователь', 'Администратор telegram-бота'): 
        TELEGRAM_ENTITY_TYPE = 'user'
    elif tet == 'Telegram-группа': 
        TELEGRAM_ENTITY_TYPE = 'group'

    try:
        cnxn = await aioodbc.connect(dsn=config['telegram_bot']['db_connection_string'], loop=loop_admin)
        cursor = await cnxn.cursor()
    except:
        lbl_chat_detect_msg['text'] = "Подключение к базе данных  -  ошибка"
        print("Подключение к базе данных  -  ошибка")  ###
        return 1

    et = 'пользователем' if TELEGRAM_ENTITY_TYPE=='user' else 'группой'
    check_res = await check_telegram_entity_in_db(TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME, cnxn, cursor)
    if check_res == 2:
        await cursor.close()
        await cnxn.close()
        return 1
    if check_res:  # проверяет наличие сущности в бд
        lbl_chat_detect_msg['text'] = (f"chat_id бота {config['telegram_bot']['bot_name']} с {et} {TELEGRAM_ENTITY_NAME} " + 
            "уже записан в базу данных.")
        print(f"chat_id бота {config['telegram_bot']['bot_name']} с {et} {TELEGRAM_ENTITY_NAME} уже записан в базу данных.") ###
        await cursor.close()
        await cnxn.close()
        return 0

    if tet == 'Администратор telegram-бота':
        check_res = await check_telegram_admin_exists(cnxn, cursor)
        if check_res == 2:
            await cursor.close()
            await cnxn.close()
            return 1
        print(check_res)
        if check_res:
            lbl_chat_detect_msg['text'] = "В базе данных уже создан чат с администратором."
            await cursor.close()
            await cnxn.close()
            return 0

    lbl_chat_detect_msg['text'] = 'Отправка запроса api.telegram.org на получение событий бота .....'
    print('Отправка запроса api.telegram.org на получение событий бота .....')
    url = f"https://api.telegram.org/bot{config['telegram_bot']['bot_token']}/getUpdates"  # запрос обновлений бота
    try:
        res = requests.get(url).json()
    except Exception as e:
        lbl_chat_detect_msg['text'] = 'Ошибка: ', e
        print('Ошибка: ', e) ###
        await cursor.close()
        await cnxn.close()
        return 1
    if res['ok'] == False:
        lbl_chat_detect_msg['text'] = f'Получен ответ об ошибке:\n{res}'
        print('Получен ответ об ошибке:')  ###
        print(res)  ###
        await cursor.close()
        await cnxn.close()
        return 1

    # итерация с конца обновлений к началу, если бот удален из группы, выводит сообщение и заканчивает итерацию
    if 'result' not in res:
        lbl_chat_detect_msg['text'] = 'Ошибка, получен некорректный ответ, отсутствует поле result.'
        print('Ошибка, получен некорректный ответ, отсутствует поле result.')  ###
        return 1

    chat_id = ''

    for i in range(len(res['result']), 0, -1):
        s = res['result'][i-1]

        if TELEGRAM_ENTITY_TYPE == 'group':
            if 'my_chat_member' in s and s['my_chat_member']['new_chat_member']['status'] == 'left':
                break
            if ('my_chat_member' in s 
                and s['my_chat_member']['chat']['title'] == TELEGRAM_ENTITY_NAME 
                and s['my_chat_member']['new_chat_member']['status'] == 'member'):
                chat_id = int(s['my_chat_member']['chat']['id'])
                print('group_chat_id:', chat_id)
        elif TELEGRAM_ENTITY_TYPE == 'user':
            if ('message' in s and 'chat' in s['message'] and 'username' in s['message']['chat']
                    and s['message']['chat']['username'] == TELEGRAM_ENTITY_NAME):
                chat_id = s['message']['chat']['id']
                print('user_chat_id:', chat_id)
            
        if chat_id:
            save_res = await save_telegram_chat_id_to_db(config['telegram_bot']['bot_name'], TELEGRAM_ENTITY_TYPE, 
                TELEGRAM_ENTITY_NAME, tet, chat_id, cnxn, cursor)
            if save_res == 2:
                await cursor.close()
                await cnxn.close()
                return 1
            lbl_chat_detect_msg['text'] = (f"chat_id бота {config['telegram_bot']['bot_name']} с {et} {TELEGRAM_ENTITY_NAME} " + 
                "записан в базу данных.")
            print(f"chat_id бота {config['telegram_bot']['bot_name']} с {et} {TELEGRAM_ENTITY_NAME} записан в базу данных.")  ###
            await cursor.close()
            await cnxn.close()
            return 0
    
    if TELEGRAM_ENTITY_TYPE == 'group':
        lbl_chat_detect_msg['text'] = (f'Ошибка определения чата с группой {TELEGRAM_ENTITY_NAME}.\n' + 
            'Активируйте бота в telegram (если он не активирован) и/или добавьте в группу.\n' +
            'Если бот в группе, попробуйте удалить его и добавить заново.')
        print(f'Ошибка определения чата с группой {TELEGRAM_ENTITY_NAME}.')  ###
        print('Активируйте бота в telegram (если он не активирован) и/или добавьте в группу.') ###
        print('Если бот в группе, попробуйте удалить его и добавить заново.') ###
    elif TELEGRAM_ENTITY_TYPE == 'user':
        lbl_chat_detect_msg['text'] = (f'Ошибка определения чата с пользователем {TELEGRAM_ENTITY_NAME}.\n' +
        f'Активируйте бота в telegram (если он не активирован)и/или отправьте ему сообщение\nот {TELEGRAM_ENTITY_NAME}.')
        print(f'Ошибка определения чата с пользователем {TELEGRAM_ENTITY_NAME}.\n')  ###
        print(f'Активируйте бота в telegram (если он не активирован) и/или отправьте ему сообщение от {TELEGRAM_ENTITY_NAME}.') ###

    await cursor.close()
    await cnxn.close()


async def save_telegram_chat_id_to_db(BOT_NAME, TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME, tet, chat_id, cnxn, cursor):
    # сохраняет id и контакт telegram-чата в базу данных
    try:
        tet_for_insert = 'administrator' if tet == 'Администратор telegram-бота' else TELEGRAM_ENTITY_TYPE
        query = f"""insert into {config['telegram_bot']['db_table_chats']} (chat_id, entity_name, entity_type, bot_name) values (
            {chat_id}, '{TELEGRAM_ENTITY_NAME}', '{tet_for_insert}', '{BOT_NAME}')"""
        await cursor.execute(query)
        await cnxn.commit()
    except Exception as e:
        lbl_chat_detect_msg['text'] = 'Ошибка записи в базу данных.'
        print('Ошибка записи в базу данных.', e)  ###
        await cursor.close()
        await cnxn.close()
        return 2


async def check_telegram_entity_in_db(TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME, cnxn, cursor):
    # проверка наличия telegram-сущности в базе данных
    try:
        entity_types = "'user', 'administrator'" if TELEGRAM_ENTITY_TYPE == 'user' else "'group'"
        query = f"""select id from {config['telegram_bot']['db_table_chats']} where entity_name='{TELEGRAM_ENTITY_NAME}' 
                and entity_type in ({entity_types}) and bot_name='{config['telegram_bot']['bot_name']}'"""
        await cursor.execute(query)
        rows = await cursor.fetchall()
    except Exception as e:
        print('Ошибка чтения из базы данных.', e)  ##
        lbl_chat_detect_msg['text'] = 'Ошибка чтения из базы данных.'
        await cursor.close()
        await cnxn.close()
        return 2
    r = True if len(rows) > 0 else False
    return r


async def check_telegram_admin_exists(cnxn, cursor):
    # проверка наличия чата с администратором в базе данных
    try:
        query = f"""select id from {config['telegram_bot']['db_table_chats']} where entity_type = 'administrator'
            and bot_name='{config['telegram_bot']['bot_name']}'"""
        await cursor.execute(query)
        rows = await cursor.fetchall()
    except Exception as e:
        print('Ошибка чтения из базы данных.', e)  ###
        lbl_chat_detect_msg['text'] = 'Ошибка чтения из базы данных.'
        await cursor.close()
        await cnxn.close()
        return 2
    r = True if len(rows) > 0 else False
    return r

# ============== window sign in
root = tk.Tk()
root.resizable(0, 0)  # делает неактивной кнопку Развернуть
root.title('MessageSender administration')
frm = tk.Frame(bg=THEME_COLOR, width=400, height=400)
#lbl_sign = tk.Label(master=frm, text='', bg=LBL_COLOR, font=("Arial", 15), width=21, height=2)  #bg=LBL_COLOR
lbl_user = tk.Label(master=frm, text='Username', bg=LBL_COLOR, font=("Arial", 12), anchor='w', width=25, height=2)
ent_user = tk.Entry(master=frm, bg=ENT_COLOR, font=("Arial", 12), width=25, )
lbl_password = tk.Label(master=frm, text='Password', bg=LBL_COLOR, font=("Arial", 12), anchor='w', width=25, height=2)
ent_password = tk.Entry(master=frm, show = '*', bg=ENT_COLOR, font=("Arial", 12), width=25, )

cbt_sign_show_pwd_v1 = tk.IntVar(value = 0)
cbt_sign_show_pwd = tk.Checkbutton(frm, bg=THEME_COLOR, text='Show password', variable=cbt_sign_show_pwd_v1, onvalue=1, offvalue=0, 
                                    command=lambda: loop.create_task(show_password_signin()))

btn_sign = tk.Button(master=frm, bg=BTN_COLOR, fg='White', text='Sign in', font=("Arial", 12, "bold"), 
                    width=22, height=1, command=lambda: loop.create_task(btn_sign_click()))
lbl_msg_sign = tk.Label(master=frm, bg=LBL_COLOR, fg='PaleVioletRed', font=("Arial", 12), width=25, height=2)

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

# ============== window admin
root_admin = tk.Tk()
root_admin.resizable(0, 0)  # делает неактивной кнопку Развернуть
root_admin.title('MessageSender administration')
notebook = ttk.Notebook(root_admin)

# в каждом Frame (section) создаются подфреймы с Labels и Enrty (key и key-value)
frm, frm_params, frm_test, lbl, ent = {}, {}, {}, {}, {}

for s in config.sections():
    frm[s], lbl[s], ent[s] = tk.Frame(notebook, bg=THEME_COLOR, width=400, ), {}, {}
    
    frm_params[s] = tk.Frame(frm[s], bg=THEME_COLOR, width=400, )
    frm_test[s] = tk.Frame(frm[s], bg=THEME_COLOR, width=400, )

    for k, v in config.items(s):
        if k == 'section_description':
            lbl[s][k] = tk.Label(frm_params[s], bg=THEME_COLOR, text = v, font=('Segoe UI', 10, 'bold'))
            continue
        lbl[s][k] = tk.Label(frm_params[s], bg=THEME_COLOR,
            text = config_show[s][k][1] if len(config_show[s][k]) > 1 else v, width=27, anchor='w', )
        ent[s][k] = tk.Entry(frm_params[s], width=30, highlightthickness=1, highlightcolor = "Gainsboro", )

# формирование виджетов для полей со скрытием контента
cbt_v1, cbt = {}, {}
for s, k in password_section_key_list:
    cbt_v1[s], cbt[s] = {}, {}
    cbt_v1[s][k] = tk.IntVar(value = 0)
    cbt[s][k] = tk.Checkbutton(frm_params[s], bg=THEME_COLOR, text = 'Показать пароль', variable = cbt_v1[s][k], onvalue = 1, offvalue = 0)
    ent[s][k]['show'] = '*'
# параметры для show_password прописываются принудительно для каждой кнопки
cbt['user_credentials']['password']['command'] = lambda: loop_admin.create_task(show_password('user_credentials', 'password'))
cbt['admin_credentials']['password']['command'] = lambda: loop_admin.create_task(show_password('admin_credentials', 'password'))
cbt['email']['server_password']['command'] = lambda: loop_admin.create_task(show_password('email', 'server_password'))

# формирование элементов с функционалом тестирования
btn_test, lbl_msg_test = {}, {}
btn_test['telegram_bot'] = tk.Button(frm_test['telegram_bot'], text='Тест', width = 15, 
        command=lambda: loop_admin.create_task(btn_test_telegram_click()))
lbl_msg_test['telegram_bot'] = tk.Label(frm_test['telegram_bot'], text='', 
        bg=THEME_COLOR, width = 54, anchor='w', )
btn_test['email'] = tk.Button(frm_test['email'], text='Тест', width = 15, 
        command=lambda: loop_admin.create_task(btn_test_email_click()))
lbl_msg_test['email'] = tk.Label(frm_test['email'], text='', 
        bg=THEME_COLOR, width = 54, anchor='w', )

# формирование фрейма с общим функционалом (сохранение конфига)
frm_footer = tk.Frame(root_admin, width=400, height=280, )
btn_save_config = tk.Button(frm_footer, text='Сохранить', width=15, command=lambda: loop_admin.create_task(btn_save_config_click()))
lbl_config_msg = tk.Label(frm_footer, )

# ******** вкладка добавления telegram-чата
frm_chat_detect = tk.Frame(notebook, bg=THEME_COLOR, width=400, )
lbl_chat_detect_descr = tk.Label(frm_chat_detect, bg=THEME_COLOR, text = 'Добавление telegram-чата:', 
    font=('Segoe UI', 10, 'bold'))
lbl_chat_detect_entity_type = tk.Label(frm_chat_detect, bg=THEME_COLOR,
    text='Тип контакта', width=27, anchor='w', )
entity_type_list = ['Telegram-группа', 'Telegram-пользователь', 'Администратор telegram-бота']
var_telegram_entity_type = tk.StringVar()
cmbx_chat_detect = ttk.Combobox(frm_chat_detect, textvariable=var_telegram_entity_type, width=48, values=entity_type_list, 
    state='readonly')
lbl_chat_detect_name = tk.Label(frm_chat_detect, bg=THEME_COLOR, text='Telegram-имя контакта', width=27, anchor='w', )
ent_chat_detect_name = tk.Entry(frm_chat_detect, width=51, highlightthickness=1, highlightcolor = "Gainsboro", )
btn_chat_detect_detect = tk.Button(frm_chat_detect, text='Добавить в список чатов', width = 20, 
        command=lambda: loop_admin.create_task( detect_telegram_chat_id() ))
lbl_chat_detect_msg = tk.Label(frm_chat_detect, text='', bg=THEME_COLOR, width=73, anchor='w', justify='left')


loop_admin = asyncio.get_event_loop()
loop_admin.run_until_complete(show_admin())
