import sys, configparser, datetime, asyncio, tkinter as tk
import requests, aioodbc
from cryptography.fernet import Fernet

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

CHECK_DB_PERIOD = int(config['common']['check_db_period'].split('\t#')[0])  # период проверки новых записей в базе данных

USER_NAME = config['user_credentials']['name'].split('\t#')[0]
USER_PASSWORD = user_credentials_password

# добавить admin email

BOT_NAME = config['telegram_bot']['bot_name'].split('\t#')[0]
BOT_TOKEN = common_bot_token
TELEGRAM_DB = config['telegram_bot']['db'].split('\t#')[0]  # база данных mssql/posgres
TELEGRAM_DB_TABLE_MESSAGES = config['telegram_bot']['db_table_messages'].split('\t#')[0]  # db.schema.table
TELEGRAM_DB_TABLE_CHATS = config['telegram_bot']['db_table_chats'].split('\t#')[0]  # db.schema.table  таблица с telegram-чатами
TELEGRAM_DB_CONNECTION_STRING = config['telegram_bot']['db_connection_string'].split('\t#')[0]  # odbc driver system dsn name
ADMIN_BOT_CHAT_ID = str()  # объявление глобальной константы, которая записывается в функции load_telegram_chats_from_db

# добавить параметры email

ROBOT_START = False
ROBOT_STOP = False
APP_EXIT = False
SIGN_IN_FLAG = False
THEME_COLOR = 'Gainsboro'
LBL_COLOR = THEME_COLOR
ENT_COLOR = 'White'
BTN_COLOR = 'Green'
BTN_1_COLOR = 'IndianRed'
BTN_2_COLOR = 'OrangeRed'
BTN_3_COLOR = 'SlateGray'

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
        lbl_msg_robot["text"] = 'Запуск робота...'
    await robot()

async def btn_robot_stop_click():
    # кнопка Stop robot
    global ROBOT_START, ROBOT_STOP
    if ROBOT_START:
        lbl_msg_robot["text"] = 'Остановка робота...'
        ROBOT_STOP = True

async def window_signin():
    # рисует окно входа
    frm.pack()
    lbl_sign.place(x=95, y=30)
    lbl_user.place(x=95, y=83)
    ent_user.place(x=95, y=126)
    lbl_password.place(x=95, y=150)
    ent_password.place(x=95, y=193)
    cbt_sign_show_pwd.place(x=95, y=220)
    btn_sign.place(x=95, y=260)
    lbl_msg_sign.place(x=95, y=310)

async def window_robot():
    # рисует окно админки
    frm.pack()
    lbl_robot.place(x=95, y=30)
    btn_robot_run.place(x=95, y=93)
    btn_robot_stop.place(x=95, y=136)
    btn_exit.place(x=125, y=195)
    lbl_runner.place(x=95, y=240)
    lbl_msg_robot.place(x=95, y=280)

# === MESSENGER FUNCTIONS ===
async def robot():
    # запускает робота
    global ROBOT_START, ROBOT_STOP, ADMIN_BOT_CHAT_ID
    if ROBOT_START or ROBOT_STOP:
        return
    ROBOT_START = True  # флаг старта робота, предотвращает запуск нескольких экземпляров робота
    try:
        cnxn = await aioodbc.connect(dsn=TELEGRAM_DB_CONNECTION_STRING, loop=loop_robot)
        cursor = await cnxn.cursor()
        print(f'Создано подключение к базе данных {TELEGRAM_DB}')  ###
    except Exception as e:
        print("Подключение к базе данных  -  ошибка.", e)
        return 1
    
    # чтение из бд данных о telegram-группах
    telegram_chats, ADMIN_BOT_CHAT_ID = await load_telegram_chats_from_db(cursor)
    if telegram_chats == 1:
        await cursor.close()
        await cnxn.close()
        ROBOT_START, ROBOT_STOP = False, False
        lbl_msg_robot["text"] = 'Ошибка чтения из базы данных'
        return 1

    lbl_msg_robot["text"] = 'Робот в рабочем режиме'

    while not ROBOT_STOP:
        msg_data_records = await load_records_from_db(cursor)
        if msg_data_records == 1:
            await cursor.close()
            await cnxn.close()
            ROBOT_START, ROBOT_STOP = False, False
            lbl_msg_robot["text"] = 'Ошибка чтения из базы данных'
            return 1

        print(msg_data_records)
        print()
        if len(msg_data_records) > 0:
            await robot_send_messages(cnxn, cursor, msg_data_records, telegram_chats)
        else:
            print('Нет новых сообщений в базе данных.')  ### test

        await asyncio.sleep(CHECK_DB_PERIOD)

    #  действия после остановки робота
    await cursor.close()
    await cnxn.close()
    print("Робот остановлен")
    ROBOT_START, ROBOT_STOP = False, False
    lbl_msg_robot["text"] = 'Робот остановлен'
    if APP_EXIT:
        sys.exit()


async def robot_send_messages(cnxn, cursor, msg_data_records, telegram_chats):
    # отправляет сообщения через telegram
    for record in msg_data_records:
        # структура данных record =  (1, 'This is the test message 1!', 'test-group-1')
        print("Новая запись", record)  ###
        record_id = record[0]
        record_msg = record[1]
        record_addresses = record[2].split(';')

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
                # print(requests.get(url).json())  ###
                print(f'Сообщение {address} отправлено.\n')
            except Exception as e:
                print('Ошибка отправки:\n', e)

        await set_record_handling_time_in_db(cnxn, cursor, record_id)
        print('Запись обработана.')  ####


# === DATABASE FUNCTIONS ===
async def load_telegram_chats_from_db(cursor):
    # выборка из базы данных параметров telegram-чатов бота
    try:
        query = f"select entity_name, chat_id, entity_type from {TELEGRAM_DB_TABLE_CHATS} where bot_name='{BOT_NAME}' and is_active"
        await cursor.execute(query)
        rows = await cursor.fetchall()
        telegram_chats_dict = {row[0]: row[1] for row in rows}
        ADMIN_BOT_CHAT_ID = [row[1] for row in rows if row[2] == 'administrator'][0]
        return telegram_chats_dict, ADMIN_BOT_CHAT_ID
    except Exception as e:
        print('Ошибка чтения из базы данных.', e)
        return 1


async def load_records_from_db(cursor):
    # выборка из базы данных необработанных (новых) записей
    try:
        await cursor.execute(f'select UniqueIndexField, msg_text, adrto from {TELEGRAM_DB_TABLE_MESSAGES} where dates is null order by datep')
        rows = await cursor.fetchall()  # список кортежей
    except Exception as e:
        print('Ошибка чтения из базы данных.', e)
        return 1
    return rows

async def set_record_handling_time_in_db(cnxn, cursor, id):
    # пишет в базу дату/время отправки сообщения
    dt_string = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    await cursor.execute(f"update {TELEGRAM_DB_TABLE_MESSAGES} set dates = '{dt_string}' where UniqueIndexField = {id}")
    await cnxn.commit()

async def rec_to_log(rec):
    # пишет в лог-файл запись об отправке сообщения
    current_time = str(datetime.datetime.now())
    with open('log-mailsender.log', 'a') as f:
        f.write(f'{current_time}\t{rec}\n')


# ============== window sign in
root = tk.Tk()
root.resizable(0, 0)  # делает неактивной кнопку Развернуть
root.title('TelegramSender')
frm = tk.Frame(bg=THEME_COLOR, width=400, height=400)
lbl_sign = tk.Label(master=frm, text='Sign in to TelegramSender', bg=LBL_COLOR, font=("Arial", 15), width=21, height=2)
lbl_user = tk.Label(master=frm, text='Username', bg=LBL_COLOR, font=("Arial", 12), anchor='w', width=25, height=2)
ent_user = tk.Entry(master=frm, bg=ENT_COLOR, font=("Arial", 12), width=25, )
lbl_password = tk.Label(master=frm, text='Password', bg=LBL_COLOR, font=("Arial", 12), anchor='w', width=25, height=2)
ent_password = tk.Entry(master=frm, show='*', bg=ENT_COLOR, font=("Arial", 12), width=25, )

cbt_sign_show_pwd_v1 = tk.IntVar(value = 0)
cbt_sign_show_pwd = tk.Checkbutton(frm, bg=THEME_COLOR, text='Show password', variable=cbt_sign_show_pwd_v1, onvalue=1, offvalue=0, 
                                    command=lambda: loop.create_task(show_password_signin()))

btn_sign = tk.Button(master=frm, bg=BTN_COLOR, fg='White', text='Sign in', font=("Arial", 12, "bold"), 
                    width=22, height=1, command=lambda: loop.create_task(btn_sign_click()))
lbl_msg_sign = tk.Label(master=frm, bg=LBL_COLOR, fg='PaleVioletRed', font=("Arial", 12), width=25, height=2)

async def show():
    # показывает и обновляет окно входа
    await window_signin()
    while not SIGN_IN_FLAG:
        root.update()
        await asyncio.sleep(.1)

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
root_robot.title('TelegramSender')
frm = tk.Frame(bg=THEME_COLOR, width=400, height=400)
lbl_robot = tk.Label(master=frm, text='TelegramSender', bg=LBL_COLOR, font=("Arial", 15), width=20, height=2)

animation = "░▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒"
lbl_runner = tk.Label(master=frm, fg='DodgerBlue', text="")

btn_robot_run = tk.Button(master=frm, bg=BTN_2_COLOR, fg='White', text='Запуск робота', font=("Arial", 12, "bold"), 
                    width=22, height=1, command=lambda: loop_robot.create_task(btn_robot_run_click()))
btn_robot_stop = tk.Button(master=frm, bg=BTN_3_COLOR, fg='White', text='Остановка робота', font=("Arial", 12, "bold"), 
                    width=22, height=1, command=lambda: loop_robot.create_task(btn_robot_stop_click()))
btn_exit = tk.Button(master=frm, bg=BTN_1_COLOR, fg='Black', text='Выход', font=("Arial", 12), 
                    width=16, height=1, command=lambda: loop_robot.create_task(btn_exit_click()))
lbl_msg_robot = tk.Label(master=frm, bg=LBL_COLOR, font=("Arial", 10), width=25, height=2)

async def show_robot():
    # показывает и обновляет окно робота
    global animation

    await window_robot()
    while True:
        lbl_runner["text"] = animation
        if ROBOT_START:
            animation = animation[1:] + animation[0]

        root_robot.update()
        await asyncio.sleep(.1)

loop_robot = asyncio.get_event_loop()
loop_robot.run_until_complete(show_robot())
