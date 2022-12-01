import sys, asyncio, configparser
import aioodbc, requests
from cryptography.fernet import Fernet

# загрузка конфигурации
CONFIG_FILE = 'config.ini'
config = configparser.ConfigParser()
config.read(CONFIG_FILE, encoding='utf-8')

# загрузка ключа шифрования
with open('rec-k.txt') as f:
    rkey = f.read().encode('utf-8')
refKey = Fernet(rkey)

hashed_common_bot_token = config['telegram_bot']['bot_token'].split('\t#')[0]
common_bot_token = (refKey.decrypt(hashed_common_bot_token).decode('utf-8'))

BOT_NAME = config['telegram_bot']['bot_name'].split('\t#')[0]
BOT_TOKEN = common_bot_token
TELEGRAM_DB = config['telegram_bot']['db'].split('\t#')[0]  # база данных mssql/posgres

TELEGRAM_DB_TABLE_CHATS = config['telegram_bot']['db_table_chats'].split('\t#')[0]  # db.schema.table  таблица с telegram-чатами
TELEGRAM_DB_CONNECTION_STRING = config['telegram_bot']['db_connection_string'].split('\t#')[0]  # odbc driver system dsn name

async def detect_telegram_chat_id():
# определяет для бота и telegram-сущности id их чата, после его создания
    try:
        cnxn = await aioodbc.connect(dsn=TELEGRAM_DB_CONNECTION_STRING, loop=loop)
        cursor = await cnxn.cursor()
    except:
        print("Подключение к базе данных  -  ошибка")
        return 1

    et = 'пользователем' if TELEGRAM_ENTITY_TYPE=='user' else 'группой'
    if await check_telegram_entity_in_db(cnxn, cursor):  # проверяет наличие сущности в бд
        print(f'chat_id бота {BOT_NAME} с {et} {TELEGRAM_ENTITY_NAME} уже записан в базу данных.')
        await cursor.close()
        await cnxn.close()
        return 0

    print('Отправка запроса api.telegram.org на получение событий бота .....')
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/getUpdates'  # запрос обновлений бота
    try:
        res = requests.get(url).json()
    except Exception as e:
        print('Ошибка: ', e)
        await cursor.close()
        await cnxn.close()
        return 1
    if res['ok'] == False:
        print('Получен ответ об ошибке:')
        print(res)
        await cursor.close()
        await cnxn.close()
        return 1

    # итерация с конца обновлений к началу, если бот удален из группы, выводит сообщение и заканчивает итерацию
    if 'result' not in res:
        print('Ошибка, получен некорректный ответ, отсутствует поле result.')
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
            await save_telegram_chat_id_to_db(BOT_NAME, TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME, chat_id, cnxn, cursor)
            print(f'chat_id бота {BOT_NAME} с {et} {TELEGRAM_ENTITY_NAME} записан в базу данных.')
            await cursor.close()
            await cnxn.close()
            return 0
    
    if TELEGRAM_ENTITY_TYPE == 'group':
        print(f'Ошибка определения чата с группой {TELEGRAM_ENTITY_NAME}.')
        print('Активируйте бота в telegram (если он не активирован) и/или добавьте в группу.')
        print('Если бот в группе, попробуйте удалить его и добавить заново.')
    elif TELEGRAM_ENTITY_TYPE == 'user':
        print(f'Ошибка определения чата с пользователем {TELEGRAM_ENTITY_NAME}.')
        print('Активируйте бота в telegram (если он не активирован) и/или отправьте ему сообщение.')

    await cursor.close()
    await cnxn.close()


async def save_telegram_chat_id_to_db(BOT_NAME, TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME, chat_id, cnxn, cursor):
    # сохраняет id и контакт telegram-чата в базу данных
    try:
        tet_for_insert = 'administrator' if tet == '-ua' else TELEGRAM_ENTITY_TYPE
        query = f"""insert into {TELEGRAM_DB_TABLE_CHATS} (chat_id, entity_name, entity_type, bot_name) values (
            {chat_id}, '{TELEGRAM_ENTITY_NAME}', '{tet_for_insert}', '{BOT_NAME}')"""
        await cursor.execute(query)
        await cnxn.commit()
    except Exception as e:
        print('Ошибка записи в базу данных.', e)
        await cursor.close()
        await cnxn.close()
        sys.exit()


async def check_telegram_entity_in_db(cnxn, cursor):
    # проверка наличия telegram-сущности в базе данных
    try:
        entity_types = "'user', 'administrator'" if tet in ('-u', '-ua') else "'group'"
        query = f"""select id from {TELEGRAM_DB_TABLE_CHATS} where entity_name='{TELEGRAM_ENTITY_NAME}' 
                and entity_type in ({entity_types}) and bot_name='{BOT_NAME}' and is_active"""
        await cursor.execute(query)
        rows = await cursor.fetchall()
    except:
        print('Ошибка чтения из базы данных.')
        await cursor.close()
        await cnxn.close()
        sys.exit()
    r = True if len(rows) > 0 else False
    return r


# ================================
# приложение запускается с параметрами: тип (-u / -g) и имя сущности
if len(sys.argv) < 3:
    print("Формат запуска:  botChatIdDetect -entity_type entity_name")
    print("-entity_type: -u/-ua/-g")
    sys.exit()

tet = sys.argv[1]
if tet in ('-u', '-ua'): 
    TELEGRAM_ENTITY_NAME = sys.argv[2]
    TELEGRAM_ENTITY_TYPE = 'user'
elif tet =='-g': 
    TELEGRAM_ENTITY_NAME = sys.argv[2]
    TELEGRAM_ENTITY_TYPE = 'group'
else:
    print(f"Ошибка: некорректный ключ {tet}")
    print("Формат запуска:  botChatIdDetect -entity_type entity_name")
    print("-entity_type: -u/-ua/-g")
    sys.exit()

print(TELEGRAM_ENTITY_TYPE, TELEGRAM_ENTITY_NAME)

loop = asyncio.get_event_loop()
loop.run_until_complete(detect_telegram_chat_id())
