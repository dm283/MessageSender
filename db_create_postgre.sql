-- создает таблицу telegram-сообщений
create table if not exists telegram_messages(
	id varchar(8) NULL,
	app varchar(4) NULL,
	forms varchar(6) NULL,
	ids varchar(16) NULL,
	client int NULL,
	adrto varchar(500) NULL,
	subj varchar(100) NULL,
	msg_text varchar(600) NULL,
	attachmentfiles varchar(255) NULL,
	guid_doc varchar(36) NULL,
	datep timestamp NULL,
	dates timestamp NULL,
	datet timestamp NULL,
	datef timestamp NULL,
	fl int NULL,
	user_id varchar(3) NULL,
	status int NULL,
	UniqueIndexField serial primary key
);


-- создает таблицу чатов бота с юзерами и группами
create table if not exists telegram_chats(
	id serial primary key,
	chat_id bigint not null,
	entity_name varchar(32) not null,
	entity_type varchar(20) not null,
	bot_name varchar(32) not null,
	update_date timestamp default now() not null,
	is_active bool default True not null
);


-- создает таблицу email-сообщений
create table if not exists uemail(
	id varchar(8) NULL,
	app varchar(4) NULL,
	forms varchar(6) NULL,
	ids varchar(16) NULL,
	client int NULL,
	adrto varchar(500) NULL,
	subj varchar(100) NULL,
	textemail varchar(600) NULL,
	attachmentfiles varchar(255) NULL,
	guid_doc varchar(36) NULL,
	datep timestamp NULL,
	dates timestamp NULL,
	datet timestamp NULL,
	datef timestamp NULL,
	fl int NULL,
	user_id varchar(3) NULL,
	status int NULL,
	UniqueIndexField serial primary key
);