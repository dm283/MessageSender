-- создает таблицу email-сообщений
CREATE TABLE [uemail](
	[id] [varchar](8) NULL,
	[app] [varchar](4) NULL,
	[forms] [varchar](6) NULL,
	[ids] [varchar](16) NULL,
	[client] [int] NULL,
	[adrto] [varchar](500) NULL,
	[subj] [varchar](100) NULL,
	[textemail] [varchar](600) NULL,
	[attachmentfiles] [varchar](255) NULL,
	[guid_doc] [varchar](36) NULL,
	[datep] [datetime] NULL,
	[dates] [datetime] NULL,
	[datet] [datetime] NULL,
	[datef] [datetime] NULL,
	[fl] [int] NULL,
	[user_id] [varchar](3) NULL,
	[status] [int] NULL,
	[UniqueIndexField] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_uemail] PRIMARY KEY CLUSTERED 
(
	[UniqueIndexField] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO

-- создает таблицу telegram-сообщений
CREATE TABLE [telegram_chats](
	[id] [varchar](8) NULL,
	[app] [varchar](4) NULL,
	[forms] [varchar](6) NULL,
	[ids] [varchar](16) NULL,
	[client] [int] NULL,
	[adrto] [varchar](500) NULL,
	[subj] [varchar](100) NULL,
	[msg_text] [varchar](600) NULL,
	[attachmentfiles] [varchar](255) NULL,
	[guid_doc] [varchar](36) NULL,
	[datep] [datetime] NULL,
	[dates] [datetime] NULL,
	[datet] [datetime] NULL,
	[datef] [datetime] NULL,
	[fl] [int] NULL,
	[user_id] [varchar](3) NULL,
	[status] [int] NULL,
	[UniqueIndexField] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [pk_telegram_messages] PRIMARY KEY CLUSTERED 
(
	[UniqueIndexField] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO

-------------
-- создает таблицу чатов бота с юзерами и группами
CREATE TABLE [telegram_chats](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[chat_id] [bigint] NOT NULL,
	[entity_name] [varchar](32) NOT NULL,
	[entity_type] [varchar](20) NOT NULL,
	[bot_name] [varchar](32) NOT NULL,
	[update_date] [datetime] DEFAULT GETDATE() NOT NULL,
	[is_active] [bit] DEFAULT 1 NOT NULL,
 CONSTRAINT [pk_telegram_chats] PRIMARY KEY CLUSTERED 
([id] ASC)
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO