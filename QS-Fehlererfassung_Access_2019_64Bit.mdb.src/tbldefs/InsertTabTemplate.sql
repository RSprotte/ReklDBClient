CREATE TABLE [InsertTabTemplate] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [FieldName] VARCHAR (255),
  [FieldType] VARCHAR (255),
  [FieldValue] VARCHAR (255),
  [TableZuordnung] VARCHAR (255)
)
