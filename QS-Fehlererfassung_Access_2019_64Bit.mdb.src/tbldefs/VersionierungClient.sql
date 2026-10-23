CREATE TABLE [VersionierungClient] (
  [Version] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL CONSTRAINT [Version] UNIQUE,
  [Datum] DATETIME,
  [Änderung] LONGTEXT,
  [Kürzel] VARCHAR (50),
  [Bemerkung] LONGTEXT
)
