## Stage 2 - Set Up SQL Table
This assumes you already have an Azure SQL database set up, if not you can download the SQL Server engine and host it locally.<br>
Set up a simple table to store the data

```
CREATE TABLE [dbo].[WORK_BATTERY](
	[Percent] [tinyint] NULL,
	[Charge] [bit] NULL
) ON [PRIMARY]
GO

```
