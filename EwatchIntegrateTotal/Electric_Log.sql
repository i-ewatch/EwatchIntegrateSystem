CREATE TABLE [dbo].[Electric_Log]
(
	[ttime] NVARCHAR(8) NOT NULL , 
    [ttimen] DATETIME NOT NULL, 
    [CardNumber] NVARCHAR(6) NOT NULL DEFAULT 000000,
    [BordNumber] NVARCHAR(2) NOT NULL DEFAULT 00,
    [DeviceNumber] INT NOT NULL DEFAULT 0, 
    [Start1] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [End1] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Start2] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [End2] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Total] DECIMAL(18, 2) NOT NULL DEFAULT 0, 
    PRIMARY KEY ([ttime], [DeviceNumber], [ttimen], [CardNumber], [BordNumber]),
)
