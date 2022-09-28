CREATE TABLE [dbo].[DeviceSetting]
(
	[CardNumber] NVARCHAR(6) NOT NULL DEFAULT 000000,
    [BordNumber] NVARCHAR(2) NOT NULL DEFAULT 00,
    [DeviceNumber] INT NOT NULL DEFAULT 0,
    [DeviceName] NVARCHAR(20) NOT NULL DEFAULT '',
    [DeviceTypeEnum] INT NOT NULL DEFAULT 0,
    CONSTRAINT [PK_DeviceSetting] PRIMARY KEY ([CardNumber], [DeviceNumber], [BordNumber]),
)
