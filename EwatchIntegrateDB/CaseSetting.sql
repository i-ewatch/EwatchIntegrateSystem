CREATE TABLE [dbo].[CaseSetting]
(
	[CardNumber] NVARCHAR(6) NOT NULL DEFAULT 000000 , 
    [BordNumber] NVARCHAR(2) NOT NULL DEFAULT 00, 
    [CaseName] NVARCHAR(50) NOT NULL DEFAULT '', 
    PRIMARY KEY ([BordNumber], [CardNumber])
)
