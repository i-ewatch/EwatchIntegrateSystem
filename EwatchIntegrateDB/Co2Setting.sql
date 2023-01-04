CREATE TABLE [dbo].[Co2Setting]
(
	[Year] INT NOT NULL , 
    [Value] DECIMAL(18, 3) NULL, 
    CONSTRAINT [PK_Co2Setting] PRIMARY KEY ([Year])
)
