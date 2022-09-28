CREATE PROCEDURE [dbo].[KwhTotalProcedure]
	@nowTime CHAR(8),
	@cardNumber NVARCHAR(6),
	@bordNumber NVARCHAR(2),
	@deviceNumber INT = 0,
	@nowKwh DECIMAL(18,2) = 0
AS
	IF NOT EXISTS(SELECT * FROM sys.tables WHERE name = 'Electric_Log')
		BEGIN
		CREATE TABLE [dbo].[Electric_Log](
		[ttime] CHAR(8) NOT NULL,
		[ttimen] DATETIME NOT NULL,
		[CardNumber] NVARCHAR(6) NOT NULL,
		[BordNumber] NVARCHAR(2) NOT NULL,
		[DeviceNumber] INT NOT NULL,
		[Start1] DECIMAL(18,2) DEFAULT((0)) NOT NULL ,
		[End1] DECIMAL(18,2) DEFAULT((0)) NOT NULL ,
		[Start2] DECIMAL(18,2) DEFAULT((0)) NOT NULL ,
		[End2] DECIMAL(18,2) DEFAULT((0)) NOT NULL,
		[Total] DECIMAL(18,2) DEFAULT((0)) NOT NULL,
		CONSTRAINT [PK_Electric_Log] PRIMARY KEY([ttime],[ttimen],[CardNumber],[BordNumber],[DeviceNumber]),
		)
		END
		DECLARE @result INT
		DECLARE @nowdate CHAR(8) = CONVERT(CHAR(8),@nowTime,112)
		DECLARE @Start1 DECIMAL(18,2) = 0
		DECLARE @End1 DECIMAL(18,2) = 0
		DECLARE @Start2 DECIMAL(18,2) = 0
		DECLARE @End2 DECIMAL(18,2) = 0
		DECLARE @Total DECIMAL(18,2) = 0
		SELECT TOP 1 1 FROM Electric_Log WHERE ttime = @nowdate AND CardNumber = @cardNumber AND BordNumber = @bordNumber AND DeviceNumber = @deviceNumber
		IF @@ROWCOUNT > 0 SET @result = 1
		ELSE SET @result = 0
		IF(@result=1)
		 BEGIN
		 SELECT @Start1 = Start1,@End1 = End1,@Start2 = Start2,@End2 = End2 FROM Electric_Log WHERE ttime = @nowdate AND CardNumber = @cardNumber AND BordNumber = @bordNumber AND DeviceNumber = @deviceNumber
		 IF(@Start1 = 0 AND @End1 = 0 AND @Start2 = 0 AND @End2 = 0)
			BEGIN
				SET @Start1 = @nowKwh
				SET @End1 = @nowKwh
			END
		 ELSE IF(@End1 <= @nowKwh AND @Start2 = 0)
			BEGIN
				SET @End1 = @nowKwh
			END
		 ELSE IF(@Start2 = 0 AND @End1 > @nowKwh)
			BEGIN
				SET @Start2 = @nowKwh
				SET @End2 = @nowKwh
			END
		 ELSE IF (@Start2 <>0 AND @End2 <= @nowKwh)
			BEGIN
				SET @End2 = @nowKwh
			END
		 END
		ELSE
			BEGIN
				INSERT INTO Electric_Log (ttime,ttimen,CardNumber,BordNumber,DeviceNumber,Start1,End1) VALUES (@nowdate,@nowTime,@cardNumber,@bordNumber,@deviceNumber,@nowKwh,@nowKwh)
				SET @Start1 = @nowKwh
				SET @End1 = @nowKwh
			END
	SET @Total = (@End1-@Start1)+(@End2-@Start2)
	UPDATE Electric_Log SET [Start1] = @Start1,[End1] = @End1,[Start2] = @Start2,[End2] = @End2,[Total]=@Total WHERE ttime = @nowdate AND CardNumber = @cardNumber AND BordNumber = @bordNumber AND DeviceNumber = @deviceNumber
RETURN @Total
