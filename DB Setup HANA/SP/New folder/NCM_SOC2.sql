IF NOT EXISTS(SELECT name FROM  dbo.sysobjects WHERE xtype = 'U' AND name = '@NCM_SOC2')
 BEGIN 
	 CREATE TABLE [@NCM_SOC2]
	 (ID         NVARCHAR(8)         NOT NULL,
	 Notes      NVARCHAR(1000)      NOT NULL,
	 Image    IMAGE)
	 INSERT INTO [@NCM_SOC2]
	 VALUES (
	 '1',
	 'Note:   Any payments received after end of the month will be shown in next month''s statement.
			  paninpaaIf you do not agree with the above statement, please inform us immediately.'
	 , NULL)
 END
 