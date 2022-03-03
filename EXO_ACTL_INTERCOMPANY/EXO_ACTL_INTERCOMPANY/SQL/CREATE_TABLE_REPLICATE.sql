
CREATE COLUMN TABLE REPLICATE ( dbNameOrig NVARCHAR(100)  NOT NULL,
	 dbNameDest NVARCHAR(100)  NOT NULL,
	 tableCategory INT  NOT NULL,
	 tableName NVARCHAR(20) NOT NULL,
	 codeTable NVARCHAR(50)  NOT NULL,
	 codeTable2 NVARCHAR(50) ,
	 codeTable3 NVARCHAR(50) ,
	 codeTable4 NVARCHAR(50) ,
	 dateAdd LONGDATE PRIMARY KEY (dbNameOrig,dbNameDest,tableCategory,tableName,codeTable)) UNLOAD PRIORITY 5 AUTO MERGE 