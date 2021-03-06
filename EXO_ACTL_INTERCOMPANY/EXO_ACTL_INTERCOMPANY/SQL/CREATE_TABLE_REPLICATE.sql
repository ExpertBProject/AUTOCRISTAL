CREATE COLUMN TABLE REPLICATE ( DBNAMEORIG NVARCHAR(100)  NOT NULL,
	 DBNAMEDEST NVARCHAR(100)  NOT NULL,
	 TABLECATEGORY INT  NOT NULL,
	 TABLENAME NVARCHAR(20) NOT NULL,
	 CODETABLE NVARCHAR(50)  NOT NULL,
	 CODETABLE2 NVARCHAR(50) ,
	 CODETABLE3 NVARCHAR(50) ,
	 CODETABLE4 NVARCHAR(50) ,
	 DATEADD LONGDATE, PRIMARY KEY (DBNAMEORIG,DBNAMEDEST,TABLECATEGORY,TABLENAME,CODETABLE)) UNLOAD PRIORITY 5 AUTO MERGE 