CREATE COLUMN TABLE "INTERNETAC_NEW"."REG_CLIENTES" ("ID" INTEGER CS_INT NOT NULL , "IP" VARCHAR(20), "USUARIO" VARCHAR(200), "CUENTA" VARCHAR(200), "CLAVE" VARCHAR(200), "FECHA" LONGDATE CS_LONGDATE, PRIMARY KEY ("ID")) UNLOAD PRIORITY 5  AUTO MERGE 