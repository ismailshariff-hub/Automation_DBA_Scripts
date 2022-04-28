SET NOCOUNT ON;
Declare @dbname      sysname
DECLARE @backupStartDate datetime 
DECLARE @backup_set_id_start INT 
Declare @servername      sysname
SELECT @ServerName = @@SERVERNAME
--PRINT @ServerName
declare @backupfiles table 
(dbname  sysname,
filepath nvarchar(max)
)
Declare @DBlist table
(DbName1 Varchar(50))
Insert into @DBlist
Select name from sys.sysdatabases where name NOT in ('master','msdb','model','tempdb','DISTRIBUTION','DBA');
DECLARE cur CURSOR FOR SELECT DbName1 from @DBlist
OPEN cur

FETCH NEXT FROM cur INTO @dbname

WHILE @@FETCH_STATUS = 0 BEGIN
SELECT @backup_set_id_start = MAX(backup_set_id) 
		FROM msdb.dbo.backupset
		WHERE database_name = @dbname AND type = 'D' and is_copy_only !=1;

insert into @backupfiles
SELECT bks.database_name, 
              bkf.physical_device_name
FROM  msdb.dbo.backupmediafamily bkf INNER JOIN
      msdb.dbo.backupset bks ON bkf.media_set_id =bks.media_set_id
WHERE bks.database_name = @dbname
and server_name= @servername
--bks.backup_start_date > (GETDATE() -1) 
AND bks.type='D'
AND bks.backup_set_id = @backup_set_id_start -- it will select only the last full backup file
and is_copy_only !=1--added since the VM is being backed up with copy_only param by VSS
ORDER BY bks.backup_finish_date DESC

DECLARE       @cmd                 NVARCHAR(500),
              @backupFile          NVARCHAR(500),
              @filelocation NVARCHAR(500)


SET @filelocation=(SELECT TOP 1 CASE 
                                             WHEN (SUBSTRING(filepath,1,5)) = 'https'       THEN 'URL'--LIKE '%https%'       
                                             WHEN (SUBSTRING(filepath,1,2)) = '\\'          THEN 'DISK'
                                             WHEN (SUBSTRING(filepath,2,1)) = ':'           THEN 'DISK'
                                             END 
                                FROM @backupfiles)

DECLARE @FileList NVARCHAR(MAX), 
              @SQL NVARCHAR(MAX)
SELECT @FileList = SUBSTRING( 
( 
                 SELECT 
     ',' 
       + CHAR(10) + ' '+@filelocation+' = N''' 
    + filepath
    + ''''
    AS [text()] 
FROM @backupfiles where dbname=@dbname
                FOR XML PATH ('')
), 2 , 9999)
SELECT @SQL = ' ALTER DATABASE '+ @dbName +' SET SINGLE_USER WITH ROLLBACK IMMEDIATE;' + CHAR(10) + ' RESTORE DATABASE '+ @dbName +' FROM '  + @FileList +CHAR(10) + ' WITH REPLACE,RECOVERY' +CHAR(10) + ' GO'
Print @SQL
    FETCH NEXT FROM cur INTO @dbname
       END
CLOSE cur    
DEALLOCATE cur
