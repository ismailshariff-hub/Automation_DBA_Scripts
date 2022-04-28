-- Author:           DBA
-- Description:      move DB files to diff location, dynamically. includes set OFFLINE/ONLINE

-- Get database file information for each database 
IF     OBJECT_ID('TempDB..#holdforeachdb') IS NOT NULL
DROP TABLE #holdforeachdb;

create table #holdforeachdb
(      [databasename] [nvarchar](128) collate sql_latin1_general_cp1_ci_as not null, 
       [size] [int] not null, 
       [name] [nvarchar](128) collate sql_latin1_general_cp1_ci_as not null, 
       [filename] [nvarchar](260) collate sql_latin1_general_cp1_ci_as not null
)
    INSERT      
    INTO    #holdforeachdb exec sp_MSforeachdb 
                           'select ''?'' as databasename,
                           [?]..sysfiles.size, 
                           [?]..sysfiles.name, 
                           [?]..sysfiles.filename
                           from [?]..sysfiles '

--NEW location of DB files
DECLARE      @NewDataPath NVARCHAR(4000)='C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA1\', /*!!!!!!MODIFY ACCORDINGLY!!!!!!*/
                     @NewTlogPath NVARCHAR(4000)='C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\LOG1\'      /*!!!!!!MODIFY ACCORDINGLY!!!!!!*/

;WITH DataBasefiles (dbname, size_Gb, logical_name, Path, PhysFileName, FileType)
AS
(select       databasename ,
              (size*8.00/1024/1024) size_Gb , 
              sf.name logical_name, 
              LEFT(FileName,LEN(FileName)-CHARINDEX('\',REVERSE(FileName))+1) Path, 
              RIGHT(FileName,CHARINDEX('\',REVERSE(FileName))-1) PhysFileName,
              SUBSTRING([filename], (LEN(filename)-2), 4) AS FileType
from #holdforeachdb sf
JOIN sys.databases db on db.name=sf.databasename)

select dbname, 
              --size_Gb, 
              logical_name, 
              Path, 
              PhysFileName, 
              FileType,
              CASE 
              WHEN FileType = 'ldf'      THEN 'USE [master]; ALTER DATABASE '+'['+dbname+']'+' SET OFFLINE WITH ROLLBACK IMMEDIATE;'
              ELSE '' END                                                          AS 'SET_DB_OFFLINE',
              'USE [master]; ALTER DATABASE '+QUOTENAME(dbname)+' MODIFY FILE (Name = '+logical_name+' , FileName = N'''+CASE
                                         WHEN FileType = 'mdf'      THEN @NewDataPath
                                         WHEN FileType = 'ndf'      THEN @NewDataPath
                                         WHEN FileType = 'ldf'      THEN @NewTlogPath
                                         END +''+PhysFileName+''');' AS 'MOVE_DB_FILES_CMD',
              CASE
              WHEN FileType = 'ldf'      THEN 'USE [master]; ALTER DATABASE '+'['+dbname+']'+' SET ONLINE;'
              ELSE '' END                                                          AS 'SET_DB_ONLINE'
FROM DataBasefiles   
where dbname IN ('AdventureWorks2017','DBA','AdventureWorksDW2017') /*******add list of DBs within IN clause*******/

