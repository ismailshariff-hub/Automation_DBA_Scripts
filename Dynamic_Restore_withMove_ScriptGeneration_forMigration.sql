-- =============================================
-- Authors: <Ismail Shariff>
-- Create date: <20/05/2020>
-- Description:	This script prepares Restore script for list of databases required to be restored
-- IN PATH -backup path is taken from the last full backup from source (NOT INCLUDING 'MASTER','MSDB','MODEL','TEMPDB','DISTRIBUTION')
-- Additional details: Uses temp tables and dynamically prepares the script for restore with move,
-- we can modify making more dynamic if required by taking location from default data and log file location
-- =============================================

IF OBJECT_ID(N'tempdb..#tmp') IS NOT NULL
BEGIN
     DROP TABLE #tmp
END
go
IF OBJECT_ID(N'tempdb..#tmp1') IS NOT NULL
BEGIN
     DROP TABLE #tmp1
END
go
IF OBJECT_ID(N'tempdb..#tmp2') IS NOT NULL
BEGIN
     DROP TABLE #tmp2
END
go
DECLARE @databaseName sysname 
DECLARE @backupStartDate datetime 
DECLARE @backup_set_id_start INT 
declare @path varchar(1000)
Declare @Filelistonly varchar(1000)
Declare @RestoreString as Varchar(max)
Declare @NRestoreString as NVarchar(max)
DECLARE @LogicalName  as varchar(500)
DECLARE @pname  as NVarchar(max)
Declare @counter as int
Declare @rows as int
--Creating table to capture details of Restore Filelistonly output from backup file to be restored
create table #tmp
(
LogicalName nvarchar(128) 
,PhysicalName nvarchar(260) 
,Type char(1) 
,FileGroupName nvarchar(128) 
,Size numeric(20,0) 
,MaxSize numeric(20,0),
Fileid tinyint,
CreateLSN numeric(25,0),
DropLSN numeric(25, 0),
UniqueID uniqueidentifier,
ReadOnlyLSN numeric(25,0),
ReadWriteLSN numeric(25,0),
BackupSizeInBytes bigint,
SourceBlocSize int,
FileGroupId int,
LogGroupGUID uniqueidentifier,
DifferentialBaseLSN numeric(25,0),
DifferentialBaseGUID uniqueidentifier,
IsReadOnly bit,
IsPresent bit, 
TDEThumbPrint varchar(50),
SnapshotUrl varchar(max)
)
--Creating table for restore filelistonly has additional column in SQL verson above 2012 hence below if statement to check the version and remove the additional column SnapshotUrl in lower versions
declare @sqlver nvarchar(max)
if (select cast(left(cast(serverproperty('productversion') as varchar), 4) as decimal(5, 3))) < 12
Begin
Alter table #tmp drop Column SnapshotUrl;
end

--Creating this table to capture the databases planned to be restored which can be mentioned in where name filter or we can manually set the database name to be restored with code change
Create Table #tmp1
(DBname Varchar(500));
Insert into #tmp1
Select name from sys.sysdatabases where name NOT in ('master','msdb','model','tempdb','DISTRIBUTION')

--Creating this table to store restore script generated for all the databases
Create Table #tmp2
(RestoreString Varchar(max));

--Starting Outer cursor for getting the database names one by one to create the restore script
DECLARE My_outercursor CURSOR
FOR 
SELECT DBname from #tmp1;
OPEN My_outercursor;
FETCH NEXT FROM My_outercursor INTO @databaseName;
WHILE @@FETCH_STATUS = 0
    BEGIN
--Set @databaseName='DBA';
--Getting the last backup file id to generate the backup path to prepare the restore script

SELECT @backup_set_id_start = MAX(backup_set_id) 
		FROM msdb.dbo.backupset 
		WHERE database_name = @databaseName AND type = 'D';
		
-- Getting backup path from last full backup taken for each databases to be restored.

select @path = (SELECT '''' + mf.physical_device_name + ''''
		FROM msdb.dbo.backupset b, 
		msdb.dbo.backupmediafamily mf 
		WHERE b.media_set_id = mf.media_set_id 
		AND b.database_name = @databaseName 
		AND b.backup_set_id = @backup_set_id_start)
--print @path
-- Preparing the restore script based on the backup path 
SET @Filelistonly= ('restore filelistonly from disk =' + @path);
--print @Filelistonly
insert #tmp
EXEC(@Filelistonly)

-- setting the counter and generating the move script according to the number of the logical names for each physical files and select @Rows as [These are the number of rows]
set @counter = 1
select @rows = COUNT(*) from #tmp

--Starting inner cursor for preparing the Restore script according to the logical name of the data files
DECLARE MY_CURSOR Cursor 
FOR 
Select LogicalName From #tmp
Open My_Cursor 
FETCH NEXT FROM MY_CURSOR INTO @LogicalName
Select @RestoreString = ('RESTORE DATABASE '+ @databaseName +' FROM DISK = N'+ @path
 +' '+ ' with  ' )
While (@@FETCH_STATUS <> -1)
BEGIN
IF (@@FETCH_STATUS <> -2)
select @PName=RIGHT(physicalName, CHARINDEX('\', REVERSE(physicalName)) - 1) from #tmp where LogicalName=@LogicalName;
select @RestoreString =
case 
when @counter = 1 then 
   @RestoreString + 'move  N''' + @LogicalName + '''' + ' TO N''F:\DATA\'+
   @PName + '''' + ', '
when @counter > 1 and @counter < @rows then
   @RestoreString + 'move  N''' + @LogicalName + '''' + ' TO N''F:\DATA\'+
   @PName  + '''' + ', '
When @LogicalName like '%log%' then
   @RestoreString + 'move  N''' + @LogicalName + '''' + ' TO N''L:\TLOG\'+
   @PName +''''+';'
end
set @counter = @counter + 1
FETCH NEXT FROM MY_CURSOR INTO @LogicalName
END
set @NRestoreString = @RestoreString
insert into #tmp2
select @NRestoreString
Truncate table #tmp; -- Clearing the table to start preparing to generate restore script for next database
CLOSE MY_CURSOR
DEALLOCATE MY_CURSOR
FETCH NEXT FROM My_outercursor INTO @databaseName;
END  -- Outer cursor loop
CLOSE My_outercursor;
DEALLOCATE My_outercursor;
select * from #tmp2; -- Final output


