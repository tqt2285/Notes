GO
SET ANSI_NULLS, QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [sp_getTableColumName]
	@LikeName nvarchar(max)
WITH ENCRYPTION AS
BEGIN
	SET NOCOUNT ON;
print 'exec sp_getTableColumName N''' + @LikeName + '''
return'
	SET @LikeName = LTRIM(RTRIM(@LikeName))
	SET @LikeName = REPLACE(@LikeName,'''','''''')
	if LEFT(@LikeName,1) <> '%' and RIGHT(@LikeName,1) <> '%'
		SET @LikeName = '%'+@LikeName+'%'
	
	SET @LikeName = REPLACE(REPLACE(@LikeName,'[',''),']','')
	
	DECLARE @Called TABLE(ID int identity(1,1),Called nvarchar(500))
	DECLARE @Proc TABLE(ID int identity(1,1),ProcedureName nvarchar(500))
	DECLARE @Function TABLE(ID int identity(1,1),FunctionName nvarchar(500))
	DECLARE @Table TABLE(ID int identity(1,1),TableName nvarchar(500))
	DECLARE @Column TABLE(ID int identity(1,1),ColumnName nvarchar(500))
	DECLARE @Trigger TABLE(ID int identity(1,1),TriggerName nvarchar(500))
	CREATE TABLE #Setting (ID int identity(1,1),SettingKey nvarchar(1000),SettingTable nvarchar(500),SettingValue nvarchar(max))
	DECLARE @Query nvarchar(max),@WhereString nvarchar(max)


	if @LikeName like '[%]%' and @LikeName like '%[%]'
		SET @WhereString = N'([text] LIKE N'''+@LikeName+''')'
	else
		SET @WhereString = N'(REPLACE([text],'' '','''') LIKE REPLACE(N'''+@LikeName+''','' '',''''))'


	if @LikeName like '%&%' and @LikeName like '[%]%&%[%]'
	begin
		SET @WhereString = '('
		select @WhereString = @WhereString + N'REPLACE([text],'' '','''') LIKE REPLACE(N''%'+REPLACE(Items,'%','')+'%'','' '','''') and ' from dbo.SplitString(@LikeName,'&')
		SET @WhereString = LEFT(@WhereString,LEN(@WhereString)-4) + ')'
	end
	else if @LikeName like '%|%' and @LikeName like '[%]%|%[%]'
	begin
		SET @WhereString = '('
		select @WhereString = @WhereString + N'REPLACE([text],'' '','''') LIKE REPLACE(N''%'+REPLACE(Items,'%','')+'%'','' '','''') or ' from dbo.SplitString(@LikeName,'|')
		SET @WhereString = LEFT(@WhereString,LEN(@WhereString)-3) + ')'
	end

	SET @Query = N'SELECT OBJECT_NAME(id) FROM SYSCOMMENTS WHERE '+@WhereString+' GROUP BY OBJECT_NAME(id) order by OBJECT_NAME(id)'
	print @Query
	insert into @Called(Called)			exec sp_executesql @Query
	insert into @Proc(ProcedureName)	select [name] from sys.procedures where [name] like @LikeName order by [name]
	insert into @Function(FunctionName) select [name] from sys.objects WHERE type_desc LIKE '%FUNCTION%' and [name] like @LikeName order by [name]
	insert into @Table(TableName)		select [name] from sys.tables where [name] like @LikeName order by [name]
	insert into @Column(ColumnName)		select object_name(object_id) + '.' + [name] from sys.columns where [name] like @LikeName order by [name]
	insert into @Trigger(TriggerName)	select [name] + ' [' + object_name(parent_id) + ']' from sys.triggers where [name] like @LikeName or object_name(parent_id) like @LikeName order by [name]
	if object_id('MEN_Menu') is not null
	begin
		DECLARE @QueryInsert nvarchar(max)
		if COL_LENGTH('MEN_Menu','ClassName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select MenuID,''select * from MEN_Menu where MenuID = '''''' + MenuID +'''''''',ClassName from MEN_Menu where (ClassName like @LikeName or MenuID like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblSC_Object','ObjectName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select ObjectID,''select * from tblSC_Object where ObjectID = '' + CAST(ObjectID as varchar),ObjectName from tblSC_Object where (ObjectName like @LikeName or Description like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		
		if COL_LENGTH('tblDataSetting','TableName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ViewName,* from tblDataSetting where TableName = '''''' + TableName +'''''''',ViewName from tblDataSetting where (TableName like @LikeName or ViewName like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblDataSetting','TableEditorName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select TableEditorName,* from tblDataSetting where TableName = '''''' + TableName +'''''''',TableEditorName from tblDataSetting where (TableEditorName like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','RptTemplate') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select RptTemplate,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',RptTemplate from tblDataSetting where (RptTemplate like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','spAction') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select spAction,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',spAction from tblDataSetting where (spAction like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end		
		if COL_LENGTH('tblDataSetting','ComboboxColumns') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ComboboxColumns,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ComboboxColumns from tblDataSetting where (ComboboxColumns like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ComboboxColumn_BackupForTransfer') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ComboboxColumn_BackupForTransfer,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ComboboxColumn_BackupForTransfer from tblDataSetting where (ComboboxColumn_BackupForTransfer like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ValidationProcedures') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ValidationProcedures,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ValidationProcedures from tblDataSetting where (ValidationProcedures like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ControlStateProcedure') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ControlStateProcedure,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ControlStateProcedure from tblDataSetting where (ControlStateProcedure like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ExecuteProcBeforeLoadData') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ExecuteProcBeforeLoadData,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ExecuteProcBeforeLoadData from tblDataSetting where (ExecuteProcBeforeLoadData like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','Import') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select [Import],* from tblDataSetting where TableName = ''''''+TableName+'''''' '',[Import] from tblDataSetting where ([Import] like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ProcBeforeSave') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ProcBeforeSave,ProcAfterSave,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ISNULL(ProcBeforeSave,'''') + ISNULL('' - '' + ProcAfterSave,'''') from tblDataSetting where (ProcBeforeSave like @LikeName or ProcAfterSave like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ProcBeforeDelete') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ProcBeforeDelete,ProcAfterDelete,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ISNULL(ProcBeforeDelete,'''') + ISNULL('' - '' + ProcAfterDelete,'''') from tblDataSetting where (ProcBeforeDelete like @LikeName or ProcAfterDelete like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		if COL_LENGTH('tblDataSetting','ColumnChangeEventProc') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select TableName,''select ColumnChangeEventProc,* from tblDataSetting where TableName = ''''''+TableName+'''''' '',ColumnChangeEventProc from tblDataSetting where (ColumnChangeEventProc like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblCommonColumnCombobox','ColumnName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select ColumnName,''select ColumnName,* from tblCommonColumnCombobox where ColumnName = ''''''+ColumnName+'''''' '',ISNULL(ColumnName,'''') from tblCommonColumnCombobox where (ColumnName like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblCommonColumnCombobox','Query') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select ColumnName,''select Query,* from tblCommonColumnCombobox where ColumnName = ''''''+ColumnName+'''''' '',ISNULL(Query,'''') from tblCommonColumnCombobox where (Query like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblCommonColumnCombobox','TableEditorName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select ColumnName,''select TableEditorName,* from tblCommonColumnCombobox where ColumnName = ''''''+ColumnName+'''''' '',ISNULL(TableEditorName,'''') from tblCommonColumnCombobox where (TableEditorName like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

		if COL_LENGTH('tblCommonColumnComboboxSigned','ColumnName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select NULL,''select * from tblCommonColumnComboboxSigned where ColumnName = ''''''+ColumnName+'''''' and TableName = ''''''+TableName+'''''' '',ISNULL(TableName,'''') + ISNULL('' - '' + ColumnName,'''') from tblCommonColumnComboboxSigned where (ColumnName like @LikeName or TableName like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end
		
		insert into #Setting(SettingKey,SettingTable,SettingValue) select ExportName,'select * from tblExportList where ExportName = ''' + ExportName +'''',ProcedureName from tblExportList where (ExportName like @LikeName or ProcedureName like @LikeName or TemplateFileName like @LikeName)
		insert into #Setting(SettingKey,SettingTable,SettingValue) select ProcName,'select * from tblProcedureName where ProcID = ' + CAST(ProcID as varchar),ProcName from tblProcedureName where (ProcName like @LikeName)
		insert into #Setting(SettingKey,SettingTable,SettingValue) select ParamName,'select * from tblExportParameterList where ParamName = ''' + CAST(ParamName as varchar) + '''',Query from tblExportParameterList where (Query like @LikeName)
		insert into #Setting(SettingKey,SettingTable,SettingValue) select Code,'select * from tblParameter where Code = ''' + CAST(Code as varchar) + '''',[Value] from tblParameter where ([Code] like @LikeName or [Description] like @LikeName)
		if COL_LENGTH('TaskSchedule','FunctionName') is not null
		begin
			SET @QueryInsert = N'insert into #Setting(SettingKey,SettingTable,SettingValue) select IDTask,''select * from TaskSchedule where IDTask = '' + CAST(IDTask as varchar) + '''',ISNULL(ProducerName,'''') + ISNULL(ClassName + ''.'' + FunctionName,'''') from TaskSchedule where ([ProducerName] like @LikeName or [FunctionName] like @LikeName)'
			exec sp_executesql @QueryInsert, N'@LikeName nvarchar(max)',@LikeName
		end

	end

	select a.Called, b.ProcedureName, e.FunctionName, c.TableName, d.ColumnName, f.TriggerName, ss.SettingKey,ss.SettingTable,ss.SettingValue
		from @Called a full outer join @Proc b on a.ID = b.ID
		full outer join @Table c on a.ID = c.ID
		full outer join @Column d on a.ID = d.ID
		full outer join @Function e on a.ID = e.ID
		full outer join @Trigger f on a.ID = f.ID
		full outer join #Setting ss on a.ID = ss.ID
	drop table #Setting
END
GO
