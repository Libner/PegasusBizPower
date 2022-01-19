http://pegasus.bizpower.co.il/pegasus/getmessages.aspx?fromDate=1/14/2013 00:00&toDate=1/14/2013 23:59
http://pegasus.bizpower.co.il/pegasus/getmessages.aspx?fromDate=1/14/2013 09:00&toDate=1/14/2013 10:20

http://pegasus.bizpower.co.il/pegasus/getmessages.aspx?fromDate=1/14/2013 23:00&toDate=1/14/2013 24:00


DECLARE  @Result Table(
SPID int, Status varchar(50), Login varchar(50),
  HostName varchar(50), BlkBy varchar(50), DBName sysname NULL, Command varchar(max), CPUTIME int, DiskIO int,
    LastBatch varchar(255), ProgramName varchar(255), SPID2 int, REQUESTID int);
Insert Into @Result Exec sp_who2;
SELECT * FROM @Result WHERE dbname = 'bizpower_pegasus' and login='bizpower_pegasus_user'



Select 'Kill '+ CAST(p.spid AS VARCHAR)KillCommand into #temp
from master.dbo.sysprocesses p (nolock)
join master..sysdatabases d (nolock) on p.dbid = d.dbid
Where d.[name] = 'bizpower_pegasus'


http://pegasus.bizpower.co.il/pegasus/getmessages.aspx?fromDate=10/02/2014 08:30&toDate=10/02/2014 10:30



http://mtours2.bizpower.co.il/kishurit/getmessages.aspx?fromDate=10/01/2015&toDate=10/14/2015



http://milenium-biz.bizpower.co.il/kishurit/getmessages.aspx?fromDate=10/01/2015&toDate=10/13/2015