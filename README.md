<div align="center">

## Interbase and ADO \(open source DB for Windows\)


</div>

### Description

This tutorial will give you a crash cource in working with Interbase (A free, open source databese from Borland) and ADO. It will show you where to get the database engine and teach you some simple basics.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-17 16:07:38
**By**             |[BDCSoftware](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bdcsoftware.md)
**Level**          |Advanced
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Interbase\_4273812172001\.zip](https://github.com/Planet-Source-Code/bdcsoftware-interbase-and-ado-open-source-db-for-windows__1-29899/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<title>Interbase ADO Tutorial</title>
</head>
<body lang=EN-US link=blue vlink=purple class="Normal" bgcolor="#FFFFFF">
<h2><span style='font-size:12.0pt;'><font face="Arial, Helvetica, sans-serif" size="5">Interbase
 ADO Tutorial</font></span></h2>
<p><span style='font-family:Arial'>Has anybody ever wondered if there is an Open
 Source alternative to SQL Server or Access databases? Well, I have, and I found
 Interbase. Interbase is a Client/Server database from Borland. It is Open Source.
 It runs on Windows, Linux and bunch of other *nix platforms. It has a very small
 memory footprint and it is relatively fast. It will also support large database
 files (larger the 2 gig. I know a guy that has a 300 Gig database up and running)</span></p>
<p><span style='font-family:Arial'>Anyhow, in this article I will describe the
 issues and the necessary tools to get you up and running with Interbase. </span></p>
<p><span style='font-family:Arial'>First let me tell you about the benefits of
 Interbase:</span></p>
<ol start=1 type=1>
 <li><span
  style='font-family:Arial'>Open Source</span></li>
 <li><span
  style='font-family:Arial'>Fast</span></li>
 <li><span
  style='font-family:Arial'>Small size</span></li>
 <li><span
  style='font-family:Arial'>Very easy distribution (scripts for Wise or InstallShield
 are available)</span></li>
 <li><span
  style='font-family:Arial'>Works ADO</span></li>
 <li><span
  style='font-family:Arial'>Works with ODBC</span></li>
 <li><span
  style='font-family:Arial'>Awesome transaction management (readers never block
 writers and vice versa)</span></li>
 <li><span
  style='font-family:Arial'>Multiple platform support (Linux/Unix)</span></li>
 <li><span
  style='font-family:Arial'>Superb support for BOLB fields (Images and memo
 fields)</span></li>
 <li><span
  style='font-family:Arial'>Support for Arrays (you can store Arrays in individual
 fields)</span></li>
</ol>
<p><span style='font-family:Arial'>For starters you need to get the server and
 client software. You can get the original Open Source version (Source and Binaries)
 from Borland at: </span></p>
<p><span style='font-family:Arial'><a
href="http://www.borland.com/devsupport/interbase/opensource/">http://www.borland.com/devsupport/interbase/opensource/</a></span></p>
<p><span style='font-family:Arial'>or get it a modified version (Firebird) from:<br>
 <br>
 <a href="http://www.ibphoenix.com/ibp_download.html">http://www.ibphoenix.com/ibp_download.html</a></span></p>
<p><span style='font-family:Arial'>Download and install the server and client
 binaries. The Interbase server ships with a ODBC driver, but I hate ODBC and
 use ADO/OleDB on a day to day basis. So I had to find an OleDB driver for Interbase.
 Luckily there are numerous available. You can find a links to download sites
 on this site:<br>
 <br>
 <a href="http://www.interbase2000.org/tools_conn.htm">http://www.interbase2000.org/tools_conn.htm</a>
 </span></p>
<p><span style='font-family:Arial'>I opted for the IBProvider from <a href="http://www.lcpi.lipetsk.ru/prog/eng/index.html">http://www.lcpi.lipetsk.ru/prog/eng/index.html</a>
 because they had some VB samples of how to use the provider with ADO. The version
 that you can download is an Evaluation for 30 days. If you want a completely
 free OleDB provider then use: <a
href="http://www.oledb.net/?Page=FAQ">http://www.oledb.net/?Page=FAQ</a>. However,
 all my sample code is tested with IBProvider only.</span></p>
<p><span style='font-family:Arial'>Once you have downloaded and installed all
 the files, you are ready for development. IB (Interbase) ships with a sample
 database called employee.gdb. We will use this database as an example. (You
 can find it in &#8216;C:\Program Files\Borland\InterBase\examples\Database&#8217; , provided
 you installed the server in the default location). Anyhow, lets start with the
 basics:</span></p>
<p> </p>
<h2><span style='font-size:12.0pt;'><font size="5" face="Arial, Helvetica, sans-serif">Connecting
 to Interbase</font></span></h2>
<p><span style='font-family:Arial'>Lets establish a connection to the database.
 A sample connection:</span></p>
<p><span><font face="Courier New, Courier, mono" size="3">    </font></span><font face="Courier New, Courier, mono" size="3">Dim
 adoConn As New ADODB.Connection<br>
 <br>
 <span>    </span>adoConn.ConnectionString = "provider=LCPI.IBProvider;data
 source=localhost:C:\Interbase    DBs\Employee.gdb;ctype=win1251;user id=SYSDBA;password=masterkey"</font></p>
<p><font face="Courier New, Courier, mono" size="3"><span>    </span>adoConn.Open<span
style='font-family:Arial'></span></font></p>
<p><span style='font-family:Arial'>Ok, here are a few things to consider:<br>
 Default user name and password (like SA in SQLServer) are SYSDBA and masterkey
 (case sensitive). The &#8216;data source&#8217; parameter has a following syntax: <i>IP
 Address:file location on the remote system</i> . If you installed the server
 on your development machine then use localhost or your IP. If you installed
 it on a remote machine then use the IP Address of the machine. The <i>file location</i>
 is a bit weird. It is local to the server and you can&#8217;t use UNC paths.</span></p>
<p><span style='font-family:Arial'>Once the connection is open, we can start working
 with the database.</span></p>
<p> </p>
<h2><span style='font-size:12.0pt;'><font face="Arial, Helvetica, sans-serif" size="5">Working
 with an Interbase database</font></span></h2>
<p><span style='font-family:Arial'>For the most part, working with Interbase is
 as easy as working with SQL Server or Access. However there are a few things
 to consider: </span></p>
<p><span style='font-family:Arial'>For one, Interbase uses <i>dialects, </i>basically
 it&#8217;s the SQL syntax that you issue your commands to the database. IB 6.0 can
 use Dialect 1 (legacy) and Dialect 3. The sample databases are in written in
 Dialect 1. If you decide to use Dialect 3 (as I have), you will notice some
 weird behavior. If your database has lower case table and field names, you will
 have to surround them with double quotes. For instance: <i>Select &#8220;CompanyName&#8221;,
 &#8220;Address&#8221; from &#8220;tblCustomers&#8221;. </i>Needless to say this will create havoc with
 VB programmers </span><span>J</span><span
style='font-family:Arial'>. One workaround is to use caps for table and field
 names. (Btw, don&#8217;t ask me why this is the way it is.) For Instance: SELECT COMPAN_YNAME,
 ADDRESS FROM TBLCUSTOMERS.</span></p>
<p><span style='font-family:Arial'>The other issue that I have found is: you cannot
 use <i>adCmdStoredProc </i>as your command type. Workaround for this: use <i>adCmdText</i>.
 But more on this later.</span></p>
<p><span style='font-family:Arial'>Ok, so how would we get some data in and out
 of our database? Well, you can use your normal recordset object to execute a
 SQL statement or you can use stored procedures.</span></p>
<p><span style='font-family:Arial'>Here is a sample of a simple select statement:</span></p>
<p><span><font face="Courier New, Courier, mono" size="3">    </font></span><font face="Courier New, Courier, mono" size="3">Dim
 rst As New Recordset<br>
 <br>
 <span>    </span>rst.Source = "SELECT CUSTOMER.CONTACT_FIRST, " &
 _<br>
 <span>                </span>"CUSTOMER.CONTACT_LAST, CUSTOMER.COUNTRY "
 & _<br>
 <span>                </span>"FROM CUSTOMER"<span>              </span></font></p>
<p><font face="Courier New, Courier, mono" size="3"><span>    </span>rst.ActiveConnection
 = adoConn<br>
 <span>    </span>adoConn.BeginTrans<br>
 <span>    </span>rst.Open<br>
 <span>    </span>adoConn.CommitTrans<span style='font-family:Arial'></span></font></p>
<p><span style='font-family:Arial'>And here is a simple stored procedure execution:</span></p>
<p><span><font face="Courier New, Courier, mono" size="3">    </font></span><font face="Courier New, Courier, mono" size="3">Dim
 rst As New Recordset<br>
 <span>    </span>Dim cmd As New ADODB.Command</font></p>
<p><font face="Courier New, Courier, mono" size="3"><span>    </span>adoConn.Open<span>   
 </span></font></p>
<p><font size="3" face="Courier New, Courier, mono"><span>    </span>With cmd<br>
 <span>        </span>.ActiveConnection = adoConn<br>
 <span>        </span>.CommandText = "Select * FROM DEPT_BUDGET (100)"<br>
 <span>    </span>End With</font></p>
<p><font size="3" face="Courier New, Courier, mono"><span>    </span>adoConn.BeginTrans<br>
 <span>    </span>    Set rst = cmd.Execute<br>
 <span>    </span>adoConn.CommitTrans</font></p>
<p><span style='font-family:Arial'>Notice that if your stored procedure returns
 any rows, you have to use the &#8216;SELECT * FROM <i>stored procedure name</i>&#8217; syntax.
 If your procedure does not return any records, you can use &#8216;EXECUTE <i>stored
 procedure name</i>&#8217;.</span></p>
<p><span style='font-family:Arial'>Also, the way you pass parameters in and out
 of the procedure is a bit peculiar. Lets say you have an insert stored procedure
 that will accept 3 parameters. To pass those parameters you can use inline syntax:
 For instance, &#8216;execute procedure PROC_INSERT_TBLCUSTOMERS (<i>comma delimited
 parameter values)</i>&#8217;<i> </i>or you can use this syntax:</span></p>
<p><span style='font-size:10.0pt;font-family:"Courier New"'><font face="Courier New, Courier, mono" size="3">With
 cmd<br>
 </font></span><font face="Courier New, Courier, mono" size="3"><span>       
 </span>.ActiveConnection = adoConn<br>
 <span>        </span>.CommandText = " execute procedure PROC_INSERT_TBLCUSTOMERS
 (?,?,?)&#8221;<br>
 <span style='font-size:10.0pt;font-family:"Courier New"'>End With</span></font></p>
<p><font face="Courier New, Courier, mono" size="3"><span style='font-size:10.0pt;font-family:"Courier New"'>adoConn.BeginTrans<br>
 </span><span>        </span>cmd(0) = <i>parameter value<br>
 </i><span>        </span>cmd(1) = <i>parameter value<br>
 </i><span>        </span>cmd(2) = <i>parameter value<br>
 </i><span>        </span>cmd.Execute<br>
 <span style='font-size:10.0pt;font-family:"Courier New"'>adoConn.CommitTrans</span><span style='font-family:
Arial'></span></font></p>
<p><span style='font-family:Arial'>Anyhow, these are the basics. If you guys are
 interested in Interbase, I will write a 2<sup>nd</sup> part of the tutorial
 that will cover some advanced features like working with Images, Arrays, UDF
 functions and tools for Interbase. For now take a look at the sample code for
 this tutorial, and take a look at the sample databases that are provided by
 Borland. </span></p>
<p><span style='font-family:Arial'>Raf</span></p>
</body>
</html>

