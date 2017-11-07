<HTML>
<HEAD>
<meta http-equiv="Page-Exit" content="progid:DXImageTransform.Microsoft.Fade(duration=.5)"/>
<meta http-equiv="Page-Enter" content="Alpha(opacity=100)"/>

<title> LLAMADOR DE TURNOS ATC </title>

</HEAD>


<body>

<%


Response.Expires = 0
Response.Buffer = True
Response.AddHeader "Refresh", "4"

Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)}; Database Locking Mode=0; DBQ=C:\Inetpub\wwwroot\mostrador\turnero.mdb"

set rsOEP = server.CreateObject ("ADODB.recordset")
rsOEP.CursorType= AdOpenDynamic
rsOEP.LockType= adLockOptimistic

set rsPOSTAL = server.CreateObject ("ADODB.recordset")
rsPOSTAL.CursorType= adOpenDynamic
rsPOSTAL.LockType= AdLockOptimistic

set rsML = server.CreateObject ("ADODB.recordset")
rsML.CursorType= adOpenDynamic
rsML.LockType= AdLockOptimistic

sqlOEP= "select * from OEP order by OEP_fecha"
sqlPOSTAL= "select * from POSTAL order by POSTAL_fecha"
sqlML= "select * from ML order by ML_fecha"

rsOEP.open sqlOEP,conectarOEP
rsPOSTAL.open sqlPOSTAL, conectarOEP
rsML.open sqlML, conectarOEP


	if not rsOEP.EOF or not rsPOSTAL.EOF or not rsML.EOF then
		rsOEP.movelast
		rsPOSTAL.movelast
		rsML.movelast
		application("actualOEP")=rsOEP("OEP_nro")	
		application("actualPOSTAL")=rsPOSTAL("POSTAL_nro")
		application("actualML")=rsML("ML_nro")
		if application("actualOEP")=application("anteriorOEP") and application("actualPOSTAL")=application("anteriorPOSTAL") and application("actualML")=application("anteriorML") and application("contador") < 4 then
			select case (application("contador"))
				case 1 
				%>
				<img src="prioridad1.jpg" height="100%" width="100%">
				<%				
				case 2
				%>
				<img src="documentacion1.jpg" height="100%" width="100%">
				<%
				case 3	
				%>
				<img src="numero1.jpg" height="100%" width="100%">
				<%
			end select
			application("contador")= application("contador") + 1
					
		else

				%>

<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR>
<TH COLSPAN="2" align="center" bgcolor="red"><H1><FONT COLOR="white"><B>RETIRO PAQUETES</B></FONT></H1></TH><TH COLSPAN="2" align="center" bgcolor="green"><H1><B><FONT COLOR="white">RETIRO SOBRES</FONT><br><FONT COLOR="white">(Entradas y T.Cred.)</FONT></B></H1></TH><TH COLSPAN="2"align="center" bgcolor="blue"><H1><B><FONT COLOR="white">RECEPCION MERCADO<br><FONT COLOR="white">LIBRE</FONT></FONT></B></H1></TH>
</TR>

<TR HEIGHT="50%" FONT SIZE= "8">

<% if application("actualOEP")<> application("anteriorOEP") then

%>

<TD VALIGN="MIDDLE" bgcolor="red" ><B><H2><CENTER><p style="font-size:40px;">TURNO</P><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsOEP("OEP_nro"))%></P></H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="red"><B><H2> <CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsOEP("OEP_puesto"))%> </p></H2></B></TD>

<% 
else
%>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2><CENTER><p style="font-size:40px;">TURNO</P><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsOEP("OEP_nro"))%></P></H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2> <CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsOEP("OEP_puesto"))%> </p></H2></B></TD>

<%
end if

if application("actualPOSTAL")<> application("anteriorPOSTAL") then

%>

<TD VALIGN="MIDDLE" bgcolor="green"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsPOSTAL("POSTAL_nro"))%> </p></H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="green"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsPOSTAL("POSTAL_puesto"))%></p></H2></B></TD>

<% 
else
%>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsPOSTAL("POSTAL_nro"))%> </p></H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p style="font-size:100px;"><%response.write (rsPOSTAL("POSTAL_puesto"))%></p></H2></B></TD>

<%
end if

if application("actualML")<> application("anteriorML") then

%>

<TD VALIGN="MIDDLE" bgcolor="blue"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:100px;"><% response.write (rsML("ML_nro"))%></p> </H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="blue"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><CENTER><p style="font-size:100px;"><%response.write (rsML("ML_puesto"))%></p> </H2></B></TD>

<% 
else
%>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:100px;"><% response.write (rsML("ML_nro"))%></p> </H2></B></TD>

<TD VALIGN="MIDDLE" bgcolor="white"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><CENTER><p style="font-size:100px;"><%response.write (rsML("ML_puesto"))%></p> </H2></B></TD>

<%
end if

%>
</TR>

<TR>
<TH COLSPAN="6" BGCOLOR="SILVER" valign="middle"><p style="font-size:40px;"><B>LLAMADO ANTERIOR</B></p></TH>
</TR>

<% 

if (application("actualOEP")<> application("anteriorOEP")) or (application("actualPOSTAL")<> application("anteriorPOSTAL")) or (application("actualML")<> application("anteriorML")) then

%>
<bgsound src="oep.mp3" loop="2" volume="0" hidden="true" autostart="false" >
<%
end if


application("anteriorOEP")= rsOEP("OEP_nro")
application("anteriorPOSTAL")= rsPOSTAL("POSTAL_nro")
application("anteriorML")= rsML("ML_nro")

rsOEP.moveprevious
rsPOSTAL.moveprevious
rsML.moveprevious

%>

<TR HEIGHT="28%" BGCOLOR="SILVER"8">
<TD VALIGN="MIDDLE"><B><H2><CENTER><p style="font-size:40px;">TURNO</P><HR SIZE="1"><p style="font-size:59px;"><%response.write (rsOEP("OEP_nro"))%></P></H2></B></TD>
<TD VALIGN="MIDDLE"><B><H2> <CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p style="font-size:59px;"><%response.write(rsOEP("OEP_puesto"))%> </p></H2></B></TD>
<TD VALIGN="MIDDLE"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:59px;"><%response.write(rsPOSTAL("POSTAL_nro"))%> </p></H2></B></TD>
<TD VALIGN="MIDDLE"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><p 
style="font-size:59px;"> <%response.write(rsPOSTAL("POSTAL_puesto"))%></p></H2></B></TD>
<TD VALIGN="MIDDLE"><B><H2><CENTER><p style="font-size:40px;">TURNO</p><HR SIZE="1"><p style="font-size:59px;"><%response.write(rsML("ML_nro"))%></p> </H2></B></TD>
<TD VALIGN="MIDDLE"><B><H2><CENTER><p style="font-size:40px;">PUESTO</p><HR SIZE="1"><CENTER><p style="font-size:59px;"> <%response.write(rsML("ML_puesto"))%></p></H2></B></TD>
</TR>

</TABLE>

<%
if application("contador") = 4 then
			application("contador")=1
end if
		end if


	end if

rsOEP.close
SET rsOEP= nothing

rsPOSTAL.close
SET rsPOSTAL= nothing

rsML.close
SET rsML= nothing

conectarOEP.close
set conectarOEP= nothing

%>
<H3 align= "right">Hoy es: <%=date%></H3>
</body>



</HTML>