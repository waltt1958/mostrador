<HTML>
<HEAD>
<title> Alta inicial de turnos</title>

</HEAD>

<body bgcolor="#FF00FF" onload="maximizar()">

<%


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

		If rsOEP.RecordCount = 0 Then

			rsOEP.AddNew
			rs("oep_nro") = request.form("OEP")
			rs("OEP_puesto") = 0
			rsOEP.Update


		Else

			If rsOEP.BOF Then

				rsOEP.MoveLast

				rsOEP.AddNew
				rsOEP("oep_nro") = request.form("OEP")
				rsOEP("OEP_puesto") = 0
				rsOEP.Update



			Else

				rsOEP.MoveLast


				rsOEP.AddNew
				rsOEP("oep_nro") = request.form("OEP")
				rsOEP("OEP_puesto") = 0
				rsOEP.Update


			End If
		End If


		rsOEP.close
		SEt rsOEP= nothing



%>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Cargo el turno OEP (ROJO) nro:  <%=request.form("OEP")%></boldt></font></CENTER></H1>

<%

		rsPOSTAL.open sqlPOSTAL,conectarOEP

		If rsPOSTAL.RecordCount = 0 Then

			rsPOSTAL.AddNew
			rs("POSTAL_nro") = request.form("POSTAL")
			rs("POSTAL_puesto") = 0
			rsPOSTAL.Update


		Else

			If rsPOSTAL.BOF Then

				rsPOSTAL.MoveLast

				rsPOSTAL.AddNew
				rsPOSTAL("POSTAL_nro") = request.form("POSTAL")
				rsPOSTAL("POSTAL_puesto") = 0
				rsPOSTAL.Update


			Else

				rsPOSTAL.MoveLast

				rsPOSTAL.AddNew
				rsPOSTAL("POSTAL_nro") = request.form("POSTAL")
				rsPOSTAL("POSTAL_puesto") = 0
				rsPOSTAL.Update


			End If
		End If


		rsPOSTAL.close
		Set rsPOSTAL= nothing



%>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Cargo el turno POSTAL (verde) nro:  <%=request.form("POSTAL") %></boldt></font></CENTER></H1>

<%


		rsML.open sqlML,conectarOEP

		If rsML.RecordCount = 0 Then

			rsML.AddNew
			rs("ML_nro") = request.form("ML")
			rs("ML_puesto") = 0
			rsML.Update

		Else

			If rsML.BOF Then

				rsML.MoveLast

				rsML.AddNew
				rsML("ML_nro") = request.form("ML")
				rsML("ML_puesto") = 0
				rsML.Update

			Else

				rsML.MoveLast

				rsML.AddNew
				rsML("ML_nro") = request.form("ML")
				rsML("ML_puesto") = 0
				rsML.Update


			End If
		End If


		rsML.close
		Set rsML= nothing

%>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Cargo el turno M. LIBRE (amarillo) nro:  <%=request.form("ML") %></boldt></font></CENTER></H1>

<CENTER><INPUT type="button" value="CERRAR VENTANA" onclick="cerrar()"></CENTER>

<%


conectarOEP.close
set conectarOEP= nothing


%>


<SCRIPT Language="javascript" type="text/javascript">

function cerrar()  {
ventana= window.self;
ventana.opener= window.self;
ventana.close();
}

</SCRIPT>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>

</body>
</HTML>


