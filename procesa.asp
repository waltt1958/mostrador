<HTML>
<HEAD>
<title> Procesa turno OEP</title>

</HEAD>

<body bgcolor="#FF00FF" onload="maximizar()">

<%

if session ("puesto")<> "" then

Const adOpenDynamic = 2
Const adLockOptimistic = 3

set conectarOEP = Server.CreateObject("ADODB.Connection")


'conectarOEP.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\inetpub\wwwroot\mostrador\turnero.mdb"

conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)}; Database Locking Mode=0; DBQ=C:\inetpub\wwwroot\mostrador\turnero.mdb"



set rsOEP = server.CreateObject ("ADODB.recordset")
'rsOEP.CursorType= adOpenDynamic
'rsOEP.LockType= adLockOptimistic


set rsPOSTAL = server.CreateObject ("ADODB.recordset")
'rsPOSTAL.CursorType= adOpenDynamic
'rsPOSTAL.LockType= AdLockOptimistic

set rsML = server.CreateObject ("ADODB.recordset")
'rsML.CursorType= adOpenDynamic
'rsML.LockType= AdLockOptimistic


sqlOEP= "select * from OEP order by OEP_fecha"
sqlPOSTAL= "select * from POSTAL order by POSTAL_fecha"
sqlML= "select * from ML order by ML_fecha"

if request.form ("turno")= "LLAMAR TURNO" then

select case request.form ("tipoTurno")
		
	case "OEP"
	
		rsOEP.open sqlOEP, conectarOEP, adOpenDynamic, adLockOptimistic

		If rsOEP.RecordCount = 0 Then

			suma = suma + 1
			
			rsOEP.AddNew
			rs("oep_nro") = suma
			rs("OEP_puesto") = session("puesto")
			rsOEP.Update

			Else
			
			If rsOEP.BOF Then
				
				rsOEP.MoveLast
	
				suma = rsOEP("oep_nro")
				suma = suma + 1

				rsOEP.AddNew
				rsOEP("oep_nro") = suma
				rsOEP("OEP_puesto") = session("puesto")
				rsOEP.Update

			Else
				
				rsOEP.MoveLast

				suma = rsOEP("oep_nro")
				suma = suma + 1
				
				rsOEP.AddNew
				rsOEP("oep_nro") = suma
				rsOEP("OEP_puesto") = session("puesto")
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
<br>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Llamó al turno OEP (ROJO) nro:  <%=suma %></boldt></font></CENTER></H1>
<br>
<br>
<CENTER><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></CENTER>
<%
	case "POSTAL"
		
		rsPOSTAL.open sqlPOSTAL,conectarOEP, adOpenDynamic, adLockOptimistic		
		If rsPOSTAL.RecordCount = 0 Then

			suma = suma + 1

			rsPOSTAL.AddNew
			rs("POSTAL_nro") = suma
			rs("POSTAL_puesto") = session("puesto")
			rsPOSTAL.Update




		Else

			If rsPOSTAL.BOF Then

				rsPOSTAL.MoveLast
	
				suma = rsPOSTAL("POSTAL_nro")
				suma = suma + 1

				rsPOSTAL.AddNew
				rsPOSTAL("POSTAL_nro") = suma
				rsPOSTAL("POSTAL_puesto") = session("puesto")
				rsPOSTAL.Update



			Else

				rsPOSTAL.MoveLast

				suma = rsPOSTAL("POSTAL_nro")
				suma = suma + 1

				rsPOSTAL.AddNew
				rsPOSTAL("POSTAL_nro") = suma
				rsPOSTAL("POSTAL_puesto") = session("puesto")
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
<br>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Llamó al turno POSTAL (verde) nro:  <%=suma %></boldt></font></CENTER></H1>
<br>
<br>
<CENTER><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></CENTER>

<%

	case "M. LIBRE"

		rsML.open sqlML,conectarOEP, adOpenDynamic, adLockOptimistic		

		If rsML.RecordCount = 0 Then

			suma = suma + 1

			rsML.AddNew
			rs("ML_nro") = suma
			rs("ML_puesto") = session("puesto")
			rsML.Update

		Else

			If rsML.BOF Then

				rsML.MoveLast
	
				suma = rsML("ML_nro")
				suma = suma + 1

				rsML.AddNew
				rsML("ML_nro") = suma
				rsML("ML_puesto") = session("puesto")
				rsML.Update

			Else

				rsML.MoveLast

				suma = rsML("ML_nro")
				suma = suma + 1

				rsML.AddNew
				rsML("ML_nro") = suma
				rsML("ML_puesto") = session("puesto")
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
<br>
<br>
<br>
<br>
<br>

<H1><CENTER><font size= "7"><boldt>Llamó al turno M. LIBRE (azul) nro:  <%=suma %></boldt></font></CENTER></H1>
<br>
<br>
<CENTER><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></CENTER>

<%

	end select

else

select case request.form ("tipoTurno")

	case "OEP"

		rsOEP.open sqlOEP,conectarOEP, adOpenDynamic, adLockOptimistic

%>
<br>
<br>
<BR>
<%

		If rsOEP.eof Then

			rsOEP.Delete
		
		else

			rsOEP.movelast
			rsOEP.delete

		end if
%>

<BR>
<BR>
<BR>
<BR>
<br>
<br>
<br>
<br>
<br>
<H1><CENTER><font size= "7"><boldt>Turno OEP eliminado</boldt></font></CENTER></H1>
<br>
<br>
<BR>
<BR>
<BR>
<CENTER><b><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></b></CENTER>

<%		rsOEP.close
		Set rsOEP= nothing

	case "POSTAL"

		rsPOSTAL.open sqlPOSTAL,conectarOEP, adOpenDynamic, adLockOptimistic

		If rsPOSTAL.eof Then

			rsPOSTAL.Delete
		
		else

			rsPOSTAL.movelast
			rsPOSTAL.delete

		end if


		rsPOSTAL.close
		Set rsPOSTAL= nothing
%>
<BR>
<BR>
<BR>
<BR>
<br>
<br>
<br>
<br>
<br>
<H1><CENTER><font size= "7"><boldt>Turno POSTAL eliminado</boldt></font></CENTER></H1>
<br>
<br>
<BR>
<BR>
<BR>
<CENTER><b><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></b></CENTER>

<%
	case "M. LIBRE"

		rsML.open sqlML,conectarOEP, adOpenDynamic, adLockOptimistic
		
		If rsML.eof Then

			rsML.Delete
		
		else

			rsML.movelast
			rsML.delete

		end if


		rsML.close
		Set rsML= nothing

%>
<BR>
<BR>
<BR>
<BR>
<br>
<br>
<br>
<br>
<br>
<H1><CENTER><font size= "7"><boldt>Turno M. LIBRE eliminado</boldt></font></CENTER></H1>
<br>
<br>
<BR>
<BR>
<BR>
<CENTER><b><INPUT type="button" value="VOLVER A TURNOS" onclick="cerrar()"></b></CENTER>

<%
end select


conectarOEP.close
set conectarOEP= nothing


end if

else

response.redirect "index.asp"

end if

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


