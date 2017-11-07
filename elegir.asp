<HTML>
<HEAD>
<title> LLAMADOR DE TURNOS ATC </title>

</HEAD>
<body bgcolor="#FF00FF" onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>
<center><h1><p align="center"><u><b><font size="10"> LLAMADOR DE TURNOS ATC</font size></b></u></p> </h1></center>
<br>
<hr size= 6 color="black"></hr>

<%

session("puesto")= replace((request.form("puesto"))," ","")


%>

<br>
<br>

<FORM action="procesa.asp" method="post" TARGET="_blank">

<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR>
<TD align="center" fontsize= "8"><H2><B>PAQUETES</B></H2></TD><TD align="center"><H2><B>POSTALES</B></H2></TD><TD align="center"><H2><B>M. LIBRE</B></H2></TD>
</TR>

<TR>
<TD align="center"><INPUT TYPE="radio" NAME="tipoTurno" VALUE="OEP" checked></TD><TD align="center"><INPUT TYPE="radio" NAME="tipoTurno" VALUE="POSTAL"></TD><TD align="center"><INPUT TYPE="radio" NAME="TipoTurno" VALUE="M. LIBRE"></TD>
</TR>

</TABLE>

<BR>
<BR>
<BR>
<TABLE WIDTH= "100%">

<TR align="center">
<TD><INPUT TYPE="submit" name=turno value="LLAMAR TURNO" ></TD><TD><INPUT TYPE="submit" name=turno value="RETROCEDER TURNO"></TD>
</TR>

</TABLE>
<BR>
<BR>
<BR>
<BR>

<CENTER><INPUT type="button" value="CERRAR" onclick="cerrar()"></CENTER>

</FORM>


</body>

<SCRIPT Language="javascript">

function cerrar()   {

window.opener = self;
window.close();
}

</SCRIPT>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>
</HTML>