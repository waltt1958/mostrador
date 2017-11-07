<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> INICIALIZADOR DEL TURNERO</title>

</HEAD>
<body bgcolor="#FF00FF" onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>
<center><h1><p align="center"><u><b><font size="10"> COLOCAR EL NUMERO ANTERIOR A LLAMAR</font size></b></u></p> </h1></center>
<br>
<hr size= 6 color="black"></hr>


<br>
<br>

<FORM name="inicia" action="muestraALTA.asp" method="post" TARGET="_self">

<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR>
<TD align="center" fontsize= "8"><H2><B>PAQUETES</B></H2></TD><TD align="center"><H2><B>POSTALES</B></H2></TD><TD align="center"><H2><B>M. LIBRE</B></H2></TD>
</TR>

<TR>
<TD align="center"><INPUT TYPE="TEXT" NAME="OEP"</TD>
<TD align="center"><INPUT TYPE="TEXT" NAME="POSTAL"</TD>
<TD align="center"><INPUT TYPE="TEXT" NAME="ML"</TD>
</TR>

</TABLE>

<BR>
<BR>
<BR>
<TABLE WIDTH= "100%">

<TR align="center">
<TD><INPUT TYPE="button" value="ALTA NUMEROS" onclick="validar()" ></TD>
<TD><INPUT TYPE="reset" value="LIMPIAR"></TD>
</TR>

<TR HEIGHT= "30">
<TD></TD>
</TR>

<TR ALIGN="CENTER">
<TD><b><INPUT TYPE="button" value="CERRAR" onclick="cerrar()" ></TD></b>
</TR>

</TABLE>
<BR>
<BR>
<BR>
<BR>



</FORM>


</body>
<SCRIPT Language="javascript" type="text/javascript">

function validar() {

if (window.document.inicia.OEP.value=="") {
alert("Tiene que colocar el turno OEP previo a llamar");
window.document.inicia.OEP.focus();
return;
}

if (window.document.inicia.POSTAL.value=="") {
alert("Tiene que colocar el turno POSTAL previo a llamar");
window.document.inicia.POSTAL.focus();
return;
}

if (window.document.inicia.ML.value=="") {
alert("Tiene que colocar el turno M. LIBRE previo a llamar");
window.document.inicia.ML.focus();
return;
}

if (isNaN(window.document.inicia.OEP.value)) {
alert("Tiene que colocar solo numeros en turno OEP");
window.document.inicia.OEP.focus();
return;
}

if (isNaN(window.document.inicia.POSTAL.value)) {
alert("Tiene que colocar solo numeros en turno POSTAL");
window.document.inicia.POSTAL.focus();
return;
}

if (isNaN(window.document.inicia.ML.value)) {
alert("Tiene que colocar solo numeros en turno ML");
window.document.inicia.ML.focus();
return;
}


window.document.inicia.submit();

}

</script>

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