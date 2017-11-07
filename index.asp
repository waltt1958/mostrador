<HTML>
<HEAD>
<title> ESTABLECER NRO DE PUESTO DE ATENCION </title>

</HEAD>

<body bgcolor="#FF00FF" onload="maximizar()">
<H5 align= "right">Hoy es: <%=date%></H5>
<center><h1><p align="center"><u><b><font size="10"> COLOQUE SU NUMERO DE PUESTO DE ATENCION</font size></b></u></p> </h1></center>
<br>
<hr size= 6 color="black"></hr>

<% 
session.timeout= 480

 %>
<br>
<br>

<FORM name="numerar" action="elegir.asp" method="post">

<BR>
<BR>
<BR>
<TABLE align="center">

<TR>
<TD align="center">
<INPUT TYPE="TEXT" NAME="puesto" id="text_nombre">

</TD>
</TR>
<TR align="center">
<TD><CENTER><INPUT TYPE="button" value="CONTINUAR" onclick="validar()"></CENTER></TD><TD><INPUT TYPE="reset" value="BORRAR"></TD>
</TR>
<P><H2><u><b><CENTER>Ingrese el numero de puesto o no podra continuar</CENTER></u></b></H2>
</TABLE>

</FORM>


<SCRIPT Language="javascript" type="text/javascript">

function validar(){

if (window.document.numerar.puesto.value=="") {
alert("Tiene que colocar el nro de puesto");
window.document.numerar.puesto.focus();
return;
}

if (isNaN(window.document.numerar.puesto.value)) {

alert("Tiene que colocar solo numeros");
window.document.numerar.puesto.focus();
return;
}

window.document.numerar.submit();

}

</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>