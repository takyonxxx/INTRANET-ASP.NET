<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<script language="vb" runat="server">    
    Dim objConn As OleDbConnection
    Dim strSQL As String
    Dim name As String
    Public Function firmaad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_supplier where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
    End Function
    Public Function departmentad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_department where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
    End Function
    Public Function kisiad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_user where user_name='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name") & " " & objDR("surname")
        objConn.Close()
        Return name
    End Function
    Public Function csad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_costcenter where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
    End Function
    Public Function mchad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_machine where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
    End Function
    Public Function stokad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_stok where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
       
    End Function
    Public Function onayad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_onaydurum where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
       
    End Function
    Public Function birimad(ByVal ID As String) As String
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT * FROM tbl_birim where code='" & ID & "'"
        Dim objCmd As New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        name = objDR("name")
        objConn.Close()
        Return name
    End Function
   
</script>	 
<script language="javascript">
	//////////////////////////      
	function Upper(e,r)
	{
	if ((e.keyCode > 32 && e.keyCode < 48) || (e.keyCode > 57 && e.keyCode < 65) || (e.keyCode > 90 && e.keyCode < 97))
	e.returnValue = false;
	else
	r.value = r.value.toUpperCase();
	}
	///////////////////////////

	function onlyDigits(el,e) {
	var isIE = document.all?true:false;
	var isNS = document.layers?true:false;
	var IS_PERIOD=46;
	var PERIOD_TYPED=false;
	var _regExp=/\./;
	PERIOD_TYPED=el.value.match(_regExp);
	var _ret = true;
	if (isIE) {
	if (window.event.keyCode == IS_PERIOD) {
	if (!PERIOD_TYPED) {
	PERIOD_TYPED=true;
	} else {
	window.event.keyCode=0;
	_ret = false;
	}
	}
	if (window.event.keyCode < 46 || window.event.keyCode > 57) {
	window.event.keyCode = 0;
	_ret = false;
	}
	}
	if (isNS) {
	if (e.which == IS_PERIOD) {
	if (!PERIOD_TYPED) {
	PERIOD_TYPED=true;
	} else {
	e.which=0;
	_ret = false;
	}
	}
	if (e.which < 46 || e.which > 57) {
	e.which = 0;
	_ret = false;
	}
	}
	return (_ret); 
	}


	///////////////////////////
	function setfocus(obj){      
	var control = document.getElementById(obj);
	if( control != null ){control.focus();}
	}
	setfocus("txtTextField")
	////////////////////////////
</script>