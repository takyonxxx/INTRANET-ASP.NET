<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<HTML>
<HEAD>
<title>database</title>
<script language="vb" runat="server">
Sub Page_Load(sender as Object, e as EventArgs)
	If Not Page.IsPostBack
	BindData()
	End If
End Sub
Sub BindData()
	Dim objConn As SqlConnection
	objConn = New SqlConnection("Data Source=TJFABSR5;" _
	& "Initial Catalog=envanter;User Id=purchuser;Password=purchuser;" _
	& "Connect Timeout=15;Network Library=dbmssocn;")    
	objConn.Open()
	Const strSQL as String = "SELECT * FROM tbl_test ORDER BY id asc "
	Dim objCmd as New SqlCommand(strSQL, objConn)
	Dim objDR as SqlDataReader
        objDR = objCmd.ExecuteReader()
        dgAccts.DataSource = objDR
	dgAccts.DataBind()
End Sub
Sub dgAccts_Edit(sender As Object, e As DataGridCommandEventArgs)
	dgAccts.EditItemIndex = e.Item.ItemIndex
	BindData()
End Sub
Sub dgAccts_Cancel(sender As Object, e As DataGridCommandEventArgs)
	dgAccts.EditItemIndex = -1
	BindData()
End Sub
Sub dgAccts_Update(sender As Object, e As DataGridCommandEventArgs)
	Dim id as Integer = trim(e.Item.Cells(2).Text)
        Dim strTxt1 As String = Trim(CType(e.Item.Cells(3).Controls(0), TextBox).Text)
        Dim strTxt2 As String = Trim(CType(e.Item.Cells(4).Controls(0), TextBox).Text)
        Dim strTxt3 As String = Trim(CType(e.Item.Cells(5).Controls(0), TextBox).Text)
        Dim strSQL As String = "UPDATE [tbl_test] SET [test1] =' " & strTxt1 & " ', " & _
        "[test2] = " & strTxt2 & " ,[test3] =' " & strTxt3 & " ' " & _
        "WHERE [id] = " & id & " "
	Dim objConn As SqlConnection
	objConn = New SqlConnection("Data Source=TJFABSR5;" _
	& "Initial Catalog=envanter;User Id=purchuser;Password=purchuser;" _
	& "Connect Timeout=15;Network Library=dbmssocn;")    
	objConn.Open()
	Dim myCommand as SqlCommand = new SqlCommand(strSQL, objConn)
	myCommand.ExecuteNonQuery()  
	objConn.Close()
	dgAccts.EditItemIndex = -1
	BindData()
End Sub
Sub dgAccts_Command(sender As Object, e As DataGridCommandEventArgs)
	Select (CType(e.CommandSource, LinkButton)).CommandName
	Case "Delete"
	Dim objConn As SqlConnection
	Dim objCommand As SqlCommand
	dim srid As integer = trim(e.Item.Cells(2).Text)
	Dim strSQL As String		   
	strSQL = "Delete from tbl_test where tbl_test.id = " & srid & " "
	objConn = New SqlConnection("Data Source=TJFABSR5;" _
	& "Initial Catalog=envanter;User Id=purchuser;Password=purchuser;" _
	& "Connect Timeout=15;Network Library=dbmssocn;")
	objCommand = New SqlCommand(strSQL, objConn)
	objCommand.Connection.Open()
	objCommand.ExecuteNonQuery()
	objCommand.Connection.Close()
	dgAccts.EditItemIndex = -1
	BindData()
	Case Else
	' Do nothing.
	End Select
End Sub
Sub btnSave_OnClick(Src as object, E as EventArgs)
	Dim strSQL          as String
	Dim objConn   as SqlConnection
	Dim objCommand      as SqlCommand
        'dim text as string=trim(txtTextField.Text)
        Dim strTxt1 As String = Trim(txtTextField.Text)
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tbl_test "
        strSQL = strSQL & "(test1) " & vbCrLf
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & "'" & strTxt1 & "'"
        strSQL = strSQL & ");"
	objConn = New SqlConnection("Data Source=TJFABSR5;" _
	& "Initial Catalog=envanter;User Id=purchuser;Password=purchuser;" _
	& "Connect Timeout=15;Network Library=dbmssocn;")
	objCommand = New SqlCommand(strSQL, objConn)
	objCommand.Connection.Open()
	objCommand.ExecuteNonQuery()
	objCommand.Connection.Close()
	BindData()
End Sub
</script>	
<% if request.querystring("excel")=1 then
		response.Clear()
		response.Charset = "UTF-8"
		response.ContentType = "application/vnd.ms-excel"
		Dim objDR as SqlDataReader
		Dim dg As New DataGrid()
		Dim stringWrite As New System.IO.StringWriter()
		Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)
		Dim objConn As SqlConnection
		objConn = New SqlConnection("Data Source=TJFABSR5;" _
		& "Initial Catalog=envanter;User Id=purchuser;Password=purchuser;" _
		& "Connect Timeout=15;Network Library=dbmssocn;")    
		objConn.Open()
		dim strSQL as String = "SELECT * FROM tbl_test ORDER BY id asc "
		Dim objCmd as New SqlCommand(strSQL, objConn)
		objDR = objCmd.ExecuteReader()
		dg.datasource = objDR
        dg.GridLines = GridLines.None
		dg.HeaderStyle.Font.Bold = True
		dg.DataBind()
		dg.RenderControl(htmlWrite)
		response.Write(stringWrite.ToString)
		response.End()	
	end if %>
</HEAD>
<body>
<table align="center">
<tr><td>
<form runat="server" method="post" id="Form1">
	<asp:datagrid id="dgAccts" backcolor="#ffffff" bordercolor="black" cellpadding="3" cellspacing="0"
		font-name="Verdana" font-size="8pt" headerstyle-backcolor="#aaaadd" oneditcommand="dgAccts_Edit"
		oncancelcommand="dgAccts_Cancel" onupdatecommand="dgAccts_Update" onitemcommand="dgAccts_Command"
		autogeneratecolumns="false" runat="server">
		<HeaderStyle BackColor="#AAAADD"></HeaderStyle>
		<Columns>
		<asp:EditCommandColumn ButtonType="LinkButton" UpdateText="Update" HeaderText="Edit item" CancelText="Cancel"
		EditText="Edit">
		<HeaderStyle Wrap="False"></HeaderStyle>
		<ItemStyle Wrap="False"></ItemStyle>
		</asp:EditCommandColumn>
		<asp:ButtonColumn Text="Delete" HeaderText="Delete item" CommandName="Delete"></asp:ButtonColumn>
		<asp:BoundColumn DataField="id" ReadOnly="True" HeaderText="id"></asp:BoundColumn>
		<asp:BoundColumn DataField="test1" HeaderText="test1"></asp:BoundColumn>
		<asp:BoundColumn DataField="test2" HeaderText="test2"></asp:BoundColumn>
		<asp:BoundColumn DataField="test3" HeaderText="test3"></asp:BoundColumn>
		</Columns>
	</asp:datagrid>
	<asp:textbox onkeypress="return Upper(event,this)" id="txtTextField" maxlength="50" runat="server"
	tabindex="0" />Bazý karakterleri giremezsin
	<asp:button id="btnSave" runat="server" onclick="btnSave_OnClick" text="Add" />

	<P>
	<asp:textbox onkeypress="onlyDigits(this,event)" id="Textbox1" maxlength="50" runat="server"
	tabindex="0" />Sayý girebilirsin</P>
</form>
</td>
</tr>
</table>
<P>
<a href="javascript:if(confirm('Excele atmak istediðinizden emin misiniz?')) location.href='index.aspx?excel=1';">EXCELE AKTAR</a>
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
</P>
</body>
</HTML>
