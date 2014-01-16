<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<!--#include file="../purchasing/header.aspx"-->
<HTML>
<HEAD>
<title>database</title>
<script language="vb" runat="server">
    Dim objConn As OleDbConnection
    Dim strSQL As String
    Dim objCommand As OleDbCommand
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If Session("username") Is Nothing Then
                Response.Redirect("../index.aspx")
            End If
            txttaleptarih.Text = System.DateTime.Now.ToShortDateString()
            objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
            objConn.Open()
                                             
            strSQL = "SELECT * FROM tbl_counter"
            objCommand = New OleDbCommand(strSQL, objConn)
            Dim objDR As OleDbDataReader
            objDR = objCommand.ExecuteReader()
            objDR.Read()
            Dim sifgecicino As Integer = Convert.ToInt32(objDR("counter").ToString)
            Session("sifgecicino") = sifgecicino + 1
            txtsifno.Text = Session("sifgecicino").ToString
            strSQL = "UPDATE [tbl_counter] SET [counter] =' " & Session("sifgecicino") & " '"
            objCommand = New OleDbCommand(strSQL, objConn)
            'objCommand.Connection.Open()
            objCommand.ExecuteNonQuery()
            objCommand.Connection.Close()
                                    
            Dim myda As OleDbDataAdapter = New OleDbDataAdapter("Select * from tbl_supplier ", objConn)
            Dim ds As DataSet = New DataSet
            myda.Fill(ds, "AllTables")
            DropDownTercihfirma.DataSource = ds
            DropDownTercihfirma.DataSource = ds.Tables(0)
            DropDownTercihfirma.DataTextField = ds.Tables(0).Columns("name").ColumnName.ToString()
            DropDownTercihfirma.DataValueField = ds.Tables(0).Columns("code").ColumnName.ToString()
            DropDownTercihfirma.DataBind()
            objConn.Close()
        End If
    End Sub
   
    Sub btnSave_OnClick(ByVal sender As Object, ByVal e As EventArgs)
       

        'Dim clickedButton As Button = sender
        'clickedButton.Text = "Please Wait..."
        'clickedButton.Enabled = False
        
        Dim strHostName As String = System.Net.Dns.GetHostName()
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
       
        strSQL = "Delete from tbl_sifmain_gecici"
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        strSQL = "Delete from tbl_sifdetay_gecici"
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tbl_sifmain_gecici "
        strSQL = strSQL & "(createip,createdate,sifno,department,username,taleptarih,tercihfirma,talepno,statusmain,onaydurum,aciklama) " & vbCrLf
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & "'" & Trim(strHostName) & "',"
        strSQL = strSQL & "'" & Trim(System.DateTime.Now.ToShortDateString()) & "',"
        strSQL = strSQL & "'" & Trim(txtsifno.Text) & "',"
        strSQL = strSQL & "'" & Trim(Session("userbolum")) & "',"
        strSQL = strSQL & "'" & Trim(Session("username")) & "',"
        strSQL = strSQL & "'" & Month(txttaleptarih.Text) & "/" & Day(txttaleptarih.Text) & "/" & Year(txttaleptarih.Text) & "',"
        strSQL = strSQL & "'" & Trim(DropDownTercihfirma.Text) & "',"
        strSQL = strSQL & "'" & Trim(txttalepno.Text) & "',"
        strSQL = strSQL & "'0',"
        strSQL = strSQL & "'0',"
        strSQL = strSQL & "'" & Trim(txtaciklama.Text) & "'"
        strSQL = strSQL & ");"
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        objConn.Close()
        Response.Redirect("sifgirisdetay.aspx")
    End Sub
</script>	 
        <style type="text/css">
            .style1
            {
                background-color: #003366;
            }
            .style2
            {
                background-color: #FEFFFF;
            }
            .style4
            {
                color: #FFFFFF;
                background-color: #003366;
            }
            .style5
            {
                color: #FFFFFF;
                font-weight: bold;
            }
        </style>
</HEAD>
<body bgcolor="#EFEFEF">
<form runat="server" id="Form1">
<table align="center" bgcolor="#006699"" 
    style="font-family: Tahoma; font-size: x-small">
    <tr>
        <td style="font-weight: bold" class="style4">PO NUMBER (Temporary):</td>
        <td class="style4">
            <asp:textbox id="txtsifno" maxlength="50" runat="server" tabindex="0" 
                Width="76px" BackColor="#003366" Font-Bold="True" ForeColor="White" 
                Enabled="False" style="background-color: #FFFFFF" /></td>
    </tr>
    <tr>
        <td class="style5">REQUIRED DATE</td>
        <td><asp:textbox id="txttaleptarih" maxlength="50" runat="server" tabindex="0" Width="235px" /></td>
    </tr>
    <tr>
        <td class="style5">DEPARTMENT REF NO</td>
        <td><asp:textbox id="txttalepno" maxlength="50" runat="server" tabindex="0" Width="235px" /></td>
    </tr>
    <tr>
        <td class="style5">REQUIRED COMPANY</td>
        <td>
            <asp:DropDownList ID="DropDownTercihfirma" runat="server" Width="235px">
            </asp:DropDownList>
                </td>
    </tr>
    <tr>
        <td class="style5">EXPLANATION</td>
        <td><asp:textbox id="txtaciklama" maxlength="250" runat="server" tabindex="0" 
                Width="235px" Height="76px" TextMode="MultiLine" /></td>
    </tr>   
      <tr>
        <td bgcolor="#009933" style="color: #003366; background-color: #003366"></td>
        <td bgcolor="#009933" style="background-color: #003366"><span class="style2">
            <span class="style1"><asp:button id="btnSave" runat="server" 
                onclick="btnSave_OnClick" text="REGISTER" 
                Width="235px" Font-Bold="True" /></span></span></td>
       
    </tr>   
</table>

</form>
</body>
</HTML>
