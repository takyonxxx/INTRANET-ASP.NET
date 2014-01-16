<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<!--#include file="../purchasing/functions.aspx"-->
<!--#include file="../purchasing/header.aspx"-->
<script runat="server">
    
    Dim objDR As OleDbDataReader
    Dim show As Integer = 0
    Dim startdate, enddate As String
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If Session("username") Is Nothing Then
                Response.Redirect("../index.aspx")
            Else
                objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                objConn.Open()
                txt_startdate.Text = Date.Today.Day & "/" & Date.Today.Month & "/" & Date.Today.Year
                'Date.Today.Month & "\" & Date.Today.Day & "\" & Date.Today.Year
                txt_enddate.Text = Date.Today.Day & "/" & Date.Today.Month & "/" & Date.Today.Year
                Dim ds As DataSet
                Dim myda As OleDbDataAdapter = New OleDbDataAdapter("Select * from tbl_department ORDER BY code ASC ", objConn)
                ds = New DataSet
                myda.Fill(ds, "AllTables")
                DropDept.DataSource = ds
                DropDept.DataSource = ds.Tables(0)
                DropDept.DataTextField = Trim(ds.Tables(0).Columns("name").ColumnName.ToString())
                DropDept.DataValueField = Trim(ds.Tables(0).Columns("code").ColumnName.ToString())
                DropDept.DataBind()
                DropDept.Items.Insert(0, "")
                       
                myda = New OleDbDataAdapter("Select * from tbl_user  ORDER BY code ASC", objConn)
                ds = New DataSet
                myda.Fill(ds, "AllTables")
                DropUser.DataSource = ds
                DropUser.DataSource = ds.Tables(0)
                DropUser.DataTextField = Trim(ds.Tables(0).Columns("user_name").ColumnName.ToString())
                DropUser.DataValueField = Trim(ds.Tables(0).Columns("user_name").ColumnName.ToString())
                DropUser.DataBind()
                DropUser.Items.Insert(0, "")
               
                objConn.Close()
            End If
        End If

    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (txt_startdate.Text IsNot Nothing And txt_enddate.Text IsNot Nothing) Then
            startdate = Month(txt_startdate.Text) & "/" & Day(txt_startdate.Text) & "/" & Year(txt_startdate.Text)
            enddate = Month(txt_enddate.Text) & "/" & Day(txt_enddate.Text) & "/" & Year(txt_enddate.Text)
            objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
            objConn.Open()
        
            strSQL = strSQL & "SELECT * FROM tbl_sifmain "
            strSQL = strSQL & "WHERE (tbl_sifmain.createdate between #" & startdate & "# AND #" & enddate & "#) "
            strSQL = strSQL & "AND sifno LIKE '%" & Replace(txt_sifno.Text, "'", "''") & "%' "
            strSQL = strSQL & "AND department LIKE '%" & DropDept.SelectedItem.Value & "%' "
            strSQL = strSQL & "AND username LIKE '%" & DropUser.SelectedItem.Value & "%' "
            strSQL = strSQL & "ORDER BY sifno asc"
            Dim objCmd As New OleDbCommand(strSQL, objConn)
            objDR = objCmd.ExecuteReader()
            If (objDR.HasRows()) Then
                show = 1
            End If
        End If
    End Sub
</script>
<head>   
    <style type="text/css">
        .style1
        {
            font-family: Tahoma;
            font-size: x-small;
            font-weight: bold;
            color: #FFFFFF;
        }
        .style3
        {
            font-size: x-small;
        }
        .style4
        {
            font-family: Tahoma;
            font-size: x-small;
        }
    </style>
</head>
<body bgcolor="#EFEFEF">
<form id="form1" runat="server">
<%  If show = 0 Then%>
<table align="center" border="0" bgcolor="#006699" width="30%">
    <tr>
        <td colspan="2" 
            style="text-align: center; color: #FFFFFF; font-weight: 700; background-color: #003366;">
            SEARCH FOR PO</td>
    </tr>
    <tr>
        <td class="style1" >
            Start Date</td>
        <td>
            <asp:TextBox ID="txt_startdate" runat="server" 
                style="font-family: Tahoma; font-size: x-small"></asp:TextBox>
        </td>
    </tr>
    <tr>
        <td class="style1" >
            End Date</td>
        <td >
            <asp:TextBox ID="txt_enddate" runat="server" 
                style="font-family: Tahoma; font-size: x-small"></asp:TextBox>
        </td >
    </tr>
    <tr>
        <td class="style1" >
            Po Number</td>
        <td>
            <asp:TextBox ID="txt_sifno" runat="server" 
                style="font-family: Tahoma; font-size: x-small"></asp:TextBox>
        </td>
    </tr>
    <tr>
        <td class="style1" >
            Department</td>
        <td >
            <asp:DropDownList ID="DropDept" runat="server" 
                style="font-family: Tahoma; font-size: x-small">
            </asp:DropDownList>
        </td>
    </tr>
    <tr>
        <td class="style1" >
            User</td>
        <td >
            <asp:DropDownList ID="DropUser" runat="server" 
                style="font-family: Tahoma; font-size: x-small">
            </asp:DropDownList>
        </td>
    </tr>    
    <tr>
        
        <td  bgcolor="#003366" >&nbsp;</td>
        <td style="text-align: left" bgcolor="#003366">
            <asp:Button ID="Button1" runat="server" 
                style="margin-left: 0px; font-weight: 700;" Text="Bul" 
                Width="97px" onclick="Button1_Click" />
        </td>
    </tr>
</table>
<%  Else%>
<table  align="center" bgcolor="#CCCCCC" width="90%">
    <tr bgcolor="#006699">
        <td class="style1">
            Po Number
        </td>
        <td class="style1" >
            Date
        </td>
        <td class="style1" >
            Department
        </td>
        <td class="style1" >
            User
        </td>
        <td class="style1" >
            Invoice No
        </td>
        <td class="style1">
            Explanation
        </td>
         <td class="style1">
            Status
        </td>         
    </tr>
   <%  While (objDR.Read())%>
    <tr>
       <td class="style4" >
       <a href="sifgor.aspx?sifno=<%=objDR("sifno")%>"><span class="style3"><%=objDR("sifno")%></span></a><span 
               class="style3"> </span>      
        </td>
        <td class="style4" >
            <%=FormatDateTime(objDR("createdate"), 2)%>
        </td>
        <td class="style4" >
            <%=departmentad(objDR("department"))%>
        </td>
        <td class="style4" >
            <%=kisiad(objDR("username"))%>
        </td>
        <td class="style4">
            <%=objDR("faturano")%>
        </td>
         <td class="style4" >
            <%=objDR("aciklama")%>
        </td>
         <td class="style4" >
            <%=onayad(objDR("onaydurum"))%>
        </td>          
    </tr>  
   <% End While%> 
</table>
<%  End If
    objConn.Close()
    %>
</form>

</body>
