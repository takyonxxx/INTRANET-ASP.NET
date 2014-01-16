<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<html>
<BODY LEFTMARGIN="5" TOPMARGIN="5" bgcolor="#EFEFEF">
<%  '////////////////////////////////////////////////////////////////
    'File uploading by Türkay Biliyor, turkaybiliyor@hotmail.com
    '///////////////////////////////////////////////////////////////
    Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Files.mdb") & ";")
    Dim objCommand As OleDbCommand
    Dim myReader As OleDbDataReader
    Dim strSQLQuery As String
    Dim Cmd As OleDbDataAdapter
    Dim dtSet As DataSet
    Dim dtTable As DataTable
    Dim dtRow As DataRow
    Dim iLoop As Integer
    Dim iNRows As Integer
    Session("admin") = 1
     If Request.QueryString("mode") = "view" Then
        objConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Files.mdb") & ";")
        strSQLQuery = "SELECT * FROM Files WHERE ID =" & Request.QueryString("file_id") & " "
        objCommand = New OleDbCommand(strSQLQuery, objConnection)
        objConnection.Open()
        myReader = objCommand.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        If myReader.Read() Then
            Dim fileData() As Byte = CType(myReader.Item("FileData"), Byte())
            Response.Clear()
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + myReader.GetString(1))
            Response.ContentType = myReader.GetString(3)
            Response.OutputStream.Write(fileData, 0, fileData.Length)
        Else
            Response.Write("<p>No File to view.</p>")
        End If
        objConnection.Close()
    End If
%>  <script language="VB" runat="server">
        Sub btnCreat_OnClick(ByVal Src As Object, ByVal E As EventArgs)
            If txt_folder.Text <> "" Then
                Dim tablename As String
                Dim fieldname1 As String
                Dim fieldname2 As String
                Dim fieldname3 As String
                Dim CmdText As String
                Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Files.mdb") & ";")
                If Session("level") = 0 Then
                    tablename = "tbl_folder1"
                    fieldname1 = "Folder1_Name"
                    CmdText = "INSERT INTO " & tablename & " ( " & fieldname1 & ") VALUES (@FolderName1)"
                    Dim cmd As OleDbCommand = New OleDbCommand(CmdText, objConnection)
                    Try
                        Dim pms As OleDbParameterCollection = cmd.Parameters
                        pms.Add("@FolderName1", OleDbType.VarChar, 50)
                        pms("@FolderName1").Value = txt_folder.Text
                        pms = Nothing
                        objConnection.Open()
                        cmd.ExecuteNonQuery()
                    Finally
                        CType(cmd, IDisposable).Dispose()
                    End Try
                ElseIf Session("level") = 1 Then
                    tablename = "tbl_folder2"
                    fieldname1 = "Folder1_Name"
                    fieldname2 = "Folder2_Name"
                    CmdText = "INSERT INTO " & tablename & " ( " & fieldname1 & "," & fieldname2 & ") VALUES (@FolderName1,@FolderName2)"
                    Dim cmd As OleDbCommand = New OleDbCommand(CmdText, objConnection)
                    Try
                        Dim pms As OleDbParameterCollection = cmd.Parameters
                        pms.Add("@FolderName1", OleDbType.VarChar, 50)
                        pms.Add("@FolderName2", OleDbType.VarChar, 50)
                        pms("@FolderName1").Value = Session("folder1")
                        pms("@FolderName2").Value = txt_folder.Text
                        pms = Nothing
                        objConnection.Open()
                        cmd.ExecuteNonQuery()
                    Finally
                        CType(cmd, IDisposable).Dispose()
                    End Try
                ElseIf Session("level") = 2 Then
                    tablename = "tbl_folder3"
                    fieldname1 = "Folder1_Name"
                    fieldname2 = "Folder2_Name"
                    fieldname3 = "Folder3_Name"
                    CmdText = "INSERT INTO " & tablename & " ( " & fieldname1 & "," & fieldname2 & "," & fieldname3 & ") VALUES (@FolderName1,@FolderName2,@FolderName3)"
                    Dim cmd As OleDbCommand = New OleDbCommand(CmdText, objConnection)
                    Try
                        Dim pms As OleDbParameterCollection = cmd.Parameters
                        pms.Add("@FolderName1", OleDbType.VarChar, 50)
                        pms.Add("@FolderName2", OleDbType.VarChar, 50)
                        pms.Add("@FolderName3", OleDbType.VarChar, 50)
                        pms("@FolderName1").Value = Session("folder1")
                        pms("@FolderName2").Value = Session("folder2")
                        pms("@FolderName3").Value = txt_folder.Text
                        pms = Nothing
                        objConnection.Open()
                        cmd.ExecuteNonQuery()
                    Finally
                        CType(cmd, IDisposable).Dispose()
                    End Try
                End If
                Response.AddHeader("Refresh", "1")
            Else
            End If
        End Sub
       Sub btnUpload_OnClick(ByVal Src As Object, ByVal E As EventArgs)
            If dokumanmodulu1.HasFile Then
                Dim files As HttpFileCollection = Request.Files
                Dim i As Integer = 0
                While i < Request.Files.Count
                    Dim file As HttpPostedFile = files(i)
                    If file.ContentLength > 0 Then
                        UploadFile(file)
                    End If
                    System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
                End While
                files = Nothing
            Else
                Response.Write("<p>Chose a file.</p>")
            End If
            Response.AddHeader("Refresh", "1")
         End Sub
         Sub UploadFile(ByVal file As HttpPostedFile)
             Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Files.mdb") & ";")
            Dim CmdText As String = "INSERT INTO Files(FileName, FileSize, ContentType, FileData,Folder_Level,Folder_Name1,Folder_Name2,Folder_Name3)" & _
            "VALUES (@FileName, @FileSize, @ContentType, @FileData,@Folder_Level,@Folder_Name1,@Folder_Name2,@Folder_Name3)"
             Dim fileName As String = Nothing
             Dim contentType As String = file.ContentType
             Dim fileLength As Integer = file.ContentLength
             Dim fileData(fileLength) As Byte
             Dim lastPos As Integer = file.FileName.LastIndexOf("\"c)
             fileName = file.FileName.Substring(System.Threading.Interlocked.Increment(lastPos))
             file.InputStream.Read(fileData, 0, fileLength)
             Dim cmd As OleDbCommand = New OleDbCommand(CmdText, objConnection)
             Try
                 Dim pms As OleDbParameterCollection = cmd.Parameters
                pms.Add("@FileName", OleDbType.VarChar, 50)
                pms.Add("@FileSize", OleDbType.Integer)
                pms.Add("@ContentType", OleDbType.VarChar, 50)
                pms.Add("@FileData", OleDbType.VarBinary)
                pms.Add("@Folder_Level", OleDbType.VarChar, 50)
                pms.Add("@Folder_Name1", OleDbType.VarChar, 50)
                pms.Add("@Folder_Name2", OleDbType.VarChar, 50)
                pms.Add("@Folder_Name3", OleDbType.VarChar, 50)
                pms("@FileName").Value = fileName
                pms("@FileSize").Value = fileLength
                pms("@ContentType").Value = contentType
                pms("@FileData").Value = fileData
                pms("@Folder_Level").Value = Session("level")
                pms("@Folder_Name1").Value = Session("folder1")
                pms("@Folder_Name2").Value = Session("folder2")
                pms("@Folder_Name3").Value = Session("folder3")
                 pms = Nothing
                 objConnection.Open()
                 cmd.ExecuteNonQuery()
             Finally
                 CType(cmd, IDisposable).Dispose()
             End Try
         End Sub
   </script> 
<%
    If Request.QueryString("delfolder") <> "" Then
        If Request.QueryString("level") = 0 Then
            strSQLQuery = "Delete from tbl_folder1 where id = " & Request.QueryString("delfolder") & " "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
            strSQLQuery = "Delete from files where Folder_Name1 = '" & Request.QueryString("delfoldern") & "' "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
        ElseIf Request.QueryString("level") = 1 Then
            strSQLQuery = "Delete from tbl_folder2 where id = " & Request.QueryString("delfolder") & " "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
            strSQLQuery = "Delete from files where Folder_Name2 = '" & Request.QueryString("delfoldern") & "' "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
        ElseIf Request.QueryString("level") = 2 Then
            strSQLQuery = "Delete from tbl_folder3 where id = " & Request.QueryString("delfolder") & " "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
            strSQLQuery = "Delete from files where Folder_Name3 = '" & Request.QueryString("delfoldern") & "' "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
        End If
        
    End If
    
        If Request.QueryString("delete") <> "" Then
            strSQLQuery = "Delete from Files where ID = " & Request.QueryString("delete") & " "
            objCommand = New OleDbCommand(strSQLQuery, objConnection)
            objConnection.Open()
            objCommand.ExecuteNonQuery()
            objConnection.Close()
        End If
    
        Session("level") = 0
        If Request.QueryString("folder") <> "" Then
            Session("level") = Request.QueryString("level")
        End If
        If Session("level") = 0 Then
            strSQLQuery = "SELECT * FROM tbl_folder1"
            Session("folder1") = ""
            Session("folder2") = ""
            Session("folder3") = ""
        ElseIf Session("level") = 1 Then
            Session("folder1") = Request.QueryString("folder")
            Session("folder2") = ""
            Session("folder3") = ""
            strSQLQuery = "SELECT * FROM tbl_folder2 where Folder1_Name='" & Session("folder1") & "' "
        ElseIf Session("level") = 2 Then
            Session("folder2") = Request.QueryString("folder")
            Session("folder3") = ""
            strSQLQuery = "SELECT * FROM tbl_folder3 where Folder2_Name='" & Session("folder2") & "' "
        ElseIf Session("level") = 3 Then
            Session("folder3") = Request.QueryString("folder")
            strSQLQuery = "SELECT * FROM tbl_folder3 where Folder3_Name='" & Session("folder3") & "' "
        End If
   
        objConnection.Open()
        Cmd = New OleDbDataAdapter(strSQLQuery, objConnection)
        dtSet = New DataSet
        Cmd.Fill(dtSet)
        dtTable = New DataTable
        dtTable = dtSet.Tables(0)
        iNRows = dtTable.Rows.Count
   
%>
<table align=center  width="80%" style="background-color: gainsboro;">  
 <tr style="background-color: #336666;">
        <td colspan=5> <font face=tahoma size=3 color=white face=navy><HR /><B>DOCUMENT MANAGEMENT SYSTEM </B></font></td>
       </tr>  
        <tr style="background-color: #336666;">
        <td colspan=5>
        <a href="../index.aspx">
        <font face=tahoma size=2 color="#CCCCCC" face=navy><B>MAIN PAGE</B></font></a>--
         <a href="main.aspx?folder=&level=0">
        <font face=tahoma size=2 color="#CCCCCC" face=navy><B>ROOT</B></font></a><hr /></td>
       </tr>  
                        <%  
                    If Session("level") = 0 Then
                        For iLoop = 0 To iNRows - 1
                            dtRow = dtTable.Rows(iLoop)
                        %>
                            <tr style="background-color: cadetblue;">
                            <td><img src="../document/images/dir_dir_red.gif" height="20" width="20" />&nbsp;
                            <a href="main.aspx?folder=<%=dtRow(1)%>&level=1"><font face=tahoma size=2 color=white><b><%=dtRow(1)%></b></font></a>
                            &nbsp;
                            <%If Session("admin") = 1 Then%>
                            <a href="main.aspx?delfolder=<%=dtRow(0)%>&level=0&delfoldern=<%=dtRow(1)%>"><font face=tahoma size=2 color=yellow>Delete</font></a>
                            <%End If%></td>
                            </tr>                         
                        <% 
                        Next iLoop
                    ElseIf Session("level") = 1 Then
                        %> <tr style="background-color: cadetblue;"><td>
                        <img src="../document/images/dir_dir_red.gif" height="20" width="20" />&nbsp;
                        <font face=tahoma size=2><b><%=Session("folder1")%></b></font></td></tr>
                        <%
                             For iLoop = 0 To iNRows - 1
                                 dtRow = dtTable.Rows(iLoop)
                        %>                         
                            <tr style="background-color: cadetblue;">
                            <td>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <img src="../document/images/dir_dir_yellow.gif" height="20" width="20" />&nbsp;<a href="main.aspx?folder=<%=dtRow(2)%>&level=2"><font face=tahoma size=2 color=white><%=dtRow(2)%></font></a>
                            &nbsp;
                             <%If Session("admin") = 1 Then%>
                            <a href="main.aspx?delfolder=<%=dtRow(0)%>&level=1&folder=<%=Session("folder1")%>&delfoldern=<%=dtRow(2)%>"><font face=tahoma size=2 color=yellow>Delete</font></a>
                            <%End If%>
                            </td>
                            </tr>                         
                        <%  
                        Next iLoop
                    ElseIf Session("level") = 2 Then
                        %>  <tr style="background-color: cadetblue;"><td>
                            <img src="../document/images/dir_dir_yellow.gif" height="20" width="20" />&nbsp;<a href="main.aspx?folder=<%=Session("folder1")%>&level=1"><font face=tahoma size=2 color=white><%=Session("folder1")%></font></a></td></tr>
                            <tr style="background-color: cadetblue;"><td>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <img src="../document/images/dir_dir_red.gif" height="20" width="20" />&nbsp;<font face=tahoma size=2><b><%=Session("folder2")%></b></font></td></tr>
                        <%
                            For iLoop = 0 To iNRows - 1
                                dtRow = dtTable.Rows(iLoop)
                             %>
                            <tr style="background-color: cadetblue;">
                            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <img src="../document/images/dir_dir_yellow.gif" height="20" width="20" />&nbsp;<a href="main.aspx?folder=<%=dtRow(3)%>&level=3"><font face=tahoma size=2 color=white><%=dtRow(3)%></font></a>
                            &nbsp;
                             <%If Session("admin") = 1 Then%>
                            <a href="main.aspx?delfolder=<%=dtRow(0)%>&level=2&folder=<%=Session("folder2")%>&delfoldern=<%=dtRow(3)%>"><font face=tahoma size=2 color=yellow>Delete</font></a>
                            <%End If%>
                            </td>
                            </tr>                         
                        <% 
                        Next iLoop
                    ElseIf Session("level") = 3 Then
                        %>
                            <tr style="background-color: cadetblue;"><td>
                            <img src="../document/images/dir_dir_yellow.gif" height="20" width="20" />&nbsp;<a href="main.aspx?folder=<%=Session("folder1")%>&level=1"><font face=tahoma size=2 color=white><%=Session("folder1")%></font></a></td></tr>
                            <tr style="background-color: cadetblue;"><td>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <img src="../document/images/dir_dir_yellow.gif" height="20" width="20" />&nbsp;<a href="main.aspx?folder=<%=Session("folder2")%>&level=2"><font face=tahoma size=2 color=white><%=Session("folder2")%></font></a></td></tr>
                            <tr style="background-color: cadetblue;"><td>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <img src="../document/images/dir_dir_red.gif" height="20" width="20" />&nbsp;<font face=tahoma size=2><b><%=Session("folder3")%></b></font></td></tr>                         
                        <% 
                        End If
                        
    If Session("level") <> 0 Then
                            If Session("level") = 1 Then
                                Session("folder") = Session("folder1")
                                strSQLQuery = "SELECT * FROM files where Folder_Name1='" & Session("folder1") & "' AND Folder_Level='" & Session("level") & "' "
                            ElseIf Session("level") = 2 Then
                                Session("folder") = Session("folder2")
                                strSQLQuery = "SELECT * FROM files where Folder_Name2='" & Session("folder2") & "'"
                            ElseIf Session("level") = 3 Then
                                Session("folder") = Session("folder3")
                                strSQLQuery = "SELECT * FROM files where Folder_Name3='" & Session("folder3") & "'"
                            End If
        Cmd = New OleDbDataAdapter(strSQLQuery, objConnection)
        dtSet = New DataSet
        Cmd.Fill(dtSet)
        dtTable = New DataTable
        dtTable = dtSet.Tables(0)
        iNRows = dtTable.Rows.Count
        iLoop = 0
                   %> 
                  <%If iNRows <> 0 Then%>
                   <tr>
                   <td style="background-color: #336666;" colspan="5"><font color=white face=tahoma size=2><b>FILES</b></font></td>                                     
                   </tr> 
                   <tr>             
                   <td>
                       
                                 <tr style="background-color: #666666;">
                                    <td><font color=white face=tahoma size=2>FileName</font></td>
                                    <td><font color=white face=tahoma size=2>FileSize</font></td>
                                    <td><font color=white face=tahoma size=2>ContentType</font></td>
                                    <td><font color=white face=tahoma size=2>Folder_Level</font></td> 
                                    <td><font color=white face=tahoma size=2>Proses</font></td>     
                                  </tr>
                                 <%  For iLoop = 0 To iNRows - 1
                                         dtRow = dtTable.Rows(iLoop)%>
                                  <tr style="background-color: whitesmoke;" >
                                        <td><a href=main.aspx?file_id=<%=dtRow(0)%>&mode=view&level=<%=Session("level")%>&folder=<%=Session("folder")%>><font  color=MidnightBlue face=tahoma size=1><%=dtRow(1)%></font></a> </td>
                                        <td><font face=tahoma size=1><%=FormatNumber(dtRow(2) / 1000, 2)%>&nbsp;Kb</font></td>
                                        <td><font face=tahoma size=1><%=dtRow(3)%></font></td>
                                        <td><font face=tahoma size=1><%=dtRow(5)%></font></td>   
                                        <td>
                                         <%If Session("admin") = 1 Then%>
                                        <a href=main.aspx?delete=<%=dtRow(0)%>&level=<%=Session("level")%>&folder=<%=Session("folder")%>><font color=MidnightBlue face=tahoma size=1>Delete</font></a> </td>         
                                        <%End If%>
                                        </tr> 
                                  <%  Next iLoop
                                  End If
                                  dtTable.Dispose()
                          %>
                        </td>
                        </tr> 
    <%End If%>    
     <%If Session("admin") = 1 Then%>            
        <form id="Form1" runat="server">
        <tr style="background-color: #666666;">
        <td colspan="5">
        <asp:button id="btn_Creat" runat="server" text="Creat_Folder" tabindex="2" OnClick = "btnCreat_OnClick" Font-Size="X-Small"></asp:button>&nbsp;
            <asp:TextBox ID="txt_folder" runat="server"></asp:TextBox>&nbsp;
            <%If Session("level") <> 0 Then%>
            <asp:fileupload id="dokumanmodulu1" runat="server" />
            FileUpload&nbsp;<asp:button id="btn_Upload" runat="server" text="Upload_File" tabindex="2" OnClick = "btnUpload_OnClick" Font-Size="X-Small"></asp:button>
            <%End If%>
        </td>
        </tr>
        </form>
        <%End If%>
</table> 
</body>
</html>
<%
    dtTable.Dispose()
    objConnection.Close()
%>