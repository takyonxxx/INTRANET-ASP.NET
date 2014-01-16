<tr><td bgcolor=red colspan=4>
    <a href="main.aspx?folder=""&level=0"><font face=tahoma size=2 color=white face=navy><b>MAIN FOLDER</b></font></a>
    </td></tr>  
                        <%  
                    If Session("level") = 0 Then
                        For iLoop = 0 To iNRows - 1
                            dtRow = dtTable.Rows(iLoop)
                        %>
                            <tr>
                            <td><a href="main.aspx?folder=<%=dtRow(1)%>&level=1"><font face=tahoma size=2 color=navy><b><%=dtRow(1)%></b></font></a></td>
                            </tr>                         
                        <% 
                        Next iLoop
                    ElseIf Session("level") = 1 Then
                        %><tr><td><%=Session("folder1")%></td></tr>
                        <%
                             For iLoop = 0 To iNRows - 1
                                 dtRow = dtTable.Rows(iLoop)
                        %>                         
                            <tr>
                            <td>---><a href="main.aspx?folder=<%=dtRow(2)%>&level=2"><font face=tahoma size=2 color=navy><%=dtRow(2)%></font></a></td>
                            </tr>                         
                        <%  
                        Next iLoop
                    ElseIf Session("level") = 2 Then
                        %><tr><td><%=Session("folder1")%></td></tr><tr><td>---><%=Session("folder2")%></td></tr>
                        <%
                            For iLoop = 0 To iNRows - 1
                                dtRow = dtTable.Rows(iLoop)
                             %>
                            <tr>
                            <td>------><a href="main.aspx?folder=<%=dtRow(3)%>&level=3"><font face=tahoma size=2 color=navy><%=dtRow(3)%></font></a></td>
                            </tr>                         
                        <% 
                        Next iLoop
                    ElseIf Session("level") = 3 Then
                        %>
                            <tr><td><%=Session("folder1")%></td></tr><tr><td>---><%=Session("folder2")%></td></tr><tr><td>------><%=Session("folder3")%></td></tr>                         
                        <% 
                        End If
               
    %>
    <form id="Form1" runat="server">
        <tr bgcolor="navy">
        <td colspan="4">
        <asp:button id="btn_Upload" runat="server" text="Upload" tabindex="2" OnClick = "btnUpload_OnClick"></asp:button>&nbsp;
        <asp:button id="dokumanmodulu1" runat="server" />
        </td>
        </tr>
        </form>