<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="DocProject.WebForm1"  %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">

            <asp:Label ID="Label1" runat="server" Text="File Location"></asp:Label>
        
        <p>
            <asp:FileUpload ID="FileUpload1" runat="server" />
            </p>
            <p>
                <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" style="height: 26px; margin-left: 0px" Text="Save" Width="59px" />
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="ID" DataSourceID="SqlDataSource1" Width="116px" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
                    <Columns>
                        <asp:BoundField DataField="ID" HeaderText="ID" InsertVisible="False" ReadOnly="True" SortExpression="ID" />
                        <asp:BoundField DataField="Title" HeaderText="Title" SortExpression="Title" />
                         <asp:TemplateField ItemStyle-HorizontalAlign="Center">
            <ItemTemplate>
                 <asp:HiddenField ID="hfHtmlContent" runat="server" Value='<%# Eval("Doc") %>' />
                <asp:LinkButton ID="lnkView" runat="server" Text="View" OnClick="GridView1_SelectedIndexChanged"></asp:LinkButton>
            </ItemTemplate>
        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DocProjConnectionString %>" SelectCommand="SELECT [ID], [Title], [Doc] FROM [Doc_Table]"></asp:SqlDataSource>
            </p>
        <div id="divHtmlContent" runat="server" visible="false">
    </div>
    </form>
</body>
</html>
