<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadIndex.aspx.cs" Inherits="Uploadfiledata.UploadIndex" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div style="margin-top:300px;margin-left:650px;">
            
            <asp:FileUpload ID="FileUpload1" runat="server" />
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="上传数据" />
        </div>
        <p style="margin-left:650px;">
            <asp:Label ID="uploadmessage" runat="server" ForeColor="Red"></asp:Label>
        </p>
        <div style="margin-left:910px;margin-top:100px">
            <a href="http://101.200.35.208:8036/DownTemp/Temp.xlsx">模板下载</a>
        </div>
    </form>
</body>
</html>
