<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="ExportExcelTest.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <input type="button" onclick="Download(1)" value="下载全院综合" />
            <input type="button" onclick="Download(2)" value="下载各个科室" />
        </div>
    </form>
</body>
<script>
    function Download(Num) {
        window.location.href = "Server.ashx?Type=download&Num=" + Num;
    }
</script>
</html>