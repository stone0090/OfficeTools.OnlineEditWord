<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FileUpload.aspx.cs" Inherits="DSOframer.FileUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="script.js" type="text/javascript"></script>
    <style type="text/css">
        input[type="button"], input[type="text"], label, select {
            width: 118px;
            font-size: 12px;
        }
    </style>
</head>
<body onload="InitEvent()">
    <form id="form1" runat="server">
        <div>
            <table width="100%">
                <tr>
                    <td width="120px" style="vertical-align: top;">
                        <input type="button" id="btnOpenWebWord" value="打开服务器的Word" />
                        <input type="button" id="btnUploadWord" value="上传Word到服务器" />
                        <input type="button" id="btnCloseWord" value="关闭Word" />
                    </td>
                    <td>
                        <object classid="clsid:00460182-9E5E-11d5-B7C8-B8269041DD57" id="oframe" width="100%"
                            height="500px" codebase="ActiveX/DSOframer/DSOframer.CAB#version=1,0,0,0">
                            <param name="BorderStyle" value="1" />
                            <param name="TitlebarColor" value="52479" />
                            <param name="TitlebarTextColor" value="0" />
                            <param name="Menubar" value="1" />
                            <param name="Titlebar" value="0" />
                            <param name="Menubar" value="0" />
                        </object>
                    </td>
                </tr>
            </table>
            <div>
                <object id="WebFile" classid="clsid:B8B4E744-E5E1-4674-87C6-8914B4E3CC4B" codebase="ActiveX/WebFileHelper.cab#version=1.1.0.0" viewastext></object>
                <object id="WebFile2" classid="clsid:2D18530F-D21E-472F-99C9-96D881BD43BE" codebase="ActiveX/WebFileHelper2.cab#version=2.2.0.0" viewastext></object>
            </div>
        </div>
    </form>
    <script type="text/javascript">
        document.getElementById("btnOpenWebWord").onclick = function () {
            OpenWebWord("<%= DocUrl %>");
        }
        document.getElementById("btnUploadWord").onclick = function () {
            UploadWord("<%= Request.Url.Authority %>");
        }
        document.getElementById("btnCloseWord").onclick = function () {
            document.getElementById("oframe").close();
        }
    </script>
</body>
</html>
