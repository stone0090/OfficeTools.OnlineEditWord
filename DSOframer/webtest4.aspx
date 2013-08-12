<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="webtest4.aspx.cs" Inherits="OfficeOnline.Demo_DSOframer.webtest4" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="script.js" type="text/javascript"></script>
    <style type="text/css">
        input[type="button"], input[type="text"], label, select
        {
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
                    <label>
                        基本操作</label>
                    <input type="button" value="新建文档" onclick="NewDoc()" />
                    <input type="button" value="打开文档" onclick="OpenDoc()" />
                    <input type="button" value="文档另存" onclick="SaveCopyDoc()" />
                    <input type="button" value="页面设置" onclick="ChgLayout()" />
                    <input type="button" value="文档打印" onclick="PrintDoc()" />
                    <input type="button" value="文档属性" onclick="OpenProperty()" />
                    <input type="button" value="关闭文档" onclick="CloseDoc()" />
                    <label>
                        菜单操作</label>
                    <input type="button" value="显示标题栏" onclick="ToggleTitlebar()" />
                    <input type="button" value="显示菜单栏" onclick="ToggleMenubar()" />
                    <input type="button" value="显示工具栏" onclick="ToggleToolbars()" />
                </td>
                <td width="120px" style="vertical-align: top;">
                    <label>
                        Word相关</label>
                    <input type="button" value="新建Word" onclick="AddNewWord()" />
                    <input type="button" value="打开网络Word" onclick="OpenWebWord('<%= DocUrl %>')" />
                    <input type="button" value="是否保留痕迹" onclick="ToggleTrackRevisions()" />
                    <input type="button" value="是否显示痕迹" onclick="ToggleShowRevisions()" />
                    <input type="button" value="接受所有修订" onclick="AcceptAllRevisions()" />
                    <input type="button" value="拒绝所有修订" onclick="RejectAllChangesInDoc()" />
                    <input type="button" value="设置当前用户" onclick="SetUserName()" />
                    <input type="text" id="tUserName" value="stone" />
                    <input type="button" value="格式加锁" onclick="ProtectDoc(2)" />
                    <input type="button" value="全文加锁" onclick="ProtectDoc(3)" />
                    <input type="button" value="全部解锁" onclick="UnProtectDoc()" />
                    <select id="sViewMode" onchange="SetViewMode()">
                        <option value="1">普通视图</option>
                        <option value="2">大纲视图</option>
                        <option value="3">页面视图</option>
                        <option value="4">打印预览视图</option>
                        <option value="6">Web视图</option>
                        <option value="7">阅读视图</option>
                    </select>
                    <select id="sPageFit" onchange="SetPageFit()">
                        <option value="1">一页显示</option>
                        <option value="2">页宽显示</option>
                        <option value="3">内容居中</option>
                        
                    </select>
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
                    <div style="display: none">
                        <!-- dsoframe事件 开始 -->
                        <script type="text/javascript" language="jscript" for="oframe" event="OnDocumentOpened(str,obj)">
                            OnDocumentOpened(str,obj);
                        </script>
                        <script type="text/javascript" language="jscript" for="oframe" event="OnDocumentClosed()">
                            OnDocumentClosed();
                        </script>
                        <!-- dsoframe事件 结束 -->
                    </div>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
