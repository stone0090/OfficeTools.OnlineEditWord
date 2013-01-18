//全局变量
var oframe;
var isOpened = false;

//初始化oframe对象
function InitEvent() {
    oframe = document.getElementById("oframe");
    alert("请您将IE的文档模式调成IE7标准模式");
}

//dsoframe(打开)(关闭)事件
function OnDocumentOpened(str, obj) {
    isOpened = true;
}
function OnDocumentClosed() {
    isOpened = false;
}

function CheckFileOpened() {
    if (!isOpened)
        alert("You do not have a document open.");
    return isOpened;
}

//具体操作
function NewDoc() {
    oframe.showdialog(0);
}
function OpenDoc() {
    oframe.showdialog(1);
}
function OpenWebDoc() {
    //
}
function SaveCopyDoc() {
    if (CheckFileOpened())
        oframe.showdialog(3);
}
function ChgLayout() {
    if (CheckFileOpened())
        oframe.showdialog(5);
}
function PrintDoc() {
    if (CheckFileOpened())
        oframe.printout(true);
}
function CloseDoc() {
    if (CheckFileOpened())
        oframe.close();
}
function ToggleTitlebar() {
    if (CheckFileOpened())
        oframe.Titlebar = !oframe.Titlebar;
}
function ToggleToolbars() {
    if (CheckFileOpened())
        oframe.Toolbars = !oframe.Toolbars;
}
function ToggleMenubar() {
    if (CheckFileOpened()) {
        oframe.Menubar = !oframe.Menubar;
        oframe.Activate();
    }
}