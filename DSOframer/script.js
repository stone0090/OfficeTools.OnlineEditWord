//全局变量
var oframe;
var isOpened = false;


//初始化oframe对象
function InitEvent() {
    oframe = document.getElementById("oframe");
    //alert("请您将IE的文档模式调成IE7标准模式");
}


//检查是否打开文档
function CheckFileOpened() {
    if (!isOpened)
        alert("You do not have a document open.");
    return isOpened;
}


//dsoframe(打开)(关闭)事件
function OnDocumentOpened(str, obj) {
    isOpened = true;
    oframe.ActiveDocument.Application.UserName = document.getElementById("tUserName").value;
    oframe.ActiveDocument.Saved = true; //saved属性用来判断文档是否被修改过 ,文档打开的时候设置成ture,当文档被修改 ,自动被设置为false,该属性由office提供.
}
function OnDocumentClosed() {
    isOpened = false;
}


//基本操作
function NewDoc() {
    oframe.showdialog(0);
}
function OpenDoc() {
    oframe.showdialog(1);
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
    if (CheckFileOpened()) {
        oframe.printout(true);
        // 1.oframe.printout(true) = oframe.showdialog(4);(弹出打印设置页面)
        // 2.oframe.printout(false); (直接打印)
    }
}
function OpenProperty() {
    if (CheckFileOpened()) {
        oframe.showdialog(6);
    }
}
function CloseDoc() {
    if (CheckFileOpened())
        oframe.close();
}

//菜单操作
function ToggleTitlebar() {
    oframe.Titlebar = !oframe.Titlebar;
    oframe.Activate();
}
function ToggleToolbars() {
    oframe.Toolbars = !oframe.Toolbars;
    oframe.Activate();
}
function ToggleMenubar() {
    oframe.Menubar = !oframe.Menubar;
    oframe.Activate();
}


//Word相关
function AddNewWord() {
    oframe.CreateNew("Word.Document");
}
function OpenWebWord(url) {
    oframe.Open(url, true);
}
function SetUserName() {
    if (CheckFileOpened()) {
        oframe.ActiveDocument.Application.UserName = document.getElementById("tUserName").value;
    }
}
function ToggleTrackRevisions() { //是否保留痕迹
    if (CheckFileOpened()) {
        oframe.ActiveDocument.TrackRevisions = !oframe.ActiveDocument.TrackRevisions;
    }
}
function ToggleShowRevisions() { //是否显示痕迹
    if (CheckFileOpened()) {
        oframe.ActiveDocument.ShowRevisions = !oframe.ActiveDocument.ShowRevisions;
    }
}
function AcceptAllRevisions() { //接受所有修订
    if (CheckFileOpened()) {
        oframe.ActiveDocument.AcceptAllRevisions();
    }
}
function RejectAllChangesInDoc() { //拒绝所有修订
    if (CheckFileOpened()) {
        oframe.ActiveDocument.Application.WordBasic.RejectAllChangesInDoc();
    }
}
function ProtectDoc(type) { //文档加锁
    if (CheckFileOpened()) {
        try { oframe.ActiveDocument.Protect(type, true, ""); } catch (e) { }
        /*
        Protect(Type, NoReset, Password, UseIRM, EnforceStyleLock)
        Type:wdAllowOnlyComments = 1,wdAllowOnlyFormFields = 2,wdAllowOnlyReading = 3,wdAllowOnlyRevisions = 0,wdNoProtection = -1
        NoReset:Optional Variant. False to reset form fields to their default values. True to retain the current form field values if the specified document is protected. If Type isn't wdAllowOnlyFormFields, the NoReset argument is ignored.
        Password:Optional Variant. The password required to remove protection from the specified document. (See Remarks below.)
        UseIRM:Optional Variant. Specifies whether to use Information Rights Management (IRM) when protecting the document from changes.
        EnforceStyleLock:Optional Variant. Specifies whether formatting restrictions are enforced in a protected document.
        */
    }
}
function UnProtectDoc() { //文档解锁
    if (CheckFileOpened()) {
        try { oframe.ActiveDocument.UnProtect(); } catch (e) { }
    }
}
function SetViewMode() {
    if (CheckFileOpened()) {
        //普通视图=1, 大纲视图 =2,页面视图 =3, 打印预览视图 =4,主控视图 =5,Web 视图=6, 阅读视图 =7
        //主控视图：会导致所有文本变成formtext，用快捷键alt+f9可以把formtext切换为文字
        oframe.ActiveDocument.ActiveWindow.View.Type = document.getElementById("sViewMode").value;
    }
}
function SetPageFit() { //使页面自动适应用户的可视范围
    if (CheckFileOpened()) {
        oframe.ActiveDocument.ActiveWindow.View.Zoom.PageFit = document.getElementById("sPageFit").value;
    }
}




