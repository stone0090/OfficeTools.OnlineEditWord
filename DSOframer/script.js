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
    alert("触发了dsoframe的OnDocumentOpened事件！");
    isOpened = true;
    oframe.ActiveDocument.Application.UserName = document.getElementById("tUserName").value;
    oframe.ActiveDocument.Saved = true; //saved属性用来判断文档是否被修改过 ,文档打开的时候设置成ture,当文档被修改 ,自动被设置为false,该属性由office提供.
}
function OnDocumentClosed() {
    alert("触发了dsoframe的OnDocumentClosed事件！");
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

//检查是否安装控件
function CheckControlInstall() {
    if (typeof (oframe) === 'undefined') {
        alert("You must install dsoframe control first.");
        return false;
    } else {
        return true;
    }
}

//Word相关
function AddNewWord() {
    oframe.CreateNew("Word.Document");
}
function OpenLocalWord(path) {
    oframe.Open(path, false, "Word.Document");
}
function OpenWebWord(url) {
    oframe.Open(url + "?random=" + Math.random(), true);
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

//上传word服务器的完整范例
//1、下载word到临时文件目录（要求ie浏览器必须开启了临时文件功能，系统默认是开启的）
//2、在页面上，使用DSOframer打开临时文件目录中的word，并进行编辑
//3、保存word到临时文件目录，然后再使用webfile将word上传到服务器
//备注：使用此方法需要把网站加入受信站点，并把安全级别设置为低，再开启Internet选项-安全-受信任站点-自定义级别-对未标记为可安全执行脚本的ActiviteX控件初始化并执行
function test() {
    debugger;
    alert(GetLocalTempFileName());
}

//用于上传下载的ActiviteX控件
var WebFile2;
//var webfile1 = new ActiveXObject("WebFileHelper.WebFile.1");
//var webfile2 = new ActiveXObject("WebFileHelper2.WebFile2.1");

function UploadWord(hostUrl) { //上传本地word到服务器上
    var tempFile1 = GetLocalTempFileName();
    var tempFile2 = GetLocalTempFileName();
    try {
        oframe.Save(tempFile1, true);
        var fso = new ActiveXObject("Scripting.FileSystemObject"); //word被占用时，上传会失败，所以用fso控件copy一份出来，再上传
        fso.CopyFile(tempFile1, tempFile2, true);
        //oframe.close(); //如果担心fso有兼容性问题，就必须先关闭oframe，才能成功上传word
    } catch (e) {
        alert("请先打开一个word！");
        return;
    }
    try {
        //WebFile2.MaxFileSize = 1258290;//该属性可以限制上传的容量
        WebFile2.UploadFile(tempFile2, "http://" + hostUrl + "/FileUpload.aspx");
        alert("上传成功！");
    } catch (e) {
        var errCode = e.number >> 16 & 0xFFFF;
        if (errCode == 32778)
            alert("上传失败:word文件容量过大，必须小于1.2M！");
        else
            alert("上传失败:" + e.message + "\n" + e.description);
    }
}

function DownloadWord(hostUrl) { //下载word到临时文件目录
    try {
        var tempFileName = GetLocalTempFileName();
        WebFile2.DownloadFile("http://" + hostUrl + "/FileDownload.aspx?random=" + Math.random(), tempFileName);
        //alert("成功下载word到临时文件目录：" + tempFileName);
        return tempFileName;
    } catch (e) {
        alert("下载失败，请检查您的ie设置！");
        return "";
    }
}

function GetLocalTempFileName() {
    if (WebFile2 == null) WebFile2 = document.getElementById("WebFile2");
    var time = new Date();
    var fileNameFix = PadLeft(time.getFullYear().toString(), 4) + PadLeft((time.getMonth() + 1).toString(), 2) + PadLeft(time.getDate().toString(), 2) + PadLeft(time.getHours().toString(), 2) + PadLeft(time.getMinutes().toString(), 2) + PadLeft(time.getSeconds().toString(), 2) + ("." + Math.random() * 1000).substr(0, 4);
    var tempFileName = WebFile2.GetLocalTempFile("temp") + "." + fileNameFix + ".doc"; //temp和.doc都是可以自己随意改的
    return tempFileName;
}

//字符串长度不足，左边补0
function PadLeft(str, len) {
    str = '00000000000000000000000000000' + str;
    return str.substr(str.length - len);
}


