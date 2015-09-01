OfficeTools.OnlineEditWord
===

this tool is used for edit word/excel/ppt in web.

DSOframer 是微软提供一款开源的用于在线编辑 Word、 Excel 、PowerPoint 的 ActiveX 控件。国内很多著名的 OA 中间件，电子印章，签名留痕等大多数是依此改进而来的。虽然博主的公司已经用了 NTKO 取代了 DSOframer，但免费的控件依旧是更多人的选择，所以在此和大家分享一下 DSOframer 的常用功能。如果看完全文还是不能解决您的问题，请在[评论区](http://shijiajie.com/2013/01/28/dsoframer-introduction-resources/#ds-thread)留言，或加入QQ群(95674923)进行学习交流。

Demo演示：
---
![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/1.png)
![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/2.png)

资源介绍：
---
- DSOframer\ActiveX\DSOframer\DsoFramer_KB311765_x86.exe  
  备注：官方提供的安装包，里面包含 DSOframer.ocx 控件及源码，还有 VB版、VB.NET版、Web版 等3个Demo。

- DSOframer\ActiveX\DSOframer\DSOframer.CAB  
  备注：博主将 DsoFramer_KB311765_x86.exe 中的 DSOframer.ocx，打包成了 DSOframer.CAB，以便在 Web 中可以自动下载。可参见 [OCX打包CAB并签名过程](http://www.cnblogs.com/rushoooooo/archive/2011/06/22/2087542.html)。

- DSOframer\ActiveX\DSOframer2007\DSOframer2007.CAB  
  备注：博主公司使用的版本，貌似修复了一些office2007兼容性问题，如果上面那个用着有问题，可以试试这个。
  
- DSOframer\ActiveX\WebFileHelper.CAB  
  DSOframer\ActiveX\WebFileHelper2.CAB  
  备注：该控件只有简单的上传、下载、压缩等功能，也是博主用来上传 doc 到服务器的方法。`如果您觉得第三方 ActiveX 不安全，请不要使用这个方法`。
  
  因为该控件未签名，在部分电脑上可能会报以下错误。
  ![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/3.png)
  
  解决方案如下：  
  1.打开IE菜单 `工具->Internet选项`，选择 `安全` 选项卡，点击 `自定义级别` 按钮，将 `下载未签名的ActiveX控件（不安全）` 设置为 `启用（不安全）`。  
  2.打开IE菜单 `工具->Internet选项`，选择 `高级` 选项卡，勾选设置列表中 `允许运行或安装软件，即使签名无效`。
  ![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/4.png)

- DSOframer\OfficialDemo.htm  
  备注：官方安装包中的 Demo，代码是用 vbscript 写的，很多朋友说不能运行。
  
- DSOframer\OfficialDemo_JS.htm  
  备注：基于官方安装包的 Demo 用 javascript 重写的版本，功能跟官方 Demo 没有区别。
  
- DSOframer\CommonDemo.html  
  备注：常用功能总结，如果大家想让博主在 Demo 加入新的功能，请在[评论区](http://shijiajie.com/2013/01/28/dsoframer-introduction-resources/#ds-thread)留言。
  ![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/1.png)
  
- DSOframer\FileUpload.aspx  
  DSOframer\FileDownload.aspx  
  备注：基于 WebFileHelper2.CAB 控件的上传下载功能的 Demo，`再次重申，如果您觉得第三方 ActiveX 不安全，请不要使用这个方法`。
  ![](http://7xkhp9.com1.z0.glb.clouddn.com/blog/2013/01/28/dsoframer-introduction-resources/2.png)

- DSOframer\script.js  
  备注：大部分 DSOframer 操作都在该文件中，并写了详细的注释，请重点参考。

网上讲解 DSOframer 开发的文章有很多，个人觉得比较有价值的帖子有：  
1.[DSO(dsoframer)的接口文档](http://www.cnblogs.com/liping13599168/archive/2009/09/13/1565801.html)  
2.[DSOFramer 控件修改成功](http://www.cppblog.com/wanhhf/archive/2006/02/20/3355.html)  