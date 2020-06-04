



#### Wps 开发平台地址 ： [点击进入](https://open.wps.cn/docs/office)
##### 总结：
1. 提供了C++ 、Java 、 浏览器 接口.
1. Java运用场景应该不多，因为 **Applet** 现在已经不让用了。
1. C++多数封装成了 **Ocx** 让其与浏览器交互 （推荐，不过浏览器有限制）
1. 浏览器 用的 **NPAPI** 也被各大浏览器逐渐淘汰
1. 使用 **COM** 组件调用 浏览器使用 **activex** 调用 WPS （目前使用 ，google不能用）
****
##### 以下是功能代码
```
   var WordApp = new ActiveXObject("kwps.Application"); //  实例化一个wps对象
   WordApp.Application.Visible=false;   //设置隐藏窗口
   var Doc=WordApp.Documents.Add("c:\\tmp.doc");   //打开一个doc
   Doc.ExportAsFixedFormat("c:\\1.pdf",17);   // 导出为pdf文件
   WordApp.Quit();  //退出wps 
```
<br>
**注： 必须打开Internet 选项中对Activex的支持**




