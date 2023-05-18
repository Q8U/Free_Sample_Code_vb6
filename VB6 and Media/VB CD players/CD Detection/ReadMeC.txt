CDNotification ActiveX Control v1.0

版权所有 1999年 李海，热情软件屋
发行日期：1999年8月10日
主页: http://members.tripod.com/~zealsoft/
      http://www.nease.net/~zealsoft/indexc.html
电子邮件: haili@public.bta.net.cn

CDNotification ActiveX Control是一个明信片软件，
如果你打算使用这个软件，请寄一张明信片（不是电子
邮件）给我，并请告诉我为什么你使用这个软件和你的
建议。我的地址：
    北京理工大学133单元1607号（100081）


什么是CDNotification ActiveX Control
--------------------------------------
CDNotification ActiveX Control v1.0是一个允许你的知道
用户在什么时候插入或取出CDRom盘片控件。

一天，我得到了ComonentSource演示光盘，当我换盘后，它可
以改变显示内容。我对此十分好奇，打算把这个功能加入我的
程序中。如果你也对此感兴趣，别忘了给我发一张明信片。

如果你想得到本软件的源程序，请参考“购买源程序”部分。

新闻邮件
--------------
如果你希望在新版本或新的免费控件发行时得到通知，你可
以访问
http://www.nease.net/~zealsoft/cdnotify
订阅Free Control新闻邮件(英文)。

安装/反安装
--------------------
你可以把所有文件解压缩后拷贝到硬盘。

CDNotify.ocx和其他VB5子目录下的文件是用于Visual Basic 
5.0 Service Pack 3(SP3)版的。

CDNotify6.ocx和其他VB6子目录下的文件是用于Visual Basic
6.0版。

如果你想下载VB5和VB6的运行时间库，可以访问
http://www.nease.net/~zealsoft/cdnotify。

要删除所有文件，只需要删除所有拷贝到硬盘上的文件就可以了。

如何使用
------------

1) 属性

   Enabled 属性
      数据类型: 布尔（Boolean）
      说    明: 当这个属性设为True(默认)，控件会在插入或取出
                光盘时产生相应的事件。 

2) 事件

   Arrival事件
      语法: Private Sub obj_Arrival(ByVal Drive As String)
      说明: 当插入CDRom时发生。参数Drive表示CD-Rom盘符。

   RemoveComplete 事件
      语法: Private Sub obj_RemoveComplete(ByVal Drive As String)
      说明:  当取出CDRom时发生。参数Drive表示CD-Rom盘符。


示例
---------
Visual Basic 5.0和6.0示例在Vb5和Vb6目录中。

购买源程序
-----------------
如果你想了解这个控件是如何工作的，可以访问
http://members.tripod.com/~zealsoft/cdnotify
来购买控件的源程序(10美圆，国内用户可以使用人民币购买，具体价格
可以通过电子邮件询问)。包括VB 5.0和VB 6.0源程序。

版本历史
---------
1.0	最初的版本