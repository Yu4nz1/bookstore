## 利用 iPad 做 Win 10的拓展屏幕

今年寒假的时候入手了iPad，和小伙伴一起。

然后小伙伴又入了一个显示器，三屏幕，太爽了。孩子直呼羡慕。

太馋了！

于是准备自己也搞一个。2333

### 操作

准备：iPad、Win 10 、数据线（电脑最好要有Type-C接口，我用了转接头）

Win 10和iPad都下载Xdispaly这个软件

[将平板电脑变成显示器 | Splashtop Wired XDisplay | 使用平板电脑作为第二屏幕](https://www.splashtop.cn/cn/wiredxdisplay)

然后用数据线去连接两者

PC端还需要下载iTunes，会有提示的。

然后连接成功了。然后win+i 打开系统显示设置，将多显示设置设置成为拓展这些屏幕，其他的一些设置可以根据自己的情况定。

![image-20210803145537287](http://inews.gtimg.com/newsapp_ls/0/13843441974/0)

双屏幕，左边ipad屏，右边PC

太爽了！





### VB6.0  adodc控件出现 类未注册，查找具有CLSID的对象问题

在数据库实训的时候 使用vb6.0时出现了以下情况。

![image-20210910100245319](http://inews.gtimg.com/newsapp_ls/0/13964389549/0)

vb的部分控件没有注册上，所以在使用的时候会提示这个错误。（就挺离谱的，全班大概只有我遇到了这个问题）

> regsvr32 msstkprp.dll

用该命令尝试注册时，有出现了问题。提示说 该模块未加载，请确保该二进制存储在指定的路径……

我就挺emmmmm 累了

其实基本上我们只要在注册表上注册上就行 但就挺emmmm 

还好我在查找该问题时 发现了曾经有一个up🐖写过怎么解决该问题，于是我就开始了白嫖之旅——白嫖了他的脚本

成功解决！

[来源]:[vb6.0 解决 adodc控件等右键点属性提示 类没有注册，查找具有 CLSID 的对象：{xxx} visual basic_哔哩哔哩_bilibili](https://www.bilibili.com/video/BV1Zv41177mu/)

以上！老师已经讲到好后面了，而我

SOS！该问题的后果就是老师已经讲到好后面了，而我才刚刚做好控件。嘤嘤嘤！

另，因为该问题，我之前已经写好的部分功能的xxx系统被我一怒之下删除了！好家伙！直接一周白干！









### VB+SQL 网上书店管理系统

网上书店，是一种网上电子购物系统，让买卖双方充分利用互联网的潜力,在无限的空间里拓展营销渠道○网上书店是一个可以无限伸展的书库，可以容纳无限的图书或图样乃至于内容，检索查询不受时间空间的限制·网上书店属于电子商务的范畴，泛指利用互联网进行图书商品营销的虚拟商店，是现代信息技术应用于图书发行领域的产物○在形式上，网上书店与传统书店迥异，它没有物理意义的店面，而是借助计算机技术﹑网络技术等现代信息系统技术及相关设备向读者展示图书·在功能上，它则与传统书店一致，即让读者了解进而购买需求的图书，以此获取效益。
数据要求
（1）用户名信息
包括用户名﹑密码﹑真实姓名﹑地址﹑联系电话，权限

Users(用户名 char(20) primary key , 密码 char(20) , 姓名 char(20) , 地址 char(50) , 联系电话 char(20) , 权限 int )

（2）图书类别信息
包括类别名（例如文学﹑体育﹑经济﹑教材等）和类别概要信息。

BookClassify(类别名 char(20)  , 类别信息)

(3)图书信息
包括图书名称﹑作者、ISBN号，出版社﹑出版时间﹑发行量﹑版号﹑页数﹑内容简介﹑读者评价(可选)﹑专家推荐（可选)﹑封面图片(可选)等信息。

Book(ISBN号 char(20) primary key , 图书名称 char(20) , 作者 char(20) , 出版社 char(20) ,价格 float , 类别 char(20) , 出版时间 time(10) , 发行量 smallint , 页数 smallint , 内容简介 char(200) , 封面图片 image，foreign key (类别) references BookClassify)

(4)订单信息
包括图书名称列表﹑单价﹑总金额﹑日期﹑顾客标识发货日期﹑状态(包括等待﹑执行﹑完成)

Order(图书名称 char(20) , 单价 float , 数量 int , 总价 float, 顾客姓名 char(20) , 发货日期 time(10) , 状态(等待、执行、完成) char(5) , primary key(图书名称，顾客姓名) , foreign key (图书名称) references Book(图书名称) , foreign key (顾客姓名) references User(用户名))

```sql
create table Users(用户名 char(20) primary key , 密码 char(20) , 姓名 char(20) , 地址 char(50) , 联系电话 char(20) , 权限 int )

create table BookClassify(类别名 char(20) primary key, 类别信息 char(100))

create table Books(ISBN号 char(20)  , 图书名称 char(20) primary key , 作者 char(20) , 出版社 char(20) ,价格 float , 类别 char(20) , 出版时间 time(7) , 发行量 smallint , 页数 smallint , 内容简介 char(200) , 封面图片 image , foreign key (类别) references BookClassify(类别名))

create table Orders(图书名称 char(20) , 单价 float , 数量 int , 总价 float, 顾客姓名 char(20) , 发货日期 time(7) , 状态 char(5) , primary key(图书名称 ,顾客姓名) , foreign key (图书名称) references Books(图书名称) , foreign key (顾客姓名) references Users(用户名))


create login ydczsq with password='881212',default_database=Bookstore_Management
create user ydczsq for login ydczsq with default_schema=dbo
exec sp_addrolemember 'db_owner','ydczsq'


```

