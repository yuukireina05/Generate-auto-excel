## 公主连结-生成AUTO作业汇总表格

该段代码的作用是根据花舞攻略组发布的在线文档中记录的auto作业，生成一个一图流的auto作业汇总excel表格。如下图所示，该表格通过图片的方式可以挂在公告中，极大方便了auto作业的查阅和自行分刀的需求。

其中编号带中括号的表示该刀型为半auto刀（从文档中读取出来的，若无"半auto"则默认为auto刀）

![image-20210720112745500](https://raw.githubusercontent.com/yuukireina05/picture-repository/main/image-20210720112745500.png)

**使用方法：**

1. 打开在线文档，下载表格到本地，格式为`.xlsx`
2. 将代码文件和文档放在同一个文件夹中
3. 打开代码，修改`Filename`为下载文档的名字
4. 运行代码后打开生成的`auto_test.xlsx`