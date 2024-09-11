# excel2jsonTool
这是一个使用python编写的轻量级的数据导出工具，根据配置规范的Excel文件导出json和c#处理脚本，可以在游戏启动时进行加载和反序列化。


具有以下特色：

1.配置简单：只需配置Excel目录和几个导出目录，并安装python运行环境即可使用。

2.傻瓜式操作：只需要点击运行.bat文件即可导出所有数据。

3.数据检查：生成json文件时，会检查excel中配置的字段格式、数据类型、是否存在重复字段和重复Sheet等规范检查。

4.运行时效率高：使用了字典来存储反序列化的数据，查找效率为O(1)。

5.策划友好：策划只面向Excel，包括增删修改文件、字段和数据，无需关心程序如何使用。且int和string皆可作为主键，string会在导出时转为enum类型。

6.程序友好：程序可以通过自动生成的C#代码来查找数据，提供了多种查找方法。


最佳实践：

待补充。


提示：

如果你需要使用更多类型的数据源而不只是Excel，或需要导出更多类型的数据而不只是json，推荐使用Luban。
