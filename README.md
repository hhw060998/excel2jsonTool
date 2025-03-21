# excel2jsonTool

这是一个使用python编写的轻量级的数据导出工具，根据配置规范的Excel文件导出json和c#处理脚本，可以在游戏启动时进行加载和反序列化。

This is a lightweight data export tool written in python. It exports json and c# processing scripts according to the configuration specification Excel file, which can be loaded and deserialized when the game starts.


具有以下特色：

1. 配置简单：只需配置Excel目录、json和c#导出目录即可使用（需要python环境）。

2. 傻瓜式操作：只需要点击运行.bat文件即可导出所有数据，并生成代码文件。

3. 数据检查：生成json文件时，会检查excel中配置的字段格式、数据类型、是否存在重复字段和重复Sheet等规范检查。

4. 运行时效率：使用了字典来存储反序列化的数据，查找效率为O(1)。

5. 策划友好：策划只面向Excel，包括增删修改文件、字段和数据，无需关心程序如何使用。且int和string皆可作为主键，string会在导出时转为enum类型。

6. 程序友好：程序在游戏开始时调用加载方法，使用自动生成的C#代码来查找数据，还提供了多种查找方法。


It has the following features:

1. Simple configuration: Just configure the Excel directory, json and c# export directory to use it (requires python environment).

2. Fool-proof operation: Just click to run the .bat file to export all data and generate a code file.

3. Data check: When generating a json file, it will check the field format, data type, whether there are duplicate fields and duplicate sheets configured in excel, and other standard checks.

4. Runtime efficiency: A dictionary is used to store deserialized data, and the search efficiency is O(1).

5. Designer-friendly: Planning is only for Excel, including adding, deleting and modifying files, fields and data, and there is no need to care about how the program is used. In addition, both int and string can be used as primary keys, and strings will be converted to enum types when exported.

6. Programmer-friendly: The program calls the loading method at the start of the game, uses the automatically generated C# code to find data, and also provides a variety of search methods.

——————————————————————————————————————

设计思路及使用说明：

https://www.bilibili.com/video/BV1rz4CeJEF5

1. 首次使用需要配置python环境，并在ExcelFolder/!【导表】.bat 中配置Excel目录（即导入目录）、工具目录和各种导出目录。

2. 创建格式正确的Excel文件，参考ExcelFolder/SL示例.xlsx，文件名需要大写，Sheet名要符合c#类的命名规范（建议使用驼峰式）。

3. 运行ExcelFolder/!【导表】.bat 批处理脚本后，会将该Excel中的数据导出到指定json和c#脚本目录下。


Design Concept and Usage Instructions:

1. For the first-time use, you need to configure the Python environment and set up the Excel directory (i.e., the import directory), tool directory, and export directories in the ExcelFolder/!【导表】.bat file.

2. Create a properly formatted Excel file, referencing the ExcelFolder/SL示例.xlsx example. The file name needs to be in uppercase, and the sheet name should conform to C# class naming conventions (camelCase is recommended).

3. After running the ExcelFolder/!【导表】.bat batch script, the data from the Excel file will be exported to the specified JSON and C# script directories.

——————————————————————————————————————

提示：

以下情况推荐使用luban来处理数据，未来我可能也会根据需求对这个工具进行完善：

1. 使用更多类型的数据源而不只是Excel。

2. 需要导出更多类型的数据而不只是json。

3. 需要自定义数据结构。

4. 需要检查数据引用。


Tips:

It is recommended to use luban to process data in the following situations. I may also improve this tool according to needs in the future:

1. Use more types of data sources instead of just Excel.

2. Need to export more types of data instead of just json.

3. Need to customize the data structure.

4. Need to check data references.
