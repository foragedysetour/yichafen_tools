# yichafen\_tools

一个用于抓取易查分网页数据并保存到本地 Excel 的工具。

> \\\[!tips]
> 本软件仅供学习交流使用

### ⚙️如何使用

#### 1\. 配置文件说明 (config.json)

在项目根目录下的 `config.json` 文件中配置以下参数：

|参数|必填|说明|
|-|-|-|-|
|`base\_url`|✅|易查分网站的基本URL（不含https://）|
|`usersDB\_path(excel)`|✅|用户数据库Excel文件路径|
|`num\_threads`|❌|并发线程数，默认为4|

**config.json 示例：**

```json
{
    "base\_url": "xxxxx.yichafen.com",
    "usersDB\_path(excel)": "./settings/data.xlsx",
    "num\_threads": 4
}
```

**现在也可以直接通过设置界面进行设置**

#### 2\. 用户数据库Excel表格格式

`usersDB\_path(excel)` 指定的Excel文件需要满足以下格式要求：

##### 表格结构

* **第一行（表头）**：必须包含列名，列名需要与查询页面的字段名匹配
* **从第二行开始**：为实际的用户数据

##### 列名匹配规则

程序会读取Excel中与查询页面所需 `data-sname` 字段匹配的列。列名不区分大小写，支持模糊匹配。

例如：

* 查询页面需要 `data-sname` 为 "姓名"，则Excel表头可以是 "姓名"、"学生姓名" 等
* 查询页面需要 `data-sname` 为 "学号"，则Excel表头可以是 "学号"、"学生学号" 等

##### 示例表格

|姓名|学号|班级|
|-|-|-|
|张三|2024001|一年级一班|
|李四|2024002|一年级一班|
|王五|2024003|一年级二班|

> \[!important]
> - Excel文件必须存在且路径正确
> - 至少需要包含查询页面所需的一个字段列
> - 支持 .xlsx 格式

