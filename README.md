# Excel Tools MCP

基于大模型协议(MCP)的Excel操作工具，提供读取、创建和更新Excel文件的能力。

## 功能特性
- 📄 Excel文件读取（支持多工作表）
- ✏️ Excel文件创建（支持自定义工作表）
- 🔄 Excel文件更新（支持单元格级更新）
- 🤖 兼容MCP协议，可与大模型无缝集成

## 安装方式
```bash
npm install @qiuyurs/excel-tools-mcp -g
```

## MCP配置示例
```json
{
    "mcpServers": {
        "excel-tools": {
            "command": "npx",  
            "args": [  
                "-y",
                "@qiuyurs/excel-tools-mcp"
            ]
        }
    }
}
```

## 构建方式
```shell
npm install
npm run builld
```

## 支持工具

### excel.read
读取Excel文件内容。

#### 输入参数
- `filePath` (string): Excel文件路径。
#### 输出参数
返回数据结构：{content: Array<{type: 'text', text: string}>}, text字段包含 {SheetName: 工作表名称, title: [行号, ...表头], content: [[行号, ...数据行]]}

### excel.create
创建新的Excel文件。

#### 输入参数
- `filePath` (string): Excel文件路径。
- `sheets` (Array<{name: string, title: Array<string>, content: Array<Array<string>>}>): 工作表信息数组。
#### 输出参数
返回数据结构：{content: Array<{type: 'text', text: string}>}, text字段包含 {SheetName: 工作表名称, title: [行号,...表头], content: [[行号,...数据行]]}

### excel.update
更新Excel文件内容。

#### 输入参数
- `filePath` (string): Excel文件路径。
- `sheets` (Array<{name: string, title: Array<string>, content: Array<Array<string>>}>): 工作表信息数组。
#### 输出参数
返回数据结构：{content: Array<{type: 'text', text: string}>}, text字段包含 {SheetName: 工作表名称, title: [行号,...表头], content: [[行号,...数据行]]}

## 项目文档
Github：![https://github.com/qiuyurs/excel-tools-mcp](https://github.com/qiuyurs/excel-tools-mcp)

教程：![https://gwl1554ppni.feishu.cn/wiki/Yi5dw2N8midd8ekDOHOcefNrnVc?fromScene=spaceOverview](https://gwl1554ppni.feishu.cn/wiki/Yi5dw2N8midd8ekDOHOcefNrnVc?fromScene=spaceOverview)