# Excel Tools MCP

基于大模型协议(MCP)的Excel操作工具，提供读取、创建和更新Excel文件的能力。

## 功能特性
- 📄 Excel文件读取（支持多工作表）
- ✏️ Excel文件创建（支持自定义工作表）
- 🔄 Excel文件更新（支持单元格级更新）
- 🤖 兼容MCP协议，可与大模型无缝集成

## 安装方式
```bash
npm install excel-tools-mcp -g
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

## 项目文档
Github：![https://github.com/qiuyurs/excel-tools-mcp](https://github.com/qiuyurs/excel-tools-mcp)
教程：![https://gwl1554ppni.feishu.cn/wiki/Yi5dw2N8midd8ekDOHOcefNrnVc?fromScene=spaceOverview](https://gwl1554ppni.feishu.cn/wiki/Yi5dw2N8midd8ekDOHOcefNrnVc?fromScene=spaceOverview)