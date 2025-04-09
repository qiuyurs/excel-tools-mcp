// 修改顶部导入语句
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import ExcelJS from 'exceljs'; // 替换xlsx为exceljs
import fs from 'fs';
// Create an MCP server
const server = new McpServer({
    name: "excel",
    version: "1.0.0"
});
// 读取excel文件内容
server.tool("excel.read", {
    filePath: z.string().describe("需要读取的excel文件绝对路径"),
}, async ({ filePath }) => {
    if (!fs.existsSync(filePath)) {
        throw new Error(`Excel文件不存在: ${filePath}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheets = workbook.worksheets.map((worksheet) => {
        if (worksheet.rowCount === 0)
            return null;
        const data = [];
        let titles = [];
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            // 修正类型断言
            const rowValues = row.values; // 强制类型断言为any
            const rawValues = typeof rowValues === 'function' ? rowValues() : rowValues;
            const cleanValues = rawValues.slice(1);
            if (rowNumber === 1) {
                titles = [rowNumber, ...cleanValues];
            }
            else {
                data.push([rowNumber, ...cleanValues]);
            }
        });
        return {
            type: "text",
            text: JSON.stringify({
                SheetName: worksheet.name,
                title: titles,
                content: data
            })
        };
    }).filter(sheet => sheet !== null);
    return {
        content: sheets,
        describe: "返回数据结构：{content: Array<{type: 'text', text: string}>}, text字段包含 {SheetName: 工作表名称, title: [行号, ...表头], content: [[行号, ...数据行]]}"
    };
});
// 修改创建工具的Schema定义
server.tool("excel.create", {
    filePath: z.string().describe("要创建的excel文件绝对路径，例如：C:/path/to/file.xlsx"),
    sheets: z.array(z.object({
        SheetName: z.string().describe("工作表名称，例如：'成绩单'"),
        title: z.array(z.string())
            .default([]) // 修复64行：添加默认值
            .describe("表头数组（不需要行号），例如：['姓名', '分数']"),
        content: z.array(z.array(z.any()))
            .default([]) // 修复64行：添加默认值
            .describe("数据行数组（不需要行号），按顺序填写值，例如：[['小李', 90], ['小王', 85]]")
    })).describe("工作表数组，必须包含SheetName字段")
}, async ({ filePath, sheets }) => {
    const workbook = new ExcelJS.Workbook();
    sheets.forEach(sheet => {
        const worksheet = workbook.addWorksheet(sheet.SheetName);
        if (sheet.title.length > 0) {
            worksheet.addRow(sheet.title);
        }
        sheet.content.forEach(row => {
            worksheet.addRow(row);
        });
    });
    await workbook.xlsx.writeFile(filePath);
    return {
        content: [{
                type: "text",
                text: `Excel文件已创建: ${filePath}`
            }]
    };
});
// 修改更新工具的Schema定义
server.tool("excel.update", "更新excel文件中的指定单元格内容。", {
    filePath: z.string().describe("要更新的excel文件绝对路径（必须存在）"),
    sheetName: z.string().describe("目标工作表名称（必须存在）"),
    rowsData: z.array(z.array(z.any())
        .refine(arr => arr.length >= 2, "至少需要行号和一个数据字段")
        .refine(arr => typeof arr[0] === 'number', "第一个元素必须是数字行号")).describe(`多行更新数据数组，格式：
      [
        [行号, 字段值],  // 单字段更新模式
        [行号, 字段1索引, 字段1值, 字段2索引, 字段2值...]  // 多字段更新模式
      ]
      示例：[[10, '新名字'], [15, 2, 95, 5, '及格']]`),
    savePath: z.string().optional().describe("另存路径（可选，默认覆盖原文件）")
}, async ({ filePath, sheetName, rowsData, savePath }) => {
    if (!fs.existsSync(filePath)) {
        throw new Error(`Excel文件不存在: ${filePath}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
        throw new Error(`工作表不存在: ${sheetName}`);
    }
    const updatedCells = [];
    for (const rowData of rowsData) {
        const [rowNumber, ...updatePairs] = rowData;
        // 处理两种更新模式
        if (updatePairs.length % 2 === 0) {
            // 多字段模式 [列索引1, 值1, 列索引2, 值2...]
            for (let i = 0; i < updatePairs.length; i += 2) {
                const colNumber = updatePairs[i];
                const value = updatePairs[i + 1];
                const cell = worksheet.getCell(rowNumber, colNumber);
                cell.value = value;
                updatedCells.push(`${rowNumber}-${colNumber}`);
            }
        }
        else {
            // 单字段模式 [值]（默认更新到第二列）
            const cell = worksheet.getCell(rowNumber, 2);
            cell.value = updatePairs[0];
            updatedCells.push(`${rowNumber}-2`);
        }
    }
    await workbook.xlsx.writeFile(savePath || filePath);
    return {
        content: [{
                type: "text",
                text: JSON.stringify({
                    success: true,
                    message: `已更新 ${updatedCells.length} 个单元格`,
                    updatedCells: updatedCells,
                    describe: "更新单元格格式：行号-列号（例如：2-3 表示第2行第3列）"
                })
            }]
    };
});
// Start receiving messages on stdin and sending messages on stdout
const transport = new StdioServerTransport();
await server.connect(transport);
