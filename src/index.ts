// 导入必要的模块
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import ExcelJS from 'exceljs';  // 使用ExcelJS库替代xlsx库，提供更好的Excel操作功能
import fs from 'fs';

// 创建MCP服务器实例
const server = new McpServer({
  name: "excel_tools",  // 工具名称
  version: "1.0.0"      // 工具版本号
});

// 定义读取Excel文件的工具方法
server.tool(
  "excel.read",  // 工具方法标识符
  "读取Excel文件内容",  // 工具方法描述
  {
    // 定义输入参数Schema
    filePath: z.string().describe("需要读取的Excel文件绝对路径，例如：C:/path/to/file.xlsx"),
  },
  async ({ filePath }) => {
    // 检查文件是否存在
    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel文件不存在: ${filePath}`);
    }
    
    // 创建Workbook实例并读取文件
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // 处理所有工作表数据
    const sheets = workbook.worksheets.map((worksheet) => {
      if (worksheet.rowCount === 0) return null;  // 跳过空工作表
      
      const data: any[] = [];  // 存储数据行
      let titles: any[] = [];   // 存储表头行
      
      // 遍历工作表的每一行
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        // 获取行数据（处理ExcelJS不同版本兼容性）
        const rowValues = (row as any).values;
        const rawValues = typeof rowValues === 'function' ? rowValues() : rowValues;
        const cleanValues = (rawValues as any[]).slice(1);  // 去除第一个空值
        
        if (rowNumber === 1) {
          // 第一行作为表头
          titles = [rowNumber, ...cleanValues];
        } else {
          // 其他行作为数据
          data.push([rowNumber, ...cleanValues]);
        }
      });
      
      // 返回工作表数据
      return {
        type: "text" as const,
        text: JSON.stringify({
          SheetName: worksheet.name,  // 工作表名称
          title: titles,              // 表头数据
          content: data               // 内容数据
        })
      };
    }).filter(sheet => sheet !== null);

    // 返回处理结果
    return { 
      content: sheets,
      describe: "返回数据结构：{content: Array<{type: 'text', text: string}>}, text字段包含 {SheetName: 工作表名称, title: [行号, ...表头], content: [[行号, ...数据行]]}"
    };
  }
);

// 修改创建工具的Schema定义
server.tool(
  "excel.create",
  "创建新的Excel文件并写入数据",
  {
    filePath: z.string().describe("要创建的Excel文件绝对路径，例如：C:/path/to/file.xlsx"),
    sheets: z.array(
      z.object({
        SheetName: z.string().describe("工作表名称，例如：'成绩单'"),
        title: z.array(z.string())
          .default([])  // 默认空数组
          .describe("表头数组（不需要行号），例如：['姓名', '分数']"),
        content: z.array(z.array(z.any()))
          .default([])  // 默认空数组
          .describe("数据行数组（不需要行号），按顺序填写值，例如：[['小李', 90], ['小王', 85]]")
      })
    ).describe("工作表数组，必须包含SheetName字段")
  },
  async ({ filePath, sheets }) => {
    // 创建新的Workbook实例
    const workbook = new ExcelJS.Workbook();
    
    // 遍历所有工作表配置
    sheets.forEach(sheet => {
      // 添加新工作表
      const worksheet = workbook.addWorksheet(sheet.SheetName);
      
      // 如果有表头数据，添加表头行
      if (sheet.title.length > 0) {
        worksheet.addRow(sheet.title);
      }
      
      // 添加所有数据行
      sheet.content.forEach(row => {
        worksheet.addRow(row);
      });
    });
    
    // 将Workbook写入文件
    await workbook.xlsx.writeFile(filePath);
    
    // 返回创建成功信息
    return {
      content: [{
        type: "text",
        text: `Excel文件已创建: ${filePath}` 
      }]
    };
  }
);

// 修改更新工具的Schema定义
server.tool(
  "excel.update",  // 工具方法标识符
  "更新excel文件中的指定单元格内容。",  // 工具描述
  {
    // 输入参数Schema定义
    filePath: z.string().describe("要更新的excel文件绝对路径（必须存在）"),
    sheetName: z.string().describe("目标工作表名称（必须存在）"),
    rowsData: z.array(
      z.array(z.any())
        .refine(arr => arr.length >= 2, "至少需要行号和一个数据字段")
        .refine(arr => typeof arr[0] === 'number', "第一个元素必须是数字行号")
    ).describe(`多行更新数据数组，格式：
      [
        [行号, 字段值],  // 单字段更新模式
        [行号, 字段1索引, 字段1值, 字段2索引, 字段2值...]  // 多字段更新模式
      ]
      示例：[[10, '新名字'], [15, 2, 95, 5, '及格']]`),
    savePath: z.string().optional().describe("另存路径（可选，默认覆盖原文件）")
  },
  async ({ filePath, sheetName, rowsData, savePath }) => {
    // 检查文件是否存在
    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel文件不存在: ${filePath}`);
    }
    
    // 读取Excel文件
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // 获取指定工作表
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`工作表不存在: ${sheetName}`);
    }

    // 记录更新的单元格
    const updatedCells = [];
    // 遍历所有要更新的行数据
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
      } else {
        // 单字段模式 [值]（默认更新到第二列）
        const cell = worksheet.getCell(rowNumber, 2);
        cell.value = updatePairs[0];
        updatedCells.push(`${rowNumber}-2`);
      }
    }
    
    // 保存文件（如果指定了savePath则另存，否则覆盖原文件）
    await workbook.xlsx.writeFile(savePath || filePath);
    
    // 返回更新结果
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          success: true,
          message: `已更新 ${updatedCells.length} 个单元格`,
          updatedCells: updatedCells
        })
      }]
    };
  }
);

// 启动MCP服务
// 使用StdioServerTransport作为通信传输层
// 通过标准输入输出(stdin/stdout)与客户端通信
const transport = new StdioServerTransport();
await server.connect(transport);  // 连接并启动服务