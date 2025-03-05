//调用示例 deleteColumn filePath 指需要修改的文件位置，sheetIndex是指第几个工作表，最后一个参数是删除的列
//对下载的文件进行修改方法，调用工具类
// 对第2和第3个工作表删除两次A列（索引为0），再删除F列（索引为5）
for (int sheetIndex = 1; sheetIndex <= 2; sheetIndex++) {
	ExcelUtils.deleteColumn(filePath, sheetIndex, 0); // 删除A列（第一次）
	ExcelUtils.deleteColumn(filePath, sheetIndex, 0); // 删除A列（第二次）
	ExcelUtils.deleteColumn(filePath, sheetIndex, 5); // 删除F列
}

// 对第4个工作表只删除一次A列
ExcelUtils.deleteColumn(filePath, 3, 0);

//调用示例 deleteRowAndShiftUp filePath 指需要修改的文件位置，sheetIndex是指第几个工作表，最后一个参数是删除的行
// 对第2、3、4个工作表删除第4行
for (int sheetIndex = 1; sheetIndex <= 3; sheetIndex++) {
	ExcelUtils.deleteRowAndShiftUp(filePath, sheetIndex, 3);
}
