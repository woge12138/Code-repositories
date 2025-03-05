package nc.bs.baseapp.utl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    /**
     * 删除指定Excel文件中指定工作表的某一列
     * @param filePath 文件路径
     * @param sheetIndex 工作表索引（从0开始）
     * @param columnIndex 要删除的列索引（从0开始）
     */
    public static void deleteColumn(String filePath, int sheetIndex, int columnIndex) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            if (sheet != null) {
                removeColumn(sheet, columnIndex);
            }

            // 保存修改后的文件
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 删除指定工作表中的一列
     * @param sheet 目标工作表
     * @param columnIndex 要删除的列索引
     */
    private static void removeColumn(XSSFSheet sheet, int columnIndex) {
        if (sheet == null || columnIndex < 0) {
            return; // 验证输入参数合法性
        }

        int lastRowNum = sheet.getLastRowNum(); // 获取最后一行索引
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                int lastCellNum = row.getLastCellNum(); // 获取行内最后一列索引
                for (int j = columnIndex; j < lastCellNum - 1; j++) {
                    // 移动每格单元格的内容
                    XSSFCell currentCell = row.getCell(j);
                    XSSFCell nextCell = row.getCell(j + 1);
                    if (currentCell == null) {
                        currentCell = row.createCell(j); // 创建当前单元格
                    }
                    if (nextCell != null) {
                        copyCell(currentCell, nextCell);
                    } else {
                        currentCell.setBlank(); // 如果当前列为空，清空单元格
                    }
                }
                // 删除最后一列
                XSSFCell lastCell = row.getCell(lastCellNum - 1);
                if (lastCell != null) {
                    row.removeCell(lastCell);
                }
            }
        }
    }

    /**
     * 复制单元格的内容和样式
     * @param newCell 新单元格
     * @param oldCell 旧单元格
     */
    private static void copyCell(XSSFCell newCell, XSSFCell oldCell) {
        newCell.setCellStyle(oldCell.getCellStyle()); // 复制样式
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                newCell.setBlank();
                break;
        }
    }

    /**
     * 删除指定工作表的某一行，并将下面的行逐行上移
     * @param filePath 文件路径
     * @param sheetIndex 工作表索引（从0开始）
     * @param rowIndex 要删除的行索引（从0开始）
     */
    public static void deleteRowAndShiftUp(String filePath, int sheetIndex, int rowIndex) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            manuallyRemoveRow(sheet, rowIndex);

            // 保存修改后的文件
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 手动删除指定的行，并将下方的内容逐行上移
     * @param sheet 目标工作表
     * @param rowIndex 要删除的行索引（从0开始）
     */
    private static void manuallyRemoveRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();

        if (rowIndex >= 0 && rowIndex <= lastRowNum) {
            // 手动上移内容
            for (int i = rowIndex; i < lastRowNum; i++) {
                Row currentRow = sheet.getRow(i);
                Row nextRow = sheet.getRow(i + 1);

                if (currentRow == null) {
                    // 如果当前行为空，则创建
                    currentRow = sheet.createRow(i);
                }

                if (nextRow != null) {
                    // 将下一行内容复制到当前行
                    copyRowData(currentRow, nextRow);
                } else {
                    // 如果下一行为空，则清空当前行
                    clearRow(currentRow);
                }
            }

            // 移除最后一行
            Row lastRow = sheet.getRow(lastRowNum);
            if (lastRow != null) {
                sheet.removeRow(lastRow);
            }
        }
    }

    /**
     * 清空指定行的所有单元格内容
     * @param row 要清空的行
     */
    private static void clearRow(Row row) {
        if (row != null) {
            for (int i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                if (cell != null) {
                    row.removeCell(cell);
                }
            }
        }
    }

    /**
     * 将源行的单元格数据复制到目标行
     * @param targetRow 目标行
     * @param sourceRow 源行
     */
    private static void copyRowData(Row targetRow, Row sourceRow) {
        // 遍历源行的所有单元格
        for (int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.getCell(i);

            if (targetCell == null) {
                // 如果目标单元格不存在，则创建
                targetCell = targetRow.createCell(i);
            }

            if (sourceCell != null) {
                // 根据单元格类型复制对应数据
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        targetCell.setBlank();
                        break;
                }

                // 复制单元格样式
                targetCell.setCellStyle(sourceCell.getCellStyle());
            } else {
                targetCell.setBlank();
            }
        }
    }
}
