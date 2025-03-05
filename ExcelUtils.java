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
     * ɾ��ָ��Excel�ļ���ָ���������ĳһ��
     * @param filePath �ļ�·��
     * @param sheetIndex ��������������0��ʼ��
     * @param columnIndex Ҫɾ��������������0��ʼ��
     */
    public static void deleteColumn(String filePath, int sheetIndex, int columnIndex) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            if (sheet != null) {
                removeColumn(sheet, columnIndex);
            }

            // �����޸ĺ���ļ�
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * ɾ��ָ���������е�һ��
     * @param sheet Ŀ�깤����
     * @param columnIndex Ҫɾ����������
     */
    private static void removeColumn(XSSFSheet sheet, int columnIndex) {
        if (sheet == null || columnIndex < 0) {
            return; // ��֤��������Ϸ���
        }

        int lastRowNum = sheet.getLastRowNum(); // ��ȡ���һ������
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                int lastCellNum = row.getLastCellNum(); // ��ȡ�������һ������
                for (int j = columnIndex; j < lastCellNum - 1; j++) {
                    // �ƶ�ÿ��Ԫ�������
                    XSSFCell currentCell = row.getCell(j);
                    XSSFCell nextCell = row.getCell(j + 1);
                    if (currentCell == null) {
                        currentCell = row.createCell(j); // ������ǰ��Ԫ��
                    }
                    if (nextCell != null) {
                        copyCell(currentCell, nextCell);
                    } else {
                        currentCell.setBlank(); // �����ǰ��Ϊ�գ���յ�Ԫ��
                    }
                }
                // ɾ�����һ��
                XSSFCell lastCell = row.getCell(lastCellNum - 1);
                if (lastCell != null) {
                    row.removeCell(lastCell);
                }
            }
        }
    }

    /**
     * ���Ƶ�Ԫ������ݺ���ʽ
     * @param newCell �µ�Ԫ��
     * @param oldCell �ɵ�Ԫ��
     */
    private static void copyCell(XSSFCell newCell, XSSFCell oldCell) {
        newCell.setCellStyle(oldCell.getCellStyle()); // ������ʽ
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
     * ɾ��ָ���������ĳһ�У��������������������
     * @param filePath �ļ�·��
     * @param sheetIndex ��������������0��ʼ��
     * @param rowIndex Ҫɾ��������������0��ʼ��
     */
    public static void deleteRowAndShiftUp(String filePath, int sheetIndex, int rowIndex) {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            manuallyRemoveRow(sheet, rowIndex);

            // �����޸ĺ���ļ�
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * �ֶ�ɾ��ָ�����У������·���������������
     * @param sheet Ŀ�깤����
     * @param rowIndex Ҫɾ��������������0��ʼ��
     */
    private static void manuallyRemoveRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();

        if (rowIndex >= 0 && rowIndex <= lastRowNum) {
            // �ֶ���������
            for (int i = rowIndex; i < lastRowNum; i++) {
                Row currentRow = sheet.getRow(i);
                Row nextRow = sheet.getRow(i + 1);

                if (currentRow == null) {
                    // �����ǰ��Ϊ�գ��򴴽�
                    currentRow = sheet.createRow(i);
                }

                if (nextRow != null) {
                    // ����һ�����ݸ��Ƶ���ǰ��
                    copyRowData(currentRow, nextRow);
                } else {
                    // �����һ��Ϊ�գ�����յ�ǰ��
                    clearRow(currentRow);
                }
            }

            // �Ƴ����һ��
            Row lastRow = sheet.getRow(lastRowNum);
            if (lastRow != null) {
                sheet.removeRow(lastRow);
            }
        }
    }

    /**
     * ���ָ���е����е�Ԫ������
     * @param row Ҫ��յ���
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
     * ��Դ�еĵ�Ԫ�����ݸ��Ƶ�Ŀ����
     * @param targetRow Ŀ����
     * @param sourceRow Դ��
     */
    private static void copyRowData(Row targetRow, Row sourceRow) {
        // ����Դ�е����е�Ԫ��
        for (int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.getCell(i);

            if (targetCell == null) {
                // ���Ŀ�굥Ԫ�񲻴��ڣ��򴴽�
                targetCell = targetRow.createCell(i);
            }

            if (sourceCell != null) {
                // ���ݵ�Ԫ�����͸��ƶ�Ӧ����
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

                // ���Ƶ�Ԫ����ʽ
                targetCell.setCellStyle(sourceCell.getCellStyle());
            } else {
                targetCell.setBlank();
            }
        }
    }
}
