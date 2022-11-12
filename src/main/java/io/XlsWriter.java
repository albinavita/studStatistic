package io;

import model.Statistics;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class XlsWriter {

    private static String STATISTIC = "Статистика";
    private static String LEARNING_PROFILE = "Профиль обучения";
    private static String AVERAGE_SCOPE = "Средний балл";
    private static String NUMBER_STUDENT = "Количество студентов";
    private static String NUMBER_UNIVERSITY = "Количество университетов";
    private static String UNIVERSITY = "Университеты";


    private XlsWriter() {
    }

    public static void createXlsTable(List<Statistics> statistics, String fileName) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(STATISTIC);
        int indexRow = 0;
        Row headerRow = sheet.createRow(indexRow++);

        createCell(workbook, headerRow, 0, LEARNING_PROFILE, 11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        createCell(workbook, headerRow, 1, AVERAGE_SCOPE, 11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        createCell(workbook, headerRow, 2, NUMBER_STUDENT, 11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        createCell(workbook, headerRow, 3, NUMBER_UNIVERSITY, 11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        createCell(workbook, headerRow, 4, UNIVERSITY, 11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);

        for (Statistics statistic : statistics) {
            createRowData(sheet, indexRow++, statistic);
        }
        try (FileOutputStream stream = new FileOutputStream(fileName)){
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    /*Создаем ячейку и выравниваем ее определенным образом
    * @param wb книга Excel
    * @param row строка
    * @param column колонка
    * @param name название страницы
    * @params hpoints размер шрифта
    * @params value толщина шрифта
    * @params halign горизонтальное выравнивание
    * @params valign вертикальное выравнивание
    */
    private static void createCell(
                            Workbook wb,
                            Row row,
                            int column,
                            String name,
                            int hpoints,
                            boolean value,
                            HorizontalAlignment halign,
                            VerticalAlignment valign
                            ) {
        Font font = wb.createFont();
        font.setFontHeightInPoints( (short) hpoints);
        font.setFontName("Calibri");
        font.setBold(value);

        Cell cell = row.createCell(column);
        cell.setCellValue(name);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cellStyle.setFont(font);

        cell.setCellStyle(cellStyle);

    }

    private static void createRowData(Sheet sheet, int indexRow, Statistics statistic) {
        Row row = sheet.createRow(indexRow);
        int indexCol = 0;
        Cell learnProf = row.createCell(indexCol);
        learnProf.setCellValue(statistic.getProfile().getProfileName());
        sheet.autoSizeColumn(indexCol);
        indexCol++;
        Cell averScope = row.createCell(indexCol);
        averScope.setCellValue(statistic.getAvgExamScore());
        sheet.autoSizeColumn(indexCol);
        indexCol++;
        Cell numberStudent = row.createCell(indexCol);
        numberStudent.setCellValue(statistic.getNumberOfStudents());
        sheet.autoSizeColumn(indexCol);
        indexCol++;
        Cell numberUniversitet = row.createCell(indexCol);
        numberUniversitet.setCellValue(statistic.getNumberOfUniversities());
        sheet.autoSizeColumn(indexCol);
        indexCol++;
        Cell universitet = row.createCell(indexCol);
        universitet.setCellValue(statistic.getUniversityNames());
        sheet.autoSizeColumn(indexCol);
    }
}
