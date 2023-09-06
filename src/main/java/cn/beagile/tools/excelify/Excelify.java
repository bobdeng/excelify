package cn.beagile.tools.excelify;

import com.jayway.jsonpath.Configuration;
import com.jayway.jsonpath.JsonPath;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;
import java.util.stream.IntStream;

public class Excelify {
    private String data;
    private Object document;
    private byte[] templateFileBytes;

    public Excelify(String data, byte[] byteArray) {
        this.data = data;
        this.templateFileBytes = byteArray;
    }

    public void write(OutputStream outputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(templateFileBytes));
        expandAllArrayPlaceholders(workbook.getSheetAt(0));
        fillAllPlaceholders(workbook.getSheetAt(0));
        workbook.write(outputStream);
    }

    private void fillAllPlaceholders(XSSFSheet sheet) {
        IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .forEach(this::fillRow);
    }

    private void expandAllArrayPlaceholders(XSSFSheet sheet) {
        while (isNeedExpand(sheet)) {
            expandArray(sheet);
        }
    }

    private void expandArray(XSSFSheet sheet) {
        IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(this::isArray)
                .forEach(row -> appendArray(sheet, row.getRowNum()));
    }

    private boolean isNeedExpand(XSSFSheet sheet) {
        return IntStream.rangeClosed(0, sheet.getLastRowNum())
                .mapToObj(sheet::getRow)
                .filter(Objects::nonNull)
                .anyMatch(this::isArray);
    }

    private int getArrayLengthOfRow(XSSFRow row) {
        return IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(this::isArrayPlaceholder)
                .map(this::getArrayLength)
                .findFirst().orElse(0);
    }

    private int getArrayLength(XSSFCell cell) {
        String name = removeDecoration(cell.getStringCellValue());
        name = name.substring(0, name.indexOf("[]"));
        String jsonPath = "$." + name + ".length()";
        return JsonPath.read(getDocument(), jsonPath);
    }

    public String getSingleValueFromJson(String name) {
        return readStringFromJson(removeDecoration(name));
    }

    private String removeDecoration(String name) {
        return name.substring(1, name.length() - 1);
    }

    public String readStringFromJson(String name) {
        try {
            return JsonPath.read(getDocument(), "$." + name).toString();
        } catch (Exception e) {
            return e.getMessage();
        }
    }

    private Object getDocument() {
        if (document == null) {
            document = Configuration.defaultConfiguration().jsonProvider().parse(data);
        }
        return document;
    }

    private void fillRow(XSSFRow row) {
        IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(Objects::nonNull)
                .filter(this::isPlaceholder)
                .forEach(this::fillCell);
    }

    private void fillCell(XSSFCell cell) {
        cell.setCellValue(getSingleValueFromJson(cell.getStringCellValue()));
    }

    private boolean isPlaceholder(XSSFCell cell) {
        return cell.getStringCellValue().startsWith("≮") && cell.getStringCellValue().endsWith("≯");
    }

    private boolean isArrayPlaceholder(XSSFCell cell) {
        if (cell == null) {
            return false;
        }
        return isPlaceholder(cell) && cell.getStringCellValue().contains("[]");
    }

    private void appendArray(XSSFSheet sheet, int rowIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        IntStream.range(0, getArrayLengthOfRow(row))
                .forEach(i -> appendRow(sheet, rowIndex, i));
        removeFirstRowAndShiftBelowIt(sheet, rowIndex, row);
    }

    private void appendRow(XSSFSheet sheet, int rowIndex, int offset) {
        XSSFRow row = sheet.getRow(rowIndex);
        if (hasRowsBelow(sheet, rowIndex, offset)) {
            sheet.shiftRows(rowIndex + offset + 1, sheet.getLastRowNum(), 1, true, false);
        }
        XSSFRow newRow = sheet.createRow(rowIndex + offset + 1);
        copyArrayCells(row, newRow, offset);
    }

    private boolean hasRowsBelow(XSSFSheet sheet, int rowIndex, int offset) {
        return rowIndex + offset + 1 < sheet.getLastRowNum();
    }

    private static void removeFirstRowAndShiftBelowIt(XSSFSheet sheet, int start, XSSFRow row) {
        sheet.removeRow(row);
        sheet.shiftRows(start + 1, sheet.getLastRowNum(), -1);
    }

    private boolean isArray(XSSFRow row) {
        if (row == null) {
            return false;
        }
        return IntStream.rangeClosed(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .anyMatch(this::isArrayPlaceholder);
    }

    private void copyArrayCells(XSSFRow source, XSSFRow target, int rowIndex) {
        for (int cellIndex = 0; cellIndex < source.getLastCellNum(); cellIndex++) {
            XSSFCell cell = source.getCell(cellIndex);
            copyCell(target, rowIndex, cellIndex, cell);
        }
    }

    private void copyCell(XSSFRow target, int rowIndex, int cellIndex, XSSFCell cell) {
        if (isArrayPlaceholder(cell)) {
            copyArrayCellAddIndex(target, rowIndex, cellIndex, cell);
            return;
        }
        justCopyCellContent(target, cellIndex, cell);
    }

    private void justCopyCellContent(XSSFRow target, int cellIndex, XSSFCell cell) {
        XSSFCell targetCell = target.createCell(cellIndex);
        targetCell.copyCellFrom(cell, new CellCopyPolicy.Builder().build());
    }

    private void copyArrayCellAddIndex(XSSFRow target, int rowIndex, int cellIndex, XSSFCell cell) {
        XSSFCell targetCell = target.createCell(cellIndex);
        targetCell.setCellValue(cell.getStringCellValue().replace("[]", "[" + rowIndex + "]"));
    }
}
