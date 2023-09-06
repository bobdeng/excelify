package cn.beagile.tools.excelify;

import com.google.common.io.Resources;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExportFormTest {
    private Excelify exportForm;
    private static String tempFile = "temp.xlsx";
    private String json;

    @BeforeEach
    public void setup() {
        json = """
                {
                        "name":"海洋学院",
                        "list":[{"name":"张三","age":18,"group":{"name":"班级1"}},{"name":"李四","age":19,"group":{"name":"班级2"}}]
                }
                """;
    }

    @AfterEach
    public void tearDown() {
        new File(tempFile).delete();
    }

    @Test
    public void get_array_data_value() {

        exportForm = new Excelify(json, new byte[0]);
        assertEquals("张三", exportForm.readStringFromJson("list[0].name"));
        assertEquals("李四", exportForm.readStringFromJson("list[1].name"));
        assertEquals("", exportForm.readStringFromJson("list[1].notExist"));
        assertEquals("海洋学院", exportForm.readStringFromJson("name"));
    }

    private byte[] readTemplate(String name) throws IOException {
        return Resources.toByteArray(Resources.getResource(name + ".xlsx"));
    }

    @Test
    public void should_export_to_excel() throws IOException {
        testExport("", "测试空", "");
    }

    private void testExport(String data, String template, String expectContent) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(tempFile);
        new Excelify(data, readTemplate(template)).write(outputStream);
        outputStream.close();
        assertExcelContent(expectContent);
    }

    @Test
    public void should_export_one() throws IOException {
        testExport(json, "测试一个值", "海洋学院");
    }

    @Test
    public void should_export_deep_array() throws IOException {
        testExport(json, "测试获取深度字段", "班级1\n" +
                "班级2");
    }


    @Test
    public void should_export_array() throws IOException {

        testExport(json, "测试一个数组值", "张三\n李四");
    }

    @Test
    public void should_export_2array() throws IOException {
        testExport(json, "测试2个数组值","""
                姓名
                张三
                李四
                                
                年龄
                18
                19
                """);
    }

    @Test
    public void should_export_mix() throws IOException {
        testExport(json, "测试混合","""
                ,海洋学院
                ,姓名,年龄
                ,张三,18
                ,李四,19
                """);
    }

    private void assertExcelContent(String expectContent) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(tempFile));
        XSSFSheet sheet = workbook.getSheetAt(0);
        String sheetContent = "";
        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            XSSFRow row = sheet.getRow(i);
            sheetContent += getRowString(row);
            sheetContent += "\n";
        }
        assertEquals(expectContent.trim(), sheetContent.trim());
    }

    private String getRowString(Row row) {
        if (row == null) {
            return "";
        }
        return IntStream.range(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .map(this::getCellString)
                .collect(Collectors.joining(","));

    }

    private String getCellString(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return "";
    }
}
