package mypkg;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelReader {

    /**
     * 读取 Excel 文件并返回第一张表。
     *
     * @param filePath 文件路径
     * @return 第一张表（Sheet 对象），如果读取失败则返回 null
     */
    public static Sheet getFirstSheetFromFile(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // 获取第一张表（索引从 0 开始）
            return workbook.getSheetAt(0);
        } catch (IOException e) {
            System.err.println("读取文件失败: " + e.getMessage());
            return null;
        }
    }

    /**
     * 读取指定单元格的内容。
     *
     * @param sheet          工作表对象
     * @param cellCoordinate 单元格坐标（如 "A1"、"B2" 等）
     * @return 单元格内容的字符串表示
     */
    public static String getCellContent(Sheet sheet, String cellCoordinate) {
        if (sheet == null) {
            return ""; // 如果表为空，返回空字符串
        }

        // 将单元格坐标转换为行和列的索引
        CellReference cellReference = new CellReference(cellCoordinate);
        int rowIndex = cellReference.getRow();
        int colIndex = cellReference.getCol();

        // 获取指定行
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return ""; // 如果行不存在，返回空字符串
        }

        // 获取指定单元格
        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) {
            return "NA"; // 如果单元格不存在，返回空字符串
        }

        // 根据单元格类型获取内容
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "NA";
            default:
                return "NA";
        }
    }

    /**
     * 将 name_to_location 转换为 name_to_value。
     *
     * @param name_to_location 名称到单元格坐标的映射
     * @param sheet            工作表对象
     * @return 名称到单元格内容的映射
     */
    public static LinkedHashMap<String, String> readExcel(LinkedHashMap<String, String> name_to_location, Sheet sheet) {
        LinkedHashMap<String, String> name_to_value = new LinkedHashMap<>();

        // 遍历输入的 LinkedHashMap
        for (Map.Entry<String, String> entry : name_to_location.entrySet()) {
            String name = entry.getKey();
            String location = entry.getValue();

            // 获取单元格内容
            String value = getCellContent(sheet, location);

            // 将结果存入新的 LinkedHashMap
            name_to_value.put(name, value);
        }

        return name_to_value;
    }

    public static String readFloatingTextbox(String filePath) {

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            if (drawing != null) {
                for (XSSFShape shape : drawing.getShapes()) {
                    if (shape instanceof XSSFSimpleShape) {
                        XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
                        String text = simpleShape.getText();
                        System.out.println(text);

                    }
                }
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return "finish";
    }

    public static void main(String[] args) {
        // 文件路径
        String filePath = "C:\\Users\\AQY2SZH\\Desktop\\excelTemplate3\\Corrugated Board_A4_0513.xlsx";

        // 获取第一张表
        Sheet sheet = getFirstSheetFromFile(filePath);
        if (sheet == null) {
            System.out.println("文件读取失败，请检查文件路径或文件格式。");
            return;
        }

        // 定义 name_to_location 映射
        LinkedHashMap<String, String> name_to_location = new LinkedHashMap<>();
        name_to_location.put("parameter1", "AE4");
        name_to_location.put("parameter2", "AE5");
        name_to_location.put("parameter3", "AE6");

        // 读取 Excel 并获取 name_to_value
        LinkedHashMap<String, String> name_to_value = readExcel(name_to_location, sheet);

        // 输出结果
        System.out.println("读取结果: " + name_to_value);

        //文本框
        readFloatingTextbox(filePath);
    }
}