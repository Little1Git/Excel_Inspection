package mypkg;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelReader {

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

    public static LinkedHashMap<String, MatchResult> checkConditions(LinkedHashMap<String, ArrayList<String>> conditions, LinkedHashMap<String, String> name_to_value) {
        LinkedHashMap<String, MatchResult> res = new LinkedHashMap<>();
        for (Map.Entry<String, ArrayList<String>> entry : conditions.entrySet()) {
            String key = entry.getKey();
            ArrayList<String> expectedValues = entry.getValue();
            String actualValue = name_to_value.get(key);

            MatchResult mr = new MatchResult();
            mr.key = key;
            mr.actualValue=actualValue;
            mr.expectedValue=expectedValues;

            if (actualValue != null && expectedValues.contains(actualValue)) {
                System.out.println("pass : key: " + key + ", Get value: " + actualValue);
                mr.result="pass";
            } else {
                System.out.println("failure : key: " + key + ", Expected values: " + expectedValues + ", Actual value: " + actualValue);
                mr.result="failure";
            }
            res.put(key,mr);
        }
        return res;
    }

    public static LinkedHashMap<String, String> extractParameters(List<String> parametersToBeOutput, LinkedHashMap<String, String> nameToValue) {
        LinkedHashMap<String, String> resultMap = new LinkedHashMap<>();

        for (String parameter : parametersToBeOutput) {
            String value = nameToValue.get(parameter);
            if (value != null) {
                resultMap.put(parameter, value);
            } else {
                resultMap.put(parameter, "NA");
            }
        }
        return resultMap;
    }

    public static LinkedHashMap<String, String> checkValues(LinkedHashMap<String, String> nameToValue) {
        LinkedHashMap<String, String> resultMap = new LinkedHashMap<>();

        for (Map.Entry<String, String> entry : nameToValue.entrySet()) {
            String key = entry.getKey();
            String value = entry.getValue();

            // 判断值是否为 "NA"
            if (!"NA".equals(value)) {
                resultMap.put(key, "pass");
            } else {
                resultMap.put(key, "failure");
            }
        }

        return resultMap;
    }

    public static void main(String[] args) {
        // 文件路径
        String filePath = "C:\\Users\\AQY2SZH\\Desktop\\excelTemplate3\\Corrugated Board_A4_0513.xlsx";

        // 定义 name_to_location 映射
        LinkedHashMap<String, String> name_to_location = new LinkedHashMap<>();
        name_to_location.put("Packaging PN", "U38");
        name_to_location.put("Description", "S36");
        name_to_location.put("Material", "K35");
        name_to_location.put("Weight", "M37");
        name_to_location.put("FEFCO Type", "AE5");
        name_to_location.put("Inner Dimensions", "AE8");
        name_to_location.put("Outside Dimensions", "AE9");
        name_to_location.put("ECT", "Y23");
        name_to_location.put("BST", "AA23");
        name_to_location.put("BCT", "AE17");
        name_to_location.put("View", "AE4");
        name_to_location.put("Manufacturer's Joint", "AE6");
        name_to_location.put("Type of Joining", "AE7");
        name_to_location.put("Printing", "AE10");
        name_to_location.put("Sort and/or Flute Combination", "AB");
        name_to_location.put("(Material Thickness)", "AE13");
        name_to_location.put("Glued Moisture-Resistant", "AE14");
        name_to_location.put("PET", "AE23");
        name_to_location.put("Ind.", "H31");
        name_to_location.put("Change", "I31");
        name_to_location.put("YYYYMMDD", "N31");
        name_to_location.put("Drawn", "Q31");
        name_to_location.put("Checked", "R31");
        name_to_location.put("Release", "T31");
        name_to_location.put("Resp. dept.", "X31");

        LinkedHashMap<String, ArrayList<String>>  conditions = new LinkedHashMap<>();
        conditions.put("View", new ArrayList<>(Arrays.asList("Outside")));
        conditions.put("FEFCO Type", new ArrayList<>(Arrays.asList("0201", "0200", "0300","Special")));
        conditions.put("Manufacturer's Joint", new ArrayList<>(Arrays.asList("Inside")));
        conditions.put("Type of Joining", new ArrayList<>(Arrays.asList("Stapled", "Glued","Special")));
        conditions.put("Printing", new ArrayList<>(Arrays.asList("Yes")));
        conditions.put("Resp. dept.", new ArrayList<>(Arrays.asList("ME/LOD1-CN")));

        List<String> Parameters_to_be_output = Arrays.asList("Packaging PN", "Description", "Weight","Special","FEFCO Type","Inner Dimensions","Outside Dimensions","ECT","BST","BCT");

        // 获取第一张表
        Sheet sheet = getFirstSheetFromFile(filePath);
        if (sheet == null) {
            System.out.println("文件读取失败，请检查文件路径或文件格式。");
            return;
        }

        // 读取 Excel 并获取 name_to_value
        LinkedHashMap<String, String> name_to_value = readExcel(name_to_location, sheet);
        System.out.println("读取结果: " + name_to_value);

        //1.输出列表
        LinkedHashMap<String, String> key_value_to_print = extractParameters(Parameters_to_be_output,name_to_value);


        //2.存在列表
        LinkedHashMap<String, String> pass_or_failure = checkValues(name_to_value);


        //3.匹配列表
        LinkedHashMap<String, MatchResult> expected_value_and_actual_value = checkConditions(conditions,name_to_value);

        // 遍历 LinkedHashMap 并输出
        for (Map.Entry<String, MatchResult> entry : expected_value_and_actual_value.entrySet()) {
            String key = entry.getKey();
            MatchResult matchResult = entry.getValue();

            // 输出键和 MatchResult 对象的详细信息
            System.out.println("Key: " + key);
            System.out.println("MatchResult: " + matchResult);
            System.out.println(); // 空行分隔
        }

        System.out.println("pass or failure: "+pass_or_failure);

        System.out.println("key_value_to_print:" + key_value_to_print);

    }


}