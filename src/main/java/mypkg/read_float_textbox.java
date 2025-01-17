package mypkg;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class read_float_textbox {
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
}
