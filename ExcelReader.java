package org.excelreader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    public <T> List<T> readExcel(String filePath, Class<T> clazz) throws Exception {
        List<T> resultList = new ArrayList<>();


        try (InputStream fis = getClass().getClassLoader().getResourceAsStream(filePath)) {
            if (fis == null) {
                throw new IllegalArgumentException("File not found: " + filePath);
            }

            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            // Get headers from the first row
            Row headerRow = sheet.getRow(0);
            int colCount = headerRow.getPhysicalNumberOfCells();

            // Iterate over rows (excluding the header)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                T instance = clazz.getDeclaredConstructor().newInstance();

                for (Field field : clazz.getDeclaredFields()) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                        String headerName = annotation.header();

                        for (int j = 0; j < colCount; j++) {
                            Cell cell = row.getCell(j);
                            String cellHeader = headerRow.getCell(j).getStringCellValue();

                            if (headerName.equalsIgnoreCase(cellHeader)) {
                                field.setAccessible(true);

                                // Set field value based on the cell type
                                switch (cell.getCellType()) {
                                    case STRING:
                                        field.set(instance, cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        if (field.getType() == int.class) {
                                            field.set(instance, (int) cell.getNumericCellValue());
                                        } else {
                                            field.set(instance, cell.getNumericCellValue());
                                        }
                                        break;
                                    // Add other types (BOOLEAN, etc.) as needed
                                }
                            }
                        }
                    }
                }

                resultList.add(instance);
            }
        } catch (IOException e) {
            e.printStackTrace();
            throw new Exception("Error reading the Excel file.");
        }

        return resultList;
    }
}
