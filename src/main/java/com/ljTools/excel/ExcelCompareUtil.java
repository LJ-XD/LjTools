package com.ljTools.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelCompareUtil {

    private ExcelCompareUtil() {
    }

    public static <T> List<T> readExcel(String filePath, Class<T> clazz) throws IOException {
        List<T> objectList = new ArrayList<>();
        FileInputStream fileInputStream = new FileInputStream(filePath);

        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }

            T object = createObjectFromRow(row, clazz);
            if (object != null) {
                objectList.add(object);
            }
        }

        workbook.close();
        fileInputStream.close();

        return objectList;
    }

    private static <T> T createObjectFromRow(Row row, Class<T> clazz) {
        try {
            Constructor<T> constructor = clazz.getDeclaredConstructor();
            constructor.setAccessible(true);
            T object = constructor.newInstance();

            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                field.setAccessible(true);

                Cell cell = row.getCell(i);
                if (cell != null) {
                    setValueForField(field, object, cell);
                }
            }

            return object;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private static void setValueForField(Field field, Object object, Cell cell) throws IllegalAccessException {
        Class<?> fieldType = field.getType();

        if (fieldType == String.class) {
            field.set(object, cell.getStringCellValue());
        } else if (fieldType == int.class || fieldType == Integer.class) {
            field.set(object, (int) cell.getNumericCellValue());
        } else if (fieldType == double.class || fieldType == Double.class) {
            field.set(object, cell.getNumericCellValue());
        } else if (fieldType == boolean.class || fieldType == Boolean.class) {
            field.set(object, cell.getBooleanCellValue());
        }
    }
}
