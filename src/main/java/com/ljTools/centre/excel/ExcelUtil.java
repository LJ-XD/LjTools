package com.ljTools.centre.excel;

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

/**
 * ExcelUtil
 */
public class ExcelUtil {

    private ExcelUtil() {
    }

    /**
     * 简单excel转对象
     *
     * @param filePath 文件路径
     * @param clazz    对象类型
     * @return 对象列表
     * @throws IOException io异常
     */
    public static <T> List<T> excelToList(String filePath, Class<T> clazz) throws IOException {
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

    /**
     * 一行数据转对象
     *
     * @param row   行
     * @param clazz 对象类型
     * @return 对象类型
     */
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

    /**
     * 设置字段值
     *
     * @param field  字段
     * @param object 对象
     * @param cell   单元格
     * @throws IllegalAccessException 非法访问异常
     */
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
