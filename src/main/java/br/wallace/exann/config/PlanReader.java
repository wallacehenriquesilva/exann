package br.wallace.exann.config;


import br.wallace.exann.annotations.ExcelColumn;
import br.wallace.exann.annotations.ExcelPlan;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PlanReader {
    public static List<?> readPlan(Class clas) throws IllegalAccessException, InstantiationException, IOException, InvalidFormatException, NoSuchFieldException {
        List<Object> objects = new ArrayList<>();
        Map<Integer, String> columns = new HashMap<>();

        Workbook workbook = null;

        for (Annotation annotation : clas.getAnnotations()) {
            if (annotation instanceof ExcelPlan) {
                String path = ((ExcelPlan) annotation).path();
                workbook = WorkbookFactory.create(new File(path));
            }
        }

        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
        int coluna = 0;


        for (Cell cell : sheet.getRow(0)) {
            String cellValue = dataFormatter.formatCellValue(cell);
            columns.put(coluna, cellValue);
            coluna++;
        }

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            Object object = clas.newInstance();
            coluna = 0;
            for (Cell cell : row) {
                String cellValue = dataFormatter.formatCellValue(cell);

                Field[] fields = clas.getDeclaredFields();
                for (Field field : fields) {
                    Annotation[] annotations = field.getDeclaredAnnotations();
                    for (Annotation annotation : annotations) {
                        if (annotation instanceof ExcelColumn) {
                            String nome = ((ExcelColumn) annotation).name();
                            if (columns.get(coluna).equals(nome)) {
                                boolean accessible = field.isAccessible();
                                field.setAccessible(true);
                                field.set(object, cellValue);
                                field.setAccessible(accessible);
                            }
                        }
                    }

                }
                coluna++;
            }
            objects.add(object);
        }
        return objects;
    }
}

