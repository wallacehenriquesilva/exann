package br.wallace.exann.config;


import br.wallace.exann.annotations.ExcelColumn;
import br.wallace.exann.annotations.ExcelPlan;
import br.wallace.exann.util.PlanilhaUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class PlanReader {
    public static List<?> readPlan(Class clas) throws IllegalAccessException, InstantiationException, IOException, InvalidFormatException, ParseException {
        Workbook workbook = null;

        for (Annotation annotation : clas.getAnnotations()) {
            if (annotation instanceof ExcelPlan) {
                String path = ((ExcelPlan) annotation).path();
                workbook = WorkbookFactory.create(new File(path));
            }
        }

        return Objects.nonNull(workbook) ? read(clas, workbook) : null;
    }


    public static List<?> readPlan(String path, Class clas) throws IllegalAccessException, InstantiationException, IOException, InvalidFormatException, ParseException {
        Workbook workbook = WorkbookFactory.create(new File(path));
        return read(clas, workbook);
    }

    private static List<?> read(Class clas, Workbook workbook) throws IllegalAccessException, InstantiationException, ParseException {
        List<Object> objects = new ArrayList<>();
        Map<Integer, String> columns = new HashMap<>();

        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
        int coluna = 0;


        for (Cell cell : sheet.getRow(0)) {
            String cellValue = dataFormatter.formatCellValue(cell).trim();
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

                if (cell.getCellTypeEnum() != CellType.STRING && (cellValue.matches("[0-2]+\\/[0-3]?[0-9]?\\/\\d+")
                        || HSSFDateUtil.isCellDateFormatted(cell))) {
                    cellValue = new SimpleDateFormat("dd/MM/yyyy").format(new SimpleDateFormat("MM/dd/yy").parse(cellValue));
                }

                Field[] fields = clas.getDeclaredFields();
                for (Field field : fields) {
                    Annotation[] annotations = field.getDeclaredAnnotations();
                    for (Annotation annotation : annotations) {
                        if (annotation instanceof ExcelColumn) {
                            if (Objects.isNull(columns.get(coluna))) {
                                continue;
                            }

                            String nome = ((ExcelColumn) annotation).name();
                            boolean caseSensitive = ((ExcelColumn) annotation).caseSensitive();
                            int p = ((ExcelColumn) annotation).porcentagemSimilaridade();

                            double porcentagem = new Float(p) / 100;
                            boolean columa = false;


                            double similaridade = PlanilhaUtil.similarity(columns.get(coluna), nome.trim());


                            if (caseSensitive) {
                                if (columns.get(coluna).equals(nome.trim())) {
                                    columa = true;
                                }
                            } else {
                                if (PlanilhaUtil.normalizaTexto(columns.get(coluna))
                                        .equalsIgnoreCase(PlanilhaUtil.normalizaTexto(nome.trim()))) {
                                    columa = true;
                                }
                            }

                            if (porcentagem >= similaridade) {
                                columa = true;
                                System.out.println("Similar");
                            }

                            if (columa) {
                                boolean accessible = field.isAccessible();
                                field.setAccessible(true);
                                try {
                                    Object o = transform(cellValue, field.getType());
                                    field.set(object, o);
                                    field.setAccessible(accessible);
                                } catch (ParseException e) {
                                    e.printStackTrace();
                                }

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


    private static Object transform(String value, Class classe) throws ParseException {
        if (classe.getName().toUpperCase().contains("STRING") || classe.getName().toUpperCase().contains("FIELD")) {
            return value;
        } else if (classe.getName().toUpperCase().contains("DATE")) {
            if (value.contains("-")) {
                return new SimpleDateFormat("yyyy-MM-dd").parse(value);
            } else if (value.contains("/")) {
                return new SimpleDateFormat("dd/MM/yyyy").parse(value);
            }
        } else if (classe.getName().toUpperCase().contains("TIME")) {
            if (value.contains("-")) {
                return new Timestamp(new SimpleDateFormat("yyyy-MM-dd").parse(value).getTime());
            } else if (value.contains("/")) {
                return new Timestamp(new SimpleDateFormat("dd/MM/yyyy").parse(value).getTime());
            }
        } else if (classe.getName().toUpperCase().contains("INT")) {
            return Integer.parseInt(value);
        } else if (classe.getName().toUpperCase().contains("LONG")) {
            return Long.valueOf(value);
        }
        return classe.cast(value);
    }

}

