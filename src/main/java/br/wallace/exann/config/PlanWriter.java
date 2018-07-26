package br.wallace.exann.config;

import br.wallace.exann.annotations.ExcelColumn;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

public class PlanWriter {
    public static String createPlan(List<?> objects, String nameArq) {
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(getPath(nameArq));
            HSSFWorkbook workbook = new HSSFWorkbook();
            List<HSSFSheet> abas = new ArrayList<>();
            Map<String, Integer> column = new HashMap<>();
            Map<String, Integer> lines = new HashMap<>();


            //Cria as folhas
            getSheets(objects).stream().map(workbook::createSheet).forEach(abas::add);

            getSheets(objects).stream().forEach(x -> column.put(x, 0));

            List<HSSFRow> listCabecalhos = new ArrayList<>();
            abas.stream().map(s -> s.createRow(0)).forEach(listCabecalhos::add);

            //Cria o estilo
            HSSFCellStyle myStyle = workbook.createCellStyle();
            myStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            myStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            HSSFFont myFont = workbook.createFont();
            myFont.setBold(true);
            myStyle.setFillBackgroundColor(new HSSFColor.BLUE().getIndex());
            myStyle.setFont(myFont);

            //Cria os cabeçalhos
            for (HSSFRow cabecalho : listCabecalhos) {
                Object object = objects.get(0);
                Field[] fields = object.getClass().getDeclaredFields();
                for (Field field : fields) {
                    Annotation[] annotations = field.getDeclaredAnnotations();
                    for (Annotation annotation : annotations) {
                        if (annotation instanceof ExcelColumn) {
                            String sheet = ((ExcelColumn) annotation).sheet();
                            String name = ((ExcelColumn) annotation).name();
                            if (cabecalho.getSheet().getSheetName().equals(sheet)) {
                                int y = column.get(sheet);
                                Cell cell = cabecalho.createCell(y);
                                cell.setCellValue(name);
                                cell.setCellStyle(myStyle);
                                column.put(sheet, ++y);
                            }
                        }
                    }

                }
            }


            //Adiciona os valores as colunas
            getSheets(objects).stream().forEach(x -> lines.put(x, 1));

            int count = 1;
            for (HSSFSheet sheet : abas) {
                for (Object object : objects) {
                    getSheets(objects).stream().forEach(x -> column.put(x, 0));
                    int y = lines.get(sheet.getSheetName());
                    HSSFRow row = sheet.createRow(y);
                    Field[] fields = object.getClass().getDeclaredFields();
                    for (Field field : fields) {
                        Annotation[] annotations = field.getDeclaredAnnotations();
                        for (Annotation annotation : annotations) {
                            if (annotation instanceof ExcelColumn) {
                                String sheetAnnotation = ((ExcelColumn) annotation).sheet();
                                if (sheet.getSheetName().equals(sheetAnnotation)) {
                                    int k = column.get(sheetAnnotation);
                                    field.setAccessible(true);
                                    String txt = String.valueOf(field.get(object));
                                    if (Objects.nonNull(txt)) {
                                        row.createCell(k).setCellValue(txt);
                                        column.put(sheetAnnotation, ++k);
                                    }
                                }

                            }

                        }

                    }
                    lines.put(sheet.getSheetName(), ++y);
                }

            }

            workbook.write(fos);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                fos.flush();
                fos.close();
            } catch (Exception e) {
                System.out.println(">>>>  ERRO - ERRO AO FINALIZAR PLANILHA");
                System.out.println(">>>>  " + e.getMessage());
            }
        }
        return getPath(nameArq);
    }

    private static Set<String> getSheets(List<?> objects) {
        Set<String> sheets = new HashSet<>();
        for (Object object : objects) {
            Field[] fields = object.getClass().getDeclaredFields();
            for (Field field : fields) {
                Annotation[] annotations = field.getDeclaredAnnotations();
                for (Annotation annotation : annotations) {
                    if (annotation instanceof ExcelColumn) {
                        String sheet = ((ExcelColumn) annotation).sheet();
                        sheets.add(sheet);
                    }
                }

            }
        }
        return sheets;
    }

    private static String getPath(String name) {
        String caminho = System.getProperty("user.dir") + "/" + name;

        if (System.getProperty("os.name").toUpperCase().contains("WIN")) {
            return caminho.replaceAll("/", "\\");
        } else {
            return caminho;
        }
    }
}
