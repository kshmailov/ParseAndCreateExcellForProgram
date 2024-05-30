package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class ParseAndCreateExcell {
    public static int rowNum = 0;
    public static FileOutputStream out;
    public static XSSFWorkbook workbookOut= new XSSFWorkbook();
    public static XSSFSheet sheet = workbookOut.createSheet("table");
    protected static void parseAndCreateExcell(String path, int count) throws IOException {
        FileInputStream fis = new FileInputStream(path);
        out = new FileOutputStream("data/Table.xlsx");
        Workbook workbook;
        try {
            workbook = new XSSFWorkbook(
                    fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ArrayList<Sheet> sheets = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            sheets.add(workbook.getSheetAt(i));
        }
        if (rowNum==0){
            XSSFRow row=sheet.createRow(rowNum);
            XSSFCell cell1 = row.createCell(0);
            sheet.autoSizeColumn(0);
            cell1.setCellValue("Ступень КПР");
            XSSFCell cell2 = row.createCell(1);
            sheet.autoSizeColumn(1);
            cell2.setCellValue("Переток");
            XSSFCell cell3 = row.createCell(2);
            sheet.autoSizeColumn(2);
            cell3.setCellValue("Уставка УВ");
            XSSFCell cell4 = row.createCell(3);
            sheet.autoSizeColumn(3);
            cell4.setCellValue("Факт");
            XSSFCellStyle style1= workbookOut.createCellStyle();
            style1.setAlignment(HorizontalAlignment.CENTER);
            style1.setVerticalAlignment(VerticalAlignment.CENTER);
            style1.setBorderBottom(BorderStyle.THIN);
            style1.setBorderTop(BorderStyle.THIN);
            style1.setBorderLeft(BorderStyle.THIN);
            style1.setBorderRight(BorderStyle.THIN);
            style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell1.setCellStyle(style1);
            cell2.setCellStyle(style1);
            cell3.setCellStyle(style1);
            cell4.setCellStyle(style1);
        }
        XSSFCellStyle style = workbookOut.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        for (Sheet list : sheets) {
            String kpr = list.getRow(0).getCell(3).getStringCellValue();
            String sezon = list.getRow(1).getCell(0).getStringCellValue();
            String apnu = "";
            if (kpr.contains("Юг")) {
                apnu = "АПНУ Юг";
            } else if (kpr.contains("Кубанское")) {
                apnu = "АПНУ Кубанское";
            } else if (kpr.contains("Маныч")) {
                apnu = "АПНУ Маныч";
            }

            for (int row = 4; row<=list.getLastRowNum() ; row++) {
                rowNum++;
                String cellValue = list.getRow(row).getCell(0).getStringCellValue();
                String numberPo = Integer.toString((int) list.getRow(row).getCell(1).getNumericCellValue());
                String namePo = list.getRow(row).getCell(2).getStringCellValue();
                String po = String.join(" ", "ПО" + numberPo, namePo);
                String header = String.join("\n", apnu+" "+sezon, "Cхема: " + cellValue, po);
                XSSFRow createHeaderRow = sheet.createRow(rowNum);
                CellRangeAddress mergedRegion = new CellRangeAddress(rowNum, rowNum, 0, 3);
                sheet.addMergedRegion(mergedRegion);
                XSSFCell createHeaderRowCell = createHeaderRow.createCell(mergedRegion.getFirstColumn());
                createHeaderRowCell.setCellValue(header);
                createHeaderRowCell.setCellStyle(style);
                HashMap<String, String> kprUst = new HashMap<>(getKprUst(kpr));
                HashMap<String, String> tableStringUv = new HashMap<>(getKprUv(kpr, sezon));
                for (int i = 1; i <= 32; i++) {
                    rowNum++;
                    XSSFRow createValueRow = sheet.createRow(rowNum);
                    XSSFCell cell1 = createValueRow.createCell(0);
                    cell1.setCellValue(i);
                    cell1.setCellStyle(style);
                    XSSFCell cell2 = createValueRow.createCell(1);
                    String ust = kprUst.get(Integer.toString(i));
                    cell2.setCellValue(ust);
                    cell2.setCellStyle(style);
                    XSSFCell cell3 = createValueRow.createCell(2);
                    cell3.setCellStyle(style);
                    String uv = Integer.toString((int) list.getRow(row).getCell(i + 2).getNumericCellValue());
                    XSSFCell cell4 = createValueRow.createCell(3);
                    cell4.setCellValue("");
                    cell4.setCellStyle(style);
                    if (uv.equals("0")) {
                        cell3.setCellValue("Нет");
                    } else {
                        cell3.setCellValue(tableStringUv.get(uv));
                    }

                }
            }
        }
        fis.close();
    }
    protected static HashMap<String, String> getKprUst(String kpr){
        HashMap<String, String> kprUst = new HashMap<>();
        switch (kpr) {
            case "КПР-1_\"Юг\"": {
                int ust = 3050;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 50;
                }
                break;
            }
            case "КПР-2_\"Юг\"": {
                int ust = 2650;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 50;
                }
                break;
            }
            case "КПР-1_\"Маныч\"": {
                int ust = 1180;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 20;
                }
                break;
            }
            case "КПР-2_\"Маныч\"": {
                int ust = 560;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 20;
                }
                break;
            }
            case "КПР-3_\"Маныч\"": {
                int ust = 1540;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 20;
                }
                break;
            }
            case "КПР-4_\"Маныч\"": {
                int ust = 1120;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 20;
                }
                break;
            }
            case "КПР-1_\"Кубанское\"": {
                int ust = 2000;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 50;
                }
                break;
            }
            case "КПР-2_\"Кубанское\"": {
                int ust = 1400;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 50;
                }
                break;
            }
            case "КПР-3_\"Кубанское\"": {
                int ust = 1150;
                for (int i = 1; i <= 32; i++) {
                    kprUst.put(Integer.toString(i), Integer.toString(ust));
                    ust += 50;
                }
                break;
            }
        }
        return kprUst;
    }
    protected static HashMap<String,String> getKprUv(String kprUv, String sezon){
        HashMap<String,String> uv = new HashMap<>();
        if (kprUv.contains("Юг")&& sezon.contains("Лето")){
            uv.put("1", "ОН 1 оч. КЭ");
            uv.put("2", "ОН 1 оч. КЭ+ОН 2 оч. КЭ");
            uv.put("3", "ОН 1 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("4", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ");
            uv.put("5", "ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("6", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("7", "ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("8", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("9", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("10", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ ОН 3 оч. КЭ + ОН 5 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("11", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("12", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 4 оч. КЭ + ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("13", "ОН 1 оч. КЭ+ОН 2 оч. КЭ + ОН 3 оч. КЭ +ОН 4 оч. КЭ +ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("14", "ОН 1 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("15", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("16", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ + ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
        } else if (kprUv.contains("Юг")&& sezon.contains("Зима")) {
            uv.put("1", "ОН 1 оч. КЭ");
            uv.put("2", "ОН 1 оч. КЭ+ОН 2 оч. КЭ");
            uv.put("3", "ОН 1 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("4", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ");
            uv.put("5", "ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("6", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ");
            uv.put("7", "ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("8", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("9", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("10", "ОН 1 оч. КЭ+ОН 2 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("11", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("12", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 4 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("13", "ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("14", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("15", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("16", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("17", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("18", "ОН 1 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("19", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ + ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("20", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ + ОН 6 оч. КЭ + ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");

        }else if (kprUv.contains("Маныч")&& sezon.contains("Лето")) {
            uv.put("1", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("2", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("3", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("4", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ ");
            uv.put("5", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("6", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("7", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("8", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
        }else if (kprUv.contains("Маныч")&& sezon.contains("Зима")) {
            uv.put("1", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("2", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("3", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("4", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("5", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("6", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ");
            uv.put("7", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("8", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("9", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("10", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("11", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("12", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
            uv.put("13", "ОН 1 оч.ВЧ + ОН 2 оч. ВЧ + ОН 3 оч. ВЧ + ОН 4 оч.ВЧ + ОН 5 оч.ВЧ + ОН 6 оч.ВЧ + ОН 7 оч. ВЧ + ОН 8 оч. ВЧ + ОН 9 оч. ВЧ");
        }else if (kprUv.contains("Кубанское")&& sezon.contains("Лето")) {
            uv.put("1", "ОН 1 оч. КЭ");
            uv.put("2", "ОН 1 оч. КЭ+ ОН 3 оч. КЭ");
            uv.put("3", "ОН 1 оч. КЭ +ОН 2 оч. КЭ+ОН 3 оч. КЭ");
            uv.put("4", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 5оч. КЭ");
            uv.put("5", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("6", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("7", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("8", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ");
        }else if (kprUv.contains("Кубанское")&& sezon.contains("Зима")) {
            uv.put("1", "ОН 1 оч. КЭ");
            uv.put("2", "ОН 1 оч. КЭ+ОН 3 оч. КЭ");
            uv.put("3", "ОН 1 оч. КЭ +ОН 2 оч. КЭ+ ОН 3 оч. КЭ");
            uv.put("4", "ОН 1 оч. КЭ+ОН 2 оч. КЭ + ОН 4 оч. КЭ");
            uv.put("5", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ ОН 4 оч. КЭ");
            uv.put("6", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ");
            uv.put("7", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 6 оч. КЭ");
            uv.put("8", "ОН 1 оч. КЭ+ОН 2 оч. КЭ+ОН 3 оч. КЭ+ОН 4 оч. КЭ+ОН 5 оч. КЭ+ОН 6 оч. КЭ");
        }
        return uv;
    }
    public static void closeTable() throws IOException {
        workbookOut.write(out);
        out.close();
    }
}
