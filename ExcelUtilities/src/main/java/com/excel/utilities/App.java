package com.excel.utilities;
import java.util.stream.Stream;

public class App {

    public static void main(String[] args) {
        try {
            ExcelUtil util = new ExcelUtil("ExcelUtilities/src/main/resources/testdemoadv.xlsx");
            util.setActiveSheet("Sheet1");
            System.out.println(util.getActiveSheetRowsCount());
            util.setActiveSheet("Sheet2");
            System.out.println(util.getAllSheetCount());
            Stream.of(util.getAllSheetsNames()).forEach(System.out::println);
            System.out.println(util.getActiveSheetRowsCount());
            System.out.println(util.getCellsCount(0));
            System.out.println(util.getCellValue(0,1));
            System.out.println(util.getCellValue(1,1));
            util.setActiveSheet("Sheet3");
            System.out.println(util.getActiveSheetRowsCount());
            System.out.println(util.getCellValue(0,0));
            System.out.println(util.getCellValue(0,1));
            System.out.println(util.getCellValue(0,2));
            System.out.println(util.getCellValue(1,0));
            System.out.println(util.getCellValue(1,1));
            System.out.println(util.getCellValue(1,2));
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }
}
