package com.excel.utilities;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class ExcelUtil {

    private Workbook workbook;

    private Sheet sheet;

    public ExcelUtil(String filePath) throws Exception {
        File excelFile = new File(filePath);
        if (excelFile.exists()) {
            System.out.println("File Exists");
            importWorkbook(filePath);
        } else {
            System.out.println("File does not Exists.");
            createExcelFile(filePath);
            importWorkbook(filePath);
        }
    }

    private void importWorkbook(String filePath) throws Exception {
        System.out.println("Importing existing workbook.");
        File excelFile = new File(filePath);
        FileInputStream inputStream = new FileInputStream(excelFile);
        String fileExtension = filePath.substring(filePath.length() - 4, filePath.length());
        if (fileExtension.contains("xlsx")) {
            System.out.println("XSSF");
            workbook = new XSSFWorkbook(inputStream);
        } else {
            System.out.println("HSSF");
            workbook = new HSSFWorkbook(inputStream);
        }
    }

    private void createExcelFile(String filePath, String sheetName) throws Exception {
        System.out.println("Creating new excel workbook with a sheet named default");
        String[] fileNameSplit = filePath.split("/");
        String fileName = fileNameSplit[fileNameSplit.length - 1];
        if (fileName.contains("xlsx")) {
            System.out.println("XSSF");
            workbook = new XSSFWorkbook();
        } else {
            System.out.println("HSSF");
            workbook = new HSSFWorkbook();
        }
        addSheetToWorkbook(sheetName);
        exportSheetToWorkbook(filePath);
    }

    private void validateSheet() throws Exception {
        if (sheet == null) {
            throw new Exception("Sheet is not initialized. Select the sheet using setActiveSheet method. If already set, check the sheet name.");
        }
    }

    private void createExcelFile(String filePath) throws Exception {
        createExcelFile(filePath, "default");
    }

    public void addSheetToWorkbook(String sheetName) {
        workbook.createSheet(sheetName);
    }

    public int getAllSheetCount() {
        return workbook.getNumberOfSheets();
    }

    public String[] getAllSheetsNames() {
        ArrayList<String> sheetNames = new ArrayList<>();
        for (int i = 0; i < getAllSheetCount(); i++) {
            sheetNames.add(workbook.getSheetName(i));
        }
        return sheetNames.toArray(new String[sheetNames.size()]);
    }

    public Sheet getActiveSheet() throws Exception {
        validateSheet();
        return sheet;
    }

    public String getActiveSheetName() throws Exception {
        return getActiveSheet().getSheetName();
    }

    public void setActiveSheet(String sheetName) throws Exception {
        sheet = workbook.getSheet(sheetName);
        validateSheet();
    }

    public void setActiveSheet(int sheetIndex) throws Exception {
        sheet = workbook.getSheetAt(sheetIndex);
    }

    public Row getRowOfActiveSheet(int rowIndex) throws Exception {
        validateSheet();
        return sheet != null ? sheet.getRow(rowIndex) : null;
    }

    public int getActiveSheetRowsCount() throws Exception {
        validateSheet();
        return sheet.getLastRowNum()+1;
    }

    public int getCellsCount(int rowIndex) throws Exception {
        return getRowOfActiveSheet(rowIndex) != null ? getRowOfActiveSheet(rowIndex).getLastCellNum() : null;
    }

    public Cell getCell(int rowIndex, int cellIndex) throws Exception {
        return getRowOfActiveSheet(rowIndex) != null ? getRowOfActiveSheet(rowIndex).getCell(cellIndex) : null;
    }

    public Cell getCell(Row row, int cellIndex) {
        return row.getCell(cellIndex) == null ? null : row.getCell(cellIndex);
    }

    public Object getCellValue(int rowIndex, int cellIndex) throws Exception {
        Cell currentCell = getCell(rowIndex, cellIndex);
        Object cellValue = null;
        if (currentCell != null) {
            if (currentCell.getCellType() == CellType.NUMERIC) {
                cellValue = (Double) currentCell.getNumericCellValue();
            } else if (currentCell.getCellType() == CellType.STRING) {
                cellValue = currentCell.getStringCellValue();
            } else if (currentCell.getCellType() == CellType.BOOLEAN) {
                cellValue = (Boolean) currentCell.getBooleanCellValue();
            }
        }
        return cellValue;
    }

    public Object getCellValue(Cell cell) throws Exception {
        Object cellValue = null;
        if (cell != null) {
            if (cell.getCellType() == CellType.NUMERIC) {
                cellValue = (Double) cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                cellValue = cell.getStringCellValue();
            } else if (cell.getCellType() == CellType.BOOLEAN) {
                cellValue = (Boolean) cell.getBooleanCellValue();
            }
        }
        return cellValue;
    }

    public void exportSheetToWorkbook(String filePath) throws Exception {
        if (filePath.equals("") || filePath.trim().length() == 0) {
            throw new Exception("filePath cannot be empty value. Please pass valid string value.");
        }
        FileOutputStream fileOut = new FileOutputStream(filePath);
        workbook.write(fileOut);
    }

    public int getCellIndexByText(String text, int rowIndex, boolean exactMatch) throws Exception {
        int cellIndex = -1;
        int cellCount = getCellsCount(rowIndex);
        for(int i=0;i<cellCount;i++){
            Cell c = getCell(rowIndex, i);
            String value = getCellValue(c) instanceof String ? ((String) getCellValue(c)) : null;
            if(value != null){
                if(exactMatch){
                    cellIndex = value.equals(text) ? c.getColumnIndex() : -1;
                }else {
                    cellIndex = value.contains(text) ? c.getColumnIndex() : -1;
                }
            }
            if(cellIndex != -1){
                return cellIndex;
            }
        }
        return cellIndex;
    }

}
