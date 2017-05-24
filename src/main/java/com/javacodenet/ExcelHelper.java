package com.javacodenet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelHelper {

    private static final String XLSX_FILE_NAME = "tmp\\employees.xlsx";
    private static final String XLS_FILE_NAME = "tmp\\employees.xls";


    public static void createExcelSheet(List<Employee> employees, boolean isXlsx) {
        System.out.println(isXlsx ? "Creating .xlsx file" : "Creating .xls file");

        //Create workbook, XSSFWorkbook for xlsx files and HSSFWorkbook for xls files
        Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();
        //Create Worksheet
        Sheet employeeDetailsSheet = workbook.createSheet("Employee details");

        //Create Header ROw
        Row rowHeader = employeeDetailsSheet.createRow(0);
        CellStyle boldCellStyle = getBoldCellStyle(workbook);
        createHeaderCell(rowHeader, 0, boldCellStyle, "EmployeeId");
        createHeaderCell(rowHeader, 1, boldCellStyle, "EmployeeName");
        createHeaderCell(rowHeader, 2, boldCellStyle, "Department");
        createHeaderCell(rowHeader, 3, boldCellStyle, "Salary");

        //Create data cells
        int rowNum = 1;
        for (Employee employee : employees) {
            Row row = employeeDetailsSheet.createRow(rowNum++);
            Cell empId = row.createCell(0);
            empId.setCellValue(employee.getEmployeeId());
            Cell empName = row.createCell(1);
            empName.setCellValue(employee.getEmployeeName());
            Cell dept = row.createCell(2);
            dept.setCellValue(employee.getDepartment());
            Cell salary = row.createCell(3);
            salary.setCellValue(employee.getSalary());
        }

        //create file and directory if not present
        File f = new File(isXlsx ? XLSX_FILE_NAME : XLS_FILE_NAME);

        if (!f.exists()) {
            //If directories are not available then create it
            File parent_directory = f.getParentFile();
            if (parent_directory != null) {
                parent_directory.mkdirs();
            }
            try {
                f.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        //write to the created file
        try {
            FileOutputStream out = new FileOutputStream(f, false);
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Done");
    }

    private static void createHeaderCell(Row rowHeader, int i, CellStyle boldCellStyle, String employeeId) {
        Cell empIdHeader = rowHeader.createCell(i);
        empIdHeader.setCellStyle(boldCellStyle);
        empIdHeader.setCellValue(employeeId);
    }

    private static CellStyle getBoldCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = getBoldFont(workbook);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private static Font getBoldFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setFontName("Arial");
        font.setColor(IndexedColors.GREEN.getIndex());
        font.setBold(true);
        font.setItalic(true);
        return font;
    }

    public static void readExcelSheet(boolean isXlsx) {
        try {
            FileInputStream excelFile = new FileInputStream(new File(isXlsx ? XLSX_FILE_NAME : XLS_FILE_NAME));
            Workbook workbook = isXlsx ? new XSSFWorkbook(excelFile) : new HSSFWorkbook(excelFile);

            Sheet workSheet = workbook.getSheetAt(0);

            for (Row currentRow : workSheet) {
                for (Cell currentCell : currentRow) {
                    printCellContents(currentCell);
                }
                System.out.println();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void printCellContents(Cell currentCell) {
        switch (currentCell.getCellTypeEnum()) {
            case STRING:
                System.out.print(currentCell.getRichStringCellValue().getString() + "|");
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(currentCell)) {
                    System.out.print(currentCell.getDateCellValue() + "|");
                } else {
                    System.out.print(currentCell.getNumericCellValue() + "|");
                }
                break;
            case BOOLEAN:
                System.out.print(currentCell.getBooleanCellValue() + "|");
                break;
            case FORMULA:
                System.out.print(currentCell.getCellFormula() + "|");
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.println();
        }
    }
}
