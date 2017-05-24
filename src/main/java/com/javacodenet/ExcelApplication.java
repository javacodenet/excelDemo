package com.javacodenet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelApplication {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        List<Employee> employees = new ArrayList<Employee>();
        for (int i = 0; i < 10; i++) {
            Employee employee = new Employee("empId" + i, "empName" + i, "dept" + i, 1000 + i);
            employees.add(employee);
        }
        //Create .xlsx file
        ExcelHelper.createExcelSheet(employees, true);
        //Read .xlsx file
        ExcelHelper.readExcelSheet(true);
        //Create .xls file
        ExcelHelper.createExcelSheet(employees, false);
        //Read .xls file
        ExcelHelper.readExcelSheet(false);
    }
}
