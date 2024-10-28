package org.excelreader;

import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception {
        ExcelReader reader = new ExcelReader();
        List<Employee> employees = reader.readExcel("employees.xlsx", Employee.class);

        for (Employee emp : employees) {
            System.out.println(emp.getName() + " - " + emp.getAge() + " - " + emp.getDepartment());
        }
    }
}
