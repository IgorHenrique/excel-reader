package org.excelreader;

public class Employee {

    @ExcelColumn(header = "Name")
    private String name;

    @ExcelColumn(header = "Age")
    private int age;

    @ExcelColumn(header = "Department")
    private String department;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }
}
