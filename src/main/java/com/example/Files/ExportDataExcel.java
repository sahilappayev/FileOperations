package com.example.Files;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
@RequestMapping("/students")
public class ExportDataExcel {

    static class Student {

        private String name;
        private String surname;
        private Integer age;
        private Date createdAt;

        public Student(String name, String surname, Integer age, Date createdAt) {
            this.name = name;
            this.surname = surname;
            this.age = age;
            this.createdAt = createdAt;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getSurname() {
            return surname;
        }

        public void setSurname(String surname) {
            this.surname = surname;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }

        public Date getCreatedAt() {
            return createdAt;
        }

        public void setCreatedAt(Date createdAt) {
            this.createdAt = createdAt;
        }
    }

    @GetMapping(produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public byte[] generateExcel() {
        byte[] res = generateExcelFile();
        return res;
    }

    private byte[] generateExcelFile() {
        List<Student> students = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            Student student = new Student("Name " + i, "Surname " + i, 15 + i, new Date());
            students.add(student);
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Students " + LocalDate.now());
        writeHeaderLine(workbook, sheet);
        writeDataLines(students, workbook, sheet);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            return bos.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private void writeHeaderLine(XSSFWorkbook workbook, XSSFSheet sheet) {
        Row headerRow = sheet.createRow(0);
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(3, 5000);
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerCellStyle.setBorderRight(BorderStyle.MEDIUM);
        headerCellStyle.setBorderTop(BorderStyle.MEDIUM);
        Font font = workbook.createFont();
        font.setColor(Font.COLOR_RED);
        font.setBold(true);
        font.setFontHeightInPoints((short) 13);
        headerCellStyle.setFont(font);

        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellStyle(headerCellStyle);
        headerCell.setCellValue("Name");

        headerCell = headerRow.createCell(1);
        headerCell.setCellStyle(headerCellStyle);
        headerCell.setCellValue("Surname");

        headerCell = headerRow.createCell(2);
        headerCell.setCellStyle(headerCellStyle);
        headerCell.setCellValue("Age");

        headerCell = headerRow.createCell(3);
        headerCell.setCellStyle(headerCellStyle);
        headerCell.setCellValue("Date");
    }

    private void writeDataLines(List<Student> students, XSSFWorkbook workbook, XSSFSheet sheet) {
        int rowCount = 1;

        for (Student student : students) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;

            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(student.name);

            cell = row.createCell(columnCount++);
            cell.setCellValue(student.surname);

            cell = row.createCell(columnCount++);
            cell.setCellValue(student.age);

            CellStyle cellStyle = workbook.createCellStyle();
            CreationHelper creationHelper = workbook.getCreationHelper();
            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm"));
            cell = row.createCell(columnCount);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(student.getCreatedAt());

        }
    }


}
