package com.demo;

import com.demo.components.ApplicationWriter;
import com.demo.model.*;
import org.apache.camel.RoutesBuilder;
import org.apache.camel.builder.RouteBuilder;
import org.apache.camel.component.file.GenericFile;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@SpringBootApplication
public class DemoApplication {

    @Autowired
    private ApplicationWriter writter;

    @Bean
    RoutesBuilder routes() {
        return new RouteBuilder() {
            @Override
            public void configure() throws Exception {
                from("file://{{user.home}}/Desktop/files/text.txt")
                        .transform()
                        .body(GenericFile.class, gf -> processFile(gf, writter));
            }
        };
    }

    private static Object processFile(GenericFile gf, ApplicationWriter writter) {
        try {
            File actualFile = File.class.cast(gf.getFile());
            assert actualFile != null;
            List<Person> list = new ArrayList<Person>();
            FileInputStream excelFile = new FileInputStream(actualFile);
            Workbook workbook = new XSSFWorkbook(excelFile);

            // Retrieving the number of sheets in the Workbook
            System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

            Iterator<XSSFSheet> iterator = ((XSSFWorkbook) workbook).iterator();
            int index = -1;
            while (iterator.hasNext()) {
                index += 1;
                Sheet sheet = iterator.next();
                String sheetName = sheet.getSheetName();
                System.out.println("=> " + sheetName + "  =>  " + index);
                Iterator<Row> rowIterator = sheet.iterator();
                rowIterator.next();

                while (rowIterator.hasNext()) {

                    DataFormatter formatter = new DataFormatter();

                    Row currentRow = rowIterator.next();
                    String name = formatter.formatCellValue(currentRow.getCell(0));
                    String gender = formatter.formatCellValue(currentRow.getCell(1));
                    String age = formatter.formatCellValue(currentRow.getCell(2));
                    Person report = new Person();
                    report.setName(name);
                    report.setGender(gender);
                    report.setAge(age);
                    if (isStringInt(age))
                    {
                        int _age = Integer.parseInt(age);
                        if (_age <= 12) {
                            report.setResult(gender.toLowerCase() == "m" ? "Boy": "Girl");
                        } else if (_age > 12 && _age <=19) {
                            report.setResult("Teenager");
                        } else if (_age > 19 && _age <=25) {
                            report.setResult(gender.toLowerCase() == "m" ? "Young Man": "Young Lady");
                        } else if (_age > 25 && _age <=69) {
                            report.setResult(gender.toLowerCase() == "m" ? "Man": "Woman");
                        } else if (_age > 70) {
                            report.setResult(gender.toLowerCase() == "m" ? "Grandpa": "Grandma");
                        }
                    } else {
                        report.setResult("Age is not an integer");
                    }

                    list.add(report);
                }
            }
            writter.SaveData(list);
            System.out.println("======>   DONE READING AND SAVED TO OUTPUT FILE    <=======");
            return null;
        } catch (Exception e) {
            System.out.println(e.getMessage());
            return null;
        }
    }

    public static boolean isStringInt(String s)
    {
        try
        {
            Integer.parseInt(s);
            return true;
        } catch (NumberFormatException ex)
        {
            return false;
        }
    }

    public static void main(String[] args) {
        SpringApplication.run(DemoApplication.class, args);
    }
}
