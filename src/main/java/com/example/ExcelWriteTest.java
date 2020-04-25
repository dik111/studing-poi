package com.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Desription:
 *
 * @ClassName ExcelWriteTest
 * @Author Zhanyuwei
 * @Date 2020/4/25 14:07
 * @Version 1.0
 **/
public class ExcelWriteTest {

    String PATH = "E:\\java\\studing-poi\\";

    @Test
    public void testWrite03() throws Exception {
        // 1.创建工作簿
        Workbook workbook = new HSSFWorkbook();

        // 2.创建一个工作簿
        Sheet sheet = workbook.createSheet("统计表");
        // 3.创建一个行
        Row row1 = sheet.createRow(0);
        //创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //创建一个单元格
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("666");

        // 第二行(2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "统计表03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("统计表03 生成完毕");
    }

    @Test
    public void testWrite07() throws Exception {
        // 1.创建工作簿
        Workbook workbook = new XSSFWorkbook();

        // 2.创建一个工作簿
        Sheet sheet = workbook.createSheet("统计表");
        // 3.创建一个行
        Row row1 = sheet.createRow(0);
        //创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //创建一个单元格
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("666");

        // 第二行(2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 生成一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "统计表07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("统计表07 生成完毕");
    }
}
