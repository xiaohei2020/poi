package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SpringBootTest(classes = PoiApplication.class)
@RunWith(SpringRunner.class)
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        System.out.println("true = " + true);
    }


    @Test
    public void poiIn() {
        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("C:\\Users\\zhangjia\\Desktop\\后期项目\\aa.xls")));
            HSSFSheet sheet = workbook.getSheet("测试");
            int lastRowNum = sheet.getLastRowNum();
            System.out.println(lastRowNum);
            for (int i = 0; i < lastRowNum; i++) {

                HSSFRow row = sheet.getRow(i + 1);
                HSSFCell cell = row.getCell(0);
                String value = cell.getStringCellValue();

                HSSFCell cell1 = row.getCell(1);
                String name = cell1.getStringCellValue();

                HSSFCell cell2 = row.getCell(2);
                double value1 = cell2.getNumericCellValue();

                HSSFCell cell3 = row.getCell(3);
                Date bir = cell3.getDateCellValue();

                System.out.println(value + " " + name + " " + value1 + " " + bir);

            }


        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Test
    public void testInGrid() throws IOException {
        File file = new File("C:\\Users\\zhangjia\\Desktop\\后期项目\\aa.xls");
        /**
         * HSSFSheet sheet = workbook.getSheet("测试");
         *             int lastRowNum = sheet.getLastRowNum();
         */
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = hssfWorkbook.getSheet("测试");
        int lastRowNum = sheet.getLastRowNum();
        System.out.println("lastRowNum = " + lastRowNum);
        for (int i = 0; i < lastRowNum; i++) {
            HSSFRow row = sheet.getRow(i + 1);
            HSSFCell cell = row.getCell(0);
            String value = cell.getStringCellValue();

            HSSFCell cell1 = row.getCell(1);
            String name = cell1.getStringCellValue();

            HSSFCell cell2 = row.getCell(2);
            double value1 = cell2.getNumericCellValue();

            HSSFCell cell3 = row.getCell(3);
            Date bir = cell3.getDateCellValue();

            System.out.println(value + " " + name + " " + value1 + " " + bir);
        }

    }

    @Test
    public void testOutGrid() {
        //创建工作簿
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //获取单元格式.用来设置字体居中
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        //对于字体的调整,可以和单元格样式同时设置
        //创建字体样式
        HSSFFont font = hssfWorkbook.createFont();
        font.setColor(Font.SS_SUB);
        font.setBold(true);
        font.setFontName("楷体");
        font.setItalic(true);

        //将字体样式设置进hssfCelleStyle
        hssfCellStyle.setFont(font);

        //获取单元格样式
        HSSFCellStyle cellStyle = hssfWorkbook.createCellStyle();
        //获取时间格式
        HSSFDataFormat format = hssfWorkbook.createDataFormat();
        short format1 = format.getFormat("yyyy-MM-dd");
        //把时间格式设置进单元格样式中
        cellStyle.setDataFormat(format1);

        //创建工作表
        HSSFSheet sheet = hssfWorkbook.createSheet("测试");
        //创建行
        HSSFRow row = sheet.createRow(0);
        //设置列宽
        sheet.setColumnWidth(12, 60 * 256);

        String[] strings = {"主键", "姓名", "年龄", "生日"};
        for (int i = 0; i < strings.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(hssfCellStyle);
            cell.setCellValue(strings[i]);
        }

        //创建单元格
        //HSSFCell cell=row.createCell(0)

        List<User> users = new ArrayList<>();
        User user = new User("1", "嫖老师", 69, new Date());
        User user1 = new User("2", "刘浩", 30, new Date());
        User uesr2 = new User("3", "许婧辉", 8, new Date());
        User user3 = new User("4", "大飞", 4, new Date());
        users.add(user);
        users.add(user1);
        users.add(uesr2);
        users.add(user3);

        for (int i = 0; i < users.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            row1.createCell(0).setCellValue(users.get(i).getId());
            row1.createCell(1).setCellValue(users.get(i).getName());
            row1.createCell(2).setCellValue(users.get(i).getAge());
            HSSFCell cell = row1.createCell(3);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(users.get(i).getBir());
        }

        try {
            hssfWorkbook.write(new FileOutputStream(new File("C:\\Users\\zhangjia\\Desktop\\后期项目\\aa.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
