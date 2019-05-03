package com.lampartvn.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class UpdateExcelDemo {

    public static void main(String[] args) throws IOException {
        File file = new File(Employee.getFilePath());

        // read file xls
        FileInputStream inputStream = new FileInputStream(file);

        // create workbook for xls file
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        // get first sheet of workbook
        HSSFSheet sheet = workbook.getSheetAt(0);

        HSSFCell cell = sheet.getRow(1).getCell(2);
        cell.setCellValue(cell.getNumericCellValue() * 2);

        cell = sheet.getRow(2).getCell(2);
        cell.setCellValue(cell.getNumericCellValue() * 2);

        cell = sheet.getRow(3).getCell(2);
        cell.setCellValue(cell.getNumericCellValue() * 2);

        HSSFRow row = sheet.createRow(4);
        cell = row.createCell(2, CellType.FORMULA);
        cell.setCellFormula("SUM(C2:C4)");

        inputStream.close();

        // write file
        FileOutputStream out = new FileOutputStream(file);
        workbook.write(out);
        out.close();
    }

}
