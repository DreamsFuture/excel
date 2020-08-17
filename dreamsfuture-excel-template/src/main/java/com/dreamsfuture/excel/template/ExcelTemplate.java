package com.dreamsfuture.excel.template;

import org.apache.commons.lang.text.StrSubstitutor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Description: a class for excel template
 *
 * @author colin
 * @create 2020-07-14 19:48
 */
public class ExcelTemplate {

    private SXSSFWorkbook writeWorkbook;
    private XSSFWorkbook readWorkbook;
    private XSSFSheet readSheet;

    public ExcelTemplate(String filePath) {

        try {
            readWorkbook = new XSSFWorkbook(new FileInputStream(filePath));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public ExcelTemplate(InputStream inputStream) {
        try {
            readWorkbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public String getData(int sheetIndex, int row, int column) {

        readSheet = readWorkbook.getSheetAt(sheetIndex);
        String data = "";
        Row r = readSheet.getRow(row);
        Cell c = r.getCell(column, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

        if (c == null) {
            data = "";
        } else {
            if (r.getCell(column).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                data = r.getCell(column).getStringCellValue();
            } else if (r.getCell(column).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                int intData = (int) r.getCell(column).getNumericCellValue();
                data = Integer.toString(intData);
            }
        }
        return data;
    }


    public void replaceTemplateWithSpecifiedData(int sheetIndex, Map<String, String> replaceData) throws Exception {
        readSheet = readWorkbook.getSheetAt(sheetIndex);
        if (readSheet != null) {
            this.replaceTemplateWithSpecifiedData(readSheet.getSheetName(), replaceData);
        }
    }

    public void replaceTemplateWithSpecifiedData(String sheetName, Map<String, String> replaceData) throws Exception {

        readSheet = readWorkbook.getSheet(sheetName);
        if (readSheet == null) {
            throw new Exception("sheet with name " + sheetName + " not found");
        }
        readSheet.setForceFormulaRecalculation(true);

        StrSubstitutor substitutor = new StrSubstitutor(replaceData, "${", "}");

        Iterator<Row> rowIterator = readSheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = null;
                try {
                    cellValue = cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e1) {
                        e1.printStackTrace();
                    }
                    e.printStackTrace();
                }

                if (!StringUtils.isEmpty(cellValue)) {
                    cell.setCellValue(substitutor.replace(cellValue));
                }
            }
        }
    }


    public SXSSFWorkbook writeData(String sheetName, int rowNum, int column, String content) throws IOException {

        SXSSFSheet sxssfSheet = writeWorkbook.getSheetAt(0);

        Iterator<Row> rowIterator = sxssfSheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(cell.getStringCellValue());
                if (cell.getStringCellValue().equals("1")) {
                    cell.setCellValue("10000");
                }
            }
        }

        return writeWorkbook;

    }

    public void writeToFile(String filePath) throws IOException {
        if (StringUtils.isEmpty(filePath)) {
            return;
        }
        FileOutputStream fout = new FileOutputStream(filePath);
        this.writeToFile(fout);
    }

    public void writeToFile(FileOutputStream fileOutputStream) throws IOException {
        if (fileOutputStream == null) {
            return;
        }
        try {
            writeWorkbook = new SXSSFWorkbook(readWorkbook);
            writeWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            this.close();
        }
    }

    public void close() throws IOException {
        readWorkbook.close();
        writeWorkbook.close();
    }

    public static void main(String[] args) throws Exception {
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        InputStream fileInputStream = ExcelTemplate.class.getClassLoader().getResourceAsStream("excel_template_test.xlsx");
        ExcelTemplate excelTemplate = new ExcelTemplate(fileInputStream);
        Map<String, String> datas = new HashMap<>();
        datas.put("ProjectName", "Jurassic Park");
        datas.put("ResponsibleName", "Dreamsfuture");
        datas.put("Region", "Asia");
        datas.put("Country", "China");
        datas.put("Deadline", dateFormat.format(new Date()));
        datas.put("Cost", "$1000 Billion");
        excelTemplate.replaceTemplateWithSpecifiedData(0, datas);
        excelTemplate.writeToFile("excel_template_result.xlsx");

    }
}
