package com.lmx.poiutil.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

/**
 * @author liumingxin
 * @create 2018 25 16:03
 * @desc
 **/
public class Demo1 {


    public static void main(String[] args) throws Exception {
        Demo1 demo1 = new Demo1();
//        demo1.saveXLS();
//        demo1.saveXLSX();
//        demo1.createSheet();
//        demo1.createCell();
        demo1.readCell();
    }

    /**
     * 向 xls excel 中写入
     * @throws IOException
     */
    public void saveXLS() throws IOException {
        Workbook excel = new HSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("workbook.xls");
        excel.write(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * 向 xlsx 结尾的excel 写入
     * @throws IOException
     */
    public void saveXLSX() throws IOException {
        Workbook excel = new XSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("workbook.xlsx");
        excel.write(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * 创建一个 工作sheet
     * @throws IOException
     */
    public void createSheet() throws IOException{
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("workBook.xls");

        Sheet sheet1 = workbook.createSheet("a");
        Sheet sheet2 = workbook.createSheet("b");
        Sheet sheet3 = workbook.createSheet("c");

        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * 创建 单元格 写入数据
     * @throws IOException
     */
    public void createCell() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream("workBooke.xls");

        CreationHelper creationHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("new sheet");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        // Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
            creationHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * 读取数据使用
     * @throws IOException
     */
    public void readCell() throws IOException {
        InputStream inp = null;
        try {
            inp = new FileInputStream("workBooke.xls");
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row r = rowIterator.next();
                if (r == null) {
                    System.out.println("Empty Row");
                    continue;
                }
                for (int i = r.getFirstCellNum(); i < r.getLastCellNum(); i++) {
                    Cell cell = r.getCell(i);
                    String cellValue = "";
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            cellValue = cell.getRichStringCellValue().getString();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                cellValue = cell.getDateCellValue().toString();
                            } else {
                                cellValue = String.valueOf(cell.getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            cellValue = String.valueOf(cell.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            break;
                        default:
                    }
                    System.out.println("CellNum:" + i + " => CellValue:" + cellValue);
                }
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inp != null) {
                inp.close();
            }
        }

    }


}
