//package com.lmx.poiutil.demo;
//
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellReference;
//
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.IOException;
//import java.io.InputStream;
//
///**
// * @author liumingxin
// * @create 2018 25 17:21
// * @desc
// **/
//public class Demo2 {
//    public void readNewwCell() throws IOException {
//        InputStream inp = null;
//        try {
//            inp = new FileInputStream("workBooke.xls");
//            Workbook wb = WorkbookFactory.create(inp);
//            DataFormatter formatter = new DataFormatter();
//            Sheet sheet1 = wb.getSheetAt(0);
//            for (Row row : sheet1) {
//                for (Cell cell : row) {
//                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
//                    System.out.print(cellRef.formatAsString());
//                    System.out.print(" - ");
//
//                    // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
//                    String text = formatter.formatCellValue(cell);
//                    System.out.println(text);
//
//                    // Alternatively, get the value and format it yourself
//                    switch (CellType.;) {
//                        case  CellType.STRING:
//                            System.out.println(cell.getRichStringCellValue().getString());
//                            break;
//                        case CellType.NUMERIC:
//                            if (DateUtil.isCellDateFormatted(cell)) {
//                                System.out.println(cell.getDateCellValue());
//                            } else {
//                                System.out.println(cell.getNumericCellValue());
//                            }
//                            break;
//                        case CellType.BOOLEAN:
//                            System.out.println(cell.getBooleanCellValue());
//                            break;
//                        case CellType.FORMULA:
//                            System.out.println(cell.getCellFormula());
//                            break;
//                        case CellType.BLANK:
//                            System.out.println();
//                            break;
//                        default:
//                            System.out.println();
//                    }
//                }
//            }
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        } finally {
//            if (inp != null) {
//                inp.close();
//            }
//        }
//    }
//}
