package com.lmx.poiutil.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.bcel.classfile.ClassFormatException;

/**
 * @author liumingxin
 * @create 2018 25 17:30
 * @desc
 **/
public class WorkBookUtil {

    private static final String path = "";

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    public Workbook getWriteWorkBook(String type){
        Workbook workbook = null;
        switch (type){
            case XLS :{
                workbook = new HSSFWorkbook();
                break;
            }
            case XLSX : {
                workbook = new XSSFWorkbook();
                break;
            }
            default:{
                throw new RuntimeException( type + "格式转换错误");
            }
        }
        return workbook;
    }

}
