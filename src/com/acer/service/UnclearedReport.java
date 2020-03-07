/*
 * V1.0 新增一卡通報表
 
 */



package com.acer.service;

import com.acer.bean.UnclearedReportBean;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Christine
 */
public class UnclearedReport extends AbstractReport
{
    private static Logger log = Logger.getLogger(UnclearedReport.class);
    
    String sheetName;
    String exportPath;
    String fileName;
    
    StringBuffer mailContent;

    Map<String , String> unit_idMap = new TreeMap();
    Map<String , String> cal_dateMap = new TreeMap();
    Map<String,String> operatingNameList = new TreeMap<String,String>();

    
    public String getSheetName()
    {
        return sheetName;
    }

    public void setSheetName(String sheetName)
    {
        this.sheetName=sheetName;
    }
    
    public String getExportPath()
    {
        return exportPath;
    }

    public void setExportPath(String exportPath)
    {
        this.exportPath=exportPath;
    }
    
    public String getFileName()
    {
        return fileName;
    }

    public void setFileName(String fileName)
    {
        this.fileName=fileName;
    }
    
    public StringBuffer getMailContent()
    {
        return mailContent;
    }
    
    public void setMailContent(StringBuffer mailContent)
    {
        this.mailContent = mailContent;
    }
    
    
    
    public UnclearedReport(Connection conn) 
    {
        super.setConn(conn);
    }
    
    public UnclearedReport() 
    {
        
    }
    
    
    @Override
    public Object rowMapper(ResultSet rs) throws SQLException 
    {
        UnclearedReportBean data=new UnclearedReportBean();
        
        data.setOperating_id(rs.getString("operating_id"));
        data.setOperating_name(rs.getString("operating_name"));
        data.setCal_date(rs.getString("cal_date"));
        data.setUnit_id(rs.getString("unit_id"));
        
        data.setTxamt_all(rs.getInt("txamt_all"));
        data.setPersonal_txamt_all(rs.getInt("personal_txamt_all"));
        data.setTransfer_txamt_all(rs.getInt("transfer_txamt_all"));
        
        data.setUncleared_count_c(rs.getInt("uncleared_count_c"));
        
        data.setTxamt_c(rs.getInt("txamt_c"));
        data.setPersonal_txamt_c(rs.getInt("personal_txamt_c"));
        data.setTransfer_txamt_c(rs.getInt("transfer_txamt_c"));
        
        data.setUncleared_count_e(rs.getInt("uncleared_count_e"));
        
        data.setTxamt_e(rs.getInt("txamt_e"));
        data.setPersonal_txamt_e(rs.getInt("personal_txamt_e"));
        data.setTransfer_txamt_e(rs.getInt("transfer_txamt_e"));
        
        data.setUpdate_time(rs.getString("update_time"));
        
        cal_dateMap.put(rs.getString("cal_date") , rs.getString("cal_date"));
        
        //V1.0 Map新增票證公司名稱
        if(rs.getString("unit_id").equals("02"))
        {
            unit_idMap.put(rs.getString("unit_id") , "悠遊卡");
        }
        else if(rs.getString("unit_id").equals("05"))
        {
            unit_idMap.put(rs.getString("unit_id") ,"台智卡");
        }
        else if(rs.getString("unit_id").equals("09"))
        {
            unit_idMap.put(rs.getString("unit_id") , "一卡通");
        }
        
        operatingNameList.put(rs.getString("operating_id") , rs.getString("operating_name"));
        
        return data;
    }
    
    
    public List<UnclearedReportBean> queryTxnList() 
    {
        List<UnclearedReportBean> resultList = new ArrayList<UnclearedReportBean>();
        String sql = generateSql();               
        List<Object> list = super.query(sql);
        
        if( list != null )
        {
            for( Object obj : list)
            {
                resultList.add((UnclearedReportBean) obj);
            }              
        }   
        return resultList;
    }
    
    SimpleDateFormat sdf=new SimpleDateFormat("YYYYMMdd");
    SimpleDateFormat sdf_YYYYMM =new SimpleDateFormat("YYYYMM");
    SimpleDateFormat sdf_YYYY =new SimpleDateFormat("YYYY");
    Calendar day =Calendar.getInstance();   
    Calendar B2day =Calendar.getInstance();  
    String today=null;
    String Before2day=null;
    String YYYYMM=null; //今年今月
    String YYYYBM = null; //今年前一個月
    String YYYY=null; //今年
    String BYYYBM=null; //去年前一個月
 
    
     private String generateSql()
    {
        //Operating date    
        today=sdf.format(day.getTime()); 
        B2day.add(Calendar.DATE,-2);
        Before2day = sdf.format(B2day.getTime());
        YYYYMM = sdf_YYYYMM.format(day.getTime());
        YYYY= sdf_YYYY.format(day.getTime());
        
        StringBuffer sb = new StringBuffer();       
        
        sb.append("Select  operating_id, operating_name, cal_date, unit_id,");
        sb.append(" txamt_all, personal_txamt_all, transfer_txamt_all, uncleared_count_c, txamt_c, personal_txamt_c, transfer_txamt_c, uncleared_count_e, txamt_e, personal_txamt_e, transfer_txamt_e, update_time");
        sb.append(" from uncleared_report");
        sb.append(" where");

//        sb.append(" cal_date >='20150401' and cal_date <= '20150430'"); //測試 時間區間
        
        if(today.equals(YYYYMM+"01")) // 04/01的情況
        {
            day.add(Calendar.MONTH, -1);
            YYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" cal_date >= '"+ YYYYBM+"01" +"'"+"and cal_date <= '"+ Before2day +"'");
        }
        else if(today.equals(YYYYMM+"02")) // 04/02的情況
        {
            day.add(Calendar.MONTH, -1);
            YYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" cal_date >= '"+ YYYYBM+"01" +"'"+"and cal_date <= '"+ Before2day +"'");
        }
        else if(today.equals(YYYY+"0101")) // 2015/01/01的情況
        {
            day.add(Calendar.MONTH, -1);
            day.add(Calendar.YEAR, -1);
            BYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" cal_date >= '"+ BYYYBM+"01" +"'"+"and cal_date <= '"+ Before2day +"'");
        }
        else if(today.equals(YYYY+"0102")) // 2015/01/02的情況
        {
            day.add(Calendar.MONTH, -1);
            day.add(Calendar.YEAR, -1);
            BYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" cal_date >= '"+ BYYYBM+"01" +"'"+"and cal_date <= '"+ Before2day +"'");
        }
        else // 04/03 - 04/30的情況
        {
//            System.out.println(YYYYMM+"01" +"-" + today);
            sb.append(" cal_date >= '"+ YYYYMM+"01" +"'"+"and cal_date <= '"+ Before2day +"'");
        }
         
        sb.append(" order by operating_id , cal_date , unit_id ");
        
        log.info(sb.toString());
        
        return sb.toString();
    }
     
     public Map<String ,ArrayList<Object>> Operatring_dataMap (String operating_id , List<UnclearedReportBean> datalist)
    {
        Map<String ,ArrayList<Object>> operatring_dataMap = new  TreeMap<String ,ArrayList<Object>>();
        
        for(UnclearedReportBean data : datalist)//該業者data值
        {
            if(data.getOperating_id().equals(operating_id))
            {
                ArrayList<Object> list = new ArrayList<Object>();
                
                String dataKey = data.getOperating_id() + data.getCal_date()  + data.getUnit_id();
                
                list.add(data.getTxamt_all());
                list.add(data.getPersonal_txamt_all());
                list.add(data.getTransfer_txamt_all());
                list.add(data.getUncleared_count_c());
                list.add(data.getTxamt_c());
                list.add(data.getPersonal_txamt_c());
                list.add(data.getTransfer_txamt_c());
                list.add(data.getUncleared_count_e());
                list.add(data.getTxamt_e());
                list.add(data.getPersonal_txamt_e());
                list.add(data.getTransfer_txamt_e());
                
                operatring_dataMap.put(dataKey,list);
            }
        }
        
        return operatring_dataMap;
    }
     
     public void generateExcel(List<UnclearedReportBean> datalist) throws IOException
    {
        XSSFWorkbook  workbook = new XSSFWorkbook();

        File excelFile = new File(this.getExportPath() + File.separator + this.getFileName()); // 帶值       
        FileOutputStream fileOut = null;
        
        try
        {
            StringBuffer sb =new StringBuffer();
            
            for(String id : operatingNameList.keySet()) //資料庫有的業者逐一產生
            {
                // 目前未清分有 : 首都客運 台北客運 大都會客運 新店客運 東南客運(北部) 基市公車 福和客運 長榮巴士
                XSSFSheet sheet =workbook.createSheet(operatingNameList.get(id));// 首都客運 
                excelContent(workbook, id, sheet , datalist, sb);// 首都客運 
            }
            
            
            setMailContent(sb);

            fileOut = new FileOutputStream(excelFile);

            workbook.write(fileOut);

            log.info("Generate Report Successful !");
        }
        catch(Exception e)
        {
            log.error("Generate Report Failed ! Reason: " + e);
        }
        finally
        {
            fileOut.flush();
            fileOut.close();
        }  
        
    }
     //V1.0 新增儲存格框線Style
     public void boldRight_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
     }
     public void boldLeft_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.THIN);
     }
     public void boldTop_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
     }
     public void boldBottom_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
     }
     public void generalAll_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
     }
     public void boldAll_border(XSSFCellStyle cellStyle)
     {
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
     }
     
     public void excelContent(XSSFWorkbook  workbook , String operating_id , XSSFSheet sheet , List<UnclearedReportBean> datalist , StringBuffer sb)
    {
        //Excel樣式

        XSSFCellStyle cellStyle, cellStyle1, cellStyle2, cellStyle3, cellStyle4, cellStyle5  = null;        
        XSSFDataFormat intFormat = workbook.createDataFormat();      
        XSSFColor blue=new XSSFColor(new java.awt.Color(0,0,255));
        XSSFColor red=new XSSFColor(new java.awt.Color(255,0,0));
        XSSFColor redBrown=new XSSFColor(new java.awt.Color(226, 197, 197));//Report Title
        XSSFColor salmon1=new XSSFColor(new java.awt.Color(255,140,105));//Column Title
        XSSFColor pink=new XSSFColor(new java.awt.Color(255,192,203));
        XSSFColor peachPuff=new XSSFColor(new java.awt.Color(255,218,185));
        
        //V1.0 統一字型
        XSSFFont font, font1, font2 = null;
        font=workbook.createFont();
        font.setFontName("微軟正黑體");
        font.setFontHeightInPoints((short)11);
//        font.setColor(blue);
        
        font1=workbook.createFont();
        font1.setFontName("Arial");
        font1.setFontHeightInPoints((short)10);
        font1.setColor(blue);
        
        font2=workbook.createFont();
        font2.setFontName("Arial");
        font2.setFontHeightInPoints((short)10);
        font2.setColor(red);
        
        
        //一般儲存格(上下左右邊框 + 置中 + 黑色字體)
        cellStyle=workbook.createCellStyle();
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        generalAll_border(cellStyle);
        cellStyle.setFont(font);
        
                    
        //一般數值儲存格(上下左右邊框 + 靠右 + 藍色字體)
        cellStyle1=workbook.createCellStyle();
        cellStyle1.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        generalAll_border(cellStyle1);
        cellStyle1.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));     
        cellStyle1.setFont(font1);
        
        //標示錯誤儲存格(上下左右邊框 + 靠右 + 紅色字體 + 橘色反白)
        cellStyle2=workbook.createCellStyle();
        cellStyle2.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        generalAll_border(cellStyle2);
        cellStyle2.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));        
        cellStyle2.setFont(font2);
        cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle2.setFillForegroundColor(peachPuff);
        
        //Title分界儲存格(上下左邊框 + 右粗框 + 置中 + 黑色字體)
        cellStyle3=workbook.createCellStyle();
        cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        boldRight_border(cellStyle3);
        cellStyle3.setFont(font);
        
        //數值分界儲存格(上下左邊框 + 右粗框 + 靠右 + 藍色字體)
        cellStyle4=workbook.createCellStyle();
        cellStyle4.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        boldRight_border(cellStyle4);
        cellStyle4.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));     
        cellStyle4.setFont(font1);
        
        //標示錯誤分界儲存格(上下左邊框 + 右粗框 + 紅色字體 + 橘色反白)
        cellStyle5=workbook.createCellStyle();
        cellStyle5.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        boldRight_border(cellStyle5);
        cellStyle5.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));        
        cellStyle5.setFont(font2);
        cellStyle5.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle5.setFillForegroundColor(peachPuff);

        // Title
        int rowIndex = 0;      
        int cellIndex = 0;
        
        XSSFRow titleRow = sheet.createRow(0);
        XSSFCell titleCell ;
        CellRangeAddress report_titleMerge = new CellRangeAddress(0,0,0,2); 
        sheet.addMergedRegion(report_titleMerge );       
        for(int col=0 ; col<=2 ; col++)
        {
            titleCell = titleRow.createCell(col);
            titleCell.setCellStyle(cellStyle);
            titleCell.setCellValue("未清分統計報表(日)"); 
        }
        
        CellRangeAddress operating_titleMerge = new CellRangeAddress(0,0,3,5); 
        sheet.addMergedRegion(operating_titleMerge );
        for(int col=3 ; col<=5 ; col++)
        {
            titleCell = titleRow.createCell(col);
            titleCell.setCellStyle(cellStyle);

            titleCell.setCellValue(operatingNameList.get(operating_id));          
            
        }
        
  
        // Column
        rowIndex = 1; 
        cellIndex = 0;

        XSSFRow columnRow ;
        XSSFCell tCell ;
        CellRangeAddress dateMerge = new CellRangeAddress(1 , 4 , 0 , 0 ); 
        sheet.addMergedRegion(dateMerge );
        
        XSSFRow column1Row = null,column2Row = null,column3Row = null ,column4Row = null;

        for(int row= 1 ; row<=4 ; row++)
        {
            columnRow = sheet.createRow(row);
            switch(row) 
            {
                case 1:
                    column1Row = columnRow;
                    break;
                case 2:
                    column2Row = columnRow;
                    break;
                case 3:
                    column3Row = columnRow;
                    break;   
                case 4:
                    column4Row = columnRow;
                    break; 
            }
            
            tCell = columnRow.createCell(cellIndex);
            tCell.setCellStyle(cellStyle);
            tCell.setCellValue("清分日");
            sheet.setColumnWidth(cellIndex,(short)(13 * 256)); // 自動調整欄位寬度
            rowIndex++;
        }

        
        XSSFCell column1Cell ;
//        column1Cell.setCellStyle(cellStyle);
        cellIndex =1;
        String [] titleArray={"卡片扣款金額","社福優惠金額","轉乘優惠金額"};

        //票證公司名稱 ECC
        CellRangeAddress column1Merge = new CellRangeAddress(1 , 1 , cellIndex , cellIndex +9 ); 
        sheet.addMergedRegion(column1Merge );
        for(int col= cellIndex ; col<=cellIndex +9 ; col++)
        {
            column1Cell = column1Row.createCell(col);
            column1Cell.setCellValue("悠遊卡");
            if(col==cellIndex +9)
            {
                column1Cell.setCellStyle(cellStyle3);  
            }
            else
            {
                column1Cell.setCellStyle(cellStyle);  
            }                   
        }
   
        //CPS全部
        cellIndex =1;
        CellRangeAddress column2Merge = new CellRangeAddress(2 , 2 , cellIndex , cellIndex +2 ); 
        sheet.addMergedRegion(column2Merge );
        for(int col= cellIndex ; col<=cellIndex +2 ; col++)
        {
            column1Cell = column2Row.createCell(col);
            column1Cell.setCellValue("CPS全部");
            column1Cell.setCellStyle(cellStyle);
            
            CellRangeAddress columnMerge = new CellRangeAddress(3 , 4 , col , col ); 
            
            sheet.addMergedRegion(columnMerge );
            column1Cell = column3Row.createCell(col);
            column1Cell.setCellValue(titleArray[col-1]);  
            column1Cell.setCellStyle(cellStyle);
            
            
            column1Cell = column4Row.createCell(col);
            column1Cell.setCellValue(titleArray[col-1]);
            column1Cell.setCellStyle(cellStyle);
            
            sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
        }
        
        cellIndex =4;
        //未清分
        CellRangeAddress column3Merge = new CellRangeAddress(2 , 2 , cellIndex , cellIndex +6 ); 
        sheet.addMergedRegion(column3Merge );
        CellRangeAddress column4Merge = new CellRangeAddress(3 , 3 , cellIndex+1 , cellIndex+3 ); 
        sheet.addMergedRegion(column4Merge );
        CellRangeAddress column5Merge = new CellRangeAddress(3 , 3 , cellIndex+4 , cellIndex+6 ); 
        sheet.addMergedRegion(column5Merge );
        for(int col= cellIndex ; col<=cellIndex +6 ; col++)
        {
            column1Cell = column2Row.createCell(col);
            column1Cell.setCellValue("未清分");
            if(col==10)
            {
                column1Cell.setCellStyle(cellStyle3);  
            }
            else
            {
                column1Cell.setCellStyle(cellStyle);  
            }
            
            if(col==4)
            {
                CellRangeAddress columnMerge = new CellRangeAddress(3 , 4 , col , col ); 
                sheet.addMergedRegion(columnMerge );
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("筆數");
                column1Cell.setCellStyle(cellStyle);
                column1Cell = column4Row.createCell(col);
                column1Cell.setCellValue("筆數");
                column1Cell.setCellStyle(cellStyle);
            }
            
            else if(col>=5 && col<=7)
            {
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("已上傳未清分");
                column1Cell.setCellStyle(cellStyle);
                
                column1Cell = column4Row.createCell(col);
                column1Cell.setCellValue(titleArray[col-5]);
                column1Cell.setCellStyle(cellStyle);
                sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
            }
            
            else 
            {
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("清分錯誤");
                if(col==10)
                {
                    column1Cell.setCellStyle(cellStyle3);  
                }
                else
                {
                    column1Cell.setCellStyle(cellStyle);  
                }
                
                column1Cell = column4Row.createCell(col);
                if(col==10)
                {
                    column1Cell.setCellStyle(cellStyle3);  
                }
                else
                {
                    column1Cell.setCellStyle(cellStyle);  
                }
                column1Cell.setCellValue(titleArray[col-8]);               
                sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
            }

        }

        
        //V1.0 新增 票證公司名稱 KRTC
        cellIndex =11;
        CellRangeAddress column1Merge6 = new CellRangeAddress(1 , 1 , cellIndex , cellIndex +9 ); 
        sheet.addMergedRegion(column1Merge6 );
        for(int col= cellIndex ; col<=cellIndex +9 ; col++)
        {
            column1Cell = column1Row.createCell(col);
            column1Cell.setCellValue("一卡通");
            if(col==cellIndex +9)
            {
                column1Cell.setCellStyle(cellStyle3);  
            }
            else
            {
                column1Cell.setCellStyle(cellStyle);  
            }          
        }
   
        //CPS全部
        cellIndex =11;
        CellRangeAddress column7Merge = new CellRangeAddress(2 , 2 , cellIndex , cellIndex +2 ); //11-13
        sheet.addMergedRegion(column7Merge );
        for(int col= cellIndex ; col<=cellIndex +2 ; col++)
        {
            column1Cell = column2Row.createCell(col);
            column1Cell.setCellValue("CPS全部");
            column1Cell.setCellStyle(cellStyle);
            
            CellRangeAddress columnMerge = new CellRangeAddress(3 , 4 , col , col ); 
            
            sheet.addMergedRegion(columnMerge );
            column1Cell = column3Row.createCell(col);
            column1Cell.setCellValue(titleArray[col-11]);  
            column1Cell.setCellStyle(cellStyle);
            
            
            column1Cell = column4Row.createCell(col);
            column1Cell.setCellValue(titleArray[col-11]);
            column1Cell.setCellStyle(cellStyle);
            
            sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
        }
        
        cellIndex =14;
        //未清分
        CellRangeAddress column8Merge = new CellRangeAddress(2 , 2 , cellIndex , cellIndex +6 ); 
        sheet.addMergedRegion(column8Merge );
        CellRangeAddress column9Merge = new CellRangeAddress(3 , 3 , cellIndex+1 , cellIndex+3 ); 
        sheet.addMergedRegion(column9Merge );
        CellRangeAddress column10Merge = new CellRangeAddress(3 , 3 , cellIndex+4 , cellIndex+6 ); 
        sheet.addMergedRegion(column10Merge );
        for(int col= cellIndex ; col<=cellIndex +6 ; col++) //14-20
        {
            column1Cell = column2Row.createCell(col);
            column1Cell.setCellValue("未清分");
            if(col==20)
            {
                column1Cell.setCellStyle(cellStyle3);  
            }
            else
            {
                column1Cell.setCellStyle(cellStyle);  
            }
            
            if(col==14)
            {
                CellRangeAddress columnMerge = new CellRangeAddress(3 , 4 , col , col ); 
                sheet.addMergedRegion(columnMerge );
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("筆數");
                column1Cell.setCellStyle(cellStyle);
                column1Cell = column4Row.createCell(col);
                column1Cell.setCellValue("筆數");
                column1Cell.setCellStyle(cellStyle);
            }
            
            else if(col>=15 && col<=17)
            {
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("已上傳未清分");
                column1Cell.setCellStyle(cellStyle);
                
                column1Cell = column4Row.createCell(col);
                column1Cell.setCellValue(titleArray[col-15]);
                column1Cell.setCellStyle(cellStyle);
                sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
            }
            
            else 
            {
                column1Cell = column3Row.createCell(col);
                column1Cell.setCellValue("清分錯誤");
                if(col==20)
                {
                    column1Cell.setCellStyle(cellStyle3);  
                }
                else
                {
                    column1Cell.setCellStyle(cellStyle);  
                }
                
                column1Cell = column4Row.createCell(col);
                if(col==20)
                {
                    column1Cell.setCellStyle(cellStyle3);  
                }
                else
                {
                    column1Cell.setCellStyle(cellStyle);  
                }
                column1Cell.setCellValue(titleArray[col-18]);
                sheet.setColumnWidth(col,(short)(14 * 256)); // 自動調整欄位寬度
            }

        }
        
        //data
        Map<String ,ArrayList<Object>> dataMap =Operatring_dataMap(operating_id , datalist); //資料庫該月所有業者data
        int ECCsum_txamt_all =0, ECCsum_personal_txamt_all =0, ECCsum_transfer_txamt_all=0, ECCsum_uncleared_count =0, ECCsum_txamt_c =0, ECCsum_personal_txamt_c =0, ECCsum_transfer_txamt_c =0, ECCsum_txamt_e =0, ECCsum_personal_txamt_e =0, ECCsum_transfer_txamt_e =0;
        int KRTCsum_txamt_all =0, KRTCsum_personal_txamt_all =0, KRTCsum_transfer_txamt_all=0, KRTCsum_uncleared_count =0, KRTCsum_txamt_c =0, KRTCsum_personal_txamt_c =0, KRTCsum_transfer_txamt_c =0, KRTCsum_txamt_e =0, KRTCsum_personal_txamt_e =0, KRTCsum_transfer_txamt_e =0;

        
        Calendar b2day =Calendar.getInstance();
        b2day.add(Calendar.DATE,-2);
        Before2day = sdf.format(b2day.getTime());        
        
//        rowIndex=1;
        XSSFCell datecell;
        for(String calDate : cal_dateMap.keySet()) //搜尋日期
        {
            cellIndex = 0;
            XSSFRow dataRow = sheet.createRow(rowIndex++);
            datecell = dataRow.createCell(cellIndex++);
            datecell.setCellStyle(cellStyle);
            datecell.setCellValue(calDate);
            
            
            //V1.0 新增票證公司代碼判斷
            for(String unitId : unit_idMap.keySet()) 
            {

                if(unitId.equals("05"))//台智卡
                {
                    
                }
                else  //悠遊卡及一卡通
                {
                    String key = operating_id + calDate + unitId; //悠遊卡
                
                    if(dataMap.containsKey(key)) // 若有資料，填值
                    {
                        ArrayList list = dataMap.get(key);

                        if((int)list.get(0)==0 )
                        {                     
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,list.get(0));
                            setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(1));
                            setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(2)); 

                            if(calDate.equals(Before2day))
                            {

                                sb.append(operatingNameList.get(operating_id)+", 票證公司:"+ unit_idMap.get(unitId) + "，清分日:"+ calDate + "無清分資料。\r\n<br>");                                  

                            }                                                               
                        }
                        else
                        {
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,list.get(0));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(1));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(2));

                        }

                        if( ((int)list.get(3)+(int)list.get(7)) >0)
                        {
                            setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,(int)list.get(3)+(int)list.get(7));

                            //未清分
                            if((int)list.get(3)>0 )
                            {
                                setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(4));
                                setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(5));
                                setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(6));

                                if(calDate.equals(Before2day))
                                {

                                    sb.append(operatingNameList.get(operating_id)+", 票證公司:"+ unit_idMap.get(unitId) + "，清分日:"+ calDate + "有未清分的資料。\r\n<br>");                                  

                                }
                            }
                            else
                            {
                                setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(4));
                                setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(5));
                                setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(6));
                            }

                            //清分錯誤
                            if((int)list.get(7)>0 )
                            {
                                setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(8));
                                setCellData(workbook, dataRow , cellStyle2,cellIndex++ ,list.get(9));
                                                                
                                setCellData(workbook, dataRow , cellStyle5,cellIndex++ ,list.get(10));

                                if(calDate.equals(Before2day))
                                {

                                    sb.append(operatingNameList.get(operating_id)+", 票證公司:"+ unit_idMap.get(unitId) + "，清分日:"+ calDate + "有清分錯誤的資料。\r\n<br>");                                  

                                }
                            }
                            else
                            {
                                setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(8));
                                setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(9));
                                
                                setCellData(workbook, dataRow , cellStyle4,cellIndex++ ,list.get(10));
                            }

                        }
                        else
                        {                      
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,(int)list.get(3)+(int)list.get(7));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(4));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(5));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(6));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(8));
                            setCellData(workbook, dataRow , cellStyle1,cellIndex++ ,list.get(9));
                            setCellData(workbook, dataRow , cellStyle4,cellIndex++ ,list.get(10));
                        }

                        if(unitId.equals("02"))// 悠遊卡加總
                        {
                            ECCsum_txamt_all += (int)list.get(0);

                            ECCsum_personal_txamt_all += (int)list.get(1);

                            ECCsum_transfer_txamt_all += (int)list.get(2);

                            ECCsum_uncleared_count += (int)list.get(3) + (int)list.get(7);

                            ECCsum_txamt_c += (int)list.get(4);

                            ECCsum_personal_txamt_c += (int)list.get(5);

                            ECCsum_transfer_txamt_c += (int)list.get(6);

                            ECCsum_txamt_e += (int)list.get(8);

                            ECCsum_personal_txamt_e += (int)list.get(9);

                            ECCsum_transfer_txamt_e += (int)list.get(10);
                        }
                        
                        else if(unitId.equals("09"))// 一卡通加總
                        {
                            KRTCsum_txamt_all += (int)list.get(0);

                            KRTCsum_personal_txamt_all += (int)list.get(1);

                            KRTCsum_transfer_txamt_all += (int)list.get(2);

                            KRTCsum_uncleared_count += (int)list.get(3) + (int)list.get(7);

                            KRTCsum_txamt_c += (int)list.get(4);

                            KRTCsum_personal_txamt_c += (int)list.get(5);

                            KRTCsum_transfer_txamt_c += (int)list.get(6);

                            KRTCsum_txamt_e += (int)list.get(8);

                            KRTCsum_personal_txamt_e += (int)list.get(9);

                            KRTCsum_transfer_txamt_e += (int)list.get(10);
                        }
                        

                    }      
                    else // 若沒有資料，填0
                    {

                        if(operating_id.equals("008") && unitId.equals("02")) // 因為福和客運沒有使用ECC，所以資料庫沒有是正常。
                        {
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle1, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle4, cellIndex++ ,0);
                        }
                        else
                        {
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle2, cellIndex++ ,0);
                            setCellData(workbook, dataRow , cellStyle5, cellIndex++ ,0);
                            if(calDate.equals(Before2day))
                            {

                                sb.append(operatingNameList.get(operating_id) +", 票證公司:"+ unit_idMap.get(unitId) + "，清分日:"+ calDate + "資料庫沒有資料。\r\n<br>");                                  

                            }
                        }
                        

                        
                    }
                }
        
            }
            
        }
        
        
        
        //Sum
        cellIndex = 0;
        XSSFRow sum_dataRow = sheet.createRow(rowIndex++);
        XSSFCell sumCell = sum_dataRow.createCell(cellIndex++);
        sumCell.setCellStyle(cellStyle);
        sumCell.setCellValue("合計");
        
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_personal_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_transfer_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_uncleared_count);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_personal_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_transfer_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_txamt_e);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,ECCsum_personal_txamt_e);
        setCellData(workbook, sum_dataRow , cellStyle4, cellIndex++ ,ECCsum_transfer_txamt_e);
        
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_personal_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_transfer_txamt_all);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_uncleared_count);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_personal_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_transfer_txamt_c);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_txamt_e);
        setCellData(workbook, sum_dataRow , cellStyle1, cellIndex++ ,KRTCsum_personal_txamt_e);
        setCellData(workbook, sum_dataRow , cellStyle4, cellIndex++ ,KRTCsum_transfer_txamt_e);
        
        
        sheet.autoSizeColumn(8,true); // 自動調整欄位寬度
}

private void setCellData(XSSFWorkbook  workbook , XSSFRow row , XSSFCellStyle cellStyle,int index , Object obj)
   {      
        int cellrow = index++;
        XSSFCell cell = row.createCell(cellrow);

        cell.setCellStyle(cellStyle);

       
        
        
       if( obj instanceof Integer)
       {
           cell.setCellValue((Integer)obj);
       }
       else if( obj instanceof Double)
       {
           cell.setCellValue((Double)obj);
       }
       else if( obj instanceof Date)
       {
           cell.setCellValue((Date)obj);
       }
       else if( obj instanceof String)
       {
           cell.setCellValue((String)obj);
       }
       
       
   }
     
}
