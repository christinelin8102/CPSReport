/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.acer.service;


import com.acer.bean.RevenueReportBean;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;
import javafx.scene.paint.Color;
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
public class RevenueReport extends AbstractReport 
{
    private static Logger log = Logger.getLogger(RevenueReport.class);
    
    String sheetName;
    String exportPath;
    String fileName;
    

    Map<String , String> depotMap = new TreeMap();
    Map<String , String> businessdateMap = new TreeMap();
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
        
    
    public RevenueReport(Connection conn) 
    {
        super.setConn(conn);
    }
 
    
    
    @Override
    public Object rowMapper(ResultSet rs) throws SQLException 
    {
        RevenueReportBean data=new RevenueReportBean();
        
        data.setOperator_Id(rs.getString("operating_id"));
        data.setOperator_name(rs.getString("operating_name"));
        data.setBusiness_date(rs.getString("business_date"));
        
        data.setDepot_Id(rs.getString("depot_id"));
        
        
        data.setDepot_name(rs.getString("depot_name"));
        
        data.setTrx_count(rs.getInt("trx_count"));
        data.setTxamt(rs.getInt("txamt"));
        data.setPersonal_count(rs.getInt("personal_count"));
        data.setPersonal_txamt(rs.getInt("personal_txamt"));
        data.setTransfer_MRT_count(rs.getInt("transfer_MRT_count"));
        data.setTransfer_MRT_txamt(rs.getInt("transfer_MRT_txamt"));
        data.setTransfer_bus_count(rs.getInt("transfer_bus_count"));
        data.setTransfer_bus_txamt(rs.getInt("transfer_bus_txamt"));
        
        operatingNameList.put(rs.getString("operating_id") , rs.getString("operating_name"));
        depotMap.put(rs.getString("operating_id")+ ","+  rs.getString("depot_id"),rs.getString("depot_name"));
        businessdateMap.put(rs.getString("business_date"),rs.getString("business_date"));

        return data;
    }
    
    public List<RevenueReportBean> queryTxnList() 
    {
        List<RevenueReportBean> resultList = new ArrayList<RevenueReportBean>();
        String sql = generateSql();               
        List<Object> list = super.query(sql);
        
        if( list != null )
        {
            for( Object obj : list)
            {
                resultList.add((RevenueReportBean) obj);
            }              
        }   
        return resultList;
    }
    
    public Map<String ,ArrayList<Object>> Operatring_dataMap (String operating_id , List<RevenueReportBean> datalist)
    {
        Map<String ,ArrayList<Object>> operatring_dataMap = new  TreeMap<String ,ArrayList<Object>>();
        
        for(RevenueReportBean data : datalist)//該業者data值
        {
            if(data.getOperator_Id().equals(operating_id))
            {
                ArrayList<Object> list = new ArrayList<Object>();
                
                String dataKey = data.getBusiness_date() + "," + data.getOperator_Id() + "," + data.getDepot_Id();
                
                list.add(data.getTrx_count());
                list.add(data.getTxamt());
                list.add(data.getPersonal_count());
                list.add(data.getPersonal_txamt());
                list.add(data.getTransfer_MRT_count());
                list.add(data.getTransfer_MRT_txamt());
                list.add(data.getTransfer_bus_count());
                list.add(data.getTransfer_bus_txamt());
                operatring_dataMap.put(dataKey,list);
            }
        }
        
        return operatring_dataMap;
    }
    
    public Map<String ,ArrayList<Object>> Operatring_Sum_dataMap (String operating_id , List<RevenueReportBean> datalist)
    {
        Map<String ,ArrayList<Object>> operatring_Sum_dataMap = new  TreeMap<String ,ArrayList<Object>>();
        
      
        for(String depot : depotMap.keySet())
        {
            int sum_trxCount=0, sum_trxamt=0, sum_personalCount=0, sum_personalTxamt=0, sum_TransferMRTCount=0, sum_TransferMRTTxamt=0, sum_TransferBusCount=0, sum_TransferBusTxamt=0;
            ArrayList<Object> list = new ArrayList<Object>();
            
            for(RevenueReportBean data : datalist)
            {   
                String dataKey = data.getOperator_Id() + "," + data.getDepot_Id();
                if(data.getOperator_Id().equals(operating_id) && dataKey.equals(depot))
                {
                    sum_trxCount += data.getTrx_count();
                    sum_trxamt += data.getTxamt();
                    sum_personalCount += data.getPersonal_count();
                    sum_personalTxamt += data.getPersonal_txamt();
                    sum_TransferMRTCount += data.getTransfer_MRT_count();
                    sum_TransferMRTTxamt += data.getTransfer_MRT_txamt();
                    sum_TransferBusCount += data.getTransfer_bus_count();
                    sum_TransferBusTxamt += data.getTransfer_bus_txamt();
                }
            }
            list.add(sum_trxCount);
            list.add(sum_trxamt);
            list.add(sum_personalCount);
            list.add(sum_personalTxamt);
            list.add(sum_TransferMRTCount);
            list.add(sum_TransferMRTTxamt);
            list.add(sum_TransferBusCount);
            list.add(sum_TransferBusTxamt);

            operatring_Sum_dataMap.put(depot,list);
        }

        
        return operatring_Sum_dataMap;
    }
    
    SimpleDateFormat sdf=new SimpleDateFormat("YYYYMMdd");
    SimpleDateFormat sdf_YYYYMM =new SimpleDateFormat("YYYYMM");
    SimpleDateFormat sdf_YYYY =new SimpleDateFormat("YYYY");
    Calendar day =Calendar.getInstance();      
    String today=null;
    String YYYYMM=null; //今年今月
    String YYYYBM = null; //今年前一個月
    String YYYY=null; //今年
    String BYYYBM=null; //去年前一個月
 
    
     private String generateSql()
    {
        //Operating date    
        today=sdf.format(day.getTime());      
        YYYYMM = sdf_YYYYMM.format(day.getTime());
        YYYY= sdf_YYYY.format(day.getTime());
        
        StringBuffer sb = new StringBuffer();       
        
        sb.append("Select  operating_id, operating_name, business_date, depot_id, depot_name,");
        sb.append("trx_count, txamt, personal_count, personal_txamt, transfer_MRT_count, transfer_MRT_txamt, transfer_bus_count, transfer_bus_txamt");
        sb.append(" from revenue_report");
        sb.append(" where");

        
        if(today.equals(YYYYMM+"01")) // 04/01的情況
        {
            day.add(Calendar.MONTH, -1);
            YYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" business_date >= '"+ YYYYBM+"01" +"'"+"and business_date < '"+ today +"'");
        }
        else if(today.equals(YYYY+"0101")) // 2015/01/01的情況
        {
            day.add(Calendar.MONTH, -1);
            day.add(Calendar.YEAR, -1);
            BYYYBM = sdf_YYYYMM.format(day.getTime());
            sb.append(" business_date >= '"+ BYYYBM+"01" +"'"+"and business_date < '"+ today +"'");
        }
        else // 04/02 - 04/30的情況
        {
//            System.out.println(YYYYMM+"01" +"-" + today);
            sb.append(" business_date >= '"+ YYYYMM+"01" +"'"+"and business_date < '"+ today +"'");
        }
        
                      
        sb.append(" order by operating_id , business_date , depot_id ");
        
        log.info(sb.toString());
        
        return sb.toString();
    }
     
    public void generateExcel(List<RevenueReportBean> datalist) throws IOException
    {
        XSSFWorkbook  workbook = new XSSFWorkbook();
        
         
		
        File excelFile = new File(this.getExportPath() + File.separator + this.getFileName()); // 帶值
       
        FileOutputStream fileOut = null;
        
        try
        {

            
            for(String id : operatingNameList.keySet()) //資料庫有的業者逐一產生
            {
                // 目前營收有 : 首都客運 台北客運 大都會客運 新店客運 東南客運(北部) 基市公車 巨業客運
                XSSFSheet sheet =workbook.createSheet(operatingNameList.get(id));// 首都客運 
                excelContent(workbook, id, sheet , datalist);// 首都客運 
            }

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
     
    
    
    
    public void excelContent(XSSFWorkbook  workbook , String operating_id , XSSFSheet sheet , List<RevenueReportBean> datalist)
    {
        //Excel樣式

        XSSFCellStyle cellStyle, cellStyle1, cellStyle2, cellStyle3, cellStyle4, cellStyle5, cellStyle6, cellStyle7, cellStyle8, cellStyle9  = null;        
        XSSFDataFormat intFormat = workbook.createDataFormat();      
        XSSFColor blue=new XSSFColor(new java.awt.Color(0,0,255));
        XSSFColor redBrown=new XSSFColor(new java.awt.Color(226, 197, 197));//Report Title
        XSSFColor brightGreen=new XSSFColor(new java.awt.Color(195, 195, 136));//Column Title
        XSSFColor grey=new XSSFColor(new java.awt.Color(224, 224, 224));
        
        //一般儲存格(上下左右邊框 + 置中 + 黑色字體 ， 紅棕色反白)
        cellStyle=workbook.createCellStyle();
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(redBrown);
        
        //Column儲存格(上粗框 + 置中 + 黑色字體， 紅棕色反白)
        cellStyle5=workbook.createCellStyle();
        cellStyle5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle5.setBorderTop(BorderStyle.MEDIUM);
        cellStyle5.setBorderBottom(BorderStyle.THIN);
        cellStyle5.setBorderLeft(BorderStyle.THIN);
        cellStyle5.setBorderRight(BorderStyle.THIN);
        cellStyle5.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle5.setFillForegroundColor(redBrown);
        
        //Column儲存格(上右粗框 + 置中 + 黑色字體， 紅棕色反白)
        cellStyle6=workbook.createCellStyle();
        cellStyle6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle6.setBorderTop(BorderStyle.MEDIUM);
        cellStyle6.setBorderBottom(BorderStyle.THIN);
        cellStyle6.setBorderLeft(BorderStyle.THIN);
        cellStyle6.setBorderRight(BorderStyle.MEDIUM);
        cellStyle6.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle6.setFillForegroundColor(redBrown);
        
        //Column儲存格(下粗框 + 置中 + 黑色字體， 紅棕色反白)
        cellStyle7=workbook.createCellStyle();
        cellStyle7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle7.setBorderTop(BorderStyle.THIN);
        cellStyle7.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle7.setBorderLeft(BorderStyle.THIN);
        cellStyle7.setBorderRight(BorderStyle.THIN);
        cellStyle7.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle7.setFillForegroundColor(redBrown);
        
        //Column儲存格(下右粗框 + 置中 + 黑色字體， 紅棕色反白)
        cellStyle8=workbook.createCellStyle();
        cellStyle8.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle8.setBorderTop(BorderStyle.THIN);
        cellStyle8.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle8.setBorderLeft(BorderStyle.THIN);
        cellStyle8.setBorderRight(BorderStyle.MEDIUM);
        cellStyle8.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle8.setFillForegroundColor(redBrown);
        
        //分界儲存格(右粗框 + 置中 + 黑色字體)
        cellStyle1=workbook.createCellStyle();
        cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle1.setBorderTop(BorderStyle.THIN);
        cellStyle1.setBorderBottom(BorderStyle.THIN);
        cellStyle1.setBorderLeft(BorderStyle.THIN);
        cellStyle1.setBorderRight(BorderStyle.MEDIUM);
        
        //分界儲存格(右粗框 + 置中 + 黑色字體)
        cellStyle9=workbook.createCellStyle();
        cellStyle9.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle9.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle9.setBorderTop(BorderStyle.THIN);
        cellStyle9.setBorderBottom(BorderStyle.THIN);
        cellStyle9.setBorderLeft(BorderStyle.THIN);
        cellStyle9.setBorderRight(BorderStyle.MEDIUM);
        cellStyle9.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle9.setFillForegroundColor(grey);
        
        //Column營運日及合計儲存格(上下左右粗邊框 + 置中 + 黑色字體 + 灰色反白)
        cellStyle4=workbook.createCellStyle();
        cellStyle4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle4.setBorderTop(BorderStyle.MEDIUM);
        cellStyle4.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle4.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle4.setBorderRight(BorderStyle.MEDIUM);
        cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle4.setFillForegroundColor(grey);
        
        
        
        //數值儲存格(上下左細邊框 + 右粗邊框 + 靠右 + 藍色字體)
        cellStyle2=workbook.createCellStyle();
        cellStyle2.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        cellStyle2.setBorderTop(BorderStyle.MEDIUM);
        cellStyle2.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle2.setBorderLeft(BorderStyle.THIN);
        cellStyle2.setBorderRight(BorderStyle.THIN);
        cellStyle2.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));
        
        //數值分界儲存格(上下左細邊框 + 右粗邊框 + 靠右 + 藍色字體)
        cellStyle3=workbook.createCellStyle();
        cellStyle3.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        cellStyle3.setBorderTop(BorderStyle.MEDIUM);
        cellStyle3.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle3.setBorderLeft(BorderStyle.THIN);
        cellStyle3.setBorderRight(BorderStyle.MEDIUM);
        cellStyle3.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));
        
        XSSFFont font, font1 = null;
        
        font=workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)10);
        font.setColor(blue);
        cellStyle2.setFont(font);
        cellStyle3.setFont(font);
        
        font1=workbook.createFont();
        font1.setBold(true);
        cellStyle5.setFont(font1);
        cellStyle6.setFont(font1);
       
 
        
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
            titleCell.setCellStyle(cellStyle4);
            titleCell.setCellValue("每日營收統計報表"); 
        }
        
        CellRangeAddress operating_titleMerge = new CellRangeAddress(0,0,3,5); 
        sheet.addMergedRegion(operating_titleMerge );
        for(int col=3 ; col<=5 ; col++)
        {
            titleCell = titleRow.createCell(col);
            titleCell.setCellStyle(cellStyle4);

            titleCell.setCellValue(operatingNameList.get(operating_id));          
            
        }
        
  
        // Column
        rowIndex = 1; 
        cellIndex = 0;

        XSSFRow columnRow ;
        XSSFCell tCell ;
        CellRangeAddress dateMerge = new CellRangeAddress(1 , 3 , 0 , 0 ); 
        sheet.addMergedRegion(dateMerge );
        
        XSSFRow column1Row = null,column2Row = null,column3Row = null;

        for(int row= 1 ; row<=3 ; row++)
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
            }
            
            tCell = columnRow.createCell(cellIndex);
            tCell.setCellStyle(cellStyle4);
            tCell.setCellValue("營運日"); 
            
            rowIndex++;
        }

        XSSFCell column1Cell ;

        
        cellIndex =1;
        int column2cell =1;
        int column3cell =1;
        String [] title2Array={"卡片扣款","社福卡交易","轉乘交易(捷運)","轉乘交易(公車)"};
        String [] title3Array={"筆數","金額"};
        
        for(String deptId : depotMap.keySet()) 
        {
            String[] depot = deptId .split(",");
                
            if(depot[0].equals(operating_id)) // 限定業者代碼
            {
                //各場站名稱
                CellRangeAddress columnMerge = new CellRangeAddress(1 , 1 , cellIndex , cellIndex +7 ); 
                sheet.addMergedRegion(columnMerge );
                for(int col= cellIndex ; col<=cellIndex +7 ; col++)
                {
                    column1Cell = column1Row.createCell(col);
                    
                    if(col % 8 == 0)
                    {
                        column1Cell.setCellStyle(cellStyle6);
                    }
                    else
                    {
                        column1Cell.setCellStyle(cellStyle5);
                    }
                    
                    column1Cell.setCellValue(depot[1] + "/" + depotMap.get(deptId)); // 場站代碼+場站名稱 
                    
                }

                cellIndex= cellIndex +8;
                
                //卡片扣款,社福卡交易,轉乘交易(捷運),轉乘交易(公車)
                for(String title2 : title2Array)
                {
                    CellRangeAddress column2Merge = new CellRangeAddress(2 , 2 , column2cell , column2cell +1 ); 
                    sheet.addMergedRegion(column2Merge );
                    for(int col= column2cell ; col<=column2cell +1  ; col++)
                    {
                        column1Cell = column2Row.createCell(col);
                        
                        if(col % 8 == 0)
                        {
                            column1Cell.setCellStyle(cellStyle1);
                        }
                        else
                        {
                            column1Cell.setCellStyle(cellStyle);
                        }
                        
                        column1Cell.setCellValue(title2);
                        
                    }
                    column2cell= column2cell +2;
                    
                    //金額,筆數
                    for(String title3 : title3Array)
                    {
                        column1Cell = column3Row.createCell(column3cell);
                        if(column3cell % 8 == 0)
                        {
                            column1Cell.setCellStyle(cellStyle8);
                        }
                        else
                        {
                            column1Cell.setCellStyle(cellStyle7);
                        }
                        column1Cell.setCellValue(title3);

                        column3cell++;
                    }
                }
            }
        }

        
        
        //data
        Map<String ,ArrayList<Object>> dataMap =Operatring_dataMap(operating_id , datalist); //資料庫該月所有業者data
        for(String businessdate : businessdateMap.keySet()) //搜尋日期
        {
            cellIndex = 0;
            XSSFRow dataRow = sheet.createRow(rowIndex++);
            XSSFCell datecell = dataRow.createCell(cellIndex++);
            datecell.setCellStyle(cellStyle9);
            datecell.setCellValue(businessdate);
//            setCellData(workbook, dataRow , cellIndex++ ,businessdate);
            
            for(String deptId : depotMap.keySet()) //搜尋該月所有業者的場站
            {
                String[] depot = deptId .split(",");
                
                if(depot[0].equals(operating_id)) // 限定業者代碼
                {
                    
                    String key = businessdate + "," + deptId;
                    if(dataMap.containsKey(key)) // 若有此場站資料，填值
                    {
                        ArrayList list = dataMap.get(key);
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(0));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(1));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(2));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(3));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(4));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(5));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(6));
                        setCellData(workbook, dataRow , cellIndex++ ,list.get(7));
                       
                        
                    }      
                    else // 若沒有此場站資料，填0
                    {
                              
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        setCellData(workbook, dataRow , cellIndex++ ,0);
                        
                    }
                    
                }
            } 
            
        }    
        dataMap.clear();
        
        
        
        //Sum
        cellIndex = 0;
        XSSFRow sum_dataRow = sheet.createRow(rowIndex++);
        XSSFCell sumCell = sum_dataRow.createCell(0);
        sumCell.setCellStyle(cellStyle4);
        sumCell.setCellValue("合計");
        
        cellIndex ++;
        Map<String ,ArrayList<Object>> sum_dataMap = Operatring_Sum_dataMap(operating_id , datalist);
        
        for(String deptId : depotMap.keySet())
        {
            String[] depot = deptId .split(",");
                
            if(depot[0].equals(operating_id)) // 限定業者代碼
            {
                if(sum_dataMap.containsKey(deptId))
                {
                    ArrayList<Object> list1 = sum_dataMap.get(deptId);
                    for(Object obj : list1)
                    {
                        sumCell = sum_dataRow.createCell(cellIndex);
                        if(cellIndex % 8 == 0)
                        {
                            sumCell.setCellStyle(cellStyle3);
                        }
                        else
                        {
                            sumCell.setCellStyle(cellStyle2);
                        }
                        sumCell.setCellValue(( Integer)obj );

                        cellIndex++;
                    }               
                }
            }
        }
        sum_dataMap.clear();
        

    }
    
   
    

     
     private void setCellData(XSSFWorkbook  workbook , XSSFRow row , int index , Object obj)
   {
       //Excel樣式
        XSSFCellStyle cellStyle,cellStyle1 = null;
        XSSFFont font = null;
        XSSFDataFormat intFormat = workbook.createDataFormat();      
        XSSFColor blue=new XSSFColor(new java.awt.Color(0,0,255));
        
        //一般儲存格(上下左右邊框 + 靠右 + 藍色字體)
        cellStyle=workbook.createCellStyle();
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));
        
        //分界儲存格(上下左細邊框 + 右粗邊框 + 靠右 + 藍色字體)
        cellStyle1=workbook.createCellStyle();
        cellStyle1.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
        cellStyle1.setBorderTop(BorderStyle.THIN);
        cellStyle1.setBorderBottom(BorderStyle.THIN);
        cellStyle1.setBorderLeft(BorderStyle.THIN);
        cellStyle1.setBorderRight(BorderStyle.MEDIUM);
        cellStyle1.setDataFormat(intFormat.getFormat("#,###,###,###,##0"));
        
        font=workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short)10);
        font.setColor(blue);
        cellStyle.setFont(font);
        cellStyle1.setFont(font);
       
        int cellrow = index++;
        XSSFCell cell = row.createCell(cellrow);
        if(cellrow % 8 == 0)
        {
            cell.setCellStyle(cellStyle1);
        }
        else
        {
            cell.setCellStyle(cellStyle);
        }
       
        
        
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
