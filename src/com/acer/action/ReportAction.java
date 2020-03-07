/*
 * 5/4 Christine修改報表檔名及Mail主旨。 A1
 *
 *
 */

package com.acer.action;

import com.acer.bean.RevenueReportBean;
import com.acer.bean.UnclearedReportBean;
import com.acer.service.ImportData;
import com.acer.service.Mail;
import com.acer.service.RevenueReport;
import com.acer.service.UnclearedReport;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.Properties;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;


/**
 *
 * @author Christine
 */
public class ReportAction 
{
    private static Logger log = Logger.getLogger(ReportAction.class);
    private static UnclearedReport _unclearedReport;
    
    public static void main(String args[]) throws FileNotFoundException, IOException, ClassNotFoundException
    {
        
        
        PropertyConfigurator.configure (System.getProperty("user.dir") +  File.separator +"RevenueLog.properties");	
        
        
//        Connection conn = null;
        Properties info = new Properties();
        Properties prop = new Properties(System.getProperties());
   
        InputStream input = null;     
        input = new FileInputStream("config.properties");      
        prop.load(input);
        
        
        String driver= prop.getProperty("Driver"); 
        String url = prop.getProperty("URL");  
        info.put("user",prop.getProperty("User"));
        info.put("password", prop.getProperty("Password"));
        
        String receiveCsvZip_rev= prop.getProperty("ReceiveCsvZip_rev"); //上傳的壓縮檔
        String csvImportPath_rev= prop.getProperty("CsvImportPath_rev"); //匯入DB MySQL的資料
        String csvZipBackup_rev= prop.getProperty("CsvZipBackup_rev"); //匯入資料後備份
        
        String receiveCsvZip_unc= prop.getProperty("ReceiveCsvZip_unc"); //上傳的壓縮檔
        String csvImportPath_unc= prop.getProperty("CsvImportPath_unc"); //匯入DB MySQL的資料
        String csvZipBackup_unc= prop.getProperty("CsvZipBackup_unc"); //匯入資料後備份
        
        

        // Import Data to database
          
        Connection conn =null;
        PreparedStatement ps =null;
        
        ImportData importTest_rev= null;
        List<String> fileList_rev= null;
        String sql_rev = null;
        
        if(prop.getProperty("Switch_Rev").equals("0"))
        {
            log.info("Revenue unzip　start !");
            importTest_rev = new ImportData(); //營收
            sql_rev = "replace into revenue_report( operating_id, operating_name, business_date, depot_id, depot_name, trx_count, txamt, personal_count, personal_txamt, transfer_MRT_count, transfer_MRT_txamt, transfer_bus_count, transfer_bus_txamt, update_time) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)";   
            importTest_rev.unzipFiles(receiveCsvZip_rev, csvImportPath_rev, csvZipBackup_rev); //解壓縮
            fileList_rev = importTest_rev.getFilesList(csvImportPath_rev); //產生檔案名單
        }
        
        
        ImportData importTest_unc= null;
        List<String> fileList_unc= null;
        String sql_unc= null;
        
        if(prop.getProperty("Switch_Unc").equals("0"))
        {
            log.info("Uncleared unzip　start !");
            importTest_unc = new ImportData(); //未清分
            sql_unc = "replace into uncleared_report ( operating_id, operating_name, cal_date, unit_id, txamt_all, personal_txamt_all, transfer_txamt_all, uncleared_count_c, txamt_c, personal_txamt_c, transfer_txamt_c, uncleared_count_e, txamt_e, personal_txamt_e, transfer_txamt_e, update_time) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";         
            importTest_unc.unzipFiles(receiveCsvZip_unc, csvImportPath_unc, csvZipBackup_unc); //解壓縮
            fileList_unc = importTest_unc.getFilesList(csvImportPath_unc); //產生要匯入的檔案名單
        }
        
    
        try
        {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, info);           
            
            if(prop.getProperty("Switch_Rev").equals("0"))
            {
                //營收           
                if(fileList_rev.isEmpty())
                {
                    log.error("The revenue data csv folder has no files ! ");
                }
                else // Data匯入DB MySQL
                {
                    log.info("The revenue data csv import starts ! ");
                    for(String fileName : fileList_rev)
                    {    
                        ps=conn.prepareStatement(sql_rev); 

                        importTest_rev.ImportData_Rev(ps, fileName, csvImportPath_rev); //匯入資料                   

                        ps.close();
                    }

                }
            }
            
            
            if(prop.getProperty("Switch_Unc").equals("0"))
            {
                //未清分
            
                if(fileList_unc.isEmpty())
                {
                    log.error("The uncleared data csv folder has no files ! ");
                }
                else // Data匯入DB MySQL
                {
                    log.info("The uncleared data csv import starts ! ");
                    for(String fileName : fileList_unc)
                    {    
                        ps=conn.prepareStatement(sql_unc); 

                        importTest_unc.ImportData_Unc(ps, fileName, csvImportPath_unc); //匯入資料                   

                        ps.close();
                    }

                }
            }
        
            conn.close();

        }
        catch (SQLException connectProblem) 
        {
            log.error("doesn't connect! The problem is: "+connectProblem);
        }
        finally
        {
            if( ps != null )
            {
                try 
                {                      
                    ps.close();
                } 
                catch (SQLException preparedStatementProblem)
                {
                    log.error("The PreparedStatement doesn't close! The problem is: "+preparedStatementProblem);
                }
            }

            if( conn != null )
            {
                try 
                {
                    conn.close();
                } 
                catch (SQLException connectionProblem) 
                {
                    log.error("The connection doesn't close! The problem is: "+connectionProblem);
                }
            }
        }
        
        
        

        // Generate Report
        SimpleDateFormat sdf=new SimpleDateFormat("YYYYMMdd");
        SimpleDateFormat sdf_YYYYMM =new SimpleDateFormat("YYYYMM");
        SimpleDateFormat sdf_MM =new SimpleDateFormat("MM");
        Calendar day =Calendar.getInstance();      
        String today=null;
        String YYYYMM=null; //今年今月
        String MM=null;
        String BM=null;
        today =sdf.format(day.getTime());
        YYYYMM = sdf_YYYYMM.format(day.getTime());
        MM =sdf_MM.format(day.getTime());
        
        
        
        
        try
        {
            Class.forName(driver);//DriverManager.registerDriver(new com.mysql.jdbc.Driver()) 或是 Class.forName("com.mysql.jdbc.Driver");
            conn = DriverManager.getConnection(url, info);
            
            day.add(Calendar.MONTH, -1);
            BM = sdf_MM.format(day.getTime());
            
            if(prop.getProperty("Switch_Rev").equals("0"))
            {
                //營收報表
                log.info("Generate revenue report start !");
                RevenueReport revenueReport = new RevenueReport(conn);  
                try
                {
                    List<RevenueReportBean> listA = revenueReport.queryTxnList();      
    /* A1 當每月1日，檔名顯示前一個月份*/ 
                    //檔名
                    if(today.equals(YYYYMM+"01"))
                    {                        
                        revenueReport.setFileName("TCPS北區業者"+BM+"月[帳務統計]報表_"+today+"_V1.0.xlsx");
                    }
                    else
                    {
                        revenueReport.setFileName("TCPS北區業者"+MM+"月[帳務統計]報表_"+today+"_V1.0.xlsx");
                    }  

                    revenueReport.setExportPath(prop.getProperty("RevReportPath"));                
                    revenueReport.generateExcel(listA);
                }
                catch(Exception error)
                {
                    log.error("Generate revenue report failed! Reason: " + error);
                }
            }
            
            
            if(prop.getProperty("Switch_Unc").equals("0"))
            {
                //未清分報表
                log.info("Generate uncleared report start !");
                _unclearedReport = new UnclearedReport(conn);
                try
                {
                    List<UnclearedReportBean> listB = _unclearedReport.queryTxnList();

                    if(today.equals(YYYYMM+"01"))
                    {
                        _unclearedReport.setFileName("TCPS北區業者"+BM+"月[未清分統計]報表_"+today+"_V1.0.xlsx");
                    }
                    else if(today.equals(YYYYMM+"02")) //因未清分報表是查詢前兩天的資料
                    {
                        _unclearedReport.setFileName("TCPS北區業者"+BM+"月[未清分統計]報表_"+today+"_V1.0.xlsx");
                    }
                    else
                    {
                        _unclearedReport.setFileName("TCPS北區業者"+MM+"月[未清分統計]報表_"+today+"_V1.0.xlsx");
                    }  

                    _unclearedReport.setExportPath(prop.getProperty("UncReportPath"));
                    _unclearedReport.generateExcel(listB);
                }
                catch(Exception ex)
                {
                    log.error("Generate uncleared report failed! Reason: " + ex);
                }
            }
 
        }
        catch(Exception connError)
        {
            log.error("Connection error! Reason: " + connError);
        }
        finally
        {
            if( conn != null )
            {
                try 
                {
                    conn.close();
                } 
                catch (SQLException ex) 
                {
                    System.out.println(ex);
                    log.error("Connection can't close! Reason: " + ex);
                }
            }
        }
        
        

        // Send Mail       
        String gmailAccId="cpsReportFrom@gmail.com"; //寄件者信箱(授權)
        String gmailAccPwd="cpsreport"; //寄件者密碼(授權)
              
        if(prop.getProperty("Switch_Rev_Mail").equals("0")) //Send Revenue Report Mail
        {         
            //營收
            log.info("Send revenue report mail start !");
            
            String mailSubject_rev= null;            
            /* A1 當每月1日，檔名顯示前一個月份*/         
            //檔名
            if(today.equals(YYYYMM+"01"))
            {
                mailSubject_rev="TCPS北區業者"+BM+"月[帳務統計]報表_"+today+"_V1.0"; //訊息主旨
            }
            else
            {
                mailSubject_rev="TCPS北區業者"+MM+"月[帳務統計]報表_"+today+"_V1.0"; //訊息主旨
            }

            String bodyMessage="Dears: Please check the detail, thks!"; //訊息內容
            String fromMail="cpsReportFrom@gmail.com"; //寄件者
            String fromMailTitle="TCPS_Report"; //發信人(顯示的文字)
            String[] receiveMail={"christinelin8102@gmail.com","Christine.Lin@acer.com","Nick.Chen@acer.com","ying7933@gmail.com","Tweety.Chen@acer.com","cpsreportall@gmail.com"}; //收件者(可以設定多個)
            //,"Christine.Lin@acer.com","Nick.Chen@acer.com","ying7933@gmail.com","Tweety.Chen@acer.com","cpsreportall@gmail.com"
            String[] attachments={"C:\\TCPS_Report\\RevenueReport\\"+mailSubject_rev+".xlsx"}; //附加檔案
            
    //        //參數: final String gmailAccId, final String gmailAccPwd, String mailSubject, String bodyMessage, String fromMail, String fromMailTitle, String receiveMail, String[] attachments           
            Mail mail = new Mail();
            mail.SendMailByGmailSMTPWithAttachments(gmailAccId, gmailAccPwd, mailSubject_rev, bodyMessage, fromMail, fromMailTitle, receiveMail, attachments);
                    
        }

        
        if(prop.getProperty("Switch_Unc_Mail").equals("0")) //Send Uncleared Report Mail
        {
            //未清分
            log.info("Send uncleared report mail start!");
            
            String mailSubject_unc= null;
            if(today.equals(YYYYMM+"01"))
            {
                mailSubject_unc="TCPS北區業者"+BM+"月[未清分統計]報表_"+today+"_V1.0";
            }
            else if(today.equals(YYYYMM+"02"))
            {
                mailSubject_unc="TCPS北區業者"+BM+"月[未清分統計]報表_"+today+"_V1.0"; //訊息主旨
            }
            else
            {
                mailSubject_unc="TCPS北區業者"+MM+"月[未清分統計]報表_"+today+"_V1.0";
            }

            StringBuffer mailContent = _unclearedReport.getMailContent();

            String bodyMessage_unc= "Dear all: <br>" +mailContent; //訊息內容
            String fromMail_unc= "cpsReportFrom@gmail.com"; //寄件者
            String fromMailTitle_unc= "TCPS_Report"; //發信人(顯示的文字)
            String[] receiveMail_unc= {"christinelin8102@gmail.com","Christine.Lin@acer.com" , "cpsreportall@gmail.com" , "Jonny.Wu@acer.com" , "Sam.Hs.Wang@acer.com" , "Tweety.Chen@acer.com" , "Bow.Su@acer.com" , "Hsieh.Ken@acer.com" , "Jerry.chao@acer.com" ,"Nick.Chen@acer.com","ying7933@gmail.com", "Dacy.Yang@acer.com" , "Jack.Chen@acer.com" , "Sandy.Hsu@acer.com" , "Terry.T.Chang@acer.com" , "Casper.Chan@acer.com" , "Kathy.Lu@acer.com" , "Odin.Liao@acer.com" ,"Angela.Tsai@acer.com"}; //收件者(可以設定多個)
            //主寄: ,"Christine.Lin@acer.com" , "cpsreportall@gmail.com" , "Jonny.Wu@acer.com" , "Sam.Hs.Wang@acer.com" , "Tweety.Chen@acer.com" , "Bow.Su@acer.com" , "Hsieh.Ken@acer.com" , "Jerry.chao@acer.com" ,"Nick.Chen@acer.com","ying7933@gmail.com", "Dacy.Yang@acer.com" , "Jack.Chen@acer.com" , "Sandy.Hsu@acer.com" , "Terry.T.Chang@acer.com" , "Casper.Chan@acer.com" , "Kathy.Lu@acer.com" , "Odin.Liao@acer.com" ,"Angela.Tsai@acer.com"
            String[] attachments_unc= {"C:\\TCPS_Report\\UnclearedReport\\"+mailSubject_unc+".xlsx"}; //附加檔案

            Mail mail_unc = new Mail();
            mail_unc.SendMailByGmailSMTPWithAttachments(gmailAccId, gmailAccPwd, mailSubject_unc, bodyMessage_unc, fromMail_unc, fromMailTitle_unc, receiveMail_unc, attachments_unc);
        }
        
   
     }        
}
    
    
    
   
    
    
    

