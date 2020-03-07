
package com.acer.action;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Properties;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

/**
 *
 * @author Christine
 */
public class ImportTest 
{
    private static Logger log = Logger.getLogger(ImportTest.class);
    
    
    SimpleDateFormat sdf = new SimpleDateFormat("YYYYMMdd");
    SimpleDateFormat sdf_detail = new SimpleDateFormat("YYYYMMdd hh:mm:ss");
    Calendar day = Calendar.getInstance();
    String today = sdf.format(day.getTime());
    boolean unzipSuccessful=false;
    

        
    public void unzipFiles(String srcFolder, String targetFolder, String backupFolder)
    {
        File srcZip = new File(srcFolder);
        File[] srcZipList = srcZip.listFiles();
            
        byte[] buffer = new byte[1024];
 
        try
        {
            //create output directory is not exists
            File folder = new File(targetFolder + File.separator + today);
            if( !folder.exists())
            {
    		folder.mkdir();
            }
 
            for(File zipFileName : srcZipList)
            {
                ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFileName)); //get the zip file content               
                ZipEntry ze = zis.getNextEntry(); //get the zipped file list entry     
 
                while(ze != null)
                { 
                    String fileName = ze.getName();
                    File outputFile = new File( folder + File.separator + fileName );
                    
                    log.info("File Unzip : "+ outputFile.getAbsoluteFile());
 
                    //create all non exists folders
                    //else you will hit FileNotFoundException for compressed folder
                
                    new File(outputFile.getParent()).mkdirs();
 
                    FileOutputStream fos = new FileOutputStream(outputFile);             
 
                    int len;
                    while ((len = zis.read(buffer)) > 0) 
                    {
                        fos.write(buffer, 0, len);
                    }
 
                    fos.close();   
                    ze = zis.getNextEntry();
                }
 
                zis.closeEntry();
                zis.close();
                     
            }
            
            log.info("Unzip Successfully !");
            unzipSuccessful=true;
            
 
        }
        catch(IOException unzipError)
        {
            log.error("Unzip Error ! Reason is : " + unzipError);
        }
        
        // Zip備份
        if(unzipSuccessful=true)
        {
            backup(srcFolder,backupFolder);
        }
    }
    
    
    
    public void backup(String srcFolder, String targetFolder)
    {
        File srcf = new File(srcFolder);
        File[] fl = srcf.listFiles();
        
        for(File files : fl)
        {           
            String opId = files.getName().substring(0, 3);  //業者代號                                  
            File backup = new File(targetFolder + File.separator + opId + File.separator + files.getName());              
            
            try
            {
                files.renameTo(backup);
                log.info(files.getName() + ", Backup Successful !");
            }
            catch(Exception backupError)
            {
                backup.getParentFile().mkdir();
                log.error("Backup Error ! Reason: " +  backupError);
            }
            
        }
    }

    
    // 取得檔名清單
    List<String>  getFilesList(String csvImportPath)
    {
        List<String> fileList = new ArrayList<String>();
        try
        {    
            File folder = new File(csvImportPath + File.separator + today);            
            String[] list = folder.list();           
            for(String fileName : list)
            {
                fileList.add(fileName);
            }
        }
        catch(Exception e)
        {
            log.error("'"+csvImportPath + File.separator + today +"' the folder is not exist.");
        }
        
        return fileList;
    }
       
    
    public void ImportData(PreparedStatement ps, String fileName, String csvImportPath) throws UnsupportedEncodingException, FileNotFoundException, IOException, SQLException
    {
        String csvData=null;
        int lines=1;
        
        File inputFile=new File(csvImportPath + File.separator + today + File.separator + fileName);  
        InputStreamReader reader = new InputStreamReader (new FileInputStream(inputFile),"MS950");
        BufferedReader br =new BufferedReader(reader);
            
        try      
        {
            while((csvData=br.readLine()) != null) //行數
            {                                  
//              System.out.print("第"+lines+"行 ");                            
                String[] csvList = csvData.split(",");
                       
                if(lines > 2) //3行以後印出
                {                    
                    int i = 1;
                    int j = 0;
                    ps.setString(i++, csvList[j++]); //業者代碼
                    ps.setString(i++, csvList[j++]); //業者名稱
                    ps.setString(i++, csvList[j++]); //營運日期
                    ps.setString(i++, csvList[j++]); //場站代碼
                    ps.setString(i++, csvList[j++]); //場站名稱
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //卡片扣款筆數
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //卡片扣款金額
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //社福旅次量
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //社福金額
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //捷運轉乘筆數
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //捷運轉乘金額
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //公車轉乘筆數
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); //公車轉乘金額
                    ps.setString(i++, sdf_detail.format(Calendar.getInstance().getTime())); //更新時間
                            
                    ps.addBatch();
                }
        
                ps.executeBatch();                       
                ps.clearBatch();
                lines++;     
            }
            reader.close();
            br.close();
            log.info("Now Import file is : " + fileName + ", Import Data Successfully ! ");
        }
        catch(Exception importError)
        {
            log.error(", Import Error ! Reason :" + importError);
        }
              
    }
    
    public static void main(String args[]) throws ClassNotFoundException, UnsupportedEncodingException, IOException 
    {
        PropertyConfigurator.configure (System.getProperty("user.dir") +  File.separator +"RevenueLog.properties");	
        
        
        // Get Connection and Directory Path Information from " config.properties "
        Properties prop = new Properties(System.getProperties());
        InputStream input = null;
        input = new FileInputStream("config.properties");
        prop.load(input);
        
        String driver= prop.getProperty("Driver"); 
        String url= prop.getProperty("URL");    //jdbc:mysql://localhost:3306/CPS_Report
        String user= prop.getProperty("User"); 
        String password= prop.getProperty("Password");
        
        String receiveCsvZip= prop.getProperty("ReceiveCsvZip"); //上傳的壓縮檔
        String csvImportPath= prop.getProperty("CsvImportPath"); //匯入DB MySQL的資料
        String csvZipBackup= prop.getProperty("CsvZipBackup"); //匯入資料後備份
  
        // Start Mission
        ImportTest importTest = new ImportTest();
             
        Connection conn =null;
        PreparedStatement ps =null;
        String sql = "replace into revenue_report( operating_id, operating_name, business_date, depot_id, depot_name, trx_count, txamt, personal_count, personal_txamt, transfer_MRT_count, transfer_MRT_txamt, transfer_bus_count, transfer_bus_txamt, update_time) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        
        importTest.unzipFiles(receiveCsvZip, csvImportPath, csvZipBackup); //解壓縮
        
        List<String> fileList = importTest.getFilesList(csvImportPath); //產生檔案名單
        
        if(fileList.isEmpty())
        {
            log.error("The folder has no files ! ");
        }
        else // Data匯入DB MySQL
        { 
            try
            {
                Class.forName(driver);
                conn = DriverManager.getConnection(url, user,password);           
                    
                for(String fileName : fileList)
                {    
                    ps=conn.prepareStatement(sql); 
                    
                    importTest.ImportData(ps, fileName, csvImportPath); //匯入資料                   
                    
                    ps.close();
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
        }          
     }       
}
