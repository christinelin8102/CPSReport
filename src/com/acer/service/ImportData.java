
package com.acer.service;

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
public class ImportData 
{
    private static Logger log = Logger.getLogger(ImportData.class);
    
    
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
 
        if(srcZipList.length == 0)
        {
            log.error("No zip files");
        }
        else
        {
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
                File targetfile = new File(backupFolder); 
                if(!targetfile.exists())
                {
                    targetfile.mkdir();
                }
                backup(srcFolder,backupFolder);
            }
        }
        
        
    }
    
    
    
    public void backup(String srcFolder, String targetFolder)
    {
        File srcf = new File(srcFolder);
        File[] fl = srcf.listFiles();
        boolean backupSuccessful = false;
        
        File pathParent ;
        
        for(File files : fl)
        {           
            String opId = files.getName().substring(0, 3);  //業者代號                                  
            File renFile = new File(targetFolder + File.separator + opId + File.separator + files.getName());              
            
            pathParent = new File(renFile.getParent());
            
            if(!pathParent.exists())
            {
                pathParent.mkdirs();
            }
            
            try
            {
                if( renFile.exists() == true )
                {
                    renFile.delete();
                }
                
                backupSuccessful = files.renameTo( renFile );
                log.info(files.getName() + ", Backup Successful !");
            }
            catch(Exception backupError)
            {
                log.error("Backup Error ! Reason: " +  backupError);
            }
            
        }
        
        if(backupSuccessful == true)
        {
            for(File files : fl)
            {
                files.delete();
            }
        }
        
    }

    
    // 取得檔名清單
    public List<String>  getFilesList(String csvImportPath)
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
       
    
    public void ImportData_Rev(PreparedStatement ps, String fileName, String csvImportPath) throws UnsupportedEncodingException, FileNotFoundException, IOException, SQLException //營收
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
            log.info("Now import revenue file is : " + fileName + ", Import data successfully ! ");
        }
        catch(Exception importError)
        {
            log.error(" Import revenue file error ! Reason :" + importError);
        }
              
    }
    
    public void ImportData_Unc(PreparedStatement ps, String fileName, String csvImportPath) throws UnsupportedEncodingException, FileNotFoundException, IOException, SQLException //未清分
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
                    ps.setString(i++, csvList[j++]); //清分日期
                    ps.setString(i++, csvList[j++]); //票證公司
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 卡片扣款金額(CPS全部)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 社福優惠金額(CPS全部)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 轉乘優惠金額(CPS全部)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 已上傳未清分筆數
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 卡片扣款金額(已上傳未清分)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 社福優惠金額(已上傳未清分)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 轉乘優惠金額(已上傳未清分)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 清分錯誤筆數
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 卡片扣款金額(清分錯誤)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 社福優惠金額(清分錯誤)
                    ps.setInt(i++, Integer.valueOf(csvList[j++])); // 轉乘優惠金額(清分錯誤)
                    ps.setString(i++, sdf_detail.format(Calendar.getInstance().getTime())); //更新時間
                            
                    ps.addBatch();
                }
        
                ps.executeBatch();                       
                ps.clearBatch();
                lines++;     
            }
            reader.close();
            br.close();
            log.info("Now import uncleared file is : " + fileName + ", Import data successfully ! ");
        }
        catch(Exception importError)
        {
            log.error("Import uncleared file error ! Reason :" + importError);
        }
              
    }
    
          
}
