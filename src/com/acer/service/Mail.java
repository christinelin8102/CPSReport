/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.acer.service;

import java.io.UnsupportedEncodingException;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;
import org.apache.log4j.Logger;

/**
 *
 * @author Christine
 */
public class Mail 
{
    private static Logger log = Logger.getLogger(Mail.class);
    
     public boolean SendMailByGmailSMTPWithAttachments(final String gmailAccId, final String gmailAccPwd, String mailSubject, String bodyMessage, String fromMail, String fromMailTitle, String receiveMail[], String[] attachments) throws UnsupportedEncodingException {
		String smtp_host = "smtp.gmail.com";
		String smtp_port = "587";
		
		Properties _sessionProperties = new Properties();
		_sessionProperties.put("mail.smtp.auth", "true");		
		_sessionProperties.put("mail.smtp.starttls.enable", "true");  
		_sessionProperties.put("mail.smtp.host", smtp_host);  
		_sessionProperties.put("mail.smtp.port", smtp_port);

	    Session session = Session.getInstance(_sessionProperties,new javax.mail.Authenticator() 
            {
	        protected PasswordAuthentication getPasswordAuthentication()
                {
	            return new PasswordAuthentication (gmailAccId, gmailAccPwd);
	        }
	    });
		
		//_sessionProperties.put("Mail.smtp.host", "127.0.0.1");
	
		try {
			MimeMessage mimemsg = new MimeMessage(session);
                        
			mimemsg.setSentDate(new java.util.Date()); //傳送日期
                        
			mimemsg.setFrom(new InternetAddress(fromMail, fromMailTitle, "UTF-8")); //寄件者Mail
                        
                        InternetAddress[] ToAddressArr = new InternetAddress[receiveMail.length];
                        for(int i = 0 ; i < receiveMail.length ; i++)
                        {
                            ToAddressArr[i] = new InternetAddress(receiveMail[i]);
                        }
			mimemsg.setRecipients(Message.RecipientType.TO, ToAddressArr); //收件者Mail(可以多個)
                        
			mimemsg.setSubject(mailSubject, "UTF-8"); //信件主旨
			BodyPart messageBodyPart = new MimeBodyPart(); 
			messageBodyPart.setContent(bodyMessage, "text/html; charset=utf-8"); 
			Multipart multipart = new MimeMultipart();  
			multipart.addBodyPart(messageBodyPart);
			if(attachments != null && attachments.length > 0) {
				for(int i = 0; i < attachments.length; i++) {
					addAttachment(multipart, attachments[i]);	
				}
			}
			mimemsg.setContent(multipart); 			
			//mimemsg.setContent(bodyMessage, "text/html");
			//mimemsg.setText(bodyMessage, "UTF-8");
			// Send message
			Transport.send(mimemsg);
                        log.info("Send Mail Successful !");
		} 
                catch (MessagingException e) 
                {
			//LogUtil.logWithDateAndFileNameDebugMode("SendMailByGmailSMTP ==> "+e.getMessage(), ToolsUtil.LOG_SETTLEMENT_FILE_NAME);			
			e.printStackTrace();
			//System.out.println(fromMail+ ", "+receiveMail);
                        log.info("Send Mail failed about message setting ! Reason: " + e);
			return false;      
		} 
                catch (UnsupportedEncodingException e) 
                {
			//LogUtil.logWithDateAndFileNameDebugMode("SendMailByGmailSMTP ==> "+e.getMessage(), ToolsUtil.LOG_SETTLEMENT_FILE_NAME);			
			e.printStackTrace();
			//System.out.println(fromMail+ ", "+receiveMail+" ==> unsupport Exception");
                        log.info("Send Mail failed about Encoding ! Reason: " + e);
			return false;			
		}		
		return true;
                
	}
	
	
	private void addAttachment(Multipart multipart, String filename) throws UnsupportedEncodingException
	{
	    DataSource source = new FileDataSource(filename);
	    BodyPart messageBodyPart = new MimeBodyPart();        
	    try 
            {
                messageBodyPart.setDataHandler(new DataHandler(source));
                
                //解決中文亂碼問題
                messageBodyPart.setHeader("Content-Type",  "application/octet-stream; charset=\"utf-8\"");
                messageBodyPart.setFileName(MimeUtility.encodeText(source.getName(), "UTF-8", "B"));
                
                multipart.addBodyPart(messageBodyPart);			
            } catch (MessagingException e) 
            {
                    e.printStackTrace();
            }
	}	
    
}
