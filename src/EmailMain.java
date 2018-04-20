import java.io.IOException;
import java.util.*;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;

public class EmailMain {

	public static void main(String[] args) throws IOException {
		  // Recipient's email ID
	      String to = "recipient@email.com";

	      // Sender's email ID
	      String from = "sender@email.com";

	      // Using Gmail SMTP server
	      String host = "smtp.gmail.com";

	      // Get system properties
	      Properties properties = System.getProperties();

	      // Setup mail server
	      properties.setProperty("mail.smtp.host", host);
	      properties.put("mail.smtp.auth", "true");
	      properties.put("mail.smtp.port", "465");
	      properties.put("mail.smtp.starttls.enable", "true"); 
	      properties.put("mail.smtp.socketFactory.class","javax.net.ssl.SSLSocketFactory");  
	      properties.put("mail.smtp.socketFactory.fallback", "false");  

	      // Get the default Session object.
	      Session session = Session.getDefaultInstance(properties,  
	    		    new javax.mail.Authenticator() {
	    		       protected PasswordAuthentication getPasswordAuthentication() {  
	    		       return new PasswordAuthentication("sender@email.com", "password");  
	    		   } 
	    	  });

	      try {
	         // Create a default MimeMessage object.
	         MimeMessage message = new MimeMessage(session);

	         // Set From: header field of the header.
	         message.setFrom(new InternetAddress(from));

	         // Set To: header field of the header.
	         message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

	         // Set Subject: header field
	         message.setSubject("This is a test email program");

	         // Now set the actual message
	         message.setText("This is a test message");

	         
	         //Create Excel File
	         CreateExcelFile file = new CreateExcelFile();
	         file.writeXLSXFile();
	         
	         //Set attachment
	         BodyPart messageBody = new MimeBodyPart();  
	         messageBody.setText("This is a test email program");  
	         
	         Multipart multipart = new MimeMultipart();  
	         multipart.addBodyPart(messageBody); 
	         
	         messageBody = new MimeBodyPart();
	         message.setContent(multipart);  
	         
	         DataSource source = new FileDataSource("test.xlsx");
	         messageBody.setDataHandler(new DataHandler(source)); 
	         
	         messageBody.setFileName("test.xlsx");
	         multipart.addBodyPart(messageBody);
	         
	         message.setContent(multipart);
	     
	         // Send message & attachment
	         Transport.send(message);
	         System.out.println("Sent message successfully....");
	         
	         
	      } catch (MessagingException mex) {
	         mex.printStackTrace();
	      }
		
		
	}

}
