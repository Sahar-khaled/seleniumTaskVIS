package com.vois.autotask;
import org.apache.commons.mail.*;

public class SendMail {

    public static void main(String[] args) throws EmailException {
        System.out.println("================ Test Started ==================");

        MultiPartEmail email = new MultiPartEmail();
        email.setHostName("smtp.gmail.com");
        email.setSmtpPort(465);
        email.setAuthenticator(new DefaultAuthenticator("sr6617133@gmail.com", "nmoslawvqwvrafso"));
        email.setSSLOnConnect(true);
        email.setFrom("sr6617133@gmail.com");
        email.setSubject("TestMail");
        email.setMsg("This is a test mail from selenium sahar");
        email.addTo("mohamed.almokadem@vodafone.com");

        // Create the attachment
        EmailAttachment excelSheet = new EmailAttachment();
        excelSheet.setPath("C:\\Users\\DELL\\Downloads\\sheet\\Data.xlsx"); // path of Excel sheet
        excelSheet.setDisposition(EmailAttachment.ATTACHMENT);
        excelSheet.setDescription("Excel test ");
        excelSheet.setName("Data.xlsx");
        email.attach(excelSheet);
        email.send();
        System.out.println("================ email sent ==================");

    }
}