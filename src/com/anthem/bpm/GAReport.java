package com.anthem.bpm;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GAReport {
  static String excelFilePath = "GA_Report.xlsx";
  
  public static void main(String[] args) throws ClassNotFoundException, SQLException, IOException {
    Class.forName("oracle.jdbc.driver.OracleDriver");
    Connection con = DriverManager.getConnection("jdbc:oracle:thin:@//fnetpengn-p-01.internal.das:1525/fnetcep", "SRCDATABASE", "ECMstores1#");
    Statement stmt = con.createStatement();
    String sql = "select U5F_DCN,CREATE_DATE,U7B_CLAIMNUM,U85_MEMBERID from WESTCLAIMS.DOCVERSION where OBJECT_CLASS_ID = '3B451C4C08178344A310BF81B70AD703' and trunc(CREATE_DATE) >((sysdate)-4) and U71_ROUTECODE = 'CD'";
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("G&A");
    writeHeaderLineGa(sheet);
    ResultSet rs = stmt.executeQuery(sql);
    writeDataLines(rs, workbook, sheet);
    stmt.close();
    con.close();
    sendEmail(workbook);
  }
  
  public static void sendEmail(XSSFWorkbook workbook) throws IOException {
    String[] recipient = { "mason.mele@anthem.com", "SriSaiVenkataMounica.Kollipara@anthem.com", "hariprasad.pasham@anthem.com", "Srivastava.Anurag@anthem.com", "sathishkumar.jayakanthan@anthem.com", "SaiKrishna.Chivukula@legato.com" };
    //String[] recipient = {"sakya.samanta@anthem.com"};
	  String[] recipient1 = { "DL-FileNetLightsOnSupport@anthem.com", "venkateswara.surapaneni@anthem.com", "karthik.rajaraman@anthem.com", "dayakar.malgireddy@amerigroup.com", "mamatha.chapidi@amerigroup.com", "paul.samuel@anthem.com" };
    DateFormat df = new SimpleDateFormat("ddMMM");
    String fileName = df.format(new Date());
    fileName = fileName.concat(".xlsx");
    String sender = "DL-FileNetLightsOnSupport@anthem.com";
    String host = "smtp.wellpoint.com";
    Properties properties = System.getProperties();
    properties.setProperty("mail.smtp.host", host);
    Session session = Session.getDefaultInstance(properties);
    try {
      MimeMessage message = new MimeMessage(session);
      message.setFrom((Address)new InternetAddress(sender));
      int i;
      for (i = 0; i < recipient.length; i++)
        message.addRecipient(Message.RecipientType.TO, (Address)new InternetAddress(recipient[i]));
      for (i = 0; i < recipient1.length; i++)
        message.addRecipient(Message.RecipientType.CC, (Address)new InternetAddress(recipient1[i]));
      message.setSubject("G&A Reconciliation ");
      StringBuilder builder = new StringBuilder();
      builder.append("<html><body>Hi Team,<br><br>Please find the attachment<br>");
      builder.append("<br>Regards,<br>FileNet LightsOn Support Team </body></html>");
      String result = builder.toString();
      FileOutputStream outputStream = new FileOutputStream(excelFilePath);
      workbook.write(outputStream);
      DataSource source = new FileDataSource("GA_Report.xlsx");
      MimeBodyPart mbp1 = new MimeBodyPart();
      MimeBodyPart mbp2 = new MimeBodyPart();
      mbp1.setContent(result, "text/html");
      mbp2.setDataHandler(new DataHandler(source));
      mbp2.setFileName(fileName);
      MimeMultipart mimeMultipart = new MimeMultipart();
      mimeMultipart.addBodyPart((BodyPart)mbp1);
      mimeMultipart.addBodyPart((BodyPart)mbp2);
      message.setContent((Multipart)mimeMultipart);
      message.saveChanges();
      Transport.send((Message)message);
      System.out.println("Mail successfully sent");
    } catch (MessagingException mex) {
      mex.printStackTrace();
    } 
  }
  
  public static void writeHeaderLineGa(XSSFSheet sheet) {
    XSSFRow xSSFRow = sheet.createRow(0);
    Cell headerCell = xSSFRow.createCell(0);
    headerCell.setCellValue("U5F_DCN");
    headerCell = xSSFRow.createCell(1);
    headerCell.setCellValue("CREATE_DATE");
    headerCell = xSSFRow.createCell(2);
    headerCell.setCellValue("U7B_CLAIMNUM");
    headerCell = xSSFRow.createCell(3);
    headerCell.setCellValue("U85_MEMBERID");
  }
  
  private static void writeDataLines(ResultSet rs, XSSFWorkbook workbook, XSSFSheet sheet) throws SQLException {
    int rowCount = 1;
    while (rs.next()) {
      String dcn = rs.getString(1);
      Timestamp createDate = rs.getTimestamp(2);
      String claimNum = rs.getString(3);
      String memberId = rs.getString(4);
      XSSFRow xSSFRow = sheet.createRow(rowCount++);
      int columnCount = 0;
      Cell cell = xSSFRow.createCell(columnCount++);
      cell.setCellValue(dcn);
      cell = xSSFRow.createCell(columnCount++);
      XSSFCellStyle xSSFCellStyle = workbook.createCellStyle();
      XSSFCreationHelper xSSFCreationHelper = workbook.getCreationHelper();
      xSSFCellStyle.setDataFormat(xSSFCreationHelper.createDataFormat().getFormat("MM/dd/yyyy HH:mm:ss"));
      cell.setCellStyle((CellStyle)xSSFCellStyle);
      cell.setCellValue(createDate);
      cell = xSSFRow.createCell(columnCount++);
      cell.setCellValue(claimNum);
      cell = xSSFRow.createCell(columnCount);
      cell.setCellValue(memberId);
      for (int i = 0; i < 4; i++)
        sheet.autoSizeColumn(i); 
    } 
  }
}
