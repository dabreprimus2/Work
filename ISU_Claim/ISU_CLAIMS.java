import java.beans.PropertyVetoException;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.math.BigInteger;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.nio.ByteBuffer;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
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

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ibm.as400.access.AS400;
import com.ibm.as400.access.AS400Message;
import com.ibm.as400.access.AS400SecurityException;
import com.ibm.as400.access.ErrorCompletingRequestException;
import com.ibm.as400.access.ObjectDoesNotExistException;
import com.ibm.as400.access.ProgramCall;
import com.ibm.as400.access.ProgramParameter;

public class ISU_CLAIMS{
	private static int errorcount=0;
	Properties props=null;
	/**
	 * the current library we are using 
	 */
	private static String LIB;
	/**
	 * the current client we are working on
	 */
	private static String CLIENT;
	/**
	 * global var for the inbound file that is being read from
	 */
	private static File from,out,emailfile;
	/**
	 * a flag to see if there is an error or not
	 */
	@SuppressWarnings("unused")
	private static boolean errorfound=false;
	/**
	 * this list reads in all the plans, finds the head plan then sorts all dependents under them
	 * in order of their dependent codes
	 */

	private static StringBuilder SB = new StringBuilder();
	/**
	 * Arraylist that holds the data from the inbound file
	 */
	private static ArrayList<String[]> lines =new ArrayList<String[]>();

	/**
	 * DB connection var
	 */
	private static Connection conn;
	public static void main (String args[]) throws IOException, PropertyVetoException, AS400SecurityException, ErrorCompletingRequestException, InterruptedException, ObjectDoesNotExistException{  	 
		//get input from command line
		if(args.length==1){
			LIB="HTHDATV1";
			CLIENT="IS1";
		}
		else{
			LIB="HTHDATV1";
			CLIENT="IS1";
		}
		new ISU_CLAIMS(LIB, CLIENT);
	}    

	public ISU_CLAIMS(String LIB, String CLIENT) throws IOException, PropertyVetoException, AS400SecurityException, ErrorCompletingRequestException, InterruptedException, ObjectDoesNotExistException{
		ISU_CLAIMS.LIB=LIB;
		ISU_CLAIMS.CLIENT=CLIENT;

		try {
			System.out.println("Checking to See if ISU file exisits");

			BufferedReader bi = new BufferedReader(new FileReader("C:/Users/ROB/Desktop/Primus/ISU_Claims/BCBSNE-BlueFlex-ClaimsAnalysisSummaryFile-20181023.txt"));

		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block

			//SB.append("File was not uploaded or is in an incorrect file format. Please make sure the file is named correctly and in CSV format. Example: 'Test.xlsx'");
			System.out.println(SB.toString());
			Sendemail(true,"NO");
			return;
		} 
	
		long startTime = System.currentTimeMillis();	 
		System.out.println("Connecting to the Database");	 
		String hostname="";//getting the host name of the computer that we are using 
		try {
			hostname = InetAddress.getLocalHost().getHostName();
		}catch (UnknownHostException ex){System.out.println("Hostname can not be resolved");}


		props = new Properties();//get DB info
		try {
			if (hostname.equals("PRIMUS")) { 
				System.out.println("Here");
				props.load(new FileInputStream("C:/Users/ROB/git/Hi-Tech-Health/mydb2.properties"));}
			else
				props.load(new FileInputStream("/java/mydb2.properties"));
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		

		readXLS_withblanks(out,false,0,6);



/*
		//callrpg();

		if (errorfound){
			System.out.println(SB.toString());
			Sendemail(false,"YES");
			System.out.println("DONE in: " + ((System.currentTimeMillis()-startTime)/ 1000)+"secs");
			return;
		}

		else
		{
			
			Sendemail(true,"YES");
			System.out.println("DONE in: " + ((System.currentTimeMillis()-startTime)/ 1000)+"secs");
			return;
		}*/
		
	}

	public static void readXLS_withblanks(File f, Boolean SkipFirstLine, int sheetIndex, int LineLength) throws IOException{
		System.out.println("Reading in File to Check\n\n");
		// For storing data into CSV files

		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("C:/Users/ROB/Desktop/Primus/ISU_Claims/isu_test06112016.xlsx")); 
		XSSFSheet spreadsheet = workbook.getSheet("Sheet1");
		
		String[] headers = new String[] {"NB_Group","Company/Client","Date","Amount1","Amount2","Total"};
		
		Row row6 = spreadsheet.createRow(0);
		
		
		for (int rn=0; rn<headers.length; rn++){
			row6.createCell(rn).setCellValue(headers[rn]);
		}		


		XSSFCellStyle style1 = workbook.createCellStyle();
		style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(155,194,230))); //Blue
		style1.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(169,208,142))); //Green
		style2.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style3 = workbook.createCellStyle();
		style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(244,176,132))); //Orange
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style4 = workbook.createCellStyle();
		style4.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,217,102))); //Gold
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style5 = workbook.createCellStyle();
		style5.setFillForegroundColor(new XSSFColor(new java.awt.Color(245,139,139))); //Red
		style5.setFillPattern(CellStyle.SOLID_FOREGROUND);

		Font font = workbook.createFont();
		font.setColor(IndexedColors.BLACK.getIndex());
		style1.setFont(font);
		style2.setFont(font);
		style3.setFont(font);
		style4.setFont(font);
		style5.setFont(font);
		
		try {


			//Create a blank sheet
			int i=1;
			String ii="";

			// Iterate through each rows from first sheet
			//boolean skipped=false;


			@SuppressWarnings("resource")
			BufferedReader bufferreader = new BufferedReader(new FileReader("C:/Users/ROB/Desktop/Primus/ISU_Claims/BCBSNE-BlueFlex-ClaimsAnalysisSummaryFile-20181023.txt"));
		
			String next, line = bufferreader.readLine();
			
			String NB_Group="";
			String Company="";
			String date="";
			String Amount1="";
			String Amount2="";
			String Total="";
			
			while((line = bufferreader.readLine()) != null)
			{
				NB_Group = line.substring(0,9).trim(); //9
				Company=line.substring(9,58).trim(); //32
				date=line.substring(59,63).trim(); //6
				Amount1=line.substring(63,78).trim(); //15
				Amount2=line.substring(78,93).trim(); //5
				Total=line.substring(93,108).trim(); //5
				System.out.println("Group :"+NB_Group);
				System.out.println("Company :"+Company);
				System.out.println("date :"+date);
				System.out.println("Amount1 :"+Amount1);
				System.out.println("Amount2 :"+Amount2);
				System.out.println("Total :"+Total);
				
				
				
				
				
				row6 = spreadsheet.createRow(i);

				Cell cell0 = row6.createCell(0);
				Cell cell1 = row6.createCell(1);
				Cell cell2 = row6.createCell(2);
				Cell cell3 = row6.createCell(3);
				Cell cell4 = row6.createCell(4);
				Cell cell5 = row6.createCell(5);
				
				cell0.setCellValue(NB_Group);
				cell1.setCellValue(Company);
				cell2.setCellValue(date);
				cell3.setCellValue(Amount1);
				cell4.setCellValue(Amount2);
				cell5.setCellValue(Total);
			
				if (NB_Group.length()>15 ) //Group Number Check
				{
					errorfound=true;
					errorcount++;
					cell0.setCellStyle(style1);
				}
			
			/*	int found = Integer.parseInt(DBquery("SELECT COUNT(*) FROM QTEMP/GRPMS2_"+CLIENT+" WHERE G2ALID='"+NB_Group+"' OR G2RPNO='"+NB_Group+"' ","GRPMS2"));
				
				if (found==0)
				{
					errorfound=true;
					errorcount++;
					cell0.setCellStyle(style1);
				}*/
				
				if (!isDateValid(date))
				{
					errorfound=true;
					errorcount++;
					cell2.setCellStyle(style2);
				}
				if (Total != Amount1 + Amount2)
				{
					errorfound=true;
					errorcount++;
					cell5.setCellStyle(style3);
				}
			}
		}catch (Exception ioe) {ioe.printStackTrace();}
		
		FileOutputStream outfile = new FileOutputStream(new File("C:/Users/ROB/Desktop/Primus/ISU_Claims/new.xlsx"));
		workbook.write(outfile);
		outfile.close();
		emailfile=new File("C:/Users/ROB/Desktop/Primus/ISU_Claims/new.xlsx");
		System.out.println("Writesheet.xlsx written successfully");
		
	}
	private static boolean isDateValid(String val){
		Calendar c = Calendar.getInstance();
		int year = c.get(Calendar.YEAR);
		
		String todayyear=Integer.toString(year);
		String yeartoday=todayyear.substring(2,4);
		year=Integer.parseInt(yeartoday);
		String month1 = "";
		String year1 = "";
		int yearIn=0;
		int monthIn=0;
		try {
			year1 = val.substring(2,3);
			yearIn=Integer.parseInt(val);

		} catch (Exception e) {

			//e.printStackTrace();
		}
		if(year<yearIn){return false;}
		else {return true;}
	}
	private static String DBquery(String query, String file){	
		Statement stmt = null;
		String result="";

		try {
			stmt = conn.createStatement();
			ResultSet rs = stmt.executeQuery(query);
			System.out.println(rs);
			while (rs.next()) 
				result = rs.getString(1);

			rs.close();				
			stmt.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}

		return result;
	}
	private static void Sendemail(boolean status,String file)
	{
		System.out.println("Sending Email\n\n");

		SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yy");
		Date today = Calendar.getInstance().getTime();
		String date=sdf.format(today);


		Properties props = new Properties();

		props.put("mail.smtp.auth", "true");

		props.put("mail.smtp.starttls.enable", "true");

		props.put("mail.smtp.host", "secure.emailsrvr.com");

		props.put("mail.smtp.port", "587");

		props.put("mail.debug", "false");

		Session session = Session.getInstance(props,

				new Authenticator()
		{
			protected PasswordAuthentication getPasswordAuthentication() 
			{
				return new PasswordAuthentication("cbeyer@hi-techhealth.com", "14752369Cb");
			}
		});

		try 
		{

			PrintWriter writer=null;

			File f=new File("Error_Log.txt");

			if (!f.exists())

				f.createNewFile();

			writer = new PrintWriter(f);

			writer.print(SB.toString());


			writer.close();

			//String EMAIL = Findemail("SELECT CAEML1 FROM QTEMP/BLOCK_"+CLIENT);

			Multipart multipart = new MimeMultipart();

			MimeMessage msg = new MimeMessage(session);

			MimeBodyPart textPart = new MimeBodyPart();

			MimeBodyPart messageBodyPart = new MimeBodyPart(); 

			msg.setFrom(new InternetAddress("support@hi-techhealth.com"));

			String recipientEmail = ("dabreprimus2@gmail.com");//twalsh@hi-techhealth.com
			String ccEmail=("cbeyer@hi-techhealth.com, jeller@hi-techhealth.com"); //twalsh@hi-techhealth.com
			String[] recipientList = recipientEmail.split(",");
			InternetAddress[] recipientAddress = new InternetAddress[recipientList.length];
			int counter = 0;
			for (String recipient : recipientList) {
				recipientAddress[counter] = new InternetAddress(recipient.trim());
				counter++;
			}
			msg.setRecipients(Message.RecipientType.TO, recipientAddress);

			//set CC address
			if (ccEmail.length() > 0) {
				String[] CCList = ccEmail.split(",");
				InternetAddress[] CCAddress = new InternetAddress[CCList.length];
				int c = 0;
				for (String CC : CCList) {
					CCAddress[c] = new InternetAddress(CC.trim());
					c++;
				}
				msg.setRecipients(Message.RecipientType.BCC, CCAddress);
			}

			msg.setSubject("File Upload");

			msg.setSentDate(new Date());


			MimeBodyPart messageBodyPart3 = new MimeBodyPart();
			DataSource source3 = new FileDataSource(emailfile);
			messageBodyPart3.setDataHandler( new DataHandler(source3));
			messageBodyPart3.setFileName("Errors.xlsx");

			String textContent = "";

			if(status)
			{
				textContent += "Your file submitted on "+date+" has successfully uploaded.\n";
				textContent +="\r\n******** Confidentiality Statement ************\n\nThis email, including any attachments, is for the sole use of the intended recipient and may contain confidential and privileged information.\r\nAny unauthorized review, use, disclosure or distribution is strictly prohibited.\r\nIf you are not the intended recipient please contact the sender via email and destroy all copies of the original message.";
			}
			else
			{ 
				if(file.equals("NO"))
				{
					textContent += "Your file submitted on "+date+" has failed to upload.  The file was not found or is in an incorrect format.\n";
					textContent +="\r\n******** Confidentiality Statement ************\n\nThis email, including any attachments, is for the sole use of the intended recipient and may contain confidential and privileged information.\r\nAny unauthorized review, use, disclosure or distribution is strictly prohibited.\r\nIf you are not the intended recipient please contact the sender via email and destroy all copies of the original message.";
				}
				else
				{
					textContent += "Your file submitted on "+date+" has failed to upload.  Please review the Excel File attached to this email, correct the errors indicated, and resubmit.\n";
					textContent +="\r\n******** Confidentiality Statement ************\n\nThis email, including any attachments, is for the sole use of the intended recipient and may contain confidential and privileged information.\r\nAny unauthorized review, use, disclosure or distribution is strictly prohibited.\r\nIf you are not the intended recipient please contact the sender via email and destroy all copies of the original message.";
					multipart.addBodyPart(messageBodyPart3);
				}
				//multipart.addBodyPart(messageBodyPart2);
			}

			textPart.setText(textContent);

			multipart.addBodyPart(textPart);

			msg.setContent(multipart);

			Transport.send(msg);

			System.out.println("---SENT---\n");

		}

		catch (MessagingException mex) 
		{

			mex.printStackTrace();
		} 

		catch (IOException e1) {

			e1.printStackTrace();

		}
		System.out.println("Program Closed");
		System.exit(0);

	}  

}
