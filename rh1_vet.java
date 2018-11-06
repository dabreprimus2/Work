import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
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

import java.text.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileConversionXLSToXLXS {
	private static int errorcount=0;
	private static Connection conn = null;
	public static String path="", grp="";
	private static File out,emailfile;;
	private static StringBuilder SB = new StringBuilder();
	private static boolean errorfound=false;
	private static String LIB;
	private static String CLIENT;

	static Properties props = new Properties();//get DB info
	public static void main(String[] args) throws ParseException, IOException {

		LIB="HTHDATV1";
		CLIENT="RH1";
		connect();
		FileConversionXLSToXLXS fileConversionXLSToXLXS = new FileConversionXLSToXLXS();
		String xlsFilePath = "/hthjav1/RH1_Vet/inbound/TEST_W0621_20181023.xls";
		String xlsxFilePath = fileConversionXLSToXLXS.convertXLS2XLSX(xlsFilePath);
		out=new File(xlsxFilePath);
		readXLS_withblanks(out,true,0,66);
		if(!errorfound)
		{
			DBinsert("DELETE FROM "+LIB+"/VETELG","");
			insert(out,true,0,66);
			Sendemail(true,"NO");
		}
		else
		{
			Sendemail(false,"YES");
		}
	}

	public String convertXLS2XLSX(String xlsFilePath) {
		Map cellStyleMap = new HashMap();
		String xlsxFilePath = null;
		Workbook workbookIn = null;
		File xlsxFile = null;
		Workbook workbookOut = null;
		OutputStream out = null;
		String XLSX = ".xlsx";
		try {
			InputStream inputStream = new FileInputStream(xlsFilePath);
			xlsxFilePath = xlsFilePath.substring(0, xlsFilePath.lastIndexOf('.')) + XLSX;
			workbookIn = new HSSFWorkbook(inputStream);
			xlsxFile = new File(xlsxFilePath);
			if (xlsxFile.exists())
				xlsxFile.delete();
			workbookOut = new XSSFWorkbook();
			int sheetCnt = workbookIn.getNumberOfSheets();

			for (int i = 0; i < sheetCnt; i++) {
				Sheet sheetIn = workbookIn.getSheetAt(i);
				Sheet sheetOut = workbookOut.createSheet(sheetIn.getSheetName());
				Iterator rowIt = sheetIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = (Row)rowIt.next();
					Row rowOut = sheetOut.createRow(rowIn.getRowNum());
					copyRowProperties(rowOut, rowIn,cellStyleMap);
				}
			}
			out = new BufferedOutputStream(new FileOutputStream(xlsxFile));
			workbookOut.write(out);
			System.out.println("File converted from xls to xlsx successfully");
		} catch (Exception ex) {
			System.err.println("Exception Occured inside transFormXLS2XLSX :: file Name :: " + xlsFilePath
					+ ":: reason ::" + ex.getMessage());
			ex.printStackTrace();
			xlsxFilePath = null;
		}
		return xlsxFilePath;

	}

	public static void readXLS_withblanks(File f, Boolean SkipFirstLine, int sheetIndex, int LineLength) throws IOException, ParseException{

		System.out.println("Reading in File to Check\n\n");
		// For storing data into CSV files
		ArrayList<String[]> data = new ArrayList<String[]>();

		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("/hthjav1/RH1_Vet/inbound/TEST_W0621_20181023.xlsx")); 
		XSSFSheet spreadsheet = workbook.getSheet("Sheet1");// main sheet
       // XSSFSheet sheet2 = workbook.createSheet("Sheet2");//color coding sheet
		String[] headers = new String[] {"Emp_SSN", "First_Name", "Last_Name", "MI",	"Birth_Date","Gender","Relationship","Address_1","Address_2","City","State","Zip code","Effective date",
				"tEmp_Hire_Date","Emp_Rehire_Date","Term date","Last day worked","Emp_Occupation","Emp_Hours_Per_Week","Emp_ID",	"Emp_Export_ID","Emp_Department","Emp_Location",
				"Emp_Area","Emp_Payroll_Class","Emp_UDClass_1","Emp_UDClass_2","Emp_UDClass_3","Emp_UDClass_4","Emp_UDClass_5","Emp_UD_Field",	"Emp_UDField_1","Emp_UDField_2",
				"Emp_UDField_3","Emp_UDField_4","Emp_UDField_5","Emp_Annual_Salary","Emp_Annual_Salary 2","Emp_Flex_Dollars","Emp_Demographic_Comment","Benefit_Category","Product_Code","Product_Desc","Group_Number",
				"Benefit_Code",	"Action_Code","Action_Desc","Effective_Date","Activity_Date","Termination_Date","Event_Date","Signed_Date",	"Change_Date","Coverage_Tier_Code",	
				"Coverage_Tier_Desc","Reson_Code","Reason_Desc","Provider_Code","Monthly_Premium","PerPay_Emp_Cost","Total_Annual_Volume","Approved_Annual_Volume","Volume_Factor",
				"Enroll_Comment", "Emp_Pay_Frequency","Person_System_Number"};
		Row row67 = spreadsheet.createRow(0);
	/*	// to write in sheet2
		Map<String, Object[]> data2 = new TreeMap<String, Object[]>();
		data2.put("1",new Object[] {"indicates the error with the SSN"});
		data2.put("2",new Object[] {"indicates the error with the gender field"});
		data2.put("3",new Object[] {"indicates the error with the gender field"});
		data2.put("4",new Object[] {"indicates the error with the Zip code or State field"});
		data2.put("5",new Object[] {"indicates the size of the data in corresponding field is greater than expected"});
		
		Set<String> keyset = data2.keySet(); 
        int rownum = 0; 
        for (String key : keyset) { 
            // this creates a new row in the sheet 
            Row row100 = sheet2.createRow(rownum++); 
            Object[] objArr = data2.get(key); 
            int cellnum = 0; 
            for (Object obj : objArr) { 
                // this line creates a cell in the next column of that row 
                Cell cell100 = row100.createCell(cellnum++); 
                if (obj instanceof String) 
                    cell100.setCellValue((String)obj); 
                else if (obj instanceof Integer) 
                    cell100.setCellValue((Integer)obj); 
            } 
		
        }*/
		for (int rn=0; rn<headers.length; rn++){
			row67.createCell(rn).setCellValue(headers[rn]);
		}

		XSSFCellStyle style1 = workbook.createCellStyle();
		style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(155,194,230))); //ssn
		style1.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(169,208,142))); //DOB
		style2.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style3 = workbook.createCellStyle();
		style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(244,176,132))); //gender
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style4 = workbook.createCellStyle();
		style4.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,217,102))); //Gold
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style5 = workbook.createCellStyle();
		style5.setFillForegroundColor(new XSSFColor(new java.awt.Color(245,139,139))); //state & zip
		style5.setFillPattern(CellStyle.SOLID_FOREGROUND);

		XSSFCellStyle style6 = workbook.createCellStyle();
		style5.setFillForegroundColor(new XSSFColor(new java.awt.Color(249,240,127))); //otherdates
		style5.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		
		
		Font font = workbook.createFont();
		font.setColor(IndexedColors.BLACK.getIndex());
		style1.setFont(font);
		style2.setFont(font);
		style3.setFont(font);
		style4.setFont(font);
		style5.setFont(font);

		try {
		// Get the workbook object for XLSX file
		XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(f));
		// Get first sheet from the workbook
		XSSFSheet sheet = wBook.getSheetAt(sheetIndex);
		//Create a blank sheet
		int i=1;
		String ii="";
		// Iterate through each rows from first sheet
		boolean skipped=false;
		for(Row row : sheet) {
			if (SkipFirstLine && !skipped) {
				skipped=true;
				continue;
			}

			String[] line=new String[LineLength];
			for(int cn=0; cn<row.getLastCellNum(); cn++) {

				DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
				// If the cell is missing from the file, generate a blank one
				// (Works by specifying a MissingCellPolicy)
				Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				line[cn] = formatter.formatCellValue(cell); 

				// Write the output to a file
			}

			data.add(line);

			row67 = spreadsheet.createRow(i);

			Cell cell0 = row67.createCell(0);
			Cell cell1 = row67.createCell(1);
			Cell cell2 = row67.createCell(2);
			Cell cell3 = row67.createCell(3);
			Cell cell4 = row67.createCell(4);
			Cell cell5 = row67.createCell(5);
			Cell cell6 = row67.createCell(6);
			Cell cell7 = row67.createCell(7);
			Cell cell8 = row67.createCell(8);
			Cell cell9 = row67.createCell(9);
			Cell cell10 = row67.createCell(10);
			Cell cell11 = row67.createCell(11);
			Cell cell12 = row67.createCell(12);
			Cell cell13 = row67.createCell(13);
			Cell cell14 = row67.createCell(14);
			Cell cell15 = row67.createCell(15);
			Cell cell16 = row67.createCell(16);
			Cell cell17 = row67.createCell(17);
			Cell cell18 = row67.createCell(18);
			Cell cell19 = row67.createCell(19);
			Cell cell20 = row67.createCell(20);
			Cell cell21 = row67.createCell(21);
			Cell cell22 = row67.createCell(22);
			Cell cell23 = row67.createCell(23);
			Cell cell24 = row67.createCell(24);
			Cell cell25 = row67.createCell(25);
			Cell cell26 = row67.createCell(26);
			Cell cell27 = row67.createCell(27);
			Cell cell28 = row67.createCell(28);
			Cell cell29 = row67.createCell(29);
			Cell cell30 = row67.createCell(30);
			Cell cell31 = row67.createCell(31);
			Cell cell32 = row67.createCell(32);
			Cell cell33 = row67.createCell(33);
			Cell cell34 = row67.createCell(34);
			Cell cell35 = row67.createCell(35);
			Cell cell36 = row67.createCell(36);
			Cell cell37 = row67.createCell(37);
			Cell cell38 = row67.createCell(38);
			Cell cell39 = row67.createCell(39);
			Cell cell40 = row67.createCell(40);
			Cell cell41 = row67.createCell(41);
			Cell cell42 = row67.createCell(42);
			Cell cell43 = row67.createCell(43);
			Cell cell44 = row67.createCell(44);
			Cell cell45 = row67.createCell(45);
			Cell cell46 = row67.createCell(46);
			Cell cell47 = row67.createCell(47);
			Cell cell48 = row67.createCell(48);
			Cell cell49 = row67.createCell(49);
			Cell cell50 = row67.createCell(50);
			Cell cell51 = row67.createCell(51);
			Cell cell52 = row67.createCell(52);
			Cell cell53 = row67.createCell(53);
			Cell cell54 = row67.createCell(54);
			Cell cell55 = row67.createCell(55);
			Cell cell56 = row67.createCell(56);
			Cell cell57 = row67.createCell(57);
			Cell cell58 = row67.createCell(58);
			Cell cell59 = row67.createCell(59);
			Cell cell60 = row67.createCell(60);
			Cell cell61 = row67.createCell(61);
			Cell cell62 = row67.createCell(62);
			Cell cell63 = row67.createCell(63);
			Cell cell64 = row67.createCell(64);
			Cell cell65 = row67.createCell(65);

			cell0.setCellValue(line[0]);
			cell1.setCellValue(line[1]);
			cell2.setCellValue(line[2]);
			cell3.setCellValue(line[3]);
			cell4.setCellValue(line[4]);
			cell5.setCellValue(line[5]);
			cell6.setCellValue(line[6]);
			cell7.setCellValue(line[7]);
			cell8.setCellValue(line[8]);
			cell9.setCellValue(line[9]);
			cell10.setCellValue(line[10]);
			cell11.setCellValue(line[11]);
			cell12.setCellValue(line[12]);
			cell13.setCellValue(line[13]);
			cell14.setCellValue(line[14]);
			cell15.setCellValue(line[15]);
			cell16.setCellValue(line[16]);
			cell17.setCellValue(line[17]);
			cell18.setCellValue(line[18]);
			cell19.setCellValue(line[19]);
			cell20.setCellValue(line[20]);
			cell21.setCellValue(line[21]);
			cell22.setCellValue(line[22]);
			cell23.setCellValue(line[23]);
			cell24.setCellValue(line[24]);
			cell25.setCellValue(line[25]);
			cell26.setCellValue(line[26]);
			cell27.setCellValue(line[27]);
			cell28.setCellValue(line[28]);
			cell29.setCellValue(line[29]);
			cell30.setCellValue(line[30]);
			cell31.setCellValue(line[31]);
			cell32.setCellValue(line[32]);
			cell33.setCellValue(line[33]);
			cell34.setCellValue(line[34]);
			cell35.setCellValue(line[35]);
			cell36.setCellValue(line[36]);
			cell37.setCellValue(line[37]);
			cell38.setCellValue(line[38]);
			cell39.setCellValue(line[39]);
			cell40.setCellValue(line[40]);
			cell41.setCellValue(line[41]);
			cell42.setCellValue(line[42]);
			cell43.setCellValue(line[43]);
			cell44.setCellValue(line[44]);
			cell45.setCellValue(line[45]);
			cell46.setCellValue(line[46]);
			cell47.setCellValue(line[47]);
			cell48.setCellValue(line[48]);
			cell49.setCellValue(line[49]);
			cell50.setCellValue(line[50]);
			cell51.setCellValue(line[51]);
			cell52.setCellValue(line[52]);
			cell53.setCellValue(line[53]);
			cell54.setCellValue(line[54]);
			cell55.setCellValue(line[55]);
			cell56.setCellValue(line[56]);
			cell57.setCellValue(line[57]);
			cell58.setCellValue(line[58]);
			cell59.setCellValue(line[59]);
			cell60.setCellValue(line[60]);
			cell61.setCellValue(line[61]);
			cell62.setCellValue(line[62]);
			cell63.setCellValue(line[63]);
			cell64.setCellValue(line[64]);
			cell65.setCellValue(line[65]);

			//	Cell cell67=row.createCell(i);


			if (line[1].length()>25)
			{  
				System.out.println("error"+line[1]);
				errorfound=true;
				errorcount++;
				cell1.setCellStyle(style4);
			}
			if (line[2].length()>25)
			{  System.out.println(line[2]);
			errorfound=true;
			errorcount++;
			cell2.setCellStyle(style4);
			}
			if (line[7].length()>30)
			{  System.out.println(line[7]);
			errorfound=true;
			errorcount++;
			cell7.setCellStyle(style4);
			}

			if (!isSSNValid(line[0]))
			{
				System.out.println("Error with SSN "+line[0]);
				errorfound=true;
				errorcount++;
				cell0.setCellStyle(style1);
			}

			if(line[11].length()>5)
			{

				line[11]=line[11].substring(0, 5);
			}

			if (!isAddValid(line[10],line[11]))
			{
				errorfound=true;
				errorcount++;
				cell10.setCellStyle(style5);
				cell11.setCellStyle(style5);
			}

			if (!isDateValid(line[4]) && !line[4].equals(""))
			{
				System.out.println("error date 1");
				errorfound=true;
				errorcount++;
				cell10.setCellStyle(style2);
			}

			if (!isDateValid2(line[12])&& !line[12].equals(""))
			{
				System.out.println("error date 2");

				errorfound=true;
				errorcount++;
				cell12.setCellStyle(style5);

			}	
			if (!isDateValid2(line[13])&& !line[13].equals(""))
			{
				System.out.println("error date 3");

				errorfound=true;
				errorcount++;
				cell13.setCellStyle(style5);

			}	

			if (!isDateValid2(line[14])&& !line[14].equals(""))
			{
				System.out.println("error date 4");

				errorfound=true;
				errorcount++;
				cell14.setCellStyle(style5);

			}	
			if (!isDateValid2(line[15])&& !line[15].equals(""))
			{
				System.out.println("error date 5");

				errorfound=true;
				errorcount++;
				cell15.setCellStyle(style5);

			}	
			if (!isDateValid2(line[47])&& !line[47].equals(""))
			{
				System.out.println("error date 6");

				errorfound=true;
				errorcount++;
				cell47.setCellStyle(style5);

			}	
			if (!isDateValid2(line[48])&& !line[48].equals(""))
			{
				System.out.println("error date 7");

				errorfound=true;
				errorcount++;
				cell48.setCellStyle(style5);

			}
			if (!isDateValid2(line[49])&& !line[49].equals(""))
			{
				System.out.println("error date 8");

				errorfound=true;
				errorcount++;
				cell49.setCellStyle(style5);

			}	
			if (!isDateValid2(line[50])&& !line[50].equals(""))
			{
				System.out.println("error date 9");

				errorfound=true;
				errorcount++;
				cell50.setCellStyle(style5);

			}	
			if (!isDateValid2(line[51])&& !line[51].equals(""))
			{
				System.out.println("error date 10");

				errorfound=true;
				errorcount++;
				cell51.setCellStyle(style5);

			}	
			if (!isDateValid2(line[52])&& !line[52].equals(""))
			{
				System.out.println("error date 11");

				errorfound=true;
				errorcount++;
				cell52.setCellStyle(style5);

			}	


			if (!isGenValid(line[5]))
			{
				errorfound=true;
				errorcount++;
				cell5.setCellStyle(style3);
			}


			i++;
		}
	} catch (Exception ioe) {ioe.printStackTrace();}

		//Write the workbook in file system
		FileOutputStream outfile = new FileOutputStream(new File("/hthjav1/RH1_Vet/inbound/TEST_W0621_20181023.xlsx"));
		workbook.write(outfile);
		outfile.close();
		emailfile=new File("/hthjav1/RH1_Vet/inbound/TEST_W0621_20181023.xlsx");
		System.out.println("Writesheet.xlsx written successfully");

		System.out.println("File Compiled with "+errorcount+" errors");

	}


	private static void insert(File f, Boolean SkipFirstLine, int sheetIndex, int LineLength) throws FileNotFoundException, IOException
	{
		// Get the workbook object for XLSX file
		XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(f));
		// Get first sheet from the workbook
		XSSFSheet sheet = wBook.getSheetAt(sheetIndex);
		//Create a blank sheet
		int i=1;
		String ii="";

		// Iterate through each rows from first sheet
		boolean skipped=false;
		for(Row row : sheet) {
			if (SkipFirstLine && !skipped) {
				skipped=true;
				continue;
			}

			String[] line=new String[LineLength];
			for(int cn=0; cn<row.getLastCellNum(); cn++) {

				DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
				// If the cell is missing from the file, generate a blank one
				// (Works by specifying a MissingCellPolicy)
				Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				line[cn] = formatter.formatCellValue(cell).toUpperCase(); 
				
				line[cn]=line[cn].replaceAll("'"," ");
				// Write the output to a file
			}

			line[4]=isDateValid3(line[4],"B");//dob

			line[11]=line[11].replaceAll("-", "");//zip replace

			line[12]=isDateValid3(line[12],"");//date
			line[13]=isDateValid3(line[13],"");//date
			line[14]=isDateValid3(line[14],"");//date
			line[15]=isDateValid3(line[15],"");//date
			line[16]=isDateValid3(line[16],"");//date
			line[47]=isDateValid3(line[47],"");//date
			line[48]=isDateValid3(line[48],"");//date
			line[49]=isDateValid3(line[49],"");//date
			line[50]=isDateValid3(line[50],"");//date
			line[51]=isDateValid3(line[51],"");//date
			line[52]=isDateValid3(line[52],"");//date

			line[36]=isDateValid3(line[36],"D");//annual salary
			line[37]=isDateValid3(line[37],"D");//some amount
			line[38]=isDateValid3(line[38],"D");//some amount
			line[58]=isDateValid3(line[58],"D");//some amount
			line[59]=isDateValid3(line[59],"D");//some amount
			line[60]=isDateValid3(line[60],"D");//some amount
			line[61]=isDateValid3(line[61],"D");//some amount
			line[64]=isDateValid3(line[64],"D");//some amount

			line[18]=isDateValid3(line[18],"D");//hours per week

			line[40]=isDateValid3(line[40],"D");//rben

			
			
			try {
				line[3]=line[3].substring(0, 1);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//e.printStackTrace();
			}

			if(line[53].equals("EMP"))
			{
				line[53]="1";
			}
			if(line[53].equals("ECH"))
			{
				line[53]="6";
			}
			if(line[53].equals("FAM"))
			{
				line[53]="3";
			}
			if(line[53].equals("ESP"))
			{
				line[53]="2";
			}			


			/*System.out.println("INSERT INTO CBLIB/EVERELG (RSSN,RFNAM,RLNAM,RMI,RDOB,RSEX,RREL,RADD1,RADD2,RCITY,RSTATE,RZIP,REFFD,REMPH,REMPRE,RTERMD,RLASTD,REMPOC,RHPW,REMPID,REMPXID,REDEPRT,RLCTN,RAREA,"
					+ "RPAYCL,RUDCL1,RUDCL2,RUDCL3,RUDCL4,RUDCL5,RUDFLD,RUDFLD1,RUDFLD2,RUDFLD3,RUDFLD4,RUDFLD5,RSALRY,RSALRY2,RFLEX,RDEMGR,RBEN,RPRODC,RPRODD,RGRPNO,RBENCOD,RACTCOD,RACTDES,REFF2,RACTDTE,"
					+ "RTERMD2,REVTD,RSIGND,RCHGED,RCOV,RCOVDS,RRNCDE,RRNDES,RPROVC,RMNPRM,RPPEC,RANVOL,RAPVOL,RVOLFT,RENCOM,REMPFQ,RPSN) values('"+line[0]+"','"+line[1]+"','"+line[2]+"','"+line[3]+
					"',"+line[4]+",'"+line[5]+"','"+line[6]+"','"+line[7]+"','"+line[8]+"','"+line[9]+"','"+line[10]+"','"+line[11]+"',"+line[12]+","+line[13]+","+line[14]+","+line[15]+
					","+line[16]+",'"+line[17]+"',"+line[18]+",'"+line[19]+"','"+line[20]+"','"+line[21]+"','"+line[22]+"','"+line[23]+"','"+line[24]+"','"+line[25]+"','"+line[26]+"','"
					+line[27]+"','"+line[28]+"','"+line[29]+"','"+line[30]+"','"+line[31]+"','"+line[32]+"','"+line[33]+"','"+line[34]+"','"+line[35]+"',"+line[36]+","
					+line[37]+","+line[38]+",'"+line[39]+"',"+line[40]+",'"+line[41]+"','"+line[42]+"','"+line[43]+"','"+line[44]+"','"+line[45]+"','"+line[46]+"',"
					+line[47]+","+line[48]+","+line[49]+","+line[50]+","+line[51]+","+line[52]+",'"+line[53]+"','"+line[54]+"','"+line[55]+"','"
					+line[56]+"','"+line[57]+"','"+line[58]+"','"+line[59]+"','"+line[60]+"','"+line[61]+"','"+line[62]+"','"+line[63]+"',"
					+ "'"+line[64]+"','"+line[65]+"')");*/

			DBinsert("INSERT INTO "+LIB+"/VETELGHST (RSSN,RFNAM,RLNAM,RMI,RDOB,RSEX,RREL,RADD1,RADD2,RCITY,RSTATE,RZIP,REFFD,REMPH,REMPRE,RTERMD,RLASTD,REMPOC,RHPW,REMPID,REMPXID,REDEPRT,RLCTN,RAREA,"
					+ "RPAYCL,RUDCL1,RUDCL2,RUDCL3,RUDCL4,RUDCL5,RUDFLD,RUDFLD1,RUDFLD2,RUDFLD3,RUDFLD4,RUDFLD5,RSALRY,RSALRY2,RFLEX,RDEMGR,RBEN,RPRODC,RPRODD,RGRPNO,RBENCOD,RACTCOD,RACTDES,REFF2,RACTDTE,"
					+ "RTERMD2,REVTD,RSIGND,RCHGED,RCOV,RCOVDS,RRNCDE,RRNDES,RPROVC,RMNPRM,RPPEC,RANVOL,RAPVOL,RVOLFT,RENCOM,REMPFQ,RPSN) values('"+line[0]+"','"+line[1]+"','"+line[2]+"','"+line[3]+
					"',"+line[4]+",'"+line[5]+"','"+line[6]+"','"+line[7]+"','"+line[8]+"','"+line[9]+"','"+line[10]+"','"+line[11]+"',"+line[12]+","+line[13]+","+line[14]+","+line[15]+
					","+line[16]+",'"+line[17]+"',"+line[18]+",'"+line[19]+"','"+line[20]+"','"+line[21]+"','"+line[22]+"','"+line[23]+"','"+line[24]+"','"+line[25]+"','"+line[26]+"','"
					+line[27]+"','"+line[28]+"','"+line[29]+"','"+line[30]+"','"+line[31]+"','"+line[32]+"','"+line[33]+"','"+line[34]+"','"+line[35]+"',"+line[36]+","
					+line[37]+","+line[38]+",'"+line[39]+"',"+line[40]+",'"+line[41]+"','"+line[42]+"','"+line[43]+"','"+line[44]+"','"+line[45]+"','"+line[46]+"',"
					+line[47]+","+line[48]+","+line[49]+","+line[50]+","+line[51]+","+line[52]+",'"+line[53]+"','"+line[54]+"','"+line[55]+"','"
					+line[56]+"','"+line[57]+"','"+line[58]+"','"+line[59]+"','"+line[60]+"','"+line[61]+"','"+line[62]+"','"+line[63]+"',"
					+ "'"+line[64]+"','"+line[65]+"')"," ");

			DBinsert("INSERT INTO "+LIB+"/VETELG (RSSN,RFNAM,RLNAM,RMI,RDOB,RSEX,RREL,RADD1,RADD2,RCITY,RSTATE,RZIP,REFFD,REMPH,REMPRE,RTERMD,RLASTD,REMPOC,RHPW,REMPID,REMPXID,REDEPRT,RLCTN,RAREA,"
					+ "RPAYCL,RUDCL1,RUDCL2,RUDCL3,RUDCL4,RUDCL5,RUDFLD,RUDFLD1,RUDFLD2,RUDFLD3,RUDFLD4,RUDFLD5,RSALRY,RSALRY2,RFLEX,RDEMGR,RBEN,RPRODC,RPRODD,RGRPNO,RBENCOD,RACTCOD,RACTDES,REFF2,RACTDTE,"
					+ "RTERMD2,REVTD,RSIGND,RCHGED,RCOV,RCOVDS,RRNCDE,RRNDES,RPROVC,RMNPRM,RPPEC,RANVOL,RAPVOL,RVOLFT,RENCOM,REMPFQ,RPSN) values('"+line[0]+"','"+line[1]+"','"+line[2]+"','"+line[3]+
					"',"+line[4]+",'"+line[5]+"','"+line[6]+"','"+line[7]+"','"+line[8]+"','"+line[9]+"','"+line[10]+"','"+line[11]+"',"+line[12]+","+line[13]+","+line[14]+","+line[15]+
					","+line[16]+",'"+line[17]+"',"+line[18]+",'"+line[19]+"','"+line[20]+"','"+line[21]+"','"+line[22]+"','"+line[23]+"','"+line[24]+"','"+line[25]+"','"+line[26]+"','"
					+line[27]+"','"+line[28]+"','"+line[29]+"','"+line[30]+"','"+line[31]+"','"+line[32]+"','"+line[33]+"','"+line[34]+"','"+line[35]+"',"+line[36]+","
					+line[37]+","+line[38]+",'"+line[39]+"',"+line[40]+",'"+line[41]+"','"+line[42]+"','"+line[43]+"','"+line[44]+"','"+line[45]+"','"+line[46]+"',"
					+line[47]+","+line[48]+","+line[49]+","+line[50]+","+line[51]+","+line[52]+",'"+line[53]+"','"+line[54]+"','"+line[55]+"','"
					+line[56]+"','"+line[57]+"','"+line[58]+"','"+line[59]+"','"+line[60]+"','"+line[61]+"','"+line[62]+"','"+line[63]+"',"
					+ "'"+line[64]+"','"+line[65]+"')"," ");
		}
	}

	private static void DBinsert(String query, String file){	
		Statement stmt = null;

		try {
			stmt = conn.createStatement();
			stmt.execute(query);

			stmt.close();
		} catch (SQLException e) {
			e.printStackTrace();
			System.out.println(e.getMessage());
			System.exit(-1);
		}

	}

	public void copyRowProperties(Row rowOut, Row rowIn, Map cellStyleMap) {
		rowOut.setRowNum(rowIn.getRowNum());
		rowOut.setHeight(rowIn.getHeight());
		rowOut.setHeightInPoints(rowIn.getHeightInPoints());
		rowOut.setZeroHeight(rowIn.getZeroHeight());
		Iterator cellIt = rowIn.cellIterator();
		while (cellIt.hasNext()) {
			Cell cellIn = (Cell)cellIt.next();
			Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());
			rowOut.getSheet().setColumnWidth(cellOut.getColumnIndex(),
					rowIn.getSheet().getColumnWidth(cellIn.getColumnIndex()));
			copyCellProperties(cellOut, cellIn, cellStyleMap);
		}

	}

	public void copyCellProperties(Cell cellOut, Cell cellIn, Map cellStyleMap) {

		Workbook wbOut = cellOut.getSheet().getWorkbook();
		HSSFPalette hssfPalette = ((HSSFWorkbook) cellIn.getSheet().getWorkbook()).getCustomPalette();
		switch (cellIn.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;

		case Cell.CELL_TYPE_BOOLEAN:
			cellOut.setCellValue(cellIn.getBooleanCellValue());
			break;

		case Cell.CELL_TYPE_ERROR:
			cellOut.setCellValue(cellIn.getErrorCellValue());
			break;

		case Cell.CELL_TYPE_FORMULA:
			cellOut.setCellFormula(cellIn.getCellFormula());
			break;

		case Cell.CELL_TYPE_NUMERIC:
			cellOut.setCellValue(cellIn.getNumericCellValue());
			break;

		case Cell.CELL_TYPE_STRING:
			cellOut.setCellValue(cellIn.getStringCellValue());
			break;
		}
		HSSFCellStyle styleIn = (HSSFCellStyle) cellIn.getCellStyle();
		XSSFCellStyle styleOut = null;
		if (cellStyleMap.get(styleIn.getIndex()) != null) {
			styleOut = (XSSFCellStyle) cellStyleMap.get(styleIn.getIndex());
		} else {
			styleOut = (XSSFCellStyle) wbOut.createCellStyle();
			styleOut.setAlignment(styleIn.getAlignment());
			DataFormat format = wbOut.createDataFormat();
			styleOut.setDataFormat(format.getFormat(styleIn.getDataFormatString()));
			HSSFColor forgroundColor = styleIn.getFillForegroundColorColor();
			if (forgroundColor != null) {
				short[] foregroundColorValues = forgroundColor.getTriplet();
				styleOut.setFillForegroundColor(new XSSFColor(new java.awt.Color(foregroundColorValues[0],
						foregroundColorValues[1], foregroundColorValues[2])));
				styleOut.setFillPattern(styleIn.getFillPattern());
			}
			styleOut.setFillPattern(styleIn.getFillPattern());
			styleOut.setBorderBottom(styleIn.getBorderBottom());
			styleOut.setBorderLeft(styleIn.getBorderLeft());
			styleOut.setBorderRight(styleIn.getBorderRight());
			styleOut.setBorderTop(styleIn.getBorderTop());
			HSSFColor bottom = hssfPalette.getColor(styleIn.getBottomBorderColor());
			if (bottom != null) {
				short[] bottomColorArray = bottom.getTriplet();
				styleOut.setBottomBorderColor(new XSSFColor(new java.awt.Color(bottomColorArray[0],
						bottomColorArray[1], bottomColorArray[2])));
			}
			HSSFColor top = hssfPalette.getColor(styleIn.getTopBorderColor());
			if (top != null) {
				short[] topColorArray = top.getTriplet();
				styleOut.setTopBorderColor(new XSSFColor(new java.awt.Color(topColorArray[0], topColorArray[1],
						topColorArray[2])));
			}
			HSSFColor left = hssfPalette.getColor(styleIn.getLeftBorderColor());
			if (left != null) {
				short[] leftColorArray = left.getTriplet();
				styleOut.setLeftBorderColor(new XSSFColor(new java.awt.Color(leftColorArray[0], leftColorArray[1],
						leftColorArray[2])));
			}
			HSSFColor right = hssfPalette.getColor(styleIn.getRightBorderColor());
			if (right != null) {
				short[] rightColorArray = right.getTriplet();
				styleOut.setRightBorderColor(new XSSFColor(new java.awt.Color(rightColorArray[0], rightColorArray[1],
						rightColorArray[2])));
			}
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setHidden(styleIn.getHidden());
			styleOut.setIndention(styleIn.getIndention());
			styleOut.setLocked(styleIn.getLocked());
			styleOut.setRotation(styleIn.getRotation());
			styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
			styleOut.setWrapText(styleIn.getWrapText());
			cellOut.setCellComment(cellIn.getCellComment());
			cellStyleMap.put(styleIn.getIndex(), styleOut);
		}

	}

	public static void connect(){


		System.currentTimeMillis();	 
		System.out.println("Connecting to the Database");	 
		String hostname="";//getting the host name of the computer that we are using 
		try {
			hostname = InetAddress.getLocalHost().getHostName();
		}catch (UnknownHostException ex){System.out.println("Hostname can not be resolved");}


		props = new Properties();//get DB info
		try {
			if (hostname.equals("PRIMUS")) 
				props.load(new FileInputStream("C:/Users/ROB/git/Hi-Tech-Health/mydb2.properties"));
			else
				props.load(new FileInputStream("/java/mydb2.properties"));
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}


		final String DRIVER = "com.ibm.as400.access.AS400JDBCDriver"; 
		final String URL = "jdbc:as400://"+props.getProperty("local_system").trim()+"/"+LIB+";naming=system";//jdbc:db2

		try {
			Class.forName(DRIVER); //making the connection
			conn = DriverManager.getConnection(URL, props.getProperty("userId").trim(), props.getProperty("password").trim()); 

			conn.createStatement().execute("CREATE ALIAS QTEMP/INSURE_"+CLIENT+" FOR "+LIB+"/INSURE("+CLIENT+")");
			conn.createStatement().execute("CREATE ALIAS QTEMP/EVERELG_"+CLIENT+" FOR CBLIB/EVERELG");
			conn.createStatement().execute("CREATE ALIAS QTEMP/EVERELGHST_"+CLIENT+" FOR CBLIB/EVERELGHST");
			conn.createStatement().execute("CREATE ALIAS QTEMP/ZCCITY FOR ZCLIB/ZCCITY");
			conn.createStatement().execute("CREATE ALIAS QTEMP/STECOD FOR HTHDATV1/STECOD");

			System.out.println("Connected to database");
		} catch (ClassNotFoundException e) {e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} 


		
	}

	private static boolean isSSNValid(String val) {
		//System.out.println(val);
		if (val.length() >9)
		{	
			return false;    
		}
		else if(val.length()<9){
			while(val.length()<9) {
				val ="0"+val;}
			return true;

		}
		return true;
	}

	private static boolean isAddValid(String state, String zip) throws IOException {

		int found=0;
		int foundSTE=0;

		//System.out.println(state+" "+zip);

		found=DBqueryINT("Select COUNT(*) FROM qtemp/ZCCITY where ZCODE='"+zip+"' and ZSTATE='"+state+"'");
		foundSTE=DBqueryINT("Select COUNT(*) FROM qtemp/STECOD where STATE='"+state+"'");

		if(found==0 && !zip.equals(" "))
			return false;

		else if(foundSTE==0)
			return false;

		else if(found!=0 && foundSTE!=0)
			return true;

		return errorfound;

	}
	private static boolean isGenValid(String val) {
		//System.out.println(val);
		if(val == null && !val.equals("M") && !val.equals("F")) {
			return false;
		}
		return true;
	}
	private static boolean isDateValid(String val) throws ParseException{
		//System.out.println(val);
		try {

			SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
			SimpleDateFormat outputFormat = new SimpleDateFormat("MMddyyyy");

			val=outputFormat.format(inputFormat.parse(val));		

			Date date = outputFormat.parse(val);

		} catch (ParseException e) {return false;}

		return true;
	}

	private static boolean isDateValid2(String val){

		if(val.length()>10)
		{
			return false;
		}
		else
		{	
			try {

				SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
				SimpleDateFormat outputFormat = new SimpleDateFormat("MMddyy");

				val=outputFormat.format(inputFormat.parse(val));		

				Date date = outputFormat.parse(val);

			} catch (ParseException e) {return false;}
		}
		return true;
	}

	private static String isDateValid3(String val, String type){

		try {

			SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
			SimpleDateFormat outputFormat = new SimpleDateFormat("MMddyy");
			SimpleDateFormat birthdayFormat = new SimpleDateFormat("MMddyyyy");

			if(type=="B")
			{
				val=birthdayFormat.format(inputFormat.parse(val));		
			}
			else if(type=="D")
			{
				if(val.equals(""))
					val="0";                        
				else
					return val;
			}
			else
			{
				val=outputFormat.format(inputFormat.parse(val));
			}
		} catch (ParseException e) {

			val="0";

			return val;}

		return val;
	}


	public static Integer DBqueryINT(String query){
		int result=-1;
		try {
			Statement stmt1 = conn.createStatement();
			ResultSet rs = stmt1.executeQuery(query);

			rs.next();
			result = rs.getInt(1);

			rs.close();		
			stmt1.close();
		} catch (SQLException e) {e.printStackTrace();}
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
			//String EMAIL = Findemail("SELECT CAEML1 FROM QTEMP/BLOCK_"+CLIENT);
			Multipart multipart = new MimeMultipart();
			MimeMessage msg = new MimeMessage(session);

			MimeBodyPart textPart = new MimeBodyPart();

			MimeBodyPart messageBodyPart = new MimeBodyPart(); 

			msg.setFrom(new InternetAddress("support@hi-techhealth.com"));

			String recipientEmail = ("cbeyer@hi-techhealth.com, jeller@hi-techhealth.com");//twalsh@hi-techhealth.com
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

			//Second Attachment

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
		System.out.println("Program Closed");
		System.exit(0);

	}  

}
