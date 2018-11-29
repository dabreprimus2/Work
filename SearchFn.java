package com.search;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;
import java.util.Scanner;

public class SearchFn{
	private static Connection conn;
	private static String LIB,MEMBER;
	static Properties props = new Properties();
	static ResultSet rs_block=null;
	public static void main(String[] args) throws Exception {
		LIB="HTHDATV1";	
		MEMBER="";//get memberid
		Scanner sc = new Scanner(System.in);
		System.out.print("Enter search parameters:\t");
		String input = sc.nextLine();
		sc.close();
		input = convert(input);
		connect();
		search(input,MEMBER);
	}

	static String convert(String str) 
	{ 
		char ch[] = str.toCharArray(); 
		for (int i = 0; i < str.length(); i++) { 
		 
			if (i == 0 && ch[i] != ' ' || ch[i] != ' ' && ch[i - 1] == ' ') 
			{ 
 				if (ch[i] >= 'a' && ch[i] <= 'z') 
				{ 
					ch[i] = (char)(ch[i] - 'a' + 'A'); 
				} 
			} 
			else if (ch[i] >= 'A' && ch[i] <= 'Z') 
				ch[i] = (char)(ch[i] + 'a' - 'A');			 
		}  
		String st = new String(ch); 
		return st; 
	} 
	
	private static void search(String sr, String member) throws SQLException, IOException {
	   System.out.println("Searching results relevent:\t"+sr);
	   String a= sr.toUpperCase();
	   String b= sr.toLowerCase();
	   System.out.println("Searching results relevent:\t"+a);
	   System.out.println("Searching results relevent:\t"+b);
	   String sql;
	   Statement stmt = null;
	   stmt = conn.createStatement();
	   int i=1;
	   FileWriter fstream = new FileWriter("C:/Users/Primus/Desktop/Primus/search.txt");
	   BufferedWriter out = new BufferedWriter(fstream);

	   try {
	   sql="SELECT * FROM QTEMP/SECNAM WHERE SPRGID !=''  AND SPRGNM LIKE '%"+sr+"%' OR SPRGNM LIKE'%"+a+"%' OR SPRGNM LIKE'%"+b+"%'";
	   ResultSet rs =stmt.executeQuery(sql);
	   while (rs.next()) { 
	   String id= rs.getString("SPRGID");
	   String uid= rs.getString("SMNUID");
	   String uno= rs.getString("SMNUNO");
	   String nm= rs.getString("SPRGNM");
	  
	   System.out.println("Result "+i+":"+id+" "+uid+" "+uno+" "+nm);
	   i++;
	   out.write(rs.getString("SPRGID") + " ");
	   out.write(rs.getString("SMNUID") + " ");
	   out.write(rs.getString("SMNUNO") + " ");
	   out.write(rs.getString("SPRGNM") + " ");
	   out.newLine();
	   	}
	   System.out.println("\nTotal "+(i-1)+" search results found\n\nCompleted writing search result into text file");
	   out.close();
	   }catch(Exception e) {
		   System.out.println("Couldn't find any serach result");
	   }
	}
	
	public static void connect(){

		System.currentTimeMillis();	 
		System.out.println("\nConnecting to the Database");	 
		String hostname="";//getting the host name of the computer that we are using 
		try {
			hostname = InetAddress.getLocalHost().getHostName();
		}catch (UnknownHostException ex){System.out.println("Hostname can not be resolved");}

		props = new Properties();//get DB info
		try {
			if (hostname.equals("PRIMUS")) 
				props.load(new FileInputStream("C:/Users/Primus/git/Hi-Tech-Health/mydb2.properties"));
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
			//create alias for secnam
			conn.createStatement().execute("CREATE ALIAS QTEMP/SECNAM FOR "+LIB+"/SECNAM");
			
			//create alias for sysusrp
			//conn.createStatement().execute("CREATE ALIAS QTEMP/SECNAM FOR "+LIB+"/SYSUSRP ");

			System.out.println("Connected to database");
		} catch (ClassNotFoundException e) {e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} 
	}
}
