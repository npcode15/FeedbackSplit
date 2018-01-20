import java.awt.Toolkit;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Map;
import java.util.TreeMap;
import java.util.Set;
import java.util.Iterator;

import javax.crypto.Cipher;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JButton;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.JFileChooser;
import javax.swing.UIManager;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class AFS_Sheet_2 extends JFrame implements ActionListener
{
	static JFrame frame;
	static JButton email;
	static JButton encrypt;
	static JButton extract;
	static JButton home;
	static JButton refresh;
	static JButton readfile;
	static JTextArea displayReport;
	static JComboBox<String> facultyList;

	static String subject;
	static String teacher;
	static String frameName="AFS Sheet-2";
	static String fileName="";
	static String filePath="";
	static String session="";
	static String semester="";
	static String branch="";
	static String plaintext = "";
	static String newplaintext="";
	static String finalContent;
	static String teacherToBeAssesed="";

	static File selectedFile;

	static int response=0;
	static int k=0;
	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;

	private String titlet;
	private String titles;
	private String[] content;
	private String[] temp=new String[5];
	private String[] session_Desc=new String[10];
	private String[] faculty={"Select Name of The Faculty", "Prof Abhishek Chakravorty", "Prof Jasmeet Bhatia", "Prof Rajiv Pathak", "Prof Rohit Verma", "Prof Sarang Pitale", "Prof Subhashini", "Prof T Usha"};
	private String[] cnt={"07", "08", "09", "10", "11" ,"12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "25", "26", "27", "28", "29", "30"};
	private String[] newcontent=new String[27];
	private String[] Criteria=new String[25];
	private String[] rating=new String[25];
	private String[] details=new String[2];
	private boolean profFound=false;

	public AFS_Sheet_2()
	{
		setTitle("AFS Sheet-2");
		setLayout(null);
		setSize(scrW, scrH);
		
		home=new JButton("Home");											// Redirect to Home
		home.setBounds(5, 5, 100, 40);
		add(home);
		home.setVisible(true);
		home.addActionListener(this);

		refresh=new JButton("Refresh");										// Restart Same Frame
		refresh.setBounds((scrW-5-100), 5, 100, 40);
		add(refresh);
		refresh.setVisible(true);
		refresh.addActionListener(this);

		readfile=new JButton("Start Process");								// Button for File Selection for Reading
		readfile.setBounds((scrW-750)/2, scrH/3-150, 250, 50);
		add(readfile);
		readfile.setVisible(true);
		readfile.addActionListener(this);
		
		extract=new JButton("Extract Report");								// Button for Starting the Process
		extract.setBounds((scrW-750)/2+600, scrH/3-150, 250, 50);
		add(extract);
		extract.setVisible(false);
		extract.addActionListener(this);

		encrypt=new JButton("Encrypt File Contents".toUpperCase());						// Button for Encrypting File
		encrypt.setBounds(30+(scrW-50)/2,2*scrH/3+140, (scrW-50)/2, 50);
		add(encrypt);
		encrypt.setVisible(false);
		encrypt.addActionListener(this);

		email=new JButton("E-Mail Original Contents".toUpperCase());										// Button for E-Mail
		email.setBounds(5,2*scrH/3+140, (scrW-50)/2, 50);
		add(email);
		email.setVisible(false);
		email.addActionListener(this);

		facultyList=new JComboBox<String>();
		facultyList.setBounds((scrW-750)/2+300, scrH/3-150, 250, 50);					// List of Teachers
		facultyList.setVisible(false);
		add(facultyList);
		facultyList.addActionListener(this);
		facultyList.setVisible(false);
		for(int i=0;i<faculty.length;i++)
			facultyList.addItem(faculty[i]);

		displayReport= new JTextArea();													// File Display TextArea
		displayReport.setEditable(false);  
		displayReport.setCursor(null);  
		displayReport.setOpaque(false);  
		displayReport.setFocusable(false);
		displayReport.setLineWrap(true);
		displayReport.setBounds(20, scrH/3-75, scrW-40, 460);
		displayReport.setVisible(false);
		displayReport.setWrapStyleWord(true);
		add(displayReport);

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setVisible(true);
	}

	public static void main(String [] args)
	{
		try 
		{
			UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
		} 
		catch(Exception e) 
		{
			e.printStackTrace(); 
		}
		frame = new AFS_Sheet_2();
	}

	public void actionPerformed(ActionEvent evt)
	{
		String str=evt.getActionCommand();

		profFound=false;
		
		if(str.equals("Home"))
		{
			this.dispose();
			new AFS_Home(); 
		}

		if(str.equals("Refresh"))
		{
			this.dispose();
			new AFS_Sheet_2();
		}

		if(str.equals("Encrypt File Contents".toUpperCase()))
		{
			writeExcel(subject, teacher, Criteria, rating);
			new Encrypt_Sheet(filePath, fileName, frameName);
			this.dispose();
		}
		
		if(str.equals("E-Mail Original Contents".toUpperCase()))
		{
			writeExcel(subject, teacher, Criteria, rating);
			new SendByEmail(filePath, fileName, frameName);
			this.dispose();
		}

		if(str.equals("Start Process"))
		{
			try
			{
				ReadExcel();
				facultyList.setVisible(true);
				extract.setVisible(true);
			} 
			catch (IOException e)
			{
				e.printStackTrace();
			}
		}

		newplaintext="";
		if(str.equals("Extract Report"))
		{
			teacherToBeAssesed = String.valueOf(facultyList.getSelectedItem());

			content = plaintext.split("\n");

			for(int i=0;i<content.length;i++)
			{
				if(content[i].contains("Session") && content[i].contains("Semester") && content[i].contains("Branch"))
					session_Desc=content[i].split(" ");

				if(content[i].contains(teacherToBeAssesed))
				{
					profFound=true;
					for(int j=0;j<25;j++)
					{	
						newcontent[j]=content[i];
						newplaintext+="\n"+content[i++];
					}

					for(int l=0;l<newcontent.length;l++)
						if(newcontent[l]!=null)
							System.out.println(newcontent[l]);
				}
			}

			if(!profFound)
			{
				displayReport.setText("");
				displayReport.setVisible(false);
				JOptionPane p=new JOptionPane();
				JOptionPane.showMessageDialog(p,"Professor named \""+teacherToBeAssesed+"\" not found.");
			}

			for(int i=0;i<session_Desc.length;i++)
			{
				if(session_Desc[i].contains("Session"))
					session=session_Desc[i+1];

				if(session_Desc[i].contains("Semester"))
					semester=session_Desc[i+1];

				if(session_Desc[i].contains("Branch"))
					branch=session_Desc[i+1];
			}

			int z=0;
			String tmpstr="";
			String [] Arr=new String[5];
			for(int k=2;k<newcontent.length-2;k++)
			{
				if(newcontent[k]!="" && newcontent[k].contains(":-"))
				{
					details=newcontent[k].split(":-");

					Criteria[z]=details[0];
					rating[z++]=details[1];

					if((Criteria[z-1].contains("Strength")) || (Criteria[z-1].contains("Weakness")))
					{		
						for(int index=0;index<rating[z-1].length();index++)
						{
						  if(rating[z-1].substring(index, index+1).contains(")"))
							{
								Criteria[z-1]+=rating[z-1].substring(0, index+1);
								rating[z-1]=rating[z-1].substring(index+1).trim();
							}
						}
					}
				}
				else
				{	
					tmpstr=newcontent[k];
					Arr=tmpstr.split("Number of Students who feel that");
					Arr[0]="Number of Students who feel that ";
					Criteria[z]=Arr[0];
					rating[z++]=Arr[1];
				}
			}

			temp=newcontent[0].split(":-");
			titlet = temp[0].trim();
			teacher=temp[1].trim();

			temp=newcontent[1].split(":-");
			titles = temp[0].trim();
			subject=temp[1].trim();

			displayReport.setText("");
			displayReport.setVisible(true);
			finalContent="\nSession :- "+session+"\tSemester :- "+semester+"\tBranch:- "+branch+"\n\n";		
			finalContent+=titlet+" :- "+teacher+"	"+titles+" :- "+subject+"\n\n";

			for(int i=0;i<23;i++)
				finalContent+=Criteria[i]+"	:-> "+rating[i]+"\n";
			displayReport.setText(finalContent);
			displayReport.setVisible(true);
			email.setVisible(true);
			encrypt.setVisible(true);
		}
	}

	public void ReadExcel() throws IOException
	{
		try
		{
			JFileChooser fileChooser = new JFileChooser();

			int returnValue = fileChooser.showOpenDialog(null);
			if(returnValue == JFileChooser.APPROVE_OPTION)
				selectedFile = fileChooser.getSelectedFile();

			FileInputStream file = new FileInputStream(selectedFile);

			HSSFWorkbook workbook = new HSSFWorkbook(file);

			HSSFSheet sheet = workbook.getSheetAt(1);

			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				while(cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_NUMERIC: 
						plaintext+=cell.getNumericCellValue();
						break;
					case Cell.CELL_TYPE_STRING:
						plaintext+=cell.getStringCellValue();
						break;
					}
				}
				plaintext+="\n";
			}
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	public void writeExcel(String subject, String teacher, String[] Criteria, String[] rating) 
	{
		k=0;
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		data.put("01", new Object[] {"", "", "Displaying Detailed Analysis", " ", " ", " ", " "});
		data.put("02", new Object[] {"", "", ""});
		data.put("03", new Object[] {"", "Session :- "+session, "Branch :- "+branch+" Semester :- "+semester});
		data.put("04", new Object[] {"", "", ""});
		data.put("05", new Object[] {"", titlet+" :- "+teacher, titles+" :- "+subject});
		data.put("06", new Object[] {"", "", ""});
		for(int i=0;i<23;i++)
			data.put(cnt[k++], new Object[] {"", Criteria[i]+"", rating[i]});

		/*----------------------------------------------------------------------*/   

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Assesment Report");
		
		/*----------------------------------------------------------------------*/   

		HSSFFont fonthead = workbook.createFont();
		fonthead.setFontHeight((short) 500);
		fonthead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fonthead.setColor(IndexedColors.WHITE.getIndex());

		HSSFCellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		style.setFillPattern(HSSFCellStyle.BRICKS);
		style.setFont(fonthead);

		/*----------------------------------------------------------------------*/

		HSSFFont fontsubhead = workbook.createFont();
		fontsubhead.setFontHeight((short) 350);
		fontsubhead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fontsubhead.setColor(IndexedColors.DARK_GREEN.getIndex());

		HSSFCellStyle stylesubhead = workbook.createCellStyle();
		stylesubhead.setFont(fontsubhead);

		/*----------------------------------------------------------------------*/

		HSSFFont fontsubtext = workbook.createFont();
		fontsubtext.setFontHeight((short) 250);
		fontsubtext.setColor(IndexedColors.DARK_RED.getIndex());

		HSSFCellStyle stylesubtext = workbook.createCellStyle();
		stylesubtext.setFont(fontsubtext);

		/*----------------------------------------------------------------------*/

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for(String key : keyset)
		{
			Row row = sheet.createRow(rownum++);

			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);

				if(rownum==1)
					cell.setCellStyle(style);
				else
					if(rownum==3)
					{
						cell.setCellStyle(stylesubhead);
						sheet.autoSizeColumn(2);
					}
					else
						if(rownum==5)
							cell.setCellStyle(stylesubhead);
						else
						{
							cell.setCellStyle(stylesubtext);
							sheet.autoSizeColumn(1);
						}

				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}			
		}

		try 
		{
			File f=new File(teacherToBeAssesed+"- Detailed Analysis Report.xls");
			FileOutputStream out = new FileOutputStream(f);
			workbook.write(out);

			filePath=f.getAbsolutePath();
			fileName=f.getName();
			out.close();

			System.out.println(teacherToBeAssesed+".xlsx"+" written successfully on disk.");
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}