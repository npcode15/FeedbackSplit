import java.awt.Toolkit;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.io.File;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.Iterator;

import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JButton;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JFileChooser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

public class AFS_Sheet_3 extends JFrame implements ActionListener
{
	static JFrame frame;
	static JButton email;
	static JButton encrypt;
	static JButton extract;
	static JButton readfile;
	static JButton home;
	static JButton refresh;
	static JTextArea displayReport;
	static JTextField textInput;
	static JComboBox<String> facultyList;

	static String frameName="AFS Sheet-3";
	static String fileName="";
	static String filePath="";
	static String session="";
	static String semester="";
	static String branch="";
	static String plaintext = "";
	static String newplaintext="";
	static String finalContent="";
	static String teacherToBeAssesed="";

	static File selectedFile=null;
	static FileInputStream file=null;

	static int response=0;
	static int k=0;
	static int m=0;
	static int y=0;
	static int onlyOnce=0;
	static int categoryCount=0;

	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;

	static String[] category=new String [10];
	static String[] newcontent=new String[10];
	static String[] newFaculty=new String [10];
	static String[] rank=new String [10];
	static String[] average=new String [10];
	static String[] faculty=new String [10];
	static String[] details=new String [5];
	static String[] content=new String [95];
	static String[] session_Desc;
	static String[] facultii={"Select Name of The Faculty", "Prof Abhishek Chakravorty", "Prof Jasmeet Bhatia", "Prof Rajiv Pathak", "Prof Rohit Verma", "Prof Sarang Pitale", "Prof Subhashini", "Prof T Usha"};
	static String[] cnt={"05", "06", "07", "08", "09", "10", "11" ,"12", "13", "14", "15", "16", "17", "18"};

	private boolean profFound=false;

	public AFS_Sheet_3()
	{
		setTitle("AFS Sheet-3");
		setLayout(null);
		setSize(scrW, scrH);

		home=new JButton("Home");											// Redirect to Home
		home.setBounds(15, 5, 100, 40);
		add(home);
		home.setVisible(true);
		home.addActionListener(this);

		refresh=new JButton("Refresh");										// Restart Same Frame
		refresh.setBounds((scrW-15-100), 5, 100, 40);
		add(refresh);
		refresh.setVisible(true);
		refresh.addActionListener(this);

		readfile=new JButton("Start Process");								// Button for File Selection for Reading
		readfile.setBounds((scrW-750)/2, scrH/3, 250, 50);
		add(readfile);
		readfile.setVisible(true);
		readfile.addActionListener(this);

		extract=new JButton("Extract Report");								// Button for Starting the Process
		extract.setBounds((scrW-750)/2+600, scrH/3, 250, 50);
		add(extract);
		extract.setVisible(false);
		extract.addActionListener(this);

		encrypt=new JButton("Encrypt File Contents".toUpperCase());						// Button for Encrypting File
		encrypt.setBounds(30+(scrW-50)/2,2*scrH/3+100, (scrW-50)/2, 50);
		add(encrypt);
		encrypt.setVisible(false);
		encrypt.addActionListener(this);

		email=new JButton("E-Mail Original Contents".toUpperCase());										// Button for E-Mail
		email.setBounds(5,2*scrH/3+100, (scrW-50)/2, 50);
		add(email);
		email.setVisible(false);
		email.addActionListener(this);

		facultyList=new JComboBox<String>();
		facultyList.setBounds((scrW-750)/2+300, scrH/3, 250, 50);			// List of Teachers
		facultyList.setVisible(false);
		add(facultyList);
		facultyList.addActionListener(this);
		facultyList.setVisible(false);
		for(int i=0;i<facultii.length;i++)
			facultyList.addItem(facultii[i]);

		displayReport= new JTextArea();										// File Display TextArea
		displayReport.setEditable(false);  
		displayReport.setCursor(null);  
		displayReport.setOpaque(false);  
		displayReport.setFocusable(false);
		displayReport.setLineWrap(true);
		displayReport.setBounds(5, scrH/3+80, scrW-10, 250);
		displayReport.setVisible(false);
		displayReport.setWrapStyleWord(true);
		add(displayReport);

		onlyOnce=0;

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
		frame = new AFS_Sheet_3();
	}

	public void actionPerformed(ActionEvent evt)
	{
		profFound=false;

		String str=evt.getActionCommand();

		if(str.equals("Home"))
		{
			this.dispose();
			new AFS_Home();
		}

		if(str.equals("Refresh"))
		{
			this.dispose();
			frame = new AFS_Sheet_3();
		}

		if(str.equals("Encrypt File Contents".toUpperCase()))
		{
			writeExcel();
			new Encrypt_Sheet(filePath, fileName, frameName);
			this.dispose();
		}

		if(str.equals("E-Mail Original Contents".toUpperCase()))
		{
			writeExcel();

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
				readfile.setEnabled(false);
			} 
			catch (IOException e)
			{
				e.printStackTrace();
			}
		}

		
		m=0;
		y=0;
		newplaintext="";
		categoryCount=0;
		if(str.equals("Extract Report"))
		{
			teacherToBeAssesed = String.valueOf(facultyList.getSelectedItem());

			for(int i=0;i<95;i++)
				content[i]="";

			content = plaintext.split("\n");

			for(int i=0;i<content.length;i++)
				if(content[i].contains(teacherToBeAssesed))
					profFound=true;

			if(profFound)
			{
				for(int i=0;i<content.length;i++)
				{
					if(content[i].contains("Session") && content[i].contains("Semester") && content[i].contains("Branch"))
						session_Desc=content[i].split(" ");

					if(content[i].contains("PerformanceIndex"))
					{
						System.out.println("content["+i+"] >> "+content[i]);
						if(content[i].contains(":-"))
						{
							details=content[i].split(":-");
							category[m++]=details[1];
						}

						for(int j=0;j<9;j++)
						{	
							newcontent[j]=content[i+j];
							newplaintext+="\n"+content[i+j];
						}

						int z=0;
						for(int k=1;k<newcontent.length-1;k++)
						{
							if(newcontent[k].contains(":-:"))
							{
								details=newcontent[k].split(":-:");

								rank[z]=details[0];
								average[z]=details[1];
								faculty[z++]=details[2];

								if(faculty[z-1].contains(teacherToBeAssesed))
								{
									if(rank[z-1].contains("1"))
										newFaculty[y++]=faculty[z-1]+" is Ranked "+rank[z-1].substring(0,1)+"st in "+category[categoryCount++].trim()+"("+average[z-1]+")";
									else
										if(rank[z-1].contains("2"))
											newFaculty[y++]=faculty[z-1]+" is Ranked "+rank[z-1].substring(0,1)+"nd in "+category[categoryCount++].trim()+"("+average[z-1]+")";
										else
											if(rank[z-1].contains("3"))
												newFaculty[y++]=faculty[z-1]+" is Ranked "+rank[z-1].substring(0,1)+"rd in "+category[categoryCount++].trim()+"("+average[z-1]+")";
											else
												newFaculty[y++]=faculty[z-1]+" is Ranked "+rank[z-1].substring(0,1)+"th in "+category[categoryCount++].trim()+"("+average[z-1]+")";
								}
							}
						}
					}
				}

				for(int i=0;i<session_Desc.length;i++)
				{
					if(session_Desc[i].contains("Session"))
						session=session_Desc[i+1];

					if(session_Desc[i].contains("Semester"))
						semester=session_Desc[i+1];

					if(session_Desc[i].contains("Branch"))
						branch=session_Desc[i+1].substring(0,3).trim();
				}

				displayReport.setText("");
				displayReport.setVisible(true);
				finalContent="\n"+"Session :- "+session+"\tSemester :- "+semester+"\tBranch:- "+branch+"\n\n";		

				for(int i=0;i<10;i++)
					finalContent+=newFaculty[i]+"\n";
				displayReport.setText(finalContent);

				email.setVisible(true);
				encrypt.setVisible(true);
			}
			else
				if(!profFound)
				{
					displayReport.setText("");
					displayReport.setVisible(false);
					JOptionPane p=new JOptionPane();
					JOptionPane.showMessageDialog(p,"Professor named \""+teacherToBeAssesed+"\" not found.");
				}
		}
	}

	public void ReadExcel() throws IOException
	{
		plaintext="";
		try
		{
			JFileChooser fileChooser = new JFileChooser();

			int returnValue = fileChooser.showOpenDialog(null);
			if(returnValue == JFileChooser.APPROVE_OPTION)
				selectedFile = fileChooser.getSelectedFile();

			file = new FileInputStream(selectedFile);

			HSSFWorkbook workbook = new HSSFWorkbook(file);

			HSSFSheet sheet = workbook.getSheetAt(2);

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
						plaintext+=cell.getNumericCellValue()+":-:";
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

	public void writeExcel() 
	{
		k=0;
		filePath="";
		fileName="";
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		data.put("01", new Object[] {"", "", "Displaying Performance Index wise Analysis", " ", " "});
		data.put("02", new Object[] {"", "", ""});
		data.put("03", new Object[] {"", "Session :- "+session, "Branch :- "+branch+" Semester :- "+semester});
		data.put("04", new Object[] {"", "", ""});
		for(int i=0;i<10;i++)
			data.put(cnt[k++], new Object[] {"", newFaculty[i]+""});

		/*----------------------------------------------------------------------*/   

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Assesment Report");

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

				if(rownum==4)
				{
					sheet.autoSizeColumn(1);
					sheet.autoSizeColumn(2);
				}

				if(rownum==1)
					cell.setCellStyle(style);
				else
					if(rownum==3)
						cell.setCellStyle(stylesubhead);
					else
						cell.setCellStyle(stylesubtext);

				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}			
		}

		try 
		{
			File f=new File(teacherToBeAssesed+" - Performance Index.xls");
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