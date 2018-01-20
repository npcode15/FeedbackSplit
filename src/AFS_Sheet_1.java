import java.awt.Toolkit;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import java.io.BufferedInputStream;
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
import javax.swing.JFileChooser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class AFS_Sheet_1 extends JFrame implements ActionListener
{
	static JFrame frame;
	static JButton email;
	static JButton encrypt;
	static JButton extract;
	static JButton readfile;
	static JButton home;
	static JButton refresh;
	static JTextArea displayReport;
	static JComboBox<String> facultyList;

	static String subject;
	static String teacher;
	static String frameName="AFS Sheet-1";
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
	static int c=0;
	
	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;

    static FileInputStream fileInput = null;       
    static BufferedInputStream bufferInput = null;      
    static POIFSFileSystem poiFileSystem = null;    

	private String[] content;
	private String[] session_Desc=new String[15];
	private String[] cnt={"07", "08", "09", "10", "11" ,"12", "13", "14", "15", "16", "17", "18"};
	private String[] faculty={"Select Name of The Faculty", "Prof Abhishek Chakravorty", "Prof Jasmeet Bhatia", "Prof Rajiv Pathak", "Prof Rohit Verma", "Prof Sarang Pitale", "Prof Subhashini", "Prof T Usha"};
	private String[] newcontent=new String[16];
	private String[] Criteria=new String[20];
	private String[] rating=new String[20];
	private String[] details=new String[2];
	private boolean profFound=false;
	private int count=0;

	public AFS_Sheet_1()
	{
		setTitle("AFS Sheet-1");
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
		readfile.setBounds((scrW-750)/2, scrH/3-50, 250, 50);
		add(readfile);
		readfile.setVisible(true);
		readfile.addActionListener(this);

		extract=new JButton("Extract Report");								// Button for Starting the Process
		extract.setBounds((scrW-750)/2+600, scrH/3-50, 250, 50);
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
		facultyList.setBounds((scrW-750)/2+300, scrH/3-50, 250, 50);		// List of Teachers
		facultyList.setVisible(false);
		add(facultyList);
		facultyList.addActionListener(this);
		facultyList.setVisible(false);
		for(int i=0;i<faculty.length;i++)
			facultyList.addItem(faculty[i]);

		displayReport= new JTextArea();										// File Display TextArea
		displayReport.setEditable(false);  
		displayReport.setCursor(null);  
		displayReport.setOpaque(false);  
		displayReport.setFocusable(false);
		displayReport.setLineWrap(true);
		displayReport.setBounds(5, 2*scrH/3-200, scrW-20, 300);
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
		frame = new AFS_Sheet_1();
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
			new AFS_Sheet_1();
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
				if(content[i].contains("Generating Graph for"))
					session_Desc=content[i].split(" ");

				if(content[i].contains(teacherToBeAssesed))
				{
					profFound=true;
					for(int j=0;j<15;j++)
					{	
						newcontent[j]=content[i];
						System.out.println(newcontent[j]);
						newplaintext+="\n"+content[i++];
					}
					newcontent[15]=content[i-1];
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

			String[] temp=newcontent[0].split("-:-");
			subject=temp[0].trim();
			teacher=temp[1].trim();

			int z=0;
			for(int k=2;k<newcontent.length-2;k++)
			{
				details=newcontent[k].split(":-");
				Criteria[z]=details[0];
				rating[z++]=details[1];
			}

			displayReport.setText("");
			displayReport.setVisible(true);
			finalContent="\n"+"Subject :- "+subject+"\tTaught By :- "+teacher+"\n\n"+"Session :- "+session+"\tSemester :- "+semester+"\tBranch:- "+branch+"\n\n";		

			for(int i=0;i<12;i++)
				finalContent+=Criteria[i]+"	:-> "+rating[i]+"\n";
			displayReport.setText(finalContent);

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

			HSSFSheet sheet = workbook.getSheetAt(0);

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

		data.put("01", new Object[] {"", "", "Displaying", "Graphical", "Analysis", " ", " "," "});
		data.put("02", new Object[] {"", "", "", "", "", "", ""});
		data.put("03", new Object[] {"", "Session :- "+session, "Branch :- "+branch+" Semester :- "+semester});
		data.put("04", new Object[] {"", "", ""});
		data.put("05", new Object[] {"", "Subject :- "+subject, "Taught By :- "+teacher});
		data.put("06", new Object[] {"", "", "", "", "", "", ""});

		for(int i=0;i<12;i++)
			if(Float.parseFloat(rating[i])>0 && Float.parseFloat(rating[i])<1)
			{
				data.put(cnt[k++], new Object[] {"", Criteria[i]+"", rating[i]});
				count=1;
			}
			else
				if(Float.parseFloat(rating[i])>1 && Float.parseFloat(rating[i])<2)
				{
					data.put(cnt[k++], new Object[] {"", Criteria[i]+"", "", rating[i]});
					count=2;
				}
				else
					if(Float.parseFloat(rating[i])>2 &&Float.parseFloat(rating[i])<3)
					{
						data.put(cnt[k++], new Object[] {"", Criteria[i]+"", "", "", rating[i]});
						count=3;
					}
					else
						if(Float.parseFloat(rating[i])>3 && Float.parseFloat(rating[i])<4)
						{
							data.put(cnt[k++], new Object[] {"", Criteria[i]+"", "", "", "", rating[i]});
							count=4;
						}
						else
							if(Float.parseFloat(rating[i])>4)
							{
								data.put(cnt[k++], new Object[] {"", Criteria[i]+"", "", "", "", "", rating[i]});
								count=5;
							}
		data.put("19", new Object[] {"", "--:--:--: GRADING :--:--:-- ", "1", "2", "3", "4", "5"});

		/*----------------------------------------------------------------------*/            
	
		
		/*fileInput = new FileInputStream("D:\\[Softwares]\\Coding Utilities\\Eclipse\\Eclipse Indigo\\Major_Project\\AssesmentFeedbackSplitter\\"+"Prof T Usha- Graphical Analysis Report.xls");         
        bufferInput = new BufferedInputStream(fileInput);  
		poiFileSystem = new POIFSFileSystem(bufferInput);            
        Biff8EncryptionKey.setCurrentUserPassword("secret");            
        HSSFWorkbook workbook = new HSSFWorkbook(poiFileSystem, true);*/
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Assesment Report");

		HSSFFont fonthead = workbook.createFont();
		fonthead.setFontHeight((short) 500);
		fonthead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fonthead.setColor(IndexedColors.WHITE.getIndex());

		HSSFCellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(HSSFColor.MAROON.index);
		style.setFillPattern(HSSFCellStyle.DIAMONDS);
		style.setFont(fonthead);

		/*----------------------------------------------------------------------*/

		HSSFFont fontsubhead = workbook.createFont();
		fontsubhead.setFontHeight((short) 350);
		fontsubhead.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		fontsubhead.setColor(IndexedColors.DARK_GREEN.getIndex());

		HSSFCellStyle stylesubhead = workbook.createCellStyle();
		stylesubhead.setFont(fontsubhead);

		/*stylesubhead.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		stylesubhead.setBorderTop(HSSFCellStyle.BORDER_THIN);
		stylesubhead.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		stylesubhead.setBorderRight(HSSFCellStyle.BORDER_THIN);
		stylesubhead.setBottomBorderColor(HSSFColor.BLACK.index);*/

		/*----------------------------------------------------------------------*/

		HSSFFont fontsubtext = workbook.createFont();
		fontsubtext.setFontHeight((short) 250);
		fontsubtext.setColor(IndexedColors.DARK_RED.getIndex());

		HSSFCellStyle stylesubtext = workbook.createCellStyle();
		stylesubtext.setFont(fontsubtext);

		/*stylesubtext.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		stylesubtext.setBorderTop(HSSFCellStyle.BORDER_THIN);
		stylesubtext.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		stylesubtext.setBorderRight(HSSFCellStyle.BORDER_THIN);
		stylesubtext.setBottomBorderColor(HSSFColor.BLACK.index);*/

		/*----------------------------------------------------------------------*/

		HSSFFont fontGraphColor = workbook.createFont();
		fontGraphColor.setColor(IndexedColors.WHITE.getIndex());

		HSSFCellStyle styleGraphColor[]=new HSSFCellStyle[25];
		for(int i=0;i<25;i++)
		{
			styleGraphColor[i]=workbook.createCellStyle();
			styleGraphColor[i].setFont(fontGraphColor);

			styleGraphColor[i].setAlignment((short) 3);
			styleGraphColor[i].setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
			styleGraphColor[i].setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
			styleGraphColor[i].setBottomBorderColor(HSSFColor.WHITE.index);
			styleGraphColor[i].setTopBorderColor(HSSFColor.WHITE.index);
		}

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for(String key : keyset)
		{
			Row row = sheet.createRow(rownum++);

			Object [] objArr = data.get(key);
			int cellnum = 0;
			int rowcnt=7;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);

				if(rownum>5 && cell.getColumnIndex()>1)
					for(int i=0;i<count;i++)
					{
						if(rownum==7)
						{	
							styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.BROWN.index);
							styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
							cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
						}
						else
							if(rownum==8)
							{
								styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
								styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
								cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
							}
							else
								if(rownum==9)
								{
									styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
									styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
									cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
								}
								else
									if(rownum==10)
									{
										styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.PINK.index);
										styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
										cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
									}
									else
										if(rownum==11)
										{
											styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
											styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
											cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
										}
										else
											if(rownum==12)
											{
												styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.ORANGE.index);
												styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
												cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
											}
											else
												if(rownum==13)
												{
													styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.MAROON.index);
													styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
													cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
												}
												else
													if(rownum==14)
													{
														styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.OLIVE_GREEN.index);
														styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
														cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
													}
													else
														if(rownum==15)
														{
															styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.GREEN.index);
															styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
															styleGraphColor[rownum-rowcnt].setAlignment((short)1);
															cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
														}
														else
															if(rownum==16)
															{
																styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.DARK_TEAL.index);
																styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
																cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
															}
															else
																if(rownum==17)
																{
																	styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.ORCHID.index);
																	styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
																	cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
																}
																else
																	if(rownum==18)
																	{
																		styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.DARK_YELLOW.index);
																		styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
																		cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
																	}
																	else
																		if(rownum==19)
																		{
																			cell.setCellStyle(styleGraphColor[rownum-rowcnt]);
																			styleGraphColor[rownum-rowcnt].setFillForegroundColor(HSSFColor.BLACK.index);
																			styleGraphColor[rownum-rowcnt].setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
																		}

						if(obj instanceof String)
							cell.setCellValue((String)obj);
						else 
							if(obj instanceof Integer)
								cell.setCellValue((Integer)obj);
					}
				else
				{
					if(rownum==1)
						cell.setCellStyle(style);
					else
						if(rownum==3)
						{
							cell.setCellStyle(stylesubhead);

							sheet.autoSizeColumn(0);
							sheet.autoSizeColumn(2);
							sheet.autoSizeColumn(3);
							sheet.autoSizeColumn(4);

							sheet.setColumnWidth(5, sheet.getColumnWidth(3));
							sheet.setColumnWidth(6, sheet.getColumnWidth(3));
						}
						else
							if((rownum==3 || rownum==4 || rownum==5 || rownum==6) && (cell.getColumnIndex()>0))
								cell.setCellStyle(stylesubhead);

							else
								if(rownum>6 && cell.getColumnIndex()>0)
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
		}

		try 
		{
			File f=new File(teacherToBeAssesed+" - Graphical Analysis Report.xls");
			FileOutputStream out = new FileOutputStream(f);
			workbook.write(out);
			filePath=f.getAbsolutePath();
			fileName=f.getName();
			out.close();

			System.out.println(teacherToBeAssesed+".xls"+" written successfully on disk.");
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}