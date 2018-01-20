import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.UIManager;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.crypto.spec.SecretKeySpec;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.Cipher;
import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

public class Decrypt_Sheet extends JFrame implements ActionListener
{
	static String IV = "AAAAAAAAAAAAAAAA";
	static String ciphertext = "";
	static String plaintext = "";
	static String decryptionKey= "";
	static String flag_entry=""; 
	static String title="";

	static JButton startp, startd, home, refresh;
	static JDialog option;
	static JFrame frame;
	static JTextField textinput, keyinput;
	static JLabel fileinfo, ctextinfo, keyinputinfo;
	static JTextArea cipherop;

	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;

	private static int ccnt=0;
	private static int rcnt=0;
	private static int[] colcnt=new int[30];

	private static String[] content;
	private static String[] newcontent;
	private static String[][] newColumn;
	private static String[][] column=new String[30][25];
	private static String[] cnt={"01", "02", "03", "04", "05" ,"06", "07", "08", "09", "10", "11" ,"12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"};
	private static File selectedFile;
	private static String newplaintext;

	public Decrypt_Sheet()
	{
		setTitle("Decrypt_Sheet");
		setLayout(null);
		setSize(scrW, scrH);

		home=new JButton("Home");												// Redirect to Activity
		home.setBounds(5, 5, 100, 40);
		add(home);
		home.setVisible(true);
		home.addActionListener(this);

		refresh=new JButton("Refresh");											// Restart Same Frame
		refresh.setBounds((scrW-5-100), 5, 100, 40);
		add(refresh);
		refresh.setVisible(true);
		refresh.addActionListener(this);

		startd=new JButton("Start Decryption");									// Start Decryption Button
		startd.setBounds((scrW-250)/2, scrH/2-25, 250, 50);
		add(startd);
		startd.setVisible(false);
		startd.addActionListener(this);

		startp=new JButton("Start Process");									// Start Process Button
		startp.setBounds((scrW-750)/3, scrH/3, 250, 50);
		add(startp);
		startp.setVisible(true);
		startp.addActionListener(this);

		textinput = new JTextField();											// CipherText Input TextField
		textinput.setBounds((scrW-750)/3+300, scrH/3, 250, 50);
		textinput.setVisible(false);
		textinput.addActionListener(this);
		add(textinput);

		keyinput = new JTextField();											// Key Input TextField
		keyinput.setBounds((scrW-750)/3+600, scrH/3, 250, 50);
		keyinput.setVisible(false);
		keyinput.setToolTipText("Minimum Key-Size Required : 8".toUpperCase());
		add(keyinput);

		ctextinfo =new JLabel();												// CipherText Information Label
		ctextinfo.setBounds((scrW-750)/3+300, scrH/3-40, 250, 50);
		ctextinfo.setText("Enter Text to be Decrypted");
		ctextinfo.setVisible(false);
		add(ctextinfo);

		keyinputinfo =new JLabel();												// Key Information Label
		keyinputinfo.setBounds((scrW-750)/3+600, scrH/3-40, 250, 50);
		keyinputinfo.setText("Enter Secret Key");
		keyinputinfo.setVisible(false);
		add(keyinputinfo);

		fileinfo=new JLabel();													// CipherText Location Label		
		fileinfo.setBounds((scrW-500)/2-40, scrH/3+50, 500, 50);
		fileinfo.setVisible(false);
		add(fileinfo);

		cipherop= new JTextArea();												// PlainText TextArea
		cipherop.setEditable(false);  
		cipherop.setCursor(null);  
		cipherop.setOpaque(false);  
		cipherop.setFocusable(false);
		cipherop.setLineWrap(true);
		cipherop.setBounds(0, scrH-scrH/3, scrW-20, 75);
		cipherop.setVisible(false);
		cipherop.setWrapStyleWord(true);
		add(cipherop);

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
		frame = new Decrypt_Sheet();
	}

	public void actionPerformed(ActionEvent evt)
	{
		String str=evt.getActionCommand();

		if(str.equals("Home"))
		{
			this.dispose();
			new AFS_Home();
		}

		if(str.equals("Refresh"))
		{
			this.dispose();
			new Decrypt_Sheet();
		}

		if(str.equals("Start Process"))
		{
			try
			{
				ReadExcel();
			} 
			catch (Exception e1) 
			{
				e1.printStackTrace();
			}

			startp.setEnabled(false); 
			startd.setVisible(true);
			keyinput.setVisible(true);
		}

		if(str.equals("Start Decryption"))
		{	  	
			if(!keyinput.getText().equals(""))
			{
				if(!keyinput.getText().equals(""))
					decryptionKey=keyinput.getText();
			}
			System.out.println(keyinput.getText());

			try
			{
				if(decryptionKey.length()<8)
				{
					JOptionPane jop=new JOptionPane();
					JOptionPane.showMessageDialog(jop,"Minimum 8 characters required".toUpperCase()); 
				}
				else
				{
					if((decryptionKey.length()%16)!=0 && decryptionKey.length()<16 && decryptionKey.length()>=8) 
						while((decryptionKey.length()%16)!=0)
							decryptionKey+=" ";

					if((decryptionKey.length()%16)!=0 && decryptionKey.length()>16)
						decryptionKey=decryptionKey.substring(0,16);

					process();
				} 
			} 
			catch (Exception e1) 
			{
				e1.printStackTrace();
			}
		}
	}

	public static void process()
	{	
		try
		{  
			if((decryptionKey.length()%16)!=0 && decryptionKey.length()>16)
				decryptionKey=decryptionKey.substring(0,16);

			newcontent=new String[content.length];

			for(int z=0;z<rcnt;z++)
			{
				for(int j=0;j<colcnt[rcnt-1];j++)
					System.out.print(column[z][j]+" ");
				System.out.println("\n");
			}
			System.out.println("Key : Length >> "+decryptionKey+" : "+decryptionKey.length());

			newColumn=new String[rcnt][colcnt.length];
			title="";
			newplaintext="";
			
			for(int i=0;i<rcnt;i++)
				for(int j=0;j<colcnt[i];j++)
				{
					plaintext="";
					ciphertext="";
					if(column[i][j]!=null)
					{
						ciphertext=column[i][j];

						plaintext=decrypt(ciphertext, decryptionKey);
						newplaintext+=plaintext;
						newColumn[i][j]=plaintext;

						if(i==0)
							if(newColumn[0][j]!=null)
								title+=newColumn[0][j].trim()+" ";
					}
					else
						newColumn[i][j]=" ";
				}

			cipherop.setText(newplaintext);
			cipherop.setVisible(true);
			writeExcel();
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		} 
	}

	public static String decrypt(String cipherText, String decryptionKey) throws Exception
	{
		Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
		SecretKeySpec key = new SecretKeySpec(decryptionKey.getBytes("UTF-8"), "AES");
		cipher.init(Cipher.DECRYPT_MODE, key, new IvParameterSpec(IV.getBytes("UTF-8")));
		byte[] decodedValue = new Base64().decodeBase64(cipherText);
		byte[] decValue = cipher.doFinal(decodedValue);
		String decryptedValue = new String(decValue);
		return decryptedValue;
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
			rcnt=0;
			while(rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				ccnt=0;
				while(cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_NUMERIC: 
						double temp1=cell.getNumericCellValue();
						column[rcnt][ccnt]=temp1+"";
						break;

					case Cell.CELL_TYPE_STRING:
						String temp2=cell.getStringCellValue();
						column[rcnt][ccnt]=temp2;
						break;
					}
					ccnt++;
				}
				colcnt[rcnt]=ccnt;
				rcnt++;			
			}
			file.close();	

			content=plaintext.split("\n");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	public static void writeExcel() 
	{
		int k=0;
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		System.out.println("rcnt >> "+rcnt);
		for(int i=0;i<rcnt;i++)
			if(colcnt[i]==0)
				data.put(i+"", new Object[] {""});
			else
				if(colcnt[i]==1)
					data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+""});
				else
					if(colcnt[i]==2)
						data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+""});
					else
						if(colcnt[i]==3)
							data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+""});
						else
							if(colcnt[i]==4)
								data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+""});
							else
								if(colcnt[i]==5)
									data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+""});
								else
									if(colcnt[i]==6)
										data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+"", newColumn[i][5]+""});
									else
										if(colcnt[i]==7)
											data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+"", newColumn[i][5]+"", newColumn[i][6]+""});
										else
											if(colcnt[i]==8)
												data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+"", newColumn[i][5]+"", newColumn[i][6]+"", newColumn[i][7]+""});
											else
												if(colcnt[i]==9)
													data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+"", newColumn[i][5]+"", newColumn[i][6]+"", newColumn[i][7]+"", newColumn[i][8]+""});
												else
													if(colcnt[i]==10)
														data.put(cnt[k++]+"", new Object[] {newColumn[i][0]+"", newColumn[i][1]+"", newColumn[i][2]+"", newColumn[i][3]+"", newColumn[i][4]+"", newColumn[i][5]+"", newColumn[i][6]+"", newColumn[i][7]+"", newColumn[i][8]+"", newColumn[i][9]+""});

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
		System.out.println("title >>->> "+title);
		if(title.contains("Displaying Graphical Analysis"))
		{
			System.out.println("in DGA");
			HSSFFont fontGraphColor = workbook.createFont();
			fontGraphColor.setColor(IndexedColors.WHITE.getIndex());

			HSSFCellStyle styleGraphColor[]=new HSSFCellStyle[25];
			for(int i=0;i<20;i++)
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
					}
					else
					{
						if(rownum==1)
						{
							cell.setCellStyle(style);
							sheet.autoSizeColumn(2);
							sheet.autoSizeColumn(3);
							sheet.autoSizeColumn(4);

							if(cell.getStringCellValue()!=null)
							{
								System.out.println("TRUE for cellNum : "+(cellnum-1));
								sheet.autoSizeColumn(cellnum-1);
							}

						}
						else
							if(rownum==3)
							{
								cell.setCellStyle(stylesubhead);

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
					}
					if(obj instanceof String)
						cell.setCellValue((String)obj);
					else if(obj instanceof Integer)
						cell.setCellValue((Integer)obj);
				}			
			}
		}
		else
			if(title.contains("Displaying Detailed Analysis"))
			{
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

						sheet.autoSizeColumn(0);
					}			
				}
			}
			else
				//if(title.contains("Displaying Performance-Index wise Analysis"))
			{
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
			}
		try 
		{
			File f=new File("Excel Decrypted.xls");
			FileOutputStream out = new FileOutputStream(f);
			workbook.write(out);
			out.close();

			System.out.println("Excel Decrypted.xls written successfully on disk.");
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}