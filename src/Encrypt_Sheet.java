import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.UUID;

import javax.crypto.Cipher;
import javax.crypto.SecretKey;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.UIManager;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

public class Encrypt_Sheet extends JFrame implements ActionListener
{
	static String IV = "AAAAAAAAAAAAAAAA";
	static String plaintext = "";
	static String ciphertext = "";
	static String encryptionKey="";
	static String flag_entry=""; 

	static File selectedFile;

	static SecretKey secretkey=null;
	static Cipher cipher;

	static JButton startp, starte, home, refresh, sendByMail, saveToFile, saveAutoKey;
	static JDialog option;
	static JFrame frame;
	static JTextField textinput, keyinput;
	static JLabel fileinfo, ptextinfo, keyinputinfo, keyselectinfo, choice;
	static JTextArea cipherop;
	static JCheckBox keyselect;

	static String filePath="";
	static String fileName="";
	static String frameName="";

	static int response=0;
	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;

	private static int ccnt=0;
	private static int rcnt=0;
	private static int[] colcnt=new int[30];
	private static String[][] newColumn;
	private static String[][] column=new String[30][25];
	private String[] cnt={"01", "02", "03", "04", "05" ,"06", "07", "08", "09", "10", "11" ,"12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30"};

	public Encrypt_Sheet(String filePath, String fileName, String frameName)
	{
		setTitle("Encrypt_Sheet");
		setLayout(null);
		setSize(scrW, scrH);

		home=new JButton("Home");														// Redirect to Activity
		home.setBounds(5, 5, 100, 40);
		add(home);
		home.setVisible(true);
		home.addActionListener(this);

		refresh=new JButton("Refresh");													// Restart Same Frame
		refresh.setBounds((scrW-5-100), 5, 100, 40);
		add(refresh);
		refresh.setVisible(true);
		refresh.addActionListener(this);

		sendByMail=new JButton("Send Encrypted Contents via E-Mail".toUpperCase());		// Send by Mail 
		sendByMail.setBounds(10,2*scrH/3+130, (scrW-20), 50);
		add(sendByMail);
		sendByMail.setVisible(false);
		sendByMail.setEnabled(false);
		sendByMail.addActionListener(this);

		startp=new JButton("Start Process");											// Start Process Button
		startp.setBounds((scrW-500)/3, scrH/3, 250, 50);
		add(startp);
		startp.setVisible(true);
		startp.addActionListener(this);

		starte=new JButton("Start Encryption");											// Start Encryption Button
		starte.setBounds(3*(scrW-500)/4+100, scrH/2, 250, 50);							//setBounds((scrW-250)/2, scrH/2, 250, 50);
		add(starte);
		starte.setVisible(false);
		starte.addActionListener(this);

		keyselect=new JCheckBox("Auto Generate Key");									// Auto Key CheckBox
		keyselect.setBounds((scrW-500)/3+600+50, scrH/3-15, 200, 50);
		keyselect.setVisible(false);
		keyselect.addActionListener(this);
		add(keyselect);

		keyinput = new JTextField();													// Manual Key TextField
		keyinput.setBounds((scrW-500)/3+300, scrH/3, 250, 50);
		keyinput.setVisible(false);
		keyinput.setToolTipText("Minimum Key-Size Required : 8".toUpperCase());
		add(keyinput);

		keyinputinfo =new JLabel();														// Key Information Label
		keyinputinfo.setBounds((scrW-500)/3+300+75, scrH/3+40, 250, 50);
		keyinputinfo.setText("(Enter Secret Key)");
		keyinputinfo.setVisible(false);
		add(keyinputinfo);

		choice =new JLabel();														// Key Information Label
		choice.setBounds((scrW-500)/3+575, scrH/3, 200, 50);
		choice.setText("(OR)");
		choice.setVisible(false);
		add(choice);

		keyselectinfo =new JLabel();													// Label regarding Selecting key Automatically
		keyselectinfo.setBounds((scrW-500)/3+600+25, scrH/3+5, 200, 50);
		keyselectinfo.setText("(Select to AUTOGENERATE Key)");
		keyselectinfo.setVisible(false);
		add(keyselectinfo);

		this.fileName=fileName;
		this.filePath=filePath;
		this.frameName=frameName;

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

		frame = new Encrypt_Sheet("", "", "");
	}

	public void actionPerformed(ActionEvent evt)
	{
		String str=evt.getActionCommand();

		if(str.equals("Auto Generate Key"))
		{
			if(keyselect.isSelected()==true)
			{
				keyinput.setText("");
				keyinput.setEnabled(false);
			}
			if(keyselect.isSelected()==false)
				keyinput.setEnabled(true);
		}

		if(str.equals("Home"))
		{
			this.dispose();
			new AFS_Home();
		}

		if(str.equals("Refresh"))
		{
			this.dispose();
			frame = new Encrypt_Sheet(filePath, fileName, frameName);
		}

		if(str.equals("Send Encrypted Contents via E-Mail".toUpperCase()))
		{
			writeExcel();
			new SendByEmail(filePath, fileName, frameName);
			this.dispose();
		}

		if(str.equals("Start Process"))
		{
			startp.setEnabled(false); 

			try
			{
				ReadExcel();
			} 
			catch (Exception e1) 
			{
				e1.printStackTrace();
			}

			starte.setVisible(true);
			keyinput.setVisible(true);
			keyselectinfo.setVisible(true);
			choice.setVisible(true);
			keyselect.setVisible(true);
			keyinputinfo.setVisible(true);
		}

		if(str.equals("Start Encryption"))
		{	  
			try
			{
				if(keyselect.isSelected())
				{
					String autoStr = UUID.randomUUID().toString();
					encryptionKey=autoStr.substring(0, 16);

					FileOutputStream stream = null;
					PrintStream out = null;
					try
					{
						stream = new FileOutputStream("Encryption Key"+".txt"); 
						String text = encryptionKey;
						out = new PrintStream(stream);
						out.print(text);
					} 
					catch (Exception ex)
					{
						ex.printStackTrace();
					} 
					finally 
					{
						try
						{
							if(stream!=null) stream.close();
							if(out!=null) out.close();
						}
						catch(Exception ex) 
						{
							ex.printStackTrace();
						}
					}
				}
				else
					encryptionKey=keyinput.getText();


				if(encryptionKey.length()<8)
				{
					JOptionPane jop=new JOptionPane();
					JOptionPane.showMessageDialog(jop,"Minimum 8 characters required".toUpperCase()); 
				}
				else
				{
					if((encryptionKey.length()%16)!=0 && encryptionKey.length()<16 && encryptionKey.length()>=8) 
						while((encryptionKey.length()%16)!=0)
							encryptionKey+=" ";

					if((encryptionKey.length()%16)!=0 && encryptionKey.length()>16)
						encryptionKey=encryptionKey.substring(0,16);

					process();
				} 
			}
			catch (Exception e1) 
			{
				e1.printStackTrace();
			}
		}
	}

	public static void process() throws Exception
	{ 
		try
		{   
			newColumn=new String[rcnt][colcnt.length];

			for(int i=0;i<(rcnt);i++)
			{
				for(int j=0;j<colcnt[i];j++)
				{
					plaintext="";
					ciphertext="";

					System.out.println("column["+(i)+"]["+j+"] ~> "+column[i][j]);

					if(column[i][j]!=null)
					{
						plaintext=column[i][j];
						if((plaintext.length()%16)!=0) 
							while((plaintext.length()%16)!=0)
								plaintext+=" ";

						ciphertext=encrypt(plaintext, encryptionKey);
						newColumn[i][j]=ciphertext;
					}
					else
					{
						newColumn[i][j]=" ";
						System.out.println("column["+i+"]["+j+"] >>-> "+newColumn[i][j]);
					}

					sendByMail.setVisible(true);
					sendByMail.setEnabled(true);
				}
			}
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	public static String encrypt(String plainText, String  encryptionKey) throws Exception 
	{
		Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
		SecretKey key = new SecretKeySpec(encryptionKey.getBytes("UTF-8"), "AES");
		cipher.init(Cipher.ENCRYPT_MODE, key, new IvParameterSpec(IV.getBytes("UTF-8")));
		byte[] encVal = cipher.doFinal(plainText.getBytes());
		String encryptedValue = new Base64().encodeToString(encVal);
		return encryptedValue;
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
				colcnt[rcnt++]=ccnt;
			}
			file.close();
			System.out.println("rcnt >> "+rcnt);
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

	public void writeExcel() 
	{
		int k=0;
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

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

				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}			
			sheet.autoSizeColumn(0);
		}
		try 
		{
			File f=new File("Excel Encrypted.xls");
			FileOutputStream out = new FileOutputStream(f);
			workbook.write(out);

			System.out.println("Excel Encrypted.xls written successfully on disk.");
			out.close();
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}