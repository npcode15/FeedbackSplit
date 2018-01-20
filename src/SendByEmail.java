import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import java.util.ArrayList;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.AuthenticationFailedException;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JFileChooser;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.UIManager;

public class SendByEmail  extends JFrame implements ActionListener
{
	static JFrame frame;
	static JLabel user;
	static JLabel pass;
	static JLabel emailid;
	static JButton home;
	static JButton back;
	private JButton sendmail;
	private JTextField username;
	private JPasswordField password;
	private JComboBox<String> emaillist;

	static String toMail;
	static String filePath;
	static String fileName;
	static String senderUsername;
	static String senderPassword;
	static String frameName;

	static Session mailSession;
	static MimeMessage emailMessage;
	static Properties emailProperties;

	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;
	private String[] mail={"Select_Recipient's_Email-ID", "navneet.navneet.pandey6@gmail.com","navneetpandey4@gmail.com"};

	public SendByEmail(String string)
	{
		System.out.println(string);
	}

	public SendByEmail(String fPath, String fName, String frName)
	{
		setTitle("Sending E-Mail");
		setLayout(null);
		setSize(scrW, scrH);

		home=new JButton("Home");											// Redirect to Home
		home.setBounds(15, 5, 100, 40);
		add(home);
		home.setVisible(true);
		home.addActionListener(this);

		back=new JButton("Back");											// Go Back to Previous Frame
		back.setBounds((scrW-15-100), 5, 100, 40);
		add(back);
		back.setVisible(true);
		back.addActionListener(this);

		user=new JLabel("User Name");																// Username Label
		user.setBounds((scrW-500)/2, scrH/5, 250, 50);
		user.setVisible(true);
		add(user);
		user.setVisible(true);

		pass=new JLabel("Password");																// Password Label
		pass.setBounds((scrW-500)/2, scrH/5+100, 250, 50);		
		pass.setVisible(true);
		add(pass);
		pass.setVisible(true);

		username = new JTextField();																// For Entry of Username of Sender of the Email
		username.setBounds((scrW-500)/2+300, scrH/5, 250, 50);
		username.setVisible(true);
		username.addActionListener(this);
		add(username);

		password = new JPasswordField();															// For Entry of Password of Sender of the Email
		password.setBounds((scrW-500)/2+300, scrH/5+100, 250, 50);
		password.setEchoChar('*');
		password.setVisible(true);
		password.addActionListener(this);
		add(password);

		emaillist=new JComboBox<String>();
		emaillist.setBounds((scrW-500)/2, scrH/5+200, 550, 50);										// List of E-mails
		emaillist.setVisible(true);
		add(emaillist);
		emaillist.addActionListener(this);
		emaillist.setVisible(true);
		for(int i=0;i<mail.length;i++)
			emaillist.addItem(mail[i]);

		sendmail=new JButton("Send Mail");															// Press to Send Email
		sendmail.setBounds((scrW-500)/2, scrH/5+300, 550, 50);		
		sendmail.setVisible(true);
		add(sendmail);
		sendmail.setVisible(true);
		sendmail.addActionListener(this);

		filePath=fPath;
		fileName=fName;
		frameName=frName;

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setVisible(true);
	}

	public static void main(String args[]) throws AddressException, MessagingException 
	{
		try 
		{
			UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
		} 
		catch(Exception e) 
		{
			e.printStackTrace(); 
		}

		frame = new SendByEmail("Don't Run Directly !!");
	}

	public static void setMailServerProperties() 
	{

		String smtp_Port="";
		
		//if(toMail.contains("@gmail") || toMail.contains("mail@"))
			smtp_Port = "587";													// Gmail's SMTP Port
		//else
			//smtp_Port = "587";
		
		emailProperties = System.getProperties();
		emailProperties.put("mail.smtp.auth", "true");
		emailProperties.put("mail.smtp.port", smtp_Port);
		emailProperties.put("mail.smtp.starttls.enable", "true");
		emailProperties.put("mail.smtp.ssl.trust", "smtp.gmail.com");
	}

	public static void createEmailMessage() throws AddressException,MessagingException 
	{	
		String emailSubject = "Feedback Assesment Report";
		String emailBody = "Feedback Assesment Report";

		mailSession = Session.getDefaultInstance(emailProperties, null);
		emailMessage = new MimeMessage(mailSession);

		emailMessage.addRecipient(Message.RecipientType.TO, new InternetAddress(toMail));

		emailMessage.setSubject(emailSubject);
		//emailMessage.setContent(emailBody, "text/html");					//for a HTML+Text email
	}

	public static void sendEmail() throws AddressException, MessagingException 
	{
		String emailHost = "smtp.gmail.com";
		String fromUser = senderUsername; 							//Sender's Gmail Username
		String fromUserEmailPassword = senderPassword; 		     	//Sender's Gmail Password

		BodyPart messageBodyPart = new MimeBodyPart();

		messageBodyPart.setText("Feedback Assesment Report(from HoD - I.T.)" );
		
		Multipart multipart = new MimeMultipart();

		multipart.addBodyPart(messageBodyPart);

		messageBodyPart = new MimeBodyPart();
		DataSource source = new FileDataSource(filePath);
		messageBodyPart.setDataHandler(new DataHandler(source));
		messageBodyPart.setFileName(fileName);
		multipart.addBodyPart(messageBodyPart);

		emailMessage.setContent(multipart);

		Transport transport = mailSession.getTransport("smtp");
		transport.connect(emailHost, fromUser, fromUserEmailPassword);
		transport.sendMessage(emailMessage, emailMessage.getAllRecipients());
		transport.close();

		System.out.println("Email sent successfully.");
		JOptionPane p=new JOptionPane();
		JOptionPane.showMessageDialog(p," File \""+fileName+"\"\n Located at \""+filePath+"\"\n From \""+senderUsername+"\"\n To \""+toMail+"\" "+" was sent Succesfully");
	}

	public void actionPerformed(ActionEvent evt) 
	{
		String str=evt.getActionCommand();

		toMail = String.valueOf(emaillist.getSelectedItem());

		senderUsername=username.getText().trim();
		senderPassword=password.getText().trim();

		if(str.equals("Home"))
		{
			this.dispose();
			new AFS_Home(); 
		}

		if(str.equals("Back"))
		{
			this.dispose();
			System.out.println("frameName >--> "+frameName);
			if(frameName.equals("AFS Sheet-1"))
				new AFS_Sheet_1();
			else
				if(frameName.equals("AFS Sheet-2"))
					new AFS_Sheet_2();
				else
					if(frameName.equals("AFS Sheet-3"))
						new AFS_Sheet_3();
		}

		if(str.equals("Send Mail"))
		{
			setMailServerProperties();

			try 
			{
				createEmailMessage();
			} 
			catch (AddressException e) 
			{
				e.printStackTrace();
			} 
			catch (MessagingException e)
			{
				e.printStackTrace();
			}

			try 
			{
				sendEmail();
			} 
			catch (AddressException e) 
			{
				e.printStackTrace();
			} 
			catch (AuthenticationFailedException e) 
			{
				System.out.println("Authentication Failed");
				JOptionPane p=new JOptionPane();
				JOptionPane.showMessageDialog(p,"Authentication Failed...!!");
			} 
			catch (SendFailedException e) 
			{
				System.out.println("Invalid Recipient Addresses");
				JOptionPane p=new JOptionPane();
				JOptionPane.showMessageDialog(p,"Invalid Recipient Addresses...!!");
			} 
			catch (MessagingException e)
			{
				e.printStackTrace();
			}
		}
	}
}