import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.UIManager;

public class AFS_Home extends JFrame implements ActionListener
{
	static JFrame frame;
	
	static JButton sheet1;
	static JButton sheet2;
	static JButton sheet3;
	static JButton decrypt;
	
	static JLabel desc;
	static JLabel descHome;
	
	Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
	int scrW=screenSize.width, scrH=screenSize.height;
		
	public AFS_Home()
	{
		setTitle("Assesment Report Splitter - Home ");
		setLayout(null);
		setSize(scrW, scrH);
		
		Font head=new Font("Book Antiqua", Font.BOLD, 60);
		Font subHead=new Font("Courier", Font.BOLD, 40);
		
		decrypt=new JButton("Decrypt Report File");									// Button to Process Sheet-1 of Assesment Report
		decrypt.setBounds((scrW-750)/4+900, scrH/3+100, 250, 50);
		add(decrypt);
		decrypt.setVisible(true);
		decrypt.addActionListener(this);
		
		sheet1=new JButton("Graphical Analysis");									// Button to Process Sheet-1 of Assesment Report
		sheet1.setBounds((scrW-750)/4-100, scrH/3+100, 250, 50);
		add(sheet1);
		sheet1.setVisible(true);
		sheet1.addActionListener(this);
		
		sheet2=new JButton("Ranking");												// Button to Process Sheet-2 of Assesment Report
		sheet2.setBounds((scrW-750)/4+200, scrH/3+100, 250, 50);
		add(sheet2);
		sheet2.setVisible(true);
		sheet2.addActionListener(this);
		
		sheet3=new JButton("Performance Index");									// Button to Process Sheet-3 of Assesment Report
		sheet3.setBounds((scrW-750)/4+500, scrH/3+100, 250, 50);
		add(sheet3);
		sheet3.setVisible(true);
		sheet3.addActionListener(this);
		
		desc=new JLabel("Select Type of Report");									
		desc.setBounds((scrW-600)/2-200, scrH/3, 600, 50);
		desc.setFont(subHead);
		desc.setForeground(Color.BLUE);
		add(desc);
		desc.setVisible(true);
		
		descHome=new JLabel("Report Extraction Tool");									
		descHome.setBounds((scrW-650)/2, scrH/10-50, 650, 75);
		descHome.setFont(head);
		add(descHome);
		descHome.setForeground(Color.getHSBColor(6, 1, 5));
		descHome.setVisible(true);
		
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
		frame = new AFS_Home();
	}
	
	public void actionPerformed(ActionEvent evt) 
	{
		String str=evt.getActionCommand();
		
		if(str.equals("Graphical Analysis"))
		{
			new AFS_Sheet_1();
			this.dispose();
		}
		
		if(str.equals("Ranking"))
		{
			new AFS_Sheet_2();
			this.dispose();
		}
		
		if(str.equals("Performance Index"))
		{
			new AFS_Sheet_3();
			this.dispose();
		}
		
		if(str.equals("Decrypt Report File"))
		{
			new Decrypt_Sheet();
			this.dispose();
		}
	}
}