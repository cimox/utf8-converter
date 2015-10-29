package ui;
import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import java.awt.Toolkit;

import javax.swing.border.MatteBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Color;

import javax.swing.JButton;
import javax.swing.JTextField;

import java.awt.Font;

import javax.swing.ImageIcon;
import javax.swing.JLabel;
import javax.swing.AbstractAction;

import java.awt.event.ActionEvent;

import javax.swing.Action;

import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;

import javax.swing.JFormattedTextField;
import javax.swing.SwingConstants;
import javax.swing.JTextArea;
import javax.swing.JTextPane;

import core.Convertor;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JSeparator;
import javax.swing.JInternalFrame;
import javax.swing.JToolBar;
import javax.swing.JProgressBar;
import javax.swing.border.TitledBorder;
import javax.swing.UIManager;



/**
 * @author osk10789
 *
 */
public class MainFrame extends JFrame {

	private JPanel contentPane;
	private JTextField txtFilenamecsv;
	private File[] filesToConvert;
	private JTextArea selectedFilesTextArea;
	
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {					
					MainFrame frame = new MainFrame();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});		
	}

	
	/**
	 * @param paths where to look for jre
	 * @return path to java.exe
	 */
	private String getJavaPath(File[] paths) {
		for (File path : paths) {
			System.out.println("Looking for java in path: " + path);
			File[] listOfFiles = path.listFiles();
			if (listOfFiles != null) {
				for (int i = 0; i < listOfFiles.length; i++) {
					if (listOfFiles[i].isDirectory() && listOfFiles[i].getName().contains("jre")) { // JRE is in, test version
						char c = listOfFiles[i].getName().charAt(listOfFiles[i].getName().length()-1);
						int version = c;
						if (version >= 7) { // All OK, use this JRE
							return new String(listOfFiles[i].getAbsolutePath() + "\\bin\\java.exe");
						}
					}
				}
			}
		}
		return null; // if not found
	}
		
	private String constructCommand(String javaPath, File pathToFile, String outputName) throws IOException {
		File currentPath = new File(".");
		StringBuilder cmd = new StringBuilder("\"" + javaPath + "\"" + 
				" \"-Dfile.encoding=utf-8\" -jar \"" + 
				currentPath.getCanonicalPath() + "\\lib\\excel2csv.jar\" \"" + 
				pathToFile + "\" \"" + 
				pathToFile.getParent() + "\\" + outputName + "\"" +
				" no no");
//		C:\Program Files (x86)\Java\jre7\bin "-Dfile.encoding=utf-8" -jar "excel2csv\excel2csv.jar" input.xls input.csv no no
		return cmd.toString();
		
	}
	
	/**
	 * @return state of phase converting documents
	 */
	private boolean convertToCSV(File[] filesToConvert) {
		//TODO: umoznit vybrat presnu cestu k jave - ako argument metode pojde
		//TODO: umoznit pouzivatelovi specifikovat meno vystupneho CSV
		File[] paths = new File[3];
		paths[0] = new File("."); paths[1] = new File("C:\\Program Files (x86)\\Java"); paths[2] = new File("C:\\Program Files\\Java");
		
		String javaPath = getJavaPath(paths);
		System.out.println(javaPath);			
		
		for (File file : filesToConvert) {		
//			String cmd = constructCommand(javaPath, file,	file.getName().replace("xls", "csv"));				
			
			System.setProperty("file.encoding", "UTF-8");
			final File inputFile = new File(file.toString());			
			final File outputFile = new File(file.getParent() + "\\" + file.getName().replace("xls", "csv"));				
//	            Converter.convertExcel2CSV(file, outputFile, false, false)
			Thread convert = new Thread(new Convertor() {
				
				@Override
				public void run() {
					try {
						if (Convertor.convertExcel2CSV(inputFile, outputFile, false, false) == 1)
							
						
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}						
				}
			});
            convert.start();
		}
		

		
		return false;
	}
	
	/**
	 * Create the frame.
	 */
	public MainFrame() {
		setResizable(false);
		setIconImage(Toolkit.getDefaultToolkit().getImage("C:\\Users\\osk10789\\Downloads\\orange_logo.png"));
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 640, 480);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new MatteBorder(1, 1, 1, 1, (Color) new Color(0, 0, 0)));
		setContentPane(contentPane);
		contentPane.setLayout(null);
			
		selectedFilesTextArea = new JTextArea();
		selectedFilesTextArea.setEditable(false);
//		File chooser button	
		filesToConvert = null;
		JButton fileChooseBtn = new JButton("Choose file[s]");
		fileChooseBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				chooser.setMultiSelectionEnabled(true);
			    FileNameExtensionFilter filter = new FileNameExtensionFilter(
			        "Excel files", "xls"/*, "xlsx"*/); // TODO: Dorobit konverziu XLSX		  
			    chooser.setFileFilter(filter);
			    int returnVal = chooser.showOpenDialog(MainFrame.this);
			    if(returnVal == JFileChooser.APPROVE_OPTION) {
			    	filesToConvert = chooser.getSelectedFiles();	
			    	selectedFilesTextArea.setText("");
					for (File file : filesToConvert) selectedFilesTextArea.append(file.getName() + "\n"); // print selected files to convert
			    }
			}
		});
//		fileChooseBtn.setSelectedIcon(new ImageIcon(MainFrame.class.getResource("/com/sun/javafx/scene/control/skin/caspian/images/capslock-icon.png")));
		fileChooseBtn.setBackground(new Color(153, 204, 255));
		fileChooseBtn.setFont(new Font("Arial", Font.BOLD, 12));
		fileChooseBtn.setBounds(52, 35, 160, 36);
		contentPane.add(fileChooseBtn);				
		
		
		txtFilenamecsv = new JTextField();
		txtFilenamecsv.setEditable(false);
		txtFilenamecsv.setEnabled(false);
		txtFilenamecsv.setFont(new Font("Arial", Font.PLAIN, 12));
		txtFilenamecsv.setText("filename.csv");
		txtFilenamecsv.setBounds(52, 92, 160, 38);
		contentPane.add(txtFilenamecsv);
		txtFilenamecsv.setColumns(10);
		
		// Convert to CSV
		JButton btnNewButton_1 = new JButton("Convert");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (filesToConvert == null) {
					// TODO: log errors somewhere
				} else { // convert files to CSV
					convertToCSV(filesToConvert);
				}
			}
		});
		btnNewButton_1.setBackground(new Color(102, 204, 102));
		btnNewButton_1.setFont(new Font("Arial", Font.BOLD, 12));
		btnNewButton_1.setBounds(52, 154, 160, 36);
		contentPane.add(btnNewButton_1);
		
		JLabel label = new JLabel("1.");
		label.setFont(new Font("Arial", Font.BOLD, 32));
		label.setBounds(10, 35, 32, 36);
		contentPane.add(label);
		
		JLabel label_1 = new JLabel("2.");
		label_1.setFont(new Font("Arial", Font.BOLD, 32));
		label_1.setBounds(10, 92, 32, 36);
		contentPane.add(label_1);
		
		JLabel label_2 = new JLabel("3.");
		label_2.setFont(new Font("Arial", Font.BOLD, 32));
		label_2.setBounds(10, 154, 32, 36);
		contentPane.add(label_2);
		
		JLabel lblXlsxNotSupported = new JLabel("Note: XLSX NOT SUPPORTED yet");
		lblXlsxNotSupported.setFont(new Font("Arial", Font.BOLD, 12));
		lblXlsxNotSupported.setBounds(10, 210, 202, 14);
		contentPane.add(lblXlsxNotSupported);
		
		JMenuBar menuBar = new JMenuBar();
		menuBar.setBounds(0, 0, 634, 21);
		contentPane.add(menuBar);
		
		JMenu mnNewMenu = new JMenu("File");
		mnNewMenu.setEnabled(false);
		menuBar.add(mnNewMenu);
		
		JMenu mnNewMenu_1 = new JMenu("Settings");
		mnNewMenu_1.setEnabled(false);
		menuBar.add(mnNewMenu_1);
		
		JMenu mnAbout = new JMenu("Help");
		mnAbout.setEnabled(false);
		menuBar.add(mnAbout);
		
		JMenuItem mntmHelp = new JMenuItem("Help contents");
		mntmHelp.setEnabled(false);
		mnAbout.add(mntmHelp);
		
		JSeparator separator = new JSeparator();
		mnAbout.add(separator);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mntmAbout.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// TODO: make about and help
			}
		});
		mnAbout.add(mntmAbout);
		
		JProgressBar progressBar = new JProgressBar();
		progressBar.setEnabled(false);
		progressBar.setBounds(10, 404, 614, 21);
		contentPane.add(progressBar);
		
		JLabel label_3 = new JLabel("0%");
		label_3.setBounds(321, 427, 46, 14);
		contentPane.add(label_3);
		
		JLabel lblConverting = new JLabel("Completed");
		lblConverting.setFont(new Font("Arial", Font.PLAIN, 12));
		lblConverting.setBounds(251, 426, 70, 14);
		contentPane.add(lblConverting);
		
		JSeparator separator_1 = new JSeparator();
		separator_1.setBounds(10, 391, 614, 2);
		contentPane.add(separator_1);
		
		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Selected files to convert", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(231, 35, 281, 155);
		contentPane.add(panel);
		panel.setLayout(null);
				
		selectedFilesTextArea.setFont(new Font("Arial", Font.PLAIN, 11));
		selectedFilesTextArea.setBounds(6, 16, 265, 128);
		panel.add(selectedFilesTextArea);
		selectedFilesTextArea.setLineWrap(true);
		
		JLabel lblNewLabel = new JLabel("New label");
		lblNewLabel.setBounds(205, 235, 243, 100);
		contentPane.add(lblNewLabel);
	}	
}
