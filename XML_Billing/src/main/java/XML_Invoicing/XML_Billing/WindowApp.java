package XML_Invoicing.XML_Billing;

import java.awt.EventQueue;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;


import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.Font;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.awt.event.ActionEvent;

public class WindowApp {

	private JFrame frame;
	private JTextField textField;
	private JTextField textField_1;
	private JTextField textField_2;
	private JTextField textField_3;

	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					WindowApp window = new WindowApp();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public WindowApp() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setTitle("XML_Invoicing");
		frame.setBounds(100, 100, 600, 253);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		// Create the tabbed pane
		JTabbedPane tabbedPane = new JTabbedPane();
		tabbedPane.setBounds(10, 10, 570, 200); 
		// Create the first tab panel
		JPanel tab1 = new JPanel();
		tabbedPane.addTab("XML Invoice", tab1);
		tab1.setLayout(null);

		JLabel lblNewLabel = new JLabel("Invoice Number:");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel.setBounds(30, 30, 120, 25);
		tab1.add(lblNewLabel);

		textField = new JTextField();
		textField.setBounds(160, 30, 270, 25);
		tab1.add(textField);
		textField.setColumns(10);

		JLabel lblNewLabel_1 = new JLabel("Select Excel File:");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel_1.setBounds(30, 80, 120, 25);
		tab1.add(lblNewLabel_1);

		textField_1 = new JTextField();
		textField_1.setBounds(160, 80, 270, 25);
		tab1.add(textField_1);
		textField_1.setColumns(10);
		
		

		JButton btnNewButton = new JButton("Browse");
		
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select the Excel file to process");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField_1.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(null, "Select a Valid Excel file");
				}
			}
		});
		btnNewButton.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton.setBounds(450, 80, 90, 25);
		tab1.add(btnNewButton);

		JButton btnNewButton_1 = new JButton("Start");
		btnNewButton_1.addActionListener(new ActionListener() {
		    public void actionPerformed(ActionEvent e) {
		 
		        final String excelFilePath = textField_1.getText();
		        final String invoiceNumber = textField.getText();
		        
		        if (!excelFilePath.isEmpty() && !invoiceNumber.isEmpty()) {
		        
		            File file = new File(excelFilePath);
		            if (file.exists()) {
		                XML_Invoicing.main(excelFilePath, invoiceNumber);
		                JOptionPane.showMessageDialog(textField, "Process Done");
		            } else {
		                JOptionPane.showMessageDialog(textField, "File not found. Please check the file path.");
		            }
		        } else {
		            JOptionPane.showMessageDialog(textField, "Select the Excel File and Give Invoice ID");
		        }
		    }
		});

		btnNewButton_1.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton_1.setBounds(230, 130, 90, 30);
		tab1.add(btnNewButton_1);

		JPanel tab2 = new JPanel();
		tabbedPane.addTab("Summary", tab2);
		tab2.setLayout(null);

		JLabel lblNewLabel_2 = new JLabel("Excel File Name:");
		lblNewLabel_2.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel_2.setBounds(30, 30, 120, 25);
		tab2.add(lblNewLabel_2);

		textField_2 = new JTextField();
		textField_2.setBounds(160, 30, 270, 25);
		tab2.add(textField_2);
		textField_2.setColumns(10);

		JLabel lblNewLabel_3 = new JLabel("Files Path:");
		lblNewLabel_3.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel_3.setBounds(30, 80, 120, 25);
		tab2.add(lblNewLabel_3);

		textField_3 = new JTextField();
		textField_3.setBounds(160, 80, 270, 25);
		tab2.add(textField_3);
		textField_3.setColumns(10);

		JButton btnNewButton_2 = new JButton("Browse");
		

		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Browse the folder to process");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField_3.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(null, "Enter Valid Folder Path");
				}
				
			}
		});
		btnNewButton_2.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton_2.setBounds(450, 80, 90, 25);
		tab2.add(btnNewButton_2);

		JButton btnNewButton_3 = new JButton("Start");
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				final String directoryPath=textField_3.getText();
				final String outputFileName=textField_2.getText();
				if(!textField_3.getText().isBlank() && !textField_2.getText().isBlank()) {
					try {
						Summary.processXMLFiles(directoryPath, outputFileName);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					JOptionPane.showMessageDialog(textField_3, "Process Done");
					}
				 else {
					JOptionPane.showMessageDialog(textField_3, "Select the File Path and Give Excel File Name");
				}
			}
		});
		btnNewButton_3.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton_3.setBounds(230, 130, 90, 30);
		tab2.add(btnNewButton_3);

		// Add the tabbed pane to the frame
		frame.getContentPane().add(tabbedPane);
	}
}
