package tool.Excel_to_XML;

import javax.swing.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

public class AppWindow {

	private JFrame frmWordtoxmlconversion;
	private JTextField Docfolder;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AppWindow window = new AppWindow();
					window.frmWordtoxmlconversion.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();

				}
			}
		});
	}

	public AppWindow() {
		initialize();
	}

	private void initialize() {
		frmWordtoxmlconversion = new JFrame();
		frmWordtoxmlconversion.setTitle("Excel_To_XML_Conversion");
		frmWordtoxmlconversion.setBounds(100, 100, 584, 194);
		frmWordtoxmlconversion.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmWordtoxmlconversion.getContentPane().setLayout(null);

		JLabel lblNewLabel = new JLabel("Enter Excel Filepath:");
		lblNewLabel.setBounds(10, 48, 148, 14);
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 12));
		frmWordtoxmlconversion.getContentPane().add(lblNewLabel);

		Docfolder = new JTextField();
		Docfolder.setBounds(146, 46, 314, 20);
		frmWordtoxmlconversion.getContentPane().add(Docfolder);
		Docfolder.setColumns(10);

		JButton btnBrowseInput = new JButton("Browse");
		btnBrowseInput.setBounds(470, 44, 89, 23);
		btnBrowseInput.setFont(new Font("Tahoma", Font.BOLD, 12));
		frmWordtoxmlconversion.getContentPane().add(btnBrowseInput);
		btnBrowseInput.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Browse the folder to process");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					Docfolder.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(null, "Enter Folder Path");
				}
			}
		});

		JButton btnStart = new JButton("Start");
		btnStart.setBounds(247, 93, 89, 23);
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String excelFilePath = Docfolder.getText();
				try {
					test.main(excelFilePath);
				} catch (InvalidFormatException e1) {
					
					e1.printStackTrace();
				} catch (IOException e1) {
					
					e1.printStackTrace();
				}

				if (Docfolder == null) {
					JOptionPane.showMessageDialog(frmWordtoxmlconversion, "please enter valid path");
				}
				JOptionPane.showMessageDialog(frmWordtoxmlconversion, "Process Done");

			}
		});
		btnStart.setFont(new Font("Tahoma", Font.BOLD, 12));
		frmWordtoxmlconversion.getContentPane().add(btnStart);
	}
}
