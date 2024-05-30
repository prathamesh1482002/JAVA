package SpellChecker.SpellMistake;

import java.awt.EventQueue;
import javax.swing.*;

import java.awt.Font;
import java.awt.event.ActionListener;
import java.io.File;
import java.awt.event.ActionEvent;

public class ApplicationWindow {

	private JFrame frmXmlEntityUpdate;
	private JTextField textField_2;
	private JTextField textField_3;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ApplicationWindow window = new ApplicationWindow();
					window.frmXmlEntityUpdate.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public ApplicationWindow() {
		initialize();
	}

	private void initialize() {
		frmXmlEntityUpdate = new JFrame();
		frmXmlEntityUpdate.setResizable(false);
		frmXmlEntityUpdate.setTitle("SpellMistake_Tool");
		frmXmlEntityUpdate.setBounds(100, 100, 616, 202);
		frmXmlEntityUpdate.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmXmlEntityUpdate.getContentPane().setLayout(null);

		JTabbedPane tabbedPane_1 = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane_1.setBounds(0, 0, 601, 164);
		frmXmlEntityUpdate.getContentPane().add(tabbedPane_1);

		JPanel panel = new JPanel();
		tabbedPane_1.addTab("Select File Path", null, panel, null);
		panel.setLayout(null);

		JButton btnDone_1_1 = new JButton("START");
		btnDone_1_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				String filesPath = (textField_2.getText());
				String pdfFilesPath = (textField_3.getText());
				Testing.main(filesPath, pdfFilesPath);
				if (filesPath.toString().equals("") && pdfFilesPath.toString().equals("")) {
					JOptionPane.showMessageDialog(frmXmlEntityUpdate, "Please Select the folder Path");
				} else {
					JOptionPane.showMessageDialog(frmXmlEntityUpdate, "Process Done");
				}
			}
		});
		btnDone_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		btnDone_1_1.setBounds(259, 98, 96, 25);
		panel.add(btnDone_1_1);

		JLabel lblNewLabel_1_1_1 = new JLabel("Files Path");
		lblNewLabel_1_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_1_1_1.setBounds(10, 23, 109, 20);
		panel.add(lblNewLabel_1_1_1);

		textField_2 = new JTextField();
		textField_2.setColumns(10);
		textField_2.setBounds(147, 24, 317, 20);
		panel.add(textField_2);

		JButton Brwbtn_03 = new JButton("Browse");
		Brwbtn_03.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setMultiSelectionEnabled(true);
				chooser.setDialogTitle("Select File Path");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					File[] selectedFiles = chooser.getSelectedFiles();
					if (selectedFiles.length > 0) {
						StringBuilder filesText = new StringBuilder();
						for (File file : selectedFiles) {
							filesText.append(file.getAbsolutePath()).append(";");
						}
						textField_2.setText(chooser.getSelectedFile().toString());
					}
				} else {
					JOptionPane.showMessageDialog(textField_2, "Invalid File Path");
				}
			}
		});
		Brwbtn_03.setFont(new Font("Tahoma", Font.BOLD, 13));
		Brwbtn_03.setBounds(495, 22, 89, 23);
		panel.add(Brwbtn_03);

		textField_3 = new JTextField();
		textField_3.setColumns(10);
		textField_3.setBounds(147, 71, 317, 20);
		panel.add(textField_3);

		JButton Brwbtn_03_1 = new JButton("Browse");
		Brwbtn_03_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setMultiSelectionEnabled(true);
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select PDF File Path");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					File[] selectedFiles = chooser.getSelectedFiles();
					if (selectedFiles.length > 0) {
						StringBuilder filesText = new StringBuilder();
						for (File file : selectedFiles) {
							filesText.append(file.getAbsolutePath()).append(";");
						}
						textField_3.setText(chooser.getSelectedFile().toString());
					}
				} else {
					JOptionPane.showMessageDialog(textField_3, "Invalid PDF Path");
				}
			}
		});
		Brwbtn_03_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		Brwbtn_03_1.setBounds(495, 69, 89, 23);
		panel.add(Brwbtn_03_1);

		JLabel lblNewLabel_1_1_1_1 = new JLabel("PDF Files Path");
		lblNewLabel_1_1_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_1_1_1_1.setBounds(10, 70, 127, 20);
		panel.add(lblNewLabel_1_1_1_1);
	}
}