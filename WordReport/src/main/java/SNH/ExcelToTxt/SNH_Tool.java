package SNH.ExcelToTxt;

import java.awt.EventQueue;
import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.io.File;
import java.awt.event.ActionEvent;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;

public class SNH_Tool {

	private JFrame frmXmlEntityUpdate;
	private JTextField textField;
	private JTextField textField_2;
	private JTextField textField_3;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					SNH_Tool window = new SNH_Tool();
					window.frmXmlEntityUpdate.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public SNH_Tool() {
		initialize();
	}

	private void initialize() {
		frmXmlEntityUpdate = new JFrame();
		frmXmlEntityUpdate.setResizable(false);
		frmXmlEntityUpdate.setTitle("SNH_Tool");
		frmXmlEntityUpdate.setBounds(100, 100, 633, 260);
		frmXmlEntityUpdate.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmXmlEntityUpdate.getContentPane().setLayout(null);

		JTabbedPane tabbedPane_1 = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane_1.setBounds(0, 0, 616, 220);
		frmXmlEntityUpdate.getContentPane().add(tabbedPane_1);

		JPanel panel = new JPanel();
		tabbedPane_1.addTab("Select Excel Files", null, panel, null);
		panel.setLayout(null);

		JLabel lblNewLabel_1_1 = new JLabel("Lable Request Excel Path");
		lblNewLabel_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_1_1.setBounds(10, 29, 164, 20);
		panel.add(lblNewLabel_1_1);

		textField = new JTextField();
		textField.setColumns(10);
		textField.setBounds(184, 30, 308, 20);
		panel.add(textField);

		JButton Brwbtn_02 = new JButton("Browse");
		Brwbtn_02.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select Lable Request Excel file");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(textField, "Invalid Path");
				}
			}
		});
		Brwbtn_02.setFont(new Font("Tahoma", Font.BOLD, 13));
		Brwbtn_02.setBounds(512, 28, 89, 23);
		panel.add(Brwbtn_02);

		JButton btnDone_1_1 = new JButton("START");
		btnDone_1_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// Update XML Files
				File file1 = new File(textField.getText());
				File file2 = new File(textField_2.getText());
				String path = (textField_3.getText());

				if (file1.exists() && file2.exists()) {
					if (file2.getName() != null) {
						ExcelModifier.mainmethod(file1, file2, path);
						JOptionPane.showMessageDialog(frmXmlEntityUpdate, "Process Done");
					}
				} else {
					JOptionPane.showMessageDialog(frmXmlEntityUpdate, "Please Select the Path");
				}
			}
		});
		btnDone_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		btnDone_1_1.setBounds(290, 142, 96, 25);
		panel.add(btnDone_1_1);

		JLabel lblNewLabel_1_1_1 = new JLabel("Set Up Excel Path");
		lblNewLabel_1_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_1_1_1.setBounds(10, 70, 145, 20);
		panel.add(lblNewLabel_1_1_1);

		textField_2 = new JTextField();
		textField_2.setColumns(10);
		textField_2.setBounds(184, 71, 308, 20);
		panel.add(textField_2);

		JButton Brwbtn_03 = new JButton("Browse");
		Brwbtn_03.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser chooser = new JFileChooser();

				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select Excel file");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select Set Up Excel file");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField_2.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(textField_2, "Invalid Directory Path");
				}

			}
		});
		Brwbtn_03.setFont(new Font("Tahoma", Font.BOLD, 13));
		Brwbtn_03.setBounds(512, 69, 89, 23);
		panel.add(Brwbtn_03);

		textField_3 = new JTextField();
		textField_3.setColumns(10);
		textField_3.setBounds(184, 111, 308, 20);
		panel.add(textField_3);

		JButton Brwbtn_03_1 = new JButton("Browse");
		Brwbtn_03_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setMultiSelectionEnabled(true);
				chooser.setDialogTitle("Select Input Files or Directory");
				chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

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
					JOptionPane.showMessageDialog(textField_3, "No Files Selected");
				}
			}
		});
		Brwbtn_03_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		Brwbtn_03_1.setBounds(512, 109, 89, 23);
		panel.add(Brwbtn_03_1);//

		JLabel lblNewLabel_1_1_1_1 = new JLabel("Input Path");
		lblNewLabel_1_1_1_1.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblNewLabel_1_1_1_1.setBounds(10, 114, 127, 20);
		panel.add(lblNewLabel_1_1_1_1);
	}
}
