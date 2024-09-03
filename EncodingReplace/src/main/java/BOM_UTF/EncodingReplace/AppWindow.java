package BOM_UTF.EncodingReplace;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import javax.swing.JTextField;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class AppWindow {

	private JFrame frame;
	private JTextField textField;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AppWindow window = new AppWindow();
					window.frame.setVisible(true);
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
		frame = new JFrame();
		frame.setTitle("XML Encoding Converter");
		frame.setBounds(100, 100, 637, 177);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		JLabel lblNewLabel = new JLabel("XML File Path");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblNewLabel.setBounds(24, 49, 95, 14);
		frame.getContentPane().add(lblNewLabel);

		textField = new JTextField();
		textField.setBounds(139, 47, 328, 20);
		frame.getContentPane().add(textField);
		textField.setColumns(10);

		JButton btnNewButton = new JButton("Browse");
		btnNewButton.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Browse the XML Folder");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField.setText(chooser.getSelectedFile().toString());
				} else {
					JOptionPane.showMessageDialog(null, "Select the Valid Path");
				}
			}
		});
		btnNewButton.setBounds(506, 46, 89, 23);
		frame.getContentPane().add(btnNewButton);

		JButton btnNewButton_1 = new JButton("Start");
		btnNewButton_1.setFont(new Font("Tahoma", Font.BOLD, 12));
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String directoryPath = textField.getText();
				if (directoryPath.isEmpty()) {
					JOptionPane.showMessageDialog(frame, "Please Enter Valid Path.");
					return;
				} else {
					Encoding_Replace.main(directoryPath);
					JOptionPane.showMessageDialog(frame, "Process Done");
				}
			}
		});
		btnNewButton_1.setBounds(266, 97, 89, 23);
		frame.getContentPane().add(btnNewButton_1);
	}
}