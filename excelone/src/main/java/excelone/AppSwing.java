package excelone;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GraphicsEnvironment;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.AbstractAction;
import javax.swing.ActionMap;
import javax.swing.InputMap;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.KeyStroke;

import com.formdev.flatlaf.intellijthemes.materialthemeuilite.FlatAtomOneDarkIJTheme;
import com.formdev.flatlaf.intellijthemes.materialthemeuilite.FlatGitHubDarkIJTheme;

public class AppSwing {
	
	

	public static void main(String[] args) {
		JFrame frame = getFrame();
		JPanel panel = new JPanel();
		frame.add(panel);
		
		JButton button = new JButton("BUTTON");
		panel.add(button);
		
	}
	
	
	static JFrame getFrame() {
		FlatGitHubDarkIJTheme.setup();
		JFrame frame = new JFrame("Excelone 2");
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension dimension = toolkit.getScreenSize();
		frame.setBounds(dimension.width / 2 - 250, dimension.height / 2 - 150, 500, 300);
		
		return frame;
	}
	
	
	static void viewFonts() {
		String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
		for (String string : fonts) {
			System.out.println(string);
		}
	}
	

}
