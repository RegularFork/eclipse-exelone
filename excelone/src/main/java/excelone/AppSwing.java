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

public class AppSwing {
	
	static JFrame frame = getFrame();
	static JPanel panel = new JPanel();

	public static void main(String[] args) {
			final JComponent component = new MyComponent();
			frame.add(panel);
			Font font = new Font("Arial", Font.BOLD, 20);
			//viewFonts();
			AbstractAction myAction = new MyAction();
			JButton button = new JButton(myAction);
			button.setText("Da Butn");
			panel.add(button);
			
			KeyStroke keyStroke = KeyStroke.getKeyStroke("ctrl B");
			InputMap inputMap = panel.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW);
			inputMap.put(keyStroke, "changeColor");
			ActionMap actionMap = panel.getActionMap();
			actionMap.put("changeColor", myAction);
			
			panel.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseClicked(MouseEvent e) {
					// TODO Auto-generated method stub
					super.mouseClicked(e);
					panel.setBackground(Color.YELLOW);
				}				
			});
			
			frame.add(component);
			frame.addMouseMotionListener(new MouseAdapter() {

				@Override
				public void mouseMoved(MouseEvent e) {
					super.mouseMoved(e);
					MyComponent.x = e.getX();
					MyComponent.y = e.getY();
					component.repaint();
				}
				
			});
			panel.revalidate(); // команда обновляет панель
	}
	
	static class MyAction extends AbstractAction{
		public void actionPerformed(ActionEvent e) {
			panel.setBackground(Color.GREEN);
		}
	}
	static class MyComponent extends JComponent {
		public static int x;
		public static int y;
		@Override
		protected void paintComponent(Graphics g) {
			super.paintComponent(g);
			((Graphics2D)g).drawString("Coordinates x:" + x + " y:" + y, 50, 50);
		}
		
	}

	
	
	static JFrame getFrame() {
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
