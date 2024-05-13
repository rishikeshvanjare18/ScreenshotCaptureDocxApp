package com.Rishi.DocxMaker;

import java.awt.AWTException;
import java.awt.FlowLayout;
import java.awt.GraphicsEnvironment;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jnativehook.GlobalScreen;
import org.jnativehook.NativeHookException;
import org.jnativehook.keyboard.NativeKeyEvent;
import org.jnativehook.keyboard.NativeKeyListener;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

public class ScreenshotApp extends JFrame implements NativeKeyListener {

	/**
	 * 
	 */
	private static final long serialVersionUID = -2015738342889793344L;
	private XWPFDocument doc;
	private boolean isRunning = false;
	private Robot robot;
	private JButton startButton;
	private JButton stopButton;
	private Rectangle screenRect;

	public ScreenshotApp() {

		setTitle("Screenshot Application");
		setSize(300, 200);
		setLayout(new FlowLayout());
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		startButton = new JButton("Start");
		stopButton = new JButton("Stop");

		startButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				startProcess();
			}
		});

		stopButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				stopProcess();
			}
		});

		add(startButton);
		add(stopButton);

		try {
			GlobalScreen.registerNativeHook();
		} catch (NativeHookException ex) {
			System.err.println("Failed to register native hook: " + ex.getMessage());
			System.exit(1);
		}

		// Disable logger to prevent console spam from jnativehook
		Logger logger = Logger.getLogger(GlobalScreen.class.getPackage().getName());
		logger.setLevel(Level.OFF);
		// Create the main application instance
//		ScreenshotApp app = new ScreenshotApp();
		stopButton.setVisible(false);
		// Get screen dimensions
		screenRect = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice()
				.getDefaultConfiguration().getBounds();
	}

	private void startProcess() {
		try {
			robot = new Robot();
			isRunning = true;
			// Initialize the DOCX document
			this.initDocument();

			// Add the hotkey listener
			GlobalScreen.addNativeKeyListener(this);
			startButton.setVisible(false);
			stopButton.setVisible(true);
//			while (isRunning) {
//				// Sleep for some time before taking the next screenshot (adjust as needed)
//				Thread.sleep(5000);
//			}
		} catch (AWTException ex) {
			ex.printStackTrace();
		}
	}

	private void stopProcess() {
		isRunning = false;
		// Save the document
		saveDocument("ScreenCapture-" + System.currentTimeMillis());
		robot = null;
		stopButton.setVisible(false);
		System. exit(0);
	}

	private void initDocument() {
		doc = new XWPFDocument();

		// Set left and right margins
		CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
		CTPageMar pageMar = sectPr.addNewPgMar();
		pageMar.setLeft(BigInteger.valueOf(1440)); // 1 inch in twentieths of a point
		pageMar.setRight(BigInteger.valueOf(1440)); // 1 inch in twentieths of a point

		XWPFParagraph paragraph = doc.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText("Screenshots captured:");
	}

	@Override
	public void nativeKeyPressed(NativeKeyEvent e) {
		if (isRunning) {
//			if ((e.getModifiers() & NativeKeyEvent.CTRL_L_MASK) != 0
//					&& (e.getModifiers() & NativeKeyEvent.ALT_L_MASK) != 0 && e.getKeyCode() == NativeKeyEvent.VC_S) {
			if (e.getKeyCode() == NativeKeyEvent.VC_S) {
				takeScreenshot(e);
			}
		}
	}

	private void takeScreenshot(NativeKeyEvent e) {
		// For example, if the hotkey Ctrl+Alt+S is pressed:
		System.out.println("Hotkey Ctrl+Alt+S pressed!");
		if (robot != null) {
			try {
				screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
				BufferedImage screenshot = robot.createScreenCapture(screenRect);

				// Resize screenshot to fit within document margins
				int maxWidth = 500; // Maximum width allowed in document (adjust as needed)
				int maxHeight = 400; // Maximum height allowed in document (adjust as needed)
				double widthScale = (double) maxWidth / screenshot.getWidth();
				double heightScale = (double) maxHeight / screenshot.getHeight();
				double scale = Math.min(widthScale, heightScale);
				double newWidth = screenshot.getWidth() * scale;
				double newHeight = screenshot.getHeight() * scale;

				// Save screenshot to file
				File screenshotFile = new File("screenshot_" + System.currentTimeMillis() + ".png");
				ImageIO.write(screenshot, "png", screenshotFile);

				// Add screenshot to document
				FileInputStream inputStream = new FileInputStream(screenshotFile);
				XWPFParagraph p = doc.createParagraph();
				XWPFRun r = p.createRun();
				r.setText("Screenshot captured at: " + new Date());
				r.addBreak();
				r.addPicture(inputStream, XWPFDocument.PICTURE_TYPE_PNG, screenshotFile.getName(),
						Units.toEMU(newWidth), Units.toEMU(newHeight));
				inputStream.close();

				System.out.println("Screenshot captured and added to the document");

				// Delete the screenshot file
				if (screenshotFile.exists()) {
					screenshotFile.delete();
					System.out.println("Screenshot file deleted: " + screenshotFile.getName());
				}

			} catch (IOException | InvalidFormatException ex) {
				ex.printStackTrace();
			}
		}
	}

	@Override
	public void nativeKeyReleased(NativeKeyEvent e) {
		// Handle key release events here
	}

	@Override
	public void nativeKeyTyped(NativeKeyEvent e) {
		// Handle key typed events here

	}

	// Method to save the document to a file
	public void saveDocument(String filename) {

		// Get the system username
		String username = System.getProperty("user.name");

		// Define the folder path
		String folderPath = "C:\\Users\\" + username + "\\Documents\\";
		try {
			// Create directory if it doesn't exist
			File directory = new File(folderPath + "ScreenCaptureDocx");
			if (!directory.exists()) {
				directory.mkdirs();
			}

			// Construct absolute file path
			String filePath = directory.getAbsolutePath() + File.separator + filename + ".docx";

			// Save the document
			FileOutputStream out = new FileOutputStream(filePath);
			doc.write(out);
			out.close();
			System.out.println("Document saved as: " + filePath);
			// Show popup message
			JOptionPane.showMessageDialog(null, "Document saved successfully at " + filePath, "Success",
					JOptionPane.INFORMATION_MESSAGE);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}