package com.Rishi.DocxMaker;

import javax.swing.SwingUtilities;

public class App
{
	public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
			@Override
			public void run() {
				ScreenshotApp frame = new ScreenshotApp();
			    frame.setVisible(true);
			}
		});
    }
}
