package methods;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.imageio.IIOException;
import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class reusableMethods {

	public static XWPFDocument doc;
	public static XWPFParagraph p;
	public static XWPFRun ro;
	
	public reusableMethods()
	{
			doc = new XWPFDocument();
			p = doc.createParagraph();
			ro = p.createRun();
	}
		   public static void createTestResultsinDocWithScreenshots(String testCaseName, String[] imgFileNames) throws IOException, InvalidFormatException, XmlException 
		   {
		 
			System.out.println("testcasename is : "  +testCaseName );
			System.out.println("testcasename is : "  +imgFileNames );
			System.out.println("Recieved File names : " + imgFileNames[0]);
			System.out.println("Recieved File names : " + imgFileNames[1]);

			System.out.println("inside creatdoc_2");
			XWPFDocument doc1=new XWPFDocument();
			XWPFParagraph p1=doc1.createParagraph();
			XWPFRun ro1=p1.createRun();
			
			  CTSectPr sectPr = doc1.getDocument().getBody().addNewSectPr();
			  XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc1, sectPr);
			 

			// write header content
			CTP ctpHeader = CTP.Factory.newInstance();
			CTR ctrHeader = ctpHeader.addNewR();
			CTText ctHeader = ctrHeader.addNewT();
			/*
			 * String headerText =
			 * "GDA Team - Power BI Reports Automation Test Result Screenshots";
			 * ctHeader.setStringValue(headerText);
			 */
			XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, doc1);
			XWPFRun headerRun = headerParagraph.createRun();
			headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
			headerRun.setFontSize(9);
			headerRun.setColor("808000");
			headerRun.setText("GDA Team - SSRS Reports Weekend Test Automation ");
			headerRun.addBreak();
			headerParagraph.setAlignment(ParagraphAlignment.LEFT);
			headerRun.setFontSize(9);
			headerRun.setColor("808000");
			String curr_date = getCurrentDate("yyyy-MM-dd-hh:mm:ss");
			headerRun.setText(curr_date);
			XWPFParagraph[] parsHeader = new XWPFParagraph[1];
			parsHeader[0] = headerParagraph;
			policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

			// write footer content
			CTP ctpFooter = CTP.Factory.newInstance();
			CTR ctrFooter = ctpFooter.addNewR();
			CTText ctFooter = ctrFooter.addNewT();
			String footerText = "© 2021. Confidential Do not Share this documents.";
			ctFooter.setStringValue(footerText);
			XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, doc1);
			headerParagraph.setAlignment(ParagraphAlignment.LEFT);
			XWPFParagraph[] parsFooter = new XWPFParagraph[1];
			parsFooter[0] = footerParagraph;
			policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);

			// write body content

			p1.setAlignment(ParagraphAlignment.CENTER);
			ro1.setBold(true);
			ro1.setFontFamily("Verdana");
			ro1.setText(testCaseName);
			ro1.addBreak();
			// Create a document object
			System.out.println("Inside CreateDoc Method");
			System.out.println("Recieved File names : " + imgFileNames[0]);
			System.out.println("Recieved File names : " + imgFileNames[1]);
			System.out.println("Recieved File names : " + imgFileNames);
			BufferedImage bimg1;

			for (String file : imgFileNames) {
				// System.out.println("File is : " + file); /* Uncomment this to get file names
				// of images
				// Path of the Screenshot ( Comes from the Array of file name with .jpg )
				try {
					File dest = new File(System.getProperty("user.dir") + "\\TestResults\\" + file + ".jpg");
					bimg1 = ImageIO.read(dest);
					// Set width and height of the Image before copying to Word Document
					int width = 500;
					int height = 280;

					String imgFile = dest.getName();
					int imgFormat = getImageFormat(imgFile);

					/*
					 * String p1 = "Screenshot"; ro.setText(p1);
					 */

					ro1.addBreak();
					ro1.addBreak();
					ro1.setText(file);
					ro1.addPicture(new FileInputStream(dest), imgFormat, imgFile, Units.toEMU(width), Units.toEMU(height));
				} 
				catch (IIOException e) 
				{
					continue;
				}

				
			}
			FileOutputStream out = new FileOutputStream(
					System.getProperty("user.dir") + "\\TestResults\\" + testCaseName + ".doc"); 
			doc1.write(out);
					out.close();

			System.out.println("Word document With Screenshots created successfully!");
		}
		  
		   
		   public static void captureScreenshot(String screenshotName, WebDriver driver) 
			{
				     // Cast driver object to TakesScreenshot
					System.out.println("Inside Capture Screenshot Method");
					TakesScreenshot screenshot = (TakesScreenshot) driver;
					
					// Get the screenshot as an image File
					File src = screenshot.getScreenshotAs(OutputType.FILE);
					try 
					{
						File dest = new File(System.getProperty("user.dir") + "\\TestResults\\" + screenshotName + ".jpg");
						FileUtils.copyFile(src, dest);
					} 
					catch (IOException ex) 
					{
						System.out.println(ex.getMessage());
					}
					System.out.println("Capture Screenshot Done");
			}
		   public static int getImageFormat(String imgFileName) 
			{
				// TODO Auto-generated method stub
				int format;
				 if (imgFileName.endsWith(".emf"))
				  format = XWPFDocument.PICTURE_TYPE_EMF;
				 else if (imgFileName.endsWith(".wmf"))
				  format = XWPFDocument.PICTURE_TYPE_WMF;
				 else if (imgFileName.endsWith(".pict"))
				  format = XWPFDocument.PICTURE_TYPE_PICT;
				 else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
				  format = XWPFDocument.PICTURE_TYPE_JPEG;
				 else if (imgFileName.endsWith(".png"))
				  format = XWPFDocument.PICTURE_TYPE_PNG;
				 else if (imgFileName.endsWith(".dib"))
				  format = XWPFDocument.PICTURE_TYPE_DIB;
				 else if (imgFileName.endsWith(".gif"))
				  format = XWPFDocument.PICTURE_TYPE_GIF;
				 else if (imgFileName.endsWith(".tiff"))
				  format = XWPFDocument.PICTURE_TYPE_TIFF;
				 else if (imgFileName.endsWith(".eps"))
				  format = XWPFDocument.PICTURE_TYPE_EPS;
				 else if (imgFileName.endsWith(".bmp"))
				  format = XWPFDocument.PICTURE_TYPE_BMP;
				 else if (imgFileName.endsWith(".wpg"))
				  format = XWPFDocument.PICTURE_TYPE_WPG;
				 else {
				  return 0;
				 }
				 return format;
			}

			public static String getCurrentDate(String format) 
			{
				DateFormat dateformat=new SimpleDateFormat(format);
				Date date=new Date();
				return dateformat.format(date);
			}
			
	}

