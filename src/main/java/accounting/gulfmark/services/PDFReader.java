package accounting.gulfmark.services;

import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class PDFReader {
	public static void main(String[] args) {
		String pdfPath = "C://Users//Oscar//Desktop//Wire_Detail.pdf";
		String text = readPDF(pdfPath);
		System.out.println(text);
	}
	
	
	public static String readPDF(String filepath) {
		File file = new File(filepath);
		String text = null;
		try {
			PDDocument document = PDDocument.load(file);
			PDFTextStripper pdfStripper = new PDFTextStripper();
			text = pdfStripper.getText(document);
			
		}  catch (IOException e) {
			e.printStackTrace();
		}
		return text;
	}
}
