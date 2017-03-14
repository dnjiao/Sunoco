package accounting;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import accounting.gulfmark.models.VolEntry;
import accounting.gulfmark.services.PDFReader;
import accounting.gulfmark.services.VolumeParser;
import accounting.shell.models.Settlement;
import accounting.shell.services.SettlementXls;
import accounting.shell.services.SunocoXls;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;


public class MainController {
	@FXML
	private TextField shellFilePath;
	
	@FXML
	private TextField sunocoFilePath;
	
	@FXML
	private TextArea outputShell;
	
	@FXML 
	private TextField gmFilePath;
	
	@FXML
	private TextArea outputGM;
	
	public void browseShellFile(ActionEvent event) {
		final FileChooser fileChooser = new FileChooser();
	    File file = fileChooser.showOpenDialog(new Stage());
	    if (file != null) shellFilePath.setText(file.getAbsolutePath());
	              
	}
	
	public void browseSunocoFile(ActionEvent event) {
		final FileChooser fileChooser = new FileChooser();
		File file = fileChooser.showOpenDialog(new Stage());
	    if (file != null) sunocoFilePath.setText(file.getAbsolutePath());
	}
	
	public void browseGmFile(ActionEvent event) {
		final FileChooser fileChooser = new FileChooser();
		File file = fileChooser.showOpenDialog(new Stage());
	    if (file != null) {
	    	String path = file.getAbsolutePath();
	    	gmFilePath.setText(path);
	    }
	}
	
	public void executeShell(ActionEvent event) {
		outputShell.clear();
		
		String xlsPath1 = shellFilePath.getText();
		if (xlsPath1.isEmpty() ) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Shell Xls file is not specified.");
			alert.showAndWait();
			resetShellFields();
			return;
		}
		if (!new File(xlsPath1).exists()) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Invalid file path: " + xlsPath1);
			alert.showAndWait();
			resetShellFields();
			return;
		}
		if (!xlsPath1.endsWith(".xlsx")) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Not a valid Excel file: " + xlsPath1);
			alert.showAndWait();
			resetShellFields();
			return;
		}
		
		List<List<Settlement>> lists = new ArrayList<List<Settlement>>();		
		try {	
			lists = SettlementXls.readShellXls(xlsPath1);
			outputShell.appendText("Parsing completed for " + xlsPath1);
			outputShell.appendText("\nNow appending tabs to " + xlsPath1);
			SettlementXls.writeShellXls(lists, xlsPath1);
			outputShell.appendText("\nUpdate completed for " + xlsPath1);
		} catch (Exception e) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Exception");
			alert.setContentText(e.getMessage());
			alert.showAndWait();
			resetShellFields();
			return;
		}
		
		String xlsPath2 = sunocoFilePath.getText();
		if (!xlsPath2.isEmpty()) {
			if (!new File(xlsPath2).exists()) {
				Alert alert = new Alert(AlertType.ERROR);
				alert.setTitle("Error Dialog");
				alert.setContentText("File does not exist: " + xlsPath2);
				alert.showAndWait();
				resetShellFields();
				return;
			}
			if (!xlsPath2.endsWith(".xlsx")) {
				Alert alert = new Alert(AlertType.ERROR);
				alert.setTitle("Error Dialog");
				alert.setContentText("Not a valid Excel file: " + xlsPath2);
				alert.showAndWait();
				resetShellFields();
				return;
			}
			try {
				outputShell.appendText("\nNow updating spreadsheet " + xlsPath2);
				SunocoXls.processSettlementXls(xlsPath2, lists);
				outputShell.appendText("\nUpdate completed for " + xlsPath2);
			} catch (Exception e) {
				Alert alert = new Alert(AlertType.ERROR);
				alert.setTitle("Exception");
				alert.setContentText(e.getMessage());
				alert.showAndWait();
				resetShellFields();
				return;
			}
		}
	}
	
	public void executeGM(ActionEvent event) {
		outputGM.clear();
		
		String pdfPath = gmFilePath.getText();
		if (pdfPath.isEmpty() ) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Gulfmark PDF file is not specified.");
			alert.showAndWait();
			resetShellFields();
			return;
		}
		if (!new File(pdfPath).exists()) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Invalid file path: " + pdfPath);
			alert.showAndWait();
			resetShellFields();
			return;
		}
		if (!pdfPath.endsWith(".pdf")) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setContentText("Not a valid PDF file: " + pdfPath);
			alert.showAndWait();
			resetShellFields();
			return;
		}
		
		File file = new File(pdfPath);
		String location = pdfPath.substring(0, pdfPath.lastIndexOf(File.separator));
		String inFile = file.getName();
		String outFile = inFile.substring(0, inFile.lastIndexOf(".")) + ".xls";
		outputGM.appendText("\nNow processing PDF file " + inFile + "\n");
		String text = PDFReader.readPDF(pdfPath);
		List<VolEntry> entries = VolumeParser.getEntries(text);
		String xlsPath = location + File.separator + outFile;
		outputGM.appendText("\nCreating Excel file:");
		try {
			VolumeParser.createXls(entries, xlsPath);
		} catch (Exception e) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Exception");
			alert.setContentText(e.getMessage());
			alert.showAndWait();
			resetGMFields();
			return;
		}
		outputGM.appendText("\nExcel created: " + xlsPath);
		
	}
	
	public void resetGMFields() {
		gmFilePath.clear();
		outputGM.clear();
	}

	public void resetShellFields() {
		shellFilePath.clear();
		sunocoFilePath.clear();
		outputShell.clear();
	}
	
}
