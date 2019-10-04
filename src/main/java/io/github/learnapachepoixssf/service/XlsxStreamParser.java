package io.github.learnapachepoixssf.service;

import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;

/**
 * Based from https://github.com/pjfanning/poi-shared-strings-sample
 */
public class XlsxStreamParser {

	private final OPCPackage xlsxPackage;
	private final int minimumColumnsToProcess;
	private final Callback callback;

	public XlsxStreamParser(OPCPackage xlsxPackage, int minimumColumnsToProcess, Callback callback) {
		this.xlsxPackage = xlsxPackage;
		this.minimumColumnsToProcess = minimumColumnsToProcess;
		this.callback = callback;
	}

	public void parseFirstSheetOnly()
			throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
		parse(1);
	}

	public void parse() throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
		parse(Integer.MAX_VALUE);
	}

	protected void parse(int numberOfSheetsToProcess)
			throws IOException, OpenXML4JException, SAXException, ParserConfigurationException {
		try (TempFileSharedStringsTable stringsTable = new TempFileSharedStringsTable(xlsxPackage, true)) {
			XSSFReader xssfReader = new XSSFReader(xlsxPackage);
			StylesTable styles = xssfReader.getStylesTable();
			XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
			int index = 0;
			while (iterator.hasNext() && index < numberOfSheetsToProcess) {
				try (InputStream stream = iterator.next()) {
					String sheetName = iterator.getSheetName();
					callback.beginSheet(sheetName, index);
					parseSheet(styles, stringsTable, new CallbackContentsHandler(), stream);
					callback.endSheet(sheetName, index);
				}
				index++;
			}
		}
	}

	protected void parseSheet(Styles styles, SharedStrings sharedStrings, SheetContentsHandler sheetHandler,
			InputStream sheetInputStream) throws IOException, SAXException, ParserConfigurationException {
		DataFormatter formatter = new DataFormatter();
		InputSource sheetSource = new InputSource(sheetInputStream);
		XMLReader sheetParser = SAXHelper.newXMLReader();
		ContentHandler handler = new XSSFSheetXMLHandler(styles, null, sharedStrings, sheetHandler, formatter, false);
		sheetParser.setContentHandler(handler);
		sheetParser.parse(sheetSource);
	}

	public class CallbackContentsHandler implements SheetContentsHandler {
		private int currentRowNumber = -1;
		private int currentColumnNumber = -1;

		@Override
		public void startRow(int rowNumber) {
			// Handle missing rows
			if (callback.supportEmptyRows()) {
				handleMissingRows(currentRowNumber, rowNumber - currentRowNumber - 1);
			}
			// Then setup for new row
			currentRowNumber = rowNumber;
			currentColumnNumber = -1;
			callback.beginRow(currentRowNumber);
		}

		protected void handleMissingRows(int currentRowNumber, int numberMissing) {
			for (int i = 0; i < numberMissing; i++) {
				int missingRowNumber = currentRowNumber + i + 1;
				callback.beginRow(missingRowNumber);
				for (int columnNumber = 0; columnNumber < minimumColumnsToProcess; columnNumber++) {
					callback.cellValue(missingRowNumber, columnNumber, null,
							new CellAddress(missingRowNumber, columnNumber).formatAsString(), null);
				}
				callback.endRow(missingRowNumber);
			}
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			if (cellReference == null) {
				cellReference = new CellAddress(currentRowNumber, currentColumnNumber).formatAsString();
			}

			// Handle missing columns in the middle of the row
			int newColumnNumber = (new CellReference(cellReference)).getCol();
			int numberMissing = newColumnNumber - currentColumnNumber - 1;
			for (int i = 0; i < numberMissing; i++) {
				int missingColumnNumber = currentColumnNumber + i + 1;
				callback.cellValue(currentRowNumber, missingColumnNumber, null,
						new CellAddress(currentRowNumber, missingColumnNumber).formatAsString(), null);
			}
			currentColumnNumber = newColumnNumber;

			callback.cellValue(currentRowNumber, currentColumnNumber, formattedValue, cellReference, comment);
		}

		@Override
		public void endRow(int rowNum) {
			// Make sure we handle all columns if row is short
			int numberMissing = minimumColumnsToProcess - currentColumnNumber - 1;
			for (int i = 0; i < numberMissing; i++) {
				int missingColumnNumber = currentColumnNumber + i + 1;
				callback.cellValue(currentRowNumber, missingColumnNumber + i + 1, null,
						new CellAddress(currentRowNumber, missingColumnNumber).formatAsString(), null);
			}
			callback.endRow(currentRowNumber);
		}
	}

	public interface Callback {
		default void beginSheet(String sheetName, int index) {
		}

		default boolean supportEmptyRows() {
			return true;
		}

		default void beginRow(int rowNumber) {
		}

		default void cellValue(int rowNumber, int columnNumber, String formattedValue, String cellReference,
				XSSFComment comment) {
		}

		default void endRow(int rowNumber) {
		}

		default void endSheet(String sheetName, int index) {
		}

		default void endSpreadsheet() {
		}
	}
}