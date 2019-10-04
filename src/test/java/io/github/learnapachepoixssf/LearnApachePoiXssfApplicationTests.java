package io.github.learnapachepoixssf;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import io.github.learnapachepoixssf.service.XlsxService;
import io.github.learnapachepoixssf.service.XlsxStreamParser;

@RunWith(SpringRunner.class)
@SpringBootTest
public class LearnApachePoiXssfApplicationTests {

	@Autowired
	protected XlsxService xlsxService;

	@Test
	public void convertSpreadsheetToCSV() throws Exception {
		String actualFilename = "target/actual-testdata.csv";
		final FileWriter out = new FileWriter(actualFilename);
		final CSVPrinter printer = new CSVPrinter(out, CSVFormat.DEFAULT);
		int NUMBER_OF_COLUMNS = 5;
		try (OPCPackage p = OPCPackage.open("src/test/resources/testdata.xlsx", PackageAccess.READ)) {
			XlsxStreamParser xlsxStreamParser = new XlsxStreamParser(p, NUMBER_OF_COLUMNS,
					new XlsxStreamParser.Callback() {

						List<String> values = new LinkedList<>();

						@Override
						public void cellValue(int rowNumber, int columnNumber, String formattedValue,
								String cellReference, XSSFComment comment) {
							values.add(formattedValue);
						}

						@Override
						public void endRow(int rowNumber) {
							try {
								printer.printRecord(values);
							} catch (IOException e) {
								throw new IllegalArgumentException("Cannot write record", e);
							}
							values.clear();
						}
					});
			xlsxStreamParser.parse();
			printer.close();
			boolean isSame = FileUtils.contentEquals(new File("src/test/resources/expected-testdata.csv"),
					new File(actualFilename));
			assertTrue("Result CSV file does not match expected results", isSame);
		}
	}

	@Test
	public void testSaxAndJdbc() throws Exception {
		assertUploadAndDownload(1000, "sax", "jdbc");
	}

	@Test
	public void testDomAndJdbc() throws Exception {
		assertUploadAndDownload(1000, "dom", "jdbc");
	}

	@Test
	public void testSaxAndJpa() throws Exception {
		assertUploadAndDownload(1000, "sax", "jpa");
	}

	@Test
	public void testDomAndJpa() throws Exception {
		assertUploadAndDownload(1000, "dom", "jpa");
	}

	private void assertUploadAndDownload(int batchSize, String parseType, String persistenceType) throws Exception {

		xlsxService.truncateWidgets();

		File file = new File("target/test-data.xlsx");
		file.delete();
		OutputStream out = new FileOutputStream(file);
		xlsxService.writeOutTestWidgets(out, batchSize);
		out.close();

		int rowCount = xlsxService.saveWidgets(file.toPath(), batchSize, parseType, persistenceType);
		assertEquals(batchSize, rowCount);

		file.delete();
		out = new FileOutputStream(file);
		rowCount = xlsxService.writeOutSavedWidgets(out);
		out.close();
		assertEquals(batchSize, rowCount);
	}
}
