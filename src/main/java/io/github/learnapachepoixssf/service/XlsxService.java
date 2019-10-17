package io.github.learnapachepoixssf.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.sql.DataSource;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.ParameterizedPreparedStatementSetter;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import io.github.learnapachepoixssf.model.Widget;
import io.github.learnapachepoixssf.repository.WidgetRepository;

@Service
public class XlsxService {

	static final int TOTAL_COLUMNS = 2;
	static final int ID_COLUMN = 0;
	static final int NAME_COLUMN = 1;

	@Autowired
	protected DataSource dataSource;

	@PersistenceContext
	protected EntityManager entityManager;

	@Autowired
	protected WidgetRepository widgetRepository;

	public void writeOutTestWidgets(OutputStream out, int rows) throws IOException {
		SXSSFWorkbook workbook = null;
		try {
			// keep 100 rows in memory, exceeding rows will be flushed to disk
			workbook = new SXSSFWorkbook(100);
			Sheet sheet = workbook.createSheet();
			for (int i = 0; i < rows; i++) {
				Row row = sheet.createRow(i);
				Cell column1 = row.createCell(0);
				column1.setCellValue("");
				Cell column2 = row.createCell(1);
				column2.setCellValue(UUID.randomUUID().toString());
			}
			workbook.write(out);
		} finally {
			// dispose of temporary files backing this workbook on disk
			if (workbook != null) {
				workbook.dispose();
				workbook.close();
			}
		}
	}

	@Transactional(readOnly = true)
	public int writeOutSavedWidgets(OutputStream out) throws IOException {
		AtomicInteger rowNumber = new AtomicInteger(0);
		SXSSFWorkbook workbook = null;
		try {
			// keep 100 rows in memory, exceeding rows will be flushed to disk
			workbook = new SXSSFWorkbook(100);
			Sheet sheet = workbook.createSheet();
			Stream<Widget> widgets = widgetRepository.findAll((widget, cq, cb) -> cb.conjunction());
			widgets.forEach(widget -> {
				Row row = sheet.createRow(rowNumber.get());
				Cell column1 = row.createCell(0);
				column1.setCellValue(widget.getId());
				Cell column2 = row.createCell(1);
				column2.setCellValue(widget.getName());
				rowNumber.incrementAndGet();
			});
			workbook.write(out);
		} finally {
			// dispose of temporary files backing this workbook on disk
			if (workbook != null) {
				workbook.dispose();
				workbook.close();
			}
		}
		return rowNumber.get();
	}

	@Transactional
	public void truncateWidgets() {
		widgetRepository.truncateWidgets();
	}

	@Transactional
	public int saveWidgets(Path savedFile, int batchSize, String parseType, String persistenceType) throws Exception {
		if ("dom".equalsIgnoreCase(parseType)) {
			return saveWidgetsUsingDom(savedFile, batchSize, persistenceType);
		} else {
			return saveWidgetsUsingSax(savedFile, batchSize, persistenceType);
		}
	}

	protected int saveWidgetsUsingSax(Path savedFile, final int batchSize, final String persistenceType)
			throws Exception {
		final AtomicInteger rowCount = new AtomicInteger(0);
		try (OPCPackage p = OPCPackage.open(savedFile.toString(), PackageAccess.READ)) {
			XlsxStreamParser xlsxStreamParser = new XlsxStreamParser(p, TOTAL_COLUMNS, new XlsxStreamParser.Callback() {

				Widget currentWidget = null;
				List<Widget> batchedWidgets = new LinkedList<>();

				@Override
				public void beginRow(int rowNumber) {
					currentWidget = new Widget();
				}

				@Override
				public void cellValue(int rowNumber, int columnNumber, String formattedValue, String cellReference,
						XSSFComment comment) {
					switch (columnNumber) {
					case ID_COLUMN:
						currentWidget.setId(
								(formattedValue != null && !formattedValue.isEmpty()) ? Long.valueOf(formattedValue)
										: null);
						break;
					case NAME_COLUMN:
						currentWidget.setName(formattedValue);
						break;
					default:
						break;
					}
				}

				@Override
				public void endRow(int rowNumber) {
					rowCount.incrementAndGet();
					batchedWidgets.add(currentWidget);
					if (batchedWidgets.size() == batchSize) {
						saveWidgets(batchedWidgets, batchSize, persistenceType);
						batchedWidgets.clear();
					}
				}

				@Override
				public void endSpreadsheet() {
					// Left over widgets
					saveWidgets(batchedWidgets, batchSize, persistenceType);
					batchedWidgets.clear();
				}

			});
			xlsxStreamParser.parseFirstSheetOnly();
		}
		return rowCount.get();
	}

	protected int saveWidgetsUsingDom(Path savedFile, int batchSize, String persistenceType)
			throws EncryptedDocumentException, IOException {
		InputStream inp = new FileInputStream(savedFile.toString());
		Workbook wb = WorkbookFactory.create(inp);
		Sheet sheet = wb.getSheetAt(0);
		int rowNumber = 0;
		Row row = sheet.getRow(rowNumber);
		List<Widget> batchedWidgets = new LinkedList<>();
		while (row != null) {
			Widget currentWidget = new Widget();
			Cell cell = row.getCell(ID_COLUMN);
			Long id = null;
			if (CellType.NUMERIC.equals(cell.getCellType())) {
				id = Long.valueOf(Double.valueOf(cell.getNumericCellValue()).longValue());
			} else {
				String value = cell.getStringCellValue();
				id = (value != null && !value.isEmpty()) ? Long.valueOf(value) : null;
			}
			currentWidget.setId(id);
			cell = row.getCell(NAME_COLUMN);
			String name = cell.getStringCellValue();
			currentWidget.setName(name);
			batchedWidgets.add(currentWidget);
			if (batchedWidgets.size() == batchSize) {
				saveWidgets(batchedWidgets, batchSize, persistenceType);
				batchedWidgets.clear();
			}
			row = sheet.getRow(++rowNumber);
		}
		// Left over widgets
		saveWidgets(batchedWidgets, batchSize, persistenceType);
		return rowNumber;
	}

	protected void saveWidgets(List<Widget> widgets, int batchSize, String persistenceType) {

		if ("jpa".equalsIgnoreCase(persistenceType)) {
			saveWidgetsUsingJpa(widgets, batchSize);
		} else {
			saveWidgetsUsingJdbc(widgets, batchSize);
		}
	}

	protected void saveWidgetsUsingJpa(List<Widget> widgets, int batchSize) {
		widgets.stream().forEach(widget -> {
			widgetRepository.save(widget);
		});
		entityManager.flush();
		entityManager.clear();
	}

	protected void saveWidgetsUsingJdbc(List<Widget> widgets, int batchSize) {
		JdbcTemplate jdbcTemplate = new JdbcTemplate(dataSource);
		jdbcTemplate.batchUpdate("insert into widget (id, name) values(SEQ_WIDGET.NEXTVAL, ?)",
				widgets.stream().filter(w -> w.isNew()).collect(Collectors.toList()), batchSize,
				new ParameterizedPreparedStatementSetter<Widget>() {
					@Override
					public void setValues(PreparedStatement ps, Widget widget) throws SQLException {
						ps.setString(1, widget.getName());
					}
				});
		jdbcTemplate.batchUpdate("update widget set name=? where id = ?",
				widgets.stream().filter(w -> !w.isNew()).collect(Collectors.toList()), batchSize,
				new ParameterizedPreparedStatementSetter<Widget>() {
					@Override
					public void setValues(PreparedStatement ps, Widget widget) throws SQLException {
						ps.setString(1, widget.getName());
						ps.setLong(2, widget.getId());
					}
				});
	}
}
