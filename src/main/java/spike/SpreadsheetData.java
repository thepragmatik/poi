package spike;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import spike.model.Order;
import spike.model.OrderDetails;
import spike.model.TestCaseData;

/**
 * Read data in an Excel spreadsheet and return it as a collection of objects.
 * This is designed to facilitate for parameterized tests in JUnit that get data
 * from an excel spreadsheet.
 * 
 * @author johnsmart
 * @author rahul.thakur
 */
public class SpreadsheetData {

	private transient Collection<TestCaseData> orders = new ArrayList<TestCaseData>();

	public SpreadsheetData(final InputStream excelInputStream) throws IOException {
		this.orders = loadOrdersFromSpreadSheet(excelInputStream);
	}

	public Collection<TestCaseData> getData() {
		return orders;
	}

	private Collection<TestCaseData> loadOrdersFromSpreadSheet(final InputStream excelFile)
			throws IOException, RuntimeException {
		HSSFWorkbook workbook = new HSSFWorkbook(excelFile);

		Sheet orderSheet = workbook.getSheet("Orders");
		Sheet orderDetailSheet = workbook.getSheet("OrderDetails");
		Sheet assertionSheet = workbook.getSheet("Assertions");

		List<TestCaseData> rows = new ArrayList<TestCaseData>();
		short headerIdx = orderSheet.getTopRow();
		Row header = orderSheet.getRow(headerIdx);

		for (Row row : orderSheet) {
			if (!isEmpty(row)) {
				if (!isHeaderRow(row)) {
					TestCaseData testCaseData = new TestCaseData();
					Iterator<Cell> headerIterator = header.cellIterator();

					while (headerIterator.hasNext()) {
						Cell headerCell = (Cell) headerIterator.next();
						String fieldName = headerCell.getStringCellValue();

						if (fieldName.startsWith("testCase")) {
							// testCaseData.setTestCaseId(row.getCell(0).getStringCellValue());
							testCaseData.setTestCaseId(Double.toString(row.getCell(0).getNumericCellValue())); // FIXME:
							continue;
						}

						try {
							Order order = loadOrderFromRow(row, header.cellIterator());
							testCaseData.setOrder(order);

							OrderDetails orderDetails = loadOrderDetailsFromSheetForTest(orderDetailSheet,
									testCaseData.getTestCaseId());

							testCaseData.setOrderDetails(orderDetails);

							rows.add(testCaseData);

						} catch (InstantiationException e) {
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						}
					}
				}
			}
		}

		return rows;
	}

	private OrderDetails loadOrderDetailsFromSheetForTest(Sheet sheet, String testCaseId) {
		OrderDetails orderDetails = null;
		short headerIdx = sheet.getTopRow();
		Row header = sheet.getRow(headerIdx);

		for (Row row : sheet) {
			if (!isEmpty(row)) {
				if (!isHeaderRow(row)) {
					// locate the correct test case Id and populate OrderDetail
					Double value = row.getCell(0).getNumericCellValue(); // FIXME:
																			// cleanup!

					System.out.println(String.format("Comparing %s with %s", value, testCaseId));

					if (Double.toString(value).equals(testCaseId)) {
						try {
							orderDetails = loadObject(OrderDetails.class, row, header.cellIterator());
						} catch (InstantiationException e) {
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						}
					}
				}
			}
		}
		return orderDetails;
	}

	private Order loadOrderFromRow(Row row, Iterator cellItr) throws InstantiationException, IllegalAccessException {
		Order order = loadObject(Order.class, row, cellItr);
		return order;
	}

	private <T> T loadObject(Class<T> klass, Row row, Iterator cellItr)
			throws InstantiationException, IllegalAccessException {
		System.out.println(String.format("Populating %s instance", klass.getSimpleName()));
		T obj = klass.newInstance();
		while (cellItr.hasNext()) {
			Cell cell = (Cell) cellItr.next();
			String fieldName = cell.getStringCellValue();
			if (fieldName.startsWith("testCase")) {
				// IGNORE! Already consumed.
				continue;
			}
			try {
				Field field = klass.getDeclaredField(fieldName);
				field.setAccessible(true);
				Object value = objectFrom((HSSFWorkbook) row.getSheet().getWorkbook(),
						row.getCell(cell.getColumnIndex()));
				System.out.println(String.format("setting field '%s', to value '%s'", fieldName, value.toString()));
				field.set(obj, value);
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return obj;
	}

	private boolean isEmpty(final Row row) {
		Cell firstCell = row.getCell(0);
		boolean rowIsEmpty = (firstCell == null) || (firstCell.getCellType() == Cell.CELL_TYPE_BLANK);
		return rowIsEmpty;
	}

	private boolean isHeaderRow(final Row row) {
		return (row.getSheet().getTopRow() == row.getRowNum());
	}

	/**
	 * Count the number of columns, using the number of non-empty cells in the
	 * first row.
	 */
	private int countNonEmptyColumns(final Sheet sheet) {
		Row firstRow = sheet.getRow(0);
		return firstEmptyCellPosition(firstRow);
	}

	private int firstEmptyCellPosition(final Row cells) {
		int columnCount = 0;
		for (Cell cell : cells) {
			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				break;
			}
			columnCount++;
		}
		return columnCount;
	}

	private Object objectFrom(final HSSFWorkbook workbook, final Cell cell) {
		Object cellValue = null;

		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			cellValue = cell.getRichStringCellValue().getString();
		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			cellValue = getNumericCellValue(cell);
		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			cellValue = cell.getBooleanCellValue();
		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			cellValue = evaluateCellFormula(workbook, cell);
		}

		return cellValue;

	}

	private Object getNumericCellValue(final Cell cell) {
		Object cellValue;
		if (DateUtil.isCellDateFormatted(cell)) {
			cellValue = new Date(cell.getDateCellValue().getTime());
		} else {
			cellValue = cell.getNumericCellValue();
		}
		return cellValue;
	}

	private Object evaluateCellFormula(final HSSFWorkbook workbook, final Cell cell) {
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		CellValue cellValue = evaluator.evaluate(cell);
		Object result = null;

		if (cellValue.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			result = cellValue.getBooleanValue();
		} else if (cellValue.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			result = cellValue.getNumberValue();
		} else if (cellValue.getCellType() == Cell.CELL_TYPE_STRING) {
			result = cellValue.getStringValue();
		}

		return result;
	}
}
