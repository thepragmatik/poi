package spike;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collection;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

import model.Order;
import model.TestCaseData;

/**
 * Using a parameterized test with an Excel spreadsheet.
 * 
 * @author johnsmart
 * @author rahul.thakur
 */
@RunWith(Parameterized.class)
public class OrderTests {

	private TestCaseData data;

	public OrderTests(TestCaseData o) {
		super();
		this.data = o;
	}

	@Parameters
	public static Collection<TestCaseData> spreadsheetData() throws IOException {
		return new SpreadsheetData(new FileInputStream("src/test/resources/sample-data.xls")).getData();
	}

	@Test
	public void should_calculate_order_amount() {
		// TODO:

	}

}
