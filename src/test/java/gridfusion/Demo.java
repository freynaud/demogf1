package gridfusion;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Created by freynaud on 15/09/2015.
 */
public class Demo {


  private Sheet getFirstSheet() throws IOException, InvalidFormatException {
    Workbook wb = WorkbookFactory.create(new File("src/test/resources/file.xlsx"));
    Sheet sheet = wb.getSheetAt(0);
    return sheet;
  }


  @DataProvider(name = "excelJava7")
  private Object[][] excelJava7() throws Exception {
    Object[][] res = new Object[3][3];
    for (Row row : getFirstSheet()) {
      for (Cell cell : row) {
        res[cell.getRowIndex()][cell.getColumnIndex()] = cell.getNumericCellValue();
      }
    }
    return res;
  }


  @DataProvider(name = "excelJava8")
  private Object[][] excelJava8() throws IOException, InvalidFormatException {
    return Stream(getFirstSheet())
        .map(cells ->
            Stream(cells)
                .map(cell -> cell.getNumericCellValue()).toArray(Object[]::new))
        .toArray(Object[][]::new);
  }

  @Test(dataProvider = "excelJava7")
  public void test7(double x, double y, double expected) throws IOException, InvalidFormatException {
    Assert.assertEquals(x * y, expected);

  }

  @Test(dataProvider = "excelJava8")
  public void test8(double x, double y, double expected) throws IOException, InvalidFormatException {
    Assert.assertEquals(x * y, expected);

  }


  private static <T> Stream<T> Stream(Iterable<T> iterable) {
    return StreamSupport.stream(iterable.spliterator(), false);
  }
}
