package ddtExample;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import static ddtExample.Constant.*;

@RunWith(JUnit4.class)
public class firstTest
{
    private final String filePath = PATH_TEST_DATA  + FILE_TEST_DATA;

    @Test
    public void readFile() throws Exception
    {
        FileInputStream file = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++)
        {
            for (int colNum = 0; colNum < sheet.getRow(rowNum).getPhysicalNumberOfCells(); colNum++)
            {
                XSSFCell cell = sheet.getRow(rowNum).getCell(colNum);
                String userCell = cell.getStringCellValue();
                System.out.print(userCell + "\t");
            }
            System.out.print("\n");
        }

        file.close();
    }

    @Test
    public void writeFile() throws Exception
    {
        FileInputStream file = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);

        XSSFRow row = sheet.getRow(1);
        XSSFCell cell = row.getCell(1);
        if (cell == null) {
            cell = row.createCell(1);
        }
        cell.setCellValue("bob123");

        FileOutputStream fileOut = new FileOutputStream(filePath);

        workbook.write(fileOut);
        fileOut.flush();
        fileOut.close();
        file.close();
    }
}
