import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcellUtils {
    private ArrayList<String> numbers = new ArrayList<>();

    public void readNumsFromExcell(String file) throws IOException {
        HSSFWorkbook excellBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet excellSheet = excellBook.getSheetAt(0);

        for (Row row : excellSheet) {
            for (Cell cell : row) {
                System.out.println("!");
                try {
                    if (cell.getStringCellValue().contains("+7") | cell.getStringCellValue().matches("8(.*)")) {
                        if (cell.getStringCellValue().length() == 12) {
                            String localNumber = cell.getStringCellValue().replace("+", "");
                            if (localNumber.matches("[0-9]+")) {
                                numbers.add(cell.getStringCellValue());
                                System.out.println(cell.getStringCellValue());
                            }
                        } else if (cell.getStringCellValue().length() == 11 & cell.getStringCellValue().matches("[0-9]+")) {
                            numbers.add(cell.getStringCellValue());
                            System.out.println(cell.getStringCellValue());
                        }
                    }
                } catch (Exception e) {
                }
            }
        }
    }

    public void writeNumsInExcell(String file) throws FileNotFoundException, IOException {
        HSSFWorkbook excellBook = new HSSFWorkbook();
        HSSFSheet excellSheet = excellBook.createSheet("Numbers");

        for (int i = 0; i < numbers.size(); i++) {
            Row row = excellSheet.createRow(i);
            Cell number = row.createCell(0);
            number.setCellValue(numbers.get(i));

            CellStyle style = excellBook.createCellStyle();
            style.setDataFormat((short) 0x31);
            number.setCellStyle(style);

            excellBook.write(new FileOutputStream(file));
            excellBook.close();
        }
    }
}
