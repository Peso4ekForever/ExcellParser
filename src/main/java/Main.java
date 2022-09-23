import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        try {
            ExcellUtils excellUtils = new ExcellUtils();
            excellUtils.readNumsFromExcell("./src/res/excell.xls");
            excellUtils.writeNumsInExcell("./src/res/output.xls");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
