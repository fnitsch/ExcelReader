public class Main {

    public static void main(String[] args) {

        ExcelReader excelReader = new ExcelReader();

        try {
            excelReader.openFile("Mappe1.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
