import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Integrated_Codebook_survey_2 {

    public static final String data_0y__to_6y = "/Users/kimjungi/Desktop/rex/0-6세/cocoa0y_6y_수정 사항 반영_23.10.04_조사용.xlsx";
    public static final String data_7y = "/Users/kimjungi/Desktop/rex/7세/cocoa7y_차,세브란스_23.10.06.xlsx";
    public static final String data_8y = "/Users/kimjungi/Desktop/rex/8세/cocoa8y_error5_modified_23.10.10(병원 열 추가).xlsx";
    public static final String data_9y = "/Users/kimjungi/Desktop/rex/9세/cocoa9y_error5_modified_23.10.10(병원 열 추가).xlsx";
    public static final String integrated_codebook = "/Users/kimjungi/Desktop/rex/통합 코드북 조사.xlsx";
    public static final String generated_codebook = "/Users/kimjungi/Desktop/rex/통합 코드북 조사_생성.xlsx";
    static ArrayList<ArrayList<String>> var_name_list= new ArrayList<>();
    static ArrayList<String> id_c_list = new ArrayList<>();
    static XSSFWorkbook new_workbook = new XSSFWorkbook();
    public static final File file_6y = new File(data_0y__to_6y);
    public static final File file_7y = new File(data_7y);
    public static final File file_8y = new File(data_8y);
    public static final File file_9y = new File(data_9y);
    public static final File codebook = new File(integrated_codebook);
    public static final File generated_codebook_file = new File(generated_codebook);

    public static void main(String[] args) throws Exception {

        make_sheet();

        write_values();

        save_workbook();
    }

    private static void write_values() throws FileNotFoundException {

        for(Sheet sheet : new_workbook){

            Row row = sheet.getRow(0);
            Cell cell = row.getCell(3);
            String year = cell.getStringCellValue().substring(0,2);

            switch (year){
                case "9세":
                    write_values_as_year(file_9y,sheet);

                    break;

                case "8세":
                    write_values_as_year(file_8y,sheet);

                    break;

                case "7세":
                        write_values_as_year(file_7y,sheet);
                    break;

                default:
                    write_values_as_6year(sheet);
                    break;

            }
        }

    }
    static ExcelSheetHandler excelSheetHandler;

    static {
        try {
            excelSheetHandler = ExcelSheetHandler.readExcel(file_6y);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    static List<List<String>> excelDatas = excelSheetHandler.getRows();
    private static void write_values_as_6year(Sheet sheet) {

        String var_name = sheet.getRow(0).getCell(1).getStringCellValue();

        int var_index=0;
        int row_index=1;
        Cell codebook_cell;
        boolean finding= true;

        for(List<String> row:excelDatas){
            if(row.size()==0) break;

            if (finding) {
                ListIterator<String> ListIterator = row.listIterator();
                while (ListIterator.hasNext()){

                    String data_var_name = ListIterator.next();

                    if(data_var_name.equals(var_name)){
                        finding = false;

                        System.out.println("var_id ="+data_var_name);
                        break;
                    }
                    var_index++;

                }
                continue;
            }

            System.out.println(var_name+":idc ="+row.get(11));
            if(row.size()>=var_index){
                String var_value = row.get(var_index);
                codebook_cell = sheet.getRow(row_index).createCell(1);
                codebook_cell.setCellValue(var_value);
            }


            row_index++;

        }
    }
    private static void write_values_as_year(File file, Sheet sheet) throws FileNotFoundException {

        try (FileInputStream data_fi = new FileInputStream(file)) {

            XSSFWorkbook data_workbook = new XSSFWorkbook(data_fi);
            XSSFSheet data_sheet = data_workbook.getSheetAt(0);

            String var_name =sheet.getRow(0).getCell(1).getStringCellValue();
            System.out.println(var_name);
            String data_var_name = null;
            int var_index=0;
            int row_index=1;
            Row codebook_row;
            Cell codebook_cell;
            boolean finding= true;

            for(Row data_row : data_sheet){

                if (finding) {
                    Iterator<Cell> cellIterator = data_row.cellIterator();
                    while (cellIterator.hasNext()){

                        Cell data_cell = cellIterator.next();
                        data_var_name = data_cell.getStringCellValue();

                        if(data_var_name.equals(var_name)){
                            finding = false;
                            System.out.println("var_id ="+data_var_name);
                            break;
                        }
                        var_index++;

                    }
                    continue;
                }
//                System.out.println(var_name+":idc ="+data_row.getCell(0).getStringCellValue());
                if(data_row.getCell(var_index)!=null){
                    DataFormatter formatter = new DataFormatter();

                    Cell data_cell =  data_row.getCell(var_index);
                    String var_value = formatter.formatCellValue(data_cell);
                    codebook_cell = sheet.getRow(row_index).createCell(1);
                    codebook_cell.setCellValue(var_value);
                }
                row_index++;
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static void list_var_names() throws IOException {
        try (FileInputStream codebook_fi = new FileInputStream(codebook)){

            XSSFWorkbook workbook = new XSSFWorkbook(codebook_fi);
            XSSFSheet sheet = workbook.getSheetAt(2);

            ArrayList<String> str = new ArrayList<>();

            for(Row row: sheet){
                str = new ArrayList<>();
                Cell cell1 = row.getCell(3);
                Cell cell2 = row.getCell(4);
                Cell cell3 = row.getCell(0);
                Cell cell4 = row.getCell(7);

                str.add(cell1.getStringCellValue());
                str.add(cell2.getStringCellValue());
                str.add(cell3.getStringCellValue());
                str.add(cell4.getStringCellValue());


                var_name_list.add(str);
            }
            var_name_list.remove(var_name_list.get(0));
        }
    }

    private static void make_sheet() throws Exception {

        list_var_names();
        list_id_c();

        XSSFSheet new_sheet;
        int i;
        for(ArrayList<String> s:var_name_list){
            i = 1;
            new_sheet = new_workbook.createSheet(s.get(0));
            Row row = new_sheet.createRow(0);

            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);


            cell1.setCellValue(s.get(0));
            cell2.setCellValue(s.get(1));
            cell3.setCellValue(s.get(2));
            cell4.setCellValue(s.get(3));

            for(String id:id_c_list){
                if(id.equals("id_c")){
                    Row id_c_row = new_sheet.getRow(0);
                    Cell id_c_cell = id_c_row.createCell(0);
                    id_c_cell.setCellValue(id);
                    continue;
                }
                Row id_row = new_sheet.createRow(i++);
                Cell id_cell = id_row.createCell(0);
                id_cell.setCellValue(id);
            }
        }

    }

    private static void list_id_c() throws Exception {

        ExcelSheetHandler  excelSheetHandler = ExcelSheetHandler.readExcel( file_6y );
        List<List<String>> excelDatas        = excelSheetHandler.getRows();
        String s;

        int iCol = 0;    //컬럼 구분값

        for(List<String> dataRow : excelDatas){
            for(String str : dataRow){
                if(str.equals("S10-579-C")) break;
                if(iCol == 11){
                    id_c_list.add(str);
                    break;
                }

                iCol++;
            }
            iCol = 0;
        }
    }

    private static void save_workbook() {

        try (FileOutputStream fout = new FileOutputStream(generated_codebook_file))
        {
            new_workbook.write(fout);

        } catch (IOException e){
            e.printStackTrace();
        }
    }


}
