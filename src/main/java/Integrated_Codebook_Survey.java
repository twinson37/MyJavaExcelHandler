import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Integrated_Codebook_Survey {

    public static final String data_0y__to_6y = "/Users/kimjungi/Desktop/rex/0-6세/cocoa0y_6y_수정 사항 반영_23.10.04_조사용.xlsx";
    public static final String data_7y = "/Users/kimjungi/Desktop/rex/7세/cocoa7y_차,세브란스_23.10.06.xlsx";
    public static final String data_8y = "/Users/kimjungi/Desktop/rex/8세/cocoa8y_error5_modified_23.10.10(병원 열 추가).xlsx";
    public static final String data_9y = "/Users/kimjungi/Desktop/rex/9세/cocoa9y_error5_modified_23.10.10(병원 열 추가).xlsx";
    public static final String integrated_codebook = "/Users/kimjungi/Desktop/rex/통합 코드북 조사.xlsx";
    public static final String generated_codebook = "/Users/kimjungi/Desktop/rex/통합 코드북 조사_생성.xlsx";
    static LinkedHashMap<String, ArrayList<String>> integrated_varname = new LinkedHashMap<>();
    static XSSFWorkbook new_workbook = new XSSFWorkbook();
    public static final File file_6y = new File(data_0y__to_6y);
    public static final File file_7y = new File(data_7y);
    public static final File file_8y = new File(data_8y);
    public static final File file_9y = new File(data_9y);
    public static final File codebook = new File(integrated_codebook);
    public static final File generated_codebook_file = new File(generated_codebook);


    public static void main(String[] args) throws Exception {

        make_sheet();

        save_workbook();

    }

    private static void save_workbook() {

        try (FileOutputStream fout = new FileOutputStream(generated_codebook_file))
        {
            new_workbook.write(fout);

        } catch (IOException e){
            e.printStackTrace();
        }
    }

    private static void make_sheet() throws Exception {

        list_var_names();

        ExcelSheetHandler  excelSheetHandler = ExcelSheetHandler.readExcel( file_6y );
        List<List<String>> excelDatas        = excelSheetHandler.getRows();


        write_var_names();
//        write_value();

        write_id();

    }

    private static void write_var_names(){

            XSSFSheet new_sheet;
            Row row;
            int size;
            for (Map.Entry<String, ArrayList<String>> entry : integrated_varname.entrySet()) {
                new_sheet = new_workbook.createSheet(entry.getKey());
                size = entry.getValue().size();
                row = new_sheet.createRow(0);
                for(int i = 0;i<size;i++){
                    Cell cell = row.createCell(i);
                    cell.setCellValue(entry.getValue().get(i));
                }
            }

    }

    private static void list_var_names() throws IOException {
        try (FileInputStream codebook_fi = new FileInputStream(codebook)){

            XSSFWorkbook workbook = new XSSFWorkbook(codebook_fi);
            XSSFSheet sheet = workbook.getSheetAt(2);

            String var_name;
            String int_var_name;
            for(Row Integrated_varname_row: sheet){

                Cell cell1 = Integrated_varname_row.getCell(1);
                Cell cell2 = Integrated_varname_row.getCell(3);
                int_var_name = cell1.getStringCellValue();
                var_name = cell2.getStringCellValue();

                if (integrated_varname.get(int_var_name) == null) {

                    integrated_varname.put(int_var_name,new ArrayList<>());
                    integrated_varname.get(int_var_name).add("id_c");
                }

                integrated_varname.get(int_var_name).add(var_name);
            }
            integrated_varname.remove("integrated_var_name");
        }
    }

    private static void write_id() throws Exception {

        int sheet_size = new_workbook.getNumberOfSheets();
        String[] id;

        ExcelSheetHandler excelSheetHandler = ExcelSheetHandler.readExcel(file_6y);
        List<List<String>> excelDatas = excelSheetHandler.getRows();

        int iCol = 0;    //컬럼 구분값
        int iRow = 0;    //행 구분값
        id = new String[excelDatas.size()];
        int index=0;

        for(List<String> dataRow : excelDatas){
            for(String str : dataRow){
                if(iCol == 11){
                    id[index++] = str;
                    break;
                }

                iCol++;
            }
            iCol = 0;
            iRow = 0;
        }

        XSSFSheet sheet;
        XSSFRow row;
        for (int i = 0;i<sheet_size;i++){
            sheet = new_workbook.getSheetAt(i);
            for(String s: id){
                if(s.equals("")) break;
                row = sheet.createRow(sheet.getLastRowNum()+1);
                Cell cell = row.createCell(0);
                cell.setCellValue(s);
            }
        }

    }

    private static void write_value() throws Exception {
        File[] data_list = {file_6y,file_7y,file_8y,file_9y};
//        XSSFSheet sheet;
//        XSSFRow row;
//        int sheet_index=0, row_index=1, size;

        for(File file: data_list){

            ExcelSheetHandler excelSheetHandler = ExcelSheetHandler.readExcel(file);
            List<List<String>> excelDatas = excelSheetHandler.getRows();

//            S10-579-C
            for (Sheet sheet: new_workbook){

                Row firstrow = sheet.getRow(0);
//                List<Integer> index = new ArrayList<>();
                int col_index = 0;
                int row_index = 0;
                Iterator<Cell> cellIterator = firstrow.cellIterator();

                while(cellIterator.hasNext()){

                    Cell cell = cellIterator.next();
                    List<String> data_sheet = excelDatas.get(0);

                    for(int i = 0;i<data_sheet.size();i++){
                        String tmp = cell.getStringCellValue();
                        if(data_sheet.get(i).equals(cell.getStringCellValue())){
//                            index.add(i);
                            cell.getColumnIndex();
                            for(List<String> data:excelDatas){
                                for(Row row: sheet){
                                    Cell cell1 = row.getCell(cell.getColumnIndex());
                                    if(cell1.getStringCellValue().equals("S10-579-C")) break;
                                    cell.setCellValue(data.get(i));
                                }
                            }

                            break;
                        }
                    }
                    col_index++;
                }

//                for(Row row:sheet){
//
//                    for(int i = 0;i<index.size();i++){
//
//                        for(List<String> data:excelDatas){
//                            Cell data_cell = row.getCell(i);
//                            data_cell.setCellValue(data.get(index.get(i)));
//                        }
//                    }
//
//                }


            }
//            for (Map.Entry<String, ArrayList<String>> entry : integrated_varname.entrySet()) {
//
//                size = entry.getValue().size();
//                for(String s: entry.getValue()){
//
//                    int iCol = 0;    //컬럼 구분값
//                    sheet = new_workbook.getSheetAt(sheet_index);
//                    for(List<String> dataRow : excelDatas) {
//                        for (String str : dataRow) {
//                            if (str.equals(s)) break;
//                            iCol++;
//                        }
//                        sheet.getRow(row_index++)
//
//                    }
//                }
//                sheet_index++;
//            }
        }
    }

}
