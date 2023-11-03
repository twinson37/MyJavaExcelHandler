import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class codebook_Varname {

    public static final String data = "/Users/kimjungi/Desktop/rex/0-6세/mother_father_503samples_23.10.16.xlsx";
    public static final String codebook_file = "/Users/kimjungi/Desktop/rex/0-6세/standard_codebook_0y_6y_23.10.04(cbc최대최소수정컨펌).xlsx";
    public static final File codebook = new File(codebook_file);

    public static void main(String[] args) {

        int row_num = 0;

        try(FileInputStream file = new FileInputStream(codebook)){

            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet_0 = workbook.getSheetAt(0);
            XSSFSheet sheet_1 = workbook.getSheetAt(1);
            XSSFSheet sheet_2 = workbook.createSheet();

            Row row_1 = sheet_1.getRow(0);


            XSSFRow destRowNeeded;
            Iterator<Cell> cellIterator = row_1.cellIterator();
            while (cellIterator.hasNext()){
                Cell cell_1 = cellIterator.next();
                for (Row row_0: sheet_0){

                    Cell cell_0 = row_0.getCell(7);

                    if(cell_0.getStringCellValue().equals(cell_1.getStringCellValue())){
                        // sheet_2.createRow(row_num++);//this fails because destination row number is the same as source row number
                        /**
                         * 다른 시트라도 행번호가 똑같다면 복사가 되지않는듯하다
                         */
                        if(row_num==0) {
                            destRowNeeded = sheet_2.createRow(0);
                            XSSFRow destRow = sheet_2.createRow(row_num+1);
                            destRow.copyRowFrom(row_0, new CellCopyPolicy());
                            destRowNeeded.copyRowFrom(destRow, new CellCopyPolicy());
                            //the remove wrong first destination row
                            sheet_2.removeRow(destRow);
                        }else{
                            XSSFRow destRow = sheet_2.createRow(row_num);
                            destRow.copyRowFrom(row_0, new CellCopyPolicy());
                        }
//                        XSSFRow destRow = sheet_2.createRow(row_num);
//                        destRow.copyRowFrom(row_0, new CellCopyPolicy());
//
                        row_num++;

                        break;
                    }
                }
            }


            try (FileOutputStream fout = new FileOutputStream(codebook))
            {
                //저장
                workbook.write(fout);

            } catch (IOException e){
                e.printStackTrace();
            }

        } catch (IOException e){
            e.printStackTrace();
        }

    }
}
