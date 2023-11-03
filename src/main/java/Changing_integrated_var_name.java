import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class Changing_integrated_var_name {

    public static final String path =  "/Users/kimjungi/Desktop/rex";
    public static final String  name = "통합 코드북 조사.xlsx";
    public static final File codebook = new File(path,name);

    public static void main(String[] args) throws IOException {
        StringBuilder sb = new StringBuilder();
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(System.out));

        try(FileInputStream file = new FileInputStream(codebook)){
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet changed_sheet =  workbook.cloneSheet(0);

            char first_letter;
            String str;
            String changed_Str;

            for(Row row: changed_sheet){

                Cell cell = row.getCell(2);
                str = cell.getStringCellValue();
                first_letter = str.charAt(0);
                if(first_letter=='C'){
                    changed_Str = str.replaceFirst("C", "c");
                    cell.setCellValue(changed_Str);
                    sb.append(cell.getAddress()).append(": ")
                            .append(str).append(" => ")
                            .append(changed_Str)
                            .append("\n");
                }
            }

            try (FileOutputStream fout = new FileOutputStream(codebook))
            {
                workbook.write(fout);

            } catch (IOException e){
                e.printStackTrace();
            }

            bw.write(String.valueOf(sb));
            bw.flush();
            bw.close();

        } catch (IOException e){
            e.printStackTrace();
        }

    }

}
