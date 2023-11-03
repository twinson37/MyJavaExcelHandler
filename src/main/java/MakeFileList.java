import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MakeFileList {

    static HashMap<String, ArrayList<String>> path_map = new HashMap<>();
    static Properties properties = new Properties();


    public static void main(String[] args) throws IOException {

        new MakeFileList();
        String propFile = "secret.properties";
        URL props = ClassLoader.getSystemResource(propFile);
        properties.load(props.openStream());
        XSSFWorkbook new_workbook = new XSSFWorkbook();

        XSSFSheet sheet = new_workbook.createSheet();
        int rowNum = 0;

        for(Map.Entry<String, ArrayList<String>> entry : path_map.entrySet()){

            for(String name : entry.getValue()){
                Row row = sheet.createRow(rowNum++);
                Cell idCell = row.createCell(0);
                Cell valueCell = row.createCell(1);
                idCell.setCellValue(entry.getKey());
                valueCell.setCellValue(name);
            }
        }

        String home = System.getProperty ( "user.home" );
        String generated_codebook = home.concat("/스캔본.xlsx");
        File generated_codebook_file = new File(generated_codebook);
        try (FileOutputStream fout = new FileOutputStream(generated_codebook_file))
        {
            new_workbook.write(fout);

        } catch (IOException e){
            e.printStackTrace();
        }
    }

    public MakeFileList() {
        long beforeTime = System.currentTimeMillis();
        System.out.println("path 추가 중..");

        String osName = System.getProperty("os.name").toLowerCase();
        String scan_folder_name;
        if (osName.contains("win"))
        {
            System.out.println("OS : Windows");

            scan_folder_name =properties.getProperty("ID");
        }
        else if (osName.contains("mac"))
        {
            System.out.println("OS : Mac");

            scan_folder_name = "/Users/kimjungi/Downloads/새 폴더";

        }else return;

        File scan_folder = new File(scan_folder_name);

        scanDir(scan_folder);
        System.out.println("path done");

        long afterTime = System.currentTimeMillis(); // 코드 실행 후에 시간 받아오기
        long secDiffTime = (afterTime - beforeTime)/1000; //두 시간 차 계산
        System.out.println("시간차이(m) : "+secDiffTime+"\n");


    }

    private static void scanDir(File scan_folder) {

        File[] files = scan_folder.listFiles();

        for(File f : files) {

            if(f.isDirectory()&&!f.getName().equals("새폴더")){

                scanDir(f);

            }
            String extension = f.getName().substring(f.getName().lastIndexOf(".") + 1);
            if(f.isFile()&&extension.equals("pdf")){
                String pr = f.getParentFile().getName();
                if(!path_map.containsKey(pr)){
                    path_map.put(pr,new ArrayList<>());
                }
                path_map.get(pr).add(f.getName());
                System.out.println(pr);
                System.out.println(f.getName());
            }

        }
    }
}
