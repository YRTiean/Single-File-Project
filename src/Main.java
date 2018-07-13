import java.io.*;

import com.sun.org.apache.bcel.internal.generic.BREAKPOINT;
import org.apache.poi.xwpf.usermodel.*;

public class Main {


    private static File file;
    private static String folderDir;
    private static XWPFDocument document;
    private static XWPFParagraph tmpParagraph ;
    private static XWPFRun tmpRun ;
    private static int count = 0;

    public static void main(String []args) throws Exception {

        folderDir  = "DIRECTORY_LOCATION";

        File file = new File(folderDir);
        File [] listOfFiles = file.listFiles();
        document = new XWPFDocument();

        checkFilDir(file);
        document.write(new FileOutputStream(new File("OUTPUT_FILE_ABSOLUTE_NAME")));
        System.out.println(count);
    }

    private static void checkFilDir(File file){

            if (file.isFile()) {
                readFileContent(file);
                count++;
            } else {
                for (File file1 : file.listFiles()) {
                    checkFilDir(file1);
                }
            }

    }

    private static void readFileContent(File file) {

        InputStream is = null;
        BufferedInputStream bis = null;
        DataInputStream dis = null;

        try{

            is = new FileInputStream(file.getAbsolutePath());
            bis = new BufferedInputStream(is);
            dis = new DataInputStream(bis);

            String temp = null;
            tmpParagraph = document.createParagraph();
            //tmpParagraph.setPageBreak(true);

            tmpRun = tmpParagraph.createRun();

            tmpRun.setText(file.getName().toUpperCase());
            tmpRun.setFontSize(20);
            tmpRun.setBold(true);
            tmpRun.addBreak();
            tmpParagraph.setAlignment(ParagraphAlignment.CENTER);

            tmpParagraph = document.createParagraph();
            while ((temp = dis.readLine())!=null){
                tmpRun = tmpParagraph.createRun();
                tmpRun.setColor("000000");
                tmpRun.setFontSize(10);
                tmpRun.setText(temp);
                tmpRun.addBreak();

            }
            tmpParagraph.setBorderBottom(Borders.DASH_DOT_STROKED);
            is.close();
            bis.close();
            dis.close();


        }catch (Exception e){

        }
    }


}
