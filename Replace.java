import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by Дмитрий on 22.04.2021.
 */
public class Replace {
    public static void replacer(String filename, String str_input, String str_output) {


        if (filename.contains(".docx")) {
            try {
                XWPFDocument docx = new XWPFDocument(OPCPackage.open(filename));
                for (XWPFParagraph p : docx.getParagraphs()) {
                    List<XWPFRun> runs = p.getRuns();
                    if (runs != null) {
                        for (XWPFRun r : runs) {
                            String text = r.getText(0);
                            if (text != null && text.contains(str_input)) {
                                text = text.replace(str_input, str_output);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
                String rez = filename.replace(".", "_output.");
                docx.write(new FileOutputStream(rez));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        } else {

            try {
                HWPFDocument doc = new HWPFDocument(new POIFSFileSystem(new FileInputStream(filename)));

                Range range = doc.getRange();
                range.replaceText(str_input, str_output);

                String rez = filename.replace(".", "_output.");

                OutputStream out = new FileOutputStream(rez);
                doc.write(out);

                out.flush();
                out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

}