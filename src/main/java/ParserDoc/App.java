package ParserDoc;

import java.io.File;
import java.io.PrintStream;

import static ParserDoc.WordDoc.SaveParagraphsToTableWithFormatDocx;
import static ParserDoc.WordDoc.SaveTablesWithFormatDocx;
/*
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Picture;

//*/

public class App
{

    public static void main(String[] args) {
        String[] var6 = args;
        int var5 = args.length;

        for(int var4 = 0; var4 < var5; ++var4) {
            String name = var6[var4];

            try {
                PrintStream out = new PrintStream(name.concat(".txt"));
                PrintStream err = new PrintStream(name.concat(".err"));

                try {
                    processDir(new File(name), ""/*, out, err*/);
                } finally {
                //    out.close();
                //    err.close();
                }
            } catch (Exception var11) {
                ;
            }
        }

    }

    public static void processDir(File dir, String pred/*, PrintStream out, PrintStream err*/) {
        if(dir.canRead()) {

            if(dir.isDirectory()) {
                try {
                    File[] var7;
                    int var6 = (var7 = dir.listFiles()).length;

                    for(int var5 = 0; var5 < var6; ++var5) {
                        File f = var7[var5];
                        processDir(f, dir.getName()/*, out, err*/);
                    }
                } catch (Exception var9) {
                    //var9.printStackTrace(err);
                }
            } else if(dir.getName().toLowerCase().endsWith(".doc") || dir.getName().toLowerCase().endsWith(".docx")) {
                try {

                    WordDoc doc = new WordDoc (dir.getAbsolutePath (), pred);

                    // Images
                    ImageExtract imEx = new ImageExtract (doc.getXWPFDocument ());
                    imEx.savePictureAsDocx ();
                    // Text
                    SaveParagraphsToTableWithFormatDocx (doc.getXWPFDocument (),"Text.docx");
                    // Tables
                    SaveTablesWithFormatDocx(doc.getXWPFDocument (),"Tables.docx");
                    //doc.ParseTable ();

                   //doc.ExportTextDocx();
                   // doc.ParseTable ();

                  // doc.ParseTable();
                   // doc.ParseAll ();
                   // processFile(dir, pred/*, out, err*/);
                } catch (Exception var8) {
                   // var8.printStackTrace(err);
                    var8.printStackTrace ();
                }
            }

        }
    }



}
