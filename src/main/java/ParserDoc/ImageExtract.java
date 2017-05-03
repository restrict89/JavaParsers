package ParserDoc;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Iterator;
import java.util.List;

//В файле multi.doc есть 2 изображения из папки "Мои рисунки"
//Сравниваем их наличие в файле и оригинального изображения на диске
//После чего извлекаем 1-ое изображение и сохраняем

public class ImageExtract {

    List<XWPFPictureData> allPictures;
   // XWPFPictureData docImage = null;
    byte[] origImage = null;

    public ImageExtract(XWPFDocument docA)  {

        try {
            //Создаём таблицу, где будут содержаться все изображения из документа
            allPictures = docA.getAllPictures ();
        }catch (Exception ex)
        {
            ex.printStackTrace ();
        }
    }

    //Кол-во изображений в документе
    public int getPicturesCount()
    {
        return allPictures.size ();
    }

    public List<XWPFPictureData> getListPictures()
    {
        return allPictures;
    }


    public void savePictures()
    {
           try {

               Iterator<XWPFPictureData> iterator = allPictures.iterator ();
               int i = 0;
               while (iterator.hasNext ()) {
                   XWPFPictureData pic = iterator.next ();
                   byte[] bytepic = pic.getData ();
                   BufferedImage imag = ImageIO.read (new ByteArrayInputStream (bytepic));
                   ImageIO.write (imag, "jpg", new File ("image" + (i+1) + ".jpg"));
                   i++;
               }
           }catch(Exception ex) {
                    System.out.println (ex.getMessage ());
                    ex.printStackTrace ();
                }
        }


    public void savePictureAsDocx() throws IOException {
        try {
            savePictures();

            File writeFile = new File("Pictures.docx");
            FileOutputStream fos = new FileOutputStream(writeFile.getAbsolutePath());

            XWPFDocument writeDocx = new XWPFDocument();
            XWPFParagraph par = writeDocx.createParagraph();
            XWPFRun run = par.createRun();
            run.setBold (true);
            par.setAlignment (ParagraphAlignment.CENTER);
            run.setFontSize(13);

            for(int i = 0;i< this.getPicturesCount ();i++) {
                run.setText ("image" + Integer.toString (i+1));
                run.addBreak();
                String imgFile = "image"+Integer.toString (i+1)+".jpg";
                FileInputStream is = new FileInputStream (imgFile);
                run.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200 pixels
                run.addBreak();
                is.close ();
                run.addBreak ();
            }
                writeDocx.write (fos);
                fos.close ();
                writeDocx.close();

            }
        catch(Exception ex) {
            System.out.println (ex.getMessage ());
            ex.printStackTrace ();
        }


          /*  Paragraph paragraphDoc = doc.creare ();
            XWPFRun runDoc = paragraphDoc.createRun ();

            //создание и вставка картинки
            String imgFile = "logo.png";
            FileInputStream is = new FileInputStream (imgFile);
            runDoc.addBreak ();
            runDoc.addPicture (is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU (200), Units.toEMU (200)); // 200x200 pixels
            is.close ();

            //создание колонтитула
            CTSectPr sectPr = doc.getDocument ().getBody ().addNewSectPr ();
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy (doc, sectPr);

            // создание верхнего колонтитула
            XWPFHeader header = headerFooterPolicy.createHeader (XWPFHeaderFooterPolicy.DEFAULT);
            paragraphDoc = header.getParagraphArray (0);
            paragraphDoc.setAlignment (ParagraphAlignment.LEFT);

            runDoc = paragraphDoc.createRun ();
            runDoc.setText ("Верхний колонтитул");

            // создание нижнего колонтитула
            XWPFFooter footer = headerFooterPolicy.createFooter (XWPFHeaderFooterPolicy.DEFAULT);
            paragraphDoc = footer.getParagraphArray (0);
            paragraphDoc.setAlignment (ParagraphAlignment.CENTER);

            runDoc = paragraphDoc.createRun ();
            runDoc.setText ("Нижний колонтитул");

            doc.write (new FileOutputStream ("test.docx"));
        } catch (Exception ex) {
            ex.printStackTrace ();
        }*/


/*
        //Выбираем i-ое изображение из List
        for(int i = 0; i< this.getPicturesCount (); i++)
        {
            try {
                docImage = (Picture) picCountA.get (i);
                OutputStream out = new FileOutputStream ("image" + Integer.toString (i) + ".jng");
                docImage.writeImageContent (out);
                out.flush ();
                out.close ();
            }  catch (FileNotFoundException var23)
            {
                var23.printStackTrace ();
            }
            //throw var23;
            catch (IOException var24) {
                var24.printStackTrace ();
                // throw var24;
            } finally {
           /* if(processed != 1) {
                //err.println(" !!warning ".concat(arg.getAbsolutePath()).concat(" processed not 1 Sheets"));

            }
        }
    }
*}*/}


    //Чтение изображения с диска
    private byte[] readFile(String file) throws Exception {

        ByteArrayOutputStream baos = new ByteArrayOutputStream ();
        FileInputStream fis = new FileInputStream (file);
        byte[] buffer = new byte[1024];

        int read = 0;
        while (read > -1) {
            read = fis.read (buffer);
            if (read > 0) {
                baos.write (buffer, 0, read);
            }
        }

        return baos.toByteArray ();
    }

    //Метод, сравнивающий изображения по байтно
    private boolean assertBytesSame(byte[] first, byte[] second) {
        boolean result = true;

        for (int i = 0; i < first.length; i++)
            if (first[i] != second[i])
                result = false;
        return result;
    }
}