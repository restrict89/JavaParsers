package ParserDoc;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

import static org.apache.poi.xssf.usermodel.XSSFFont.DEFAULT_FONT_SIZE;


/**
 * Created by yamashkina on 21.04.2017.
 */
@SuppressWarnings("ALL")
public class WordDoc {

    private HWPFDocument hwpfDocument;
    private XWPFDocument xwpfDocument;
    private String fileName;


    public HWPFDocument getHWPFDocument()
    {
        return hwpfDocument;
    }
    public XWPFDocument getXWPFDocument()
    {
        return xwpfDocument;
    }
    public String getFileName(){return fileName;}



    WordDoc(String file, String pred) throws IOException, InvalidFormatException {
        try {
            FileInputStream fileInputStream = new FileInputStream (file);
            if (file.endsWith (".doc")) {
                hwpfDocument = new HWPFDocument (fileInputStream);
                fileName = file;
            }
            else if (file.endsWith(".docx")) {
                xwpfDocument = new XWPFDocument (OPCPackage.open (fileInputStream));
                fileName = file;
            }
        } catch (FileNotFoundException var23) {
            var23.printStackTrace();
            //throw var23;
        } catch (IOException var24) {
            var24.printStackTrace();
            //throw var24;
        } finally {
           // hwpfDocument.close();
           // xwpfDocument.close();
        }

        }


    public static void SaveParagraphsWithFormatDocx(XWPFDocument doc, String newDocName){

        List<XWPFParagraph> paras = doc.getParagraphs();
        XWPFDocument newdoc = new XWPFDocument();

        //bullet and numering lists
        XWPFNumbering docNumbering = doc.getNumbering ();
        //XWPFNumbering newnumbering = newdoc.createNumbering ();


         newdoc.createStyles ();
        // go to paragraph
        BigInteger level = BigInteger.valueOf (0);
        for (XWPFParagraph para : paras) {

            //noinspection Since15
            XWPFStyles style = newdoc.getStyles ();

             // set DEFAULT style
            if (para.getStyleID () != null && !style.styleExist ( para.getStyleID ())) {
                style.addStyle (doc.getStyles ().getStyle ( para.getStyleID ()));//getStyle ()));
            }


            if (!para.getParagraphText ().isEmpty ()) {

                XWPFParagraph newpara = newdoc.createParagraph ();
                newpara.setStyle (para.getStyleID ());


               if(para.getNumID () != null) //(&& (para.getNumFmt ().toString () == "bullet")){
               {

                    STNumberFormat.Enum stNumberFormat = STNumberFormat.Enum.forString (para.getNumFmt ().toString ());
                    System.out.println(para.getNumIlvl () + " " + para.getNumFmt ().toString ()
                            + " " +STNumberFormat.Enum.forString (para.getNumFmt ().toString ()).toString ()
                            + " " + para.getNumID ().intValue () + "" + para.getNumLevelText ()+ " " );//*/
                   //Old  CTAbstractNum
                   XWPFNum oldXWPFNum = docNumbering.getNum (para.getNumID ());
                   BigInteger oldAbstractNumId = oldXWPFNum.getCTNum ().getAbstractNumId ().getVal ();
                   //ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                   XWPFAbstractNum oldXWPFAbstractNum = docNumbering.getAbstractNum (oldAbstractNumId);
                   CTAbstractNum oldCTAbstractNum = oldXWPFAbstractNum.getCTAbstractNum ();


                   CTAbstractNum newCTAbstractNum = CTAbstractNum.Factory.newInstance ();
                   newCTAbstractNum.setAbstractNumId (para.getNumIlvl ());//-- рабочий вариант с lvl
                   CTLvl ctLvl = newCTAbstractNum.addNewLvl ();



                   if(stNumberFormat.toString () == "bullet") {
                       ctLvl.addNewNumFmt ().setVal (STNumberFormat.BULLET);
                       ctLvl.addNewLvlText ().setVal (para.getNumLevelText ());//*/"");//
                      // ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                  }
                  else
                  {
                      ctLvl.addNewNumFmt().setVal (STNumberFormat.Enum.forString (para.getNumFmt ()));//DECIMAL_);
                      String newlvltext = para.getNumLevelText ().toString ();
                      ctLvl.addNewLvlText ().setVal ("%1.");
                      ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));//para.getNumIlvl());
                 }


                   XWPFAbstractNum abstractNum = new XWPFAbstractNum (newCTAbstractNum);//,newnumbering);
                   XWPFNumbering newnumbering = newdoc.createNumbering ();
                   BigInteger abstractNumId  = newnumbering.addAbstractNum (abstractNum);
                   BigInteger numID = newnumbering.addNum (abstractNumId);

                   System.out.println(numID.toString ()+ " " + abstractNumId.toString ());

                   newpara.setNumID (numID);

                    // работает на обычных списках
                 /*  System.out.println(para.getNumIlvl ().add (BigInteger.valueOf (1)).toString ());
                   newpara.setNumID (para.getNumIlvl ().add (BigInteger.valueOf (1)));
                   newpara.getCTP().getPPr().getNumPr().addNewIlvl().setVal(para.getNumIlvl ().add (BigInteger.valueOf (1)));
                 */

                }

                for (XWPFRun r : para.getRuns ()) {

                    XWPFRun tmpRun = newpara.createRun ();
                    tmpRun.setTextPosition (r.getTextPosition ());
                    // важно использовать именно метод toString() поскольку
                    // этот метод сохраняет возможные символы "\n", которые getText обрезает
                    tmpRun.setText (r.toString ());
                    tmpRun.setBold (r.isBold ());
                    tmpRun.setFontFamily (r.getFontFamily ());
                    tmpRun.setColor (r.getColor ());
                    tmpRun.setEmbossed (r.isEmbossed ());
                    tmpRun.setCapitalized (r.isCapitalized ());
                    tmpRun.setShadow (r.isShadowed ());
                    tmpRun.setItalic (r.isItalic ());
                    tmpRun.setStrike (r.isStrike ());
                    tmpRun.setDoubleStrikethrough (r.isDoubleStrikeThrough ());
                    if (r.isHighlighted ())
                    {
                        tmpRun.getCTR().addNewRPr().addNewHighlight().setVal(STHighlightColor.LIGHT_GRAY);
                    }
                    tmpRun.setFontSize (( r.getFontSize () == -1) ? DEFAULT_FONT_SIZE :  r.getFontSize ());

                    tmpRun.setSubscript (r.getSubscript ());
                    tmpRun.setUnderline (r.getUnderline ());
                    // метод isPageBreak всегда возвращает false,
                    // независимо от того, содержится ли разрыв страницы в параграфе или нет
                    // так что используем грязный хак
                    para.setPageBreak (r.getCTR ().toString ().contains ("<w:br w:type=\"page\"/>"));
                }
                if (para.isPageBreak ()) {
                    try {
                        newdoc.write (new FileOutputStream (newDocName));
                    } catch (IOException e) {
                        e.printStackTrace ();
                    }
                    newdoc = new XWPFDocument ();
                    // требуется версия POI  >= 3.8 чтобы сделать это
                    newdoc.createStyles ();
                }
                //System.out.println(newpara.getNumLevelText ());
                try {
                    // сохраним последний кусок в файл
                    newdoc.write (new FileOutputStream (newDocName));
                } catch (IOException e) {
                    e.printStackTrace ();
                    //copyAllRunsToAnotherParagraph(para, newpara);
                }
            }
        }
    }

    public static void SaveTablesWithFormatDocx(XWPFDocument doc, String newDocName){

        XWPFDocument newdoc = new XWPFDocument ();
        //bullet and numering lists
        XWPFNumbering docNumbering = doc.getNumbering ();

        List<XWPFTable> tables = doc.getTables ();

        int numTbl = 1;
        for ( XWPFTable tbl : tables) {

            // title table
            XWPFParagraph fPar = newdoc.createParagraph ();
            XWPFRun fParRun = fPar.createRun ();
            fParRun.setText ("Table " + numTbl);
            fParRun.setFontSize (16);
            fParRun.setFontFamily ("Times New Roman");
            fParRun.setItalic (true);
            fParRun.setBold (true);

            // create table
            XWPFTable newTable = newdoc.createTable ();
            List<XWPFTableRow> row = tbl.getRows ();

            // create row
            int numRow = 0;
            for (XWPFTableRow xwpfTableRow : row) {

                //int size = newTable.getRows ().size ();
                //XWPFTableRow row1 = newTable.getRow (0);
                //int size1 = newTable.getRows ().size ();
                //boolean b = tbl.getRows ().size () > 1;*/

                XWPFTableRow lnewRow = numRow != 0 ? newTable.createRow () : newTable.getRow (0);

                //System.out.println(numRow);

                List<XWPFTableCell> cell = xwpfTableRow.getTableCells ();
                int numCell = 0;
                for (XWPFTableCell xwpfTableCell : cell) {

                    // create cell
                    XWPFTableCell lnewCell = lnewRow.getCell (numCell) !=  null ? lnewRow.getCell (numCell): lnewRow.addNewTableCell ();
                            //numCell != 0 ? lnewRow.addNewTableCell () : lnewRow.getCell (0);

                    CTTblWidth tblWidth = lnewCell.getCTTc ().addNewTcPr ().addNewTcW ();
                    tblWidth.setW(xwpfTableCell.getCTTc ().getTcPr ().getTcW ().getW ());
                    tblWidth.setType(xwpfTableCell.getCTTc ().getTcPr ().getTcW ().getType ());//STTblWidth.DXA);


                    if (xwpfTableCell != null) {

                        // create paragraph
                        for (XWPFParagraph para: xwpfTableCell.getParagraphs()) {
                            newdoc.createStyles ();

                            /*List<XWPFParagraph> paras = */

                            XWPFParagraph lnewPara = lnewCell.addParagraph ();
                            lnewPara.setStyle (para.getStyleID ());

                            // go to paragraph
                            BigInteger level = BigInteger.valueOf (0);
                            // noinspection Since15
                            XWPFStyles style = newdoc.getStyles ();

                                // set DEFAULT style
                            if (para.getStyleID () != null && !style.styleExist (para.getStyleID ())) {
                                    style.addStyle (doc.getStyles ().getStyle (para.getStyleID ()));//getStyle ()));
                                }

                                if (!para.getParagraphText ().isEmpty ()) {

                                    //XWPFParagraph lnewPara = newdoc.createParagraph ();
                                    //newpara.setStyle (para.getStyleID ());


                                    if (para.getNumID () != null) //(&& (para.getNumFmt ().toString () == "bullet")){
                                    {

                                        STNumberFormat.Enum stNumberFormat = STNumberFormat.Enum.forString (para.getNumFmt ().toString ());
                                        System.out.println (para.getNumIlvl () + " " + para.getNumFmt ().toString ()
                                                + " " + STNumberFormat.Enum.forString (para.getNumFmt ().toString ()).toString ()
                                                + " " + para.getNumID ().intValue () + "" + para.getNumLevelText () + " ");//*/
                                        //Old  CTAbstractNum
                                        XWPFNum oldXWPFNum = docNumbering.getNum (para.getNumID ());
                                        BigInteger oldAbstractNumId = oldXWPFNum.getCTNum ().getAbstractNumId ().getVal ();
                                        //ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                                        XWPFAbstractNum oldXWPFAbstractNum = docNumbering.getAbstractNum (oldAbstractNumId);
                                        CTAbstractNum oldCTAbstractNum = oldXWPFAbstractNum.getCTAbstractNum ();


                                        CTAbstractNum newCTAbstractNum = CTAbstractNum.Factory.newInstance ();
                                        newCTAbstractNum.setAbstractNumId (para.getNumIlvl ());//-- рабочий вариант с lvl
                                        CTLvl ctLvl = newCTAbstractNum.addNewLvl ();


                                        if (stNumberFormat.toString () == "bullet") {
                                            ctLvl.addNewNumFmt ().setVal (STNumberFormat.BULLET);
                                            ctLvl.addNewLvlText ().setVal (para.getNumLevelText ());//*/"");//
                                            // ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                                        } else {
                                            ctLvl.addNewNumFmt ().setVal (STNumberFormat.Enum.forString (para.getNumFmt ()));//DECIMAL_);
                                            String newlvltext = para.getNumLevelText ().toString ();
                                            ctLvl.addNewLvlText ().setVal ("%1.");
                                            ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));//para.getNumIlvl());
                                        }


                                        XWPFAbstractNum abstractNum = new XWPFAbstractNum (newCTAbstractNum);//,newnumbering);
                                        XWPFNumbering newnumbering = newdoc.createNumbering ();
                                        BigInteger abstractNumId = newnumbering.addAbstractNum (abstractNum);
                                        BigInteger numID = newnumbering.addNum (abstractNumId);

                                        System.out.println (numID.toString () + " " + abstractNumId.toString ());

                                        lnewPara.setNumID (numID);

                                        // работает на обычных списках
                 /*  System.out.println(para.getNumIlvl ().add (BigInteger.valueOf (1)).toString ());
                   lnewPara.setNumID (para.getNumIlvl ().add (BigInteger.valueOf (1)));
                   lnewPara.getCTP().getPPr().getNumPr().addNewIlvl().setVal(para.getNumIlvl ().add (BigInteger.valueOf (1)));
                 */

                                    }

                                    for (XWPFRun r : para.getRuns ()) {

                                        XWPFRun tmpRun = lnewPara.createRun ();
                                        tmpRun.setTextPosition (r.getTextPosition ());
                                        // важно использовать именно метод toString() поскольку
                                        // этот метод сохраняет возможные символы "\n", которые getText обрезает
                                        tmpRun.setText (r.toString ());
                                        tmpRun.setBold (r.isBold ());
                                        tmpRun.setFontFamily (r.getFontFamily ());
                                        tmpRun.setColor (r.getColor ());
                                        tmpRun.setEmbossed (r.isEmbossed ());
                                        tmpRun.setCapitalized (r.isCapitalized ());
                                        tmpRun.setShadow (r.isShadowed ());
                                        tmpRun.setItalic (r.isItalic ());
                                        tmpRun.setStrike (r.isStrike ());
                                        tmpRun.setDoubleStrikethrough (r.isDoubleStrikeThrough ());
                                        if (r.isHighlighted ()) {
                                            tmpRun.getCTR ().addNewRPr ().addNewHighlight ().setVal (STHighlightColor.LIGHT_GRAY);
                                        }
                                        tmpRun.setFontSize ((r.getFontSize () == -1) ? DEFAULT_FONT_SIZE : r.getFontSize ());

                                        tmpRun.setSubscript (r.getSubscript ());
                                        tmpRun.setUnderline (r.getUnderline ());
                                        // метод isPageBreak всегда возвращает false,
                                        // независимо от того, содержится ли разрыв страницы в параграфе или нет
                                        // так что используем грязный хак
                                        para.setPageBreak (r.getCTR ().toString ().contains ("<w:br w:type=\"page\"/>"));
                                    }
                                    if (para.isPageBreak ()) {
                                        try {
                                            newdoc.write (new FileOutputStream (newDocName));
                                        } catch (IOException e) {
                                            e.printStackTrace ();
                                        }
                                        newdoc = new XWPFDocument ();
                                        // требуется версия POI  >= 3.8 чтобы сделать это
                                        newdoc.createStyles ();
                                    }
                                    //System.out.println(lnewpara.getNumLevelText ());
                                    try {
                                        // сохраним последний кусок в файл
                                        newdoc.write (new FileOutputStream (newDocName));
                                    } catch (IOException e) {
                                        e.printStackTrace ();
                                        //copyAllRunsToAnotherParagraph(para, lnewpara);
                                    }
                                }
                            }
                        }
                    ++numCell;
                    }
                ++numRow;
                }
            ++numTbl;

        }
       // }
    }

    public static void SaveParagraphsToTableWithFormatDocx(XWPFDocument doc, String newDocName){

        List<XWPFParagraph> paras = doc.getParagraphs();
        XWPFDocument newdoc = new XWPFDocument();

        //bullet and numering lists
        XWPFNumbering docNumbering = doc.getNumbering ();
        //XWPFNumbering newnumbering = newdoc.createNumbering ();
        XWPFTable newTable = newdoc.createTable ();



        newTable.getCTTbl ().addNewTblGrid ().addNewGridCol ().setW (BigInteger.valueOf (7500));
        CTTblWidth tblWidth0 = newTable.getRow (0).getCell (0).getCTTc ().addNewTcPr ().addNewTcW ();
        tblWidth0.setW(BigInteger.valueOf (7500));
        tblWidth0.setType(STTblWidth.DXA);
        CTTblWidth tblWidth1 = newTable.getRow (0).addNewTableCell ().getCTTc ().addNewTcPr ().addNewTcW ();
        tblWidth1.setW(BigInteger.valueOf (7500));
        tblWidth1.setType(STTblWidth.DXA);

        XWPFRun cellr00 = newTable.getRow (0).getCell (0).getParagraphs ().get (0).createRun ();
        cellr00.setText ("English");
        cellr00.setFontFamily ("Times New Roman");
        cellr00.setFontSize (14);
        cellr00.setBold (true);

        XWPFRun cellr01 = newTable.getRow (0).getCell (1).getParagraphs ().get (0).createRun ();
        cellr01.setText ("Перевод");
        cellr01.setFontFamily ("Times New Roman");
        cellr01.setFontSize (14);
        cellr01.setBold (true);

//*/
        newdoc.createStyles ();
        // go to paragraph
        BigInteger level = BigInteger.valueOf (0);
       for (XWPFParagraph para : paras) {

            //noinspection Since15
            XWPFStyles style = newdoc.getStyles ();

            // set DEFAULT style
            if (para.getStyleID () != null && !style.styleExist ( para.getStyleID ())) {
                style.addStyle (doc.getStyles ().getStyle ( para.getStyleID ()));//getStyle ()));
            }


            if (!para.getParagraphText ().isEmpty ()) {

                XWPFTableRow lnewRow = newTable.createRow ();
                XWPFTableCell lnewCell = lnewRow.getCell (0);
                XWPFTableCell lnewCell2 = lnewRow.getCell (1);//lnewRow.createCell ();
                //CTTblWidth lcellWidth2 = lnewCell2.getCTTc ().addNewTcPr ().addNewTcW ();
                //lcellWidth2.setW(BigInteger.valueOf (5000));
                //lcellWidth2.setType(STTblWidth.DXA);

                XWPFParagraph lnewPara = lnewCell.getParagraphs().get (0);
                XWPFParagraph lnewPara2 = lnewCell2.addParagraph ();
                lnewPara.setStyle (para.getStyleID ());
                lnewPara2.setStyle (para.getStyleID ());


                if(para.getNumID () != null) //(&& (para.getNumFmt ().toString () == "bullet")){
                {

                    STNumberFormat.Enum stNumberFormat = STNumberFormat.Enum.forString (para.getNumFmt ().toString ());
                   //System.out.println(para.getNumIlvl () + " " + para.getNumFmt ().toString ()
                   //         + " " +STNumberFormat.Enum.forString (para.getNumFmt ().toString ()).toString ()
                    //        + " " + para.getNumID ().intValue () + "" + para.getNumLevelText ()+ " " );//
                    //Old  CTAbstractNum
                    XWPFNum oldXWPFNum = docNumbering.getNum (para.getNumID ());
                    BigInteger oldAbstractNumId = oldXWPFNum.getCTNum ().getAbstractNumId ().getVal ();
                    //ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                    XWPFAbstractNum oldXWPFAbstractNum = docNumbering.getAbstractNum (oldAbstractNumId);
                    CTAbstractNum oldCTAbstractNum = oldXWPFAbstractNum.getCTAbstractNum ();


                    CTAbstractNum newCTAbstractNum = CTAbstractNum.Factory.newInstance ();
                    newCTAbstractNum.setAbstractNumId (para.getNumIlvl ());//-- рабочий вариант с lvl
                    CTLvl ctLvl = newCTAbstractNum.addNewLvl ();



                    if(stNumberFormat.toString () == "bullet") {
                        ctLvl.addNewNumFmt ().setVal (STNumberFormat.BULLET);
                        ctLvl.addNewLvlText ().setVal (para.getNumLevelText ());//"");//
                        // ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));
                    }
                    else
                    {
                        ctLvl.addNewNumFmt().setVal (STNumberFormat.Enum.forString (para.getNumFmt ()));//DECIMAL_);
                        String newlvltext = para.getNumLevelText ().toString ();
                        ctLvl.addNewLvlText ().setVal ("%1.");
                        ctLvl.addNewStart ().setVal (BigInteger.valueOf (1));//para.getNumIlvl());
                    }


                    XWPFAbstractNum abstractNum = new XWPFAbstractNum (newCTAbstractNum);//,newnumbering);
                    XWPFNumbering newnumbering = newdoc.createNumbering ();
                    BigInteger abstractNumId  = newnumbering.addAbstractNum (abstractNum);
                    BigInteger numID = newnumbering.addNum (abstractNumId);

                   // System.out.println(numID.toString ()+ " " + abstractNumId.toString ());

                    lnewPara.setNumID (numID);


                    // работает на обычных списках
                 //  System.out.println(para.getNumIlvl ().add (BigInteger.valueOf (1)).toString ());
                 //  newpara.setNumID (para.getNumIlvl ().add (BigInteger.valueOf (1)));
                 //  newpara.getCTP().getPPr().getNumPr().addNewIlvl().setVal(para.getNumIlvl ().add (BigInteger.valueOf (1)));
                ///

                }

                for (XWPFRun r : para.getRuns ()) {

                    XWPFRun tmpRun = lnewPara.createRun ();
                    tmpRun.setTextPosition (r.getTextPosition ());
                    // важно использовать именно метод toString() поскольку
                    // этот метод сохраняет возможные символы "\n", которые getText обрезает
                    tmpRun.setText (r.toString ());
                    tmpRun.setBold (r.isBold ());
                    tmpRun.setFontFamily (r.getFontFamily ());
                    tmpRun.setColor (r.getColor ());
                    tmpRun.setEmbossed (r.isEmbossed ());
                    tmpRun.setCapitalized (r.isCapitalized ());
                    tmpRun.setShadow (r.isShadowed ());
                    tmpRun.setItalic (r.isItalic ());
                    tmpRun.setStrike (r.isStrike ());
                    tmpRun.setDoubleStrikethrough (r.isDoubleStrikeThrough ());
                    if (r.isHighlighted ())
                    {
                        tmpRun.getCTR().addNewRPr().addNewHighlight().setVal(STHighlightColor.LIGHT_GRAY);
                    }
                    tmpRun.setFontSize (( r.getFontSize () == -1) ? DEFAULT_FONT_SIZE :  r.getFontSize ());

                    tmpRun.setSubscript (r.getSubscript ());
                    tmpRun.setUnderline (r.getUnderline ());
                    // метод isPageBreak всегда возвращает false,
                    // независимо от того, содержится ли разрыв страницы в параграфе или нет
                    // так что используем грязный хак
                    para.setPageBreak (r.getCTR ().toString ().contains ("<w:br w:type=\"page\"/>"));
                }
                if (para.isPageBreak ()) {
                    try {
                        newdoc.write (new FileOutputStream (newDocName));
                    } catch (IOException e) {
                        e.printStackTrace ();
                    }
                    newdoc = new XWPFDocument ();
                    // требуется версия POI  >= 3.8 чтобы сделать это
                    newdoc.createStyles ();
                }
                //System.out.println(newpara.getNumLevelText ());
                try {
                    // сохраним последний кусок в файл
                    newdoc.write (new FileOutputStream (newDocName));
                } catch (IOException e) {
                    e.printStackTrace ();
                    //copyAllRunsToAnotherParagraph(para, newpara);
                }
            }
        }
    }

    // Copy all runs from one paragraph to another, keeping the style unchanged
    @SuppressWarnings("Since15")
    public static void copyAllRunsToAnotherParagraph(XWPFParagraph oldPar, XWPFParagraph newPar) {
        final int DEFAULT_FONT_SIZE = 10;

        for (XWPFRun run : oldPar.getRuns ()) {
            String textInRun = run.getText (0);
            //noinspection Since15
            if (textInRun == null || textInRun.isEmpty ()) {
                continue;
            }

            int fontSize = run.getFontSize ();
            System.out.println ("run text = '" + textInRun + "' , fontSize = " + fontSize);

            XWPFRun newRun = newPar.createRun ();

            // Copy text
            newRun.setText (textInRun);

            // Apply the same style
            newRun.setFontSize ((fontSize == -1) ? DEFAULT_FONT_SIZE : run.getFontSize ());
            newRun.setFontFamily (run.getFontFamily ());
            newRun.setBold (run.isBold ());
            newRun.setItalic (run.isItalic ());
            newRun.setStrike (run.isStrike ());
            newRun.setColor (run.getColor ());

        }
    }



    public void ExportTextDocx()
    {
        try {
            if(xwpfDocument == null) {
                return;
            }

            File writeFile = new File("Техt.docx");
            FileOutputStream fos = new FileOutputStream(writeFile.getAbsolutePath());
            XWPFDocument writeDocx = new XWPFDocument();
            XWPFParagraph par = writeDocx.createParagraph();
            XWPFRun run = par.createRun();

            int nRows = 3;
            int nCols = 2;
            XWPFTable table = writeDocx.createTable(nRows, nCols);
            CTTblPr tblPr = table.getCTTbl().getTblPr();
            CTString styleStr = tblPr.addNewTblStyle();
            styleStr.setVal("StyledTable");

            // Get a list of the rows in the table
            List<XWPFTableRow> rows = table.getRows();
            int rowCt = 0;
            int colCt = 0;
            for (XWPFTableRow row : rows) {
                // get table row properties (trPr)
                CTTrPr trPr = row.getCtRow().addNewTrPr();
                // set row height; units = twentieth of a point, 360 = 0.25"
                CTHeight ht = trPr.addNewTrHeight();
                ht.setVal(BigInteger.valueOf(360));

                // get the cells in this row
                List<XWPFTableCell> cells = row.getTableCells();
                // add content to each cell
                for (XWPFTableCell cell : cells) {
                    // get a table cell properties element (tcPr)
                    CTTcPr tcpr = cell.getCTTc().addNewTcPr();
                    // set vertical alignment to "center"
                    CTVerticalJc va = tcpr.addNewVAlign();
                    va.setVal(STVerticalJc.CENTER);

                    // get 1st paragraph in cell's paragraph list
                    XWPFParagraph para = cell.getParagraphs().get(0);
                    // create a run to contain the content
                    XWPFRun rh = para.createRun();
                    // style cell as desired
                    if (colCt == nCols - 1) {
                        // last column is 10pt Courier
                        rh.setFontSize(10);
                        rh.setFontFamily("Courier");
                    }
                    if (rowCt == 0) {
                        // header row
                        rh.setText("header row, col " + colCt);
                        rh.setBold(true);
                        para.setAlignment(ParagraphAlignment.CENTER);
                    } else {
                        // other rows
                        rh.setText("row " + rowCt + ", col " + colCt);
                        para.setAlignment(ParagraphAlignment.LEFT);
                    }
                    colCt++;
                } // for cell
                colCt = 0;
                rowCt++;
            } // for row
            writeDocx.write (fos);
            fos.close ();
            writeDocx.close();

        }
        catch(Exception ex) {
            System.out.println (ex.getMessage ());
            ex.printStackTrace ();
        }

    }

    public void ParseTable() {
        try {
            if (hwpfDocument != null) {
                hwpfDocument.getRange ();
                Range documentRange = hwpfDocument.getRange ();
                int ic = documentRange.getEndOffset ();

                for (int i = 0; i < ic; ++i) {
                    try {
                        Paragraph startOfInnerTable = documentRange.getParagraph (i);
                        Table innerTable = documentRange.getTable (startOfInnerTable);
                        int rc = innerTable.numRows ();

                        for (int r = 0; r < rc; ++r) {
                            System.out.print (hwpfDocument.getDirectory ().getName ().replace ("\n", " ").replace ("\t", " "));
                            System.out.print ("\t");
                            System.out.print ("\t");
                            System.out.print (r);
                            System.out.print ("\t");
                            TableRow tableRow = innerTable.getRow (r);
                            int cc = tableRow.numCells ();

                            for (int c = 0; c < cc; ++c) {
                                TableCell tableCell = tableRow.getCell (c);
                                System.out.print ("\t");
                                System.out.print (tableCell.text ().replace ("\r", " ").replace ("\n", " ").replace ("\t", " ").replaceAll ("\\p{Cntrl}", ""));
                            }

                            System.out.println ();
                        }
                    } catch (Exception var22) {
                        ;
                    }
                    hwpfDocument.close();
                }
            }
              /*  if (xwpfDocument != null) {
                    try {
                        Iterator<IBodyElement> bodyElementIterator = xwpfDocument.getBodyElementsIterator ();
                        while (bodyElementIterator.hasNext ()) {
                            IBodyElement element = bodyElementIterator.next ();

                            if ("TABLE".equalsIgnoreCase (element.getElementType ().name ())) {
                                List<XWPFTable> tableList = element.getBody ().getTables ();
                                for (XWPFTable table : tableList) {
                                    System.out.println ("Total Number of Rows of Table:" + table.getNumberOfRows ());
                                    System.out.println (table.getText ());
                                }
                            }
                        }
                    } catch (Exception var22) {
                        ;
                    }
                    xwpfDocument.close();
                }*/
            } catch (FileNotFoundException var23) {
            var23.printStackTrace ();
            //throw var23;
        } catch (IOException var24) {
            var24.printStackTrace ();
           // throw var24;
        } finally {
           /* if(processed != 1) {
                //err.println(" !!warning ".concat(arg.getAbsolutePath()).concat(" processed not 1 Sheets"));
            }*/

        }
    }


    public void ParseAll()
    {
        try {

            if (hwpfDocument != null) {
                Range range = hwpfDocument.getRange ();
                for (int i = 0; i < range.numParagraphs (); i++) {
                    Paragraph par = range.getParagraph (i);

                    if (!par.isInTable ()) {
                        System.out.println ("text:" + par.text ());
                    } else {
                        Table table = range.getTable (par);
                        for (int rowIdx = 0; rowIdx < table.numRows (); rowIdx++) {
                            TableRow row = table.getRow (rowIdx);
                            for (int colIdx = 0; colIdx < row.numCells (); colIdx++) {
                                TableCell cell = row.getCell (colIdx);
                                System.out.print (" column=" + cell.getParagraph (0).text ());
                                i++;
                            }
                            System.out.println ();
                            i++;
                        }
                    }
                }
                hwpfDocument.close ();
            }
            if (xwpfDocument != null) {
                XWPFWordExtractor extractor = new XWPFWordExtractor(xwpfDocument);
                System.out.println(extractor.getText());
               /* Iterator<IBodyElement> bodyElementIterator = xwpfDocument.getBodyElementsIterator ();
                while (bodyElementIterator.hasNext ()) {
                    IBodyElement element = bodyElementIterator.next ();
                    List<XWPFParagraph> paragraphList = element.getBody ().getParagraphs ();
                    for (XWPFParagraph paragraph: paragraphList){
                        if ("paragraph".equalsIgnoreCase (element.getElementType ().name ()))
                            System.out.println(paragraph.getText());
                    }

                    if ("TABLE".equalsIgnoreCase (element.getElementType ().name ())) {
                        List<XWPFTable> tableList = element.getBody ().getTables ();
                        for (XWPFTable table : tableList) {
                            System.out.println ("Total Number of Rows of Table:" + table.getNumberOfRows ());
                            System.out.println (table.getText ());
                        }
                    }
                }*/
            }
        } catch (FileNotFoundException var23) {
            var23.printStackTrace ();
            //throw var23;
        } catch (IOException var24) {
            var24.printStackTrace ();
            // throw var24;
        } finally {
           /* if(processed != 1) {
                //err.println(" !!warning ".concat(arg.getAbsolutePath()).concat(" processed not 1 Sheets"));
            }*/

        }
    }

}
