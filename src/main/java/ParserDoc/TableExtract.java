package ParserDoc;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.IOException;
import java.util.List;

/**
 * Created by yamashkina on 24.04.2017.
 */
public class TableExtract {

    private List<XWPFTable> docTables;

    public List<XWPFTable> getTables()
    {
        return docTables;
    }

    public TableExtract(HWPFDocument doc) throws IOException, Exception {


    }
}
