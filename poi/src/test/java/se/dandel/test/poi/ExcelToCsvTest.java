package se.dandel.test.poi;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;

public class ExcelToCsvTest {

    @Test
    public void simple() throws Exception {
        InputStream is = getClass().getResourceAsStream("/poi-test.xlsx");
        OutputStream os = new ByteArrayOutputStream();
        ExcelToCsv excelToCsv = new ExcelToCsv(is, os);
        excelToCsv.readAndWrite();
    }

}
