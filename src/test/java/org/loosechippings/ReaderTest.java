package org.loosechippings;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;

public class ReaderTest {

    private static String TEST_FILE= "test.xlsx";
    private ExcelReader reader;

    @Before
    public void init() throws IOException, InvalidFormatException {
        String filePath=this.getClass().getResource(TEST_FILE).getPath();
        File f=new File(filePath);
        reader=new ExcelReader(f);
    }

    @Test
    public void getSheetNames() {
        List<String> sheetNames=reader.getSheetNames();
        assertEquals(2,sheetNames.size());
        assertEquals("Sheet1",sheetNames.get(0));
        assertEquals("Sheet2",sheetNames.get(1));
    }

    @Test
    public void canReturnHeadersAsAList() {
        List<String> headers=reader.getHeaders(0);
        List<String> expected=new ArrayList();
        expected.add("id");
        expected.add("name");
        expected.add("height");
        expected.add("job");
        assertEquals(expected,headers);
    }

    @Test
    public void canReturnDataAsJSON() {
        String expected="{\"id\":\"1.0\",\"name\":\"Andy\",\"height\":\"1.9\",\"job\":\"chef\"}";
        String actual=reader.getNextRecord();
        assertEquals(expected,actual);
    }
}
