package io.github.yorucyber.XLSXReader;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;
import java.util.HashMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Reader object that holds an Excel file and reads through it.
 *
 * @author Carlos Mendez
 * @version 0.0.1
 */
public class XLSXReader {
    /**
     * A hashmap that holds the indexes of the shared strings of the spreadsheet.
     */
    private final HashMap<Integer, String> sharedStrings = new HashMap<>();
    /**
     * Document Builder instance provided on constructor.
     */
    private DocumentBuilder documentBuilder = null;
    /**
     * The current set sheet of a workbook.
     */
    private Document currentSheet = null;
    /**
     * The provided Excel file.
     */
    private ZipFile excelFile = null;

    /**
     * Constructor of {@link XLSXReader}.
     *
     * @param documentBuilder A Document builder instance.
     * @param excelFile       An instance of {@link ZipFile} from a .xlsx file.
     */
    public XLSXReader(DocumentBuilder documentBuilder, ZipFile excelFile) {
        this.documentBuilder = documentBuilder;
        this.excelFile = excelFile;
        parseSharedStrings("xl/sharedStrings.xml");
    }

    /**
     * Reads the sharedStrings XML document and creates a hashmap with its values.
     *
     * @param path The path to the sharedStrings XML file on a standard SpreadSheet ML.
     */
    private void parseSharedStrings(String path) {
        ZipEntry entry = this.excelFile.getEntry(path);
        try {
            Document sharedStringsDoc = this.documentBuilder.parse(this.excelFile.getInputStream(entry));
            NodeList sharedStrings = sharedStringsDoc.getElementsByTagName("t");
            for (int i = 0; i < sharedStrings.getLength(); i++) {
                this.sharedStrings.put(i, sharedStrings.item(i).getTextContent());
            }
        } catch (IOException | SAXException e) {
            throw new RuntimeException(e);
        }
    }

    //FOR DEBUG!!!
    private void createXMLFile() {
        try {
            //clearly a bad practice but needed for speed
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer transformer = tf.newTransformer();
            StringWriter writer = new StringWriter();
            transformer.transform(new DOMSource(this.currentSheet), new StreamResult(writer));
            String xmlString = writer.getBuffer().toString();
            FileWriter fileWriter = new FileWriter("excel.xml");
            fileWriter.write(xmlString);
            fileWriter.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Retrieves the current sheet.
     *
     * @return If a sheet from a workbook was set, returns an instance of {@link Document} otherwise null
     */
    public Document GetCurrentSheet() {
        return this.currentSheet;
    }

    /**
     * Sets sheet if exists on provided Excel file.
     *
     * @throws IOException  if failed to read file.
     * @throws SAXException if idk.
     */
    public void SetCurrentSheet(String sheetName) throws IOException, SAXException {
        ZipEntry sheet = this.excelFile.getEntry(sheetName);
        InputStream sheetStream = this.excelFile.getInputStream(sheet);
        this.currentSheet = this.documentBuilder.parse(sheetStream);
    }

    /**
     * Finds the cell that matches the passed string.
     *
     * @param cellName The name of the cell to find, like "C2".
     * @return String with value if cell was found and was not empty
     */
    public String GetCell(String cellName) {
        Document doc = this.currentSheet;
        doc.getDocumentElement().normalize();
        NodeList sheetDataNodes = doc.getElementsByTagName("sheetData");
        // getElements as it is written in plural, returns a list even if only has one match
        Element sheetData = (Element) sheetDataNodes.item(0);
        NodeList sheetRows = sheetData.getElementsByTagName("row");
        for (int i = 0; i < sheetRows.getLength(); i++) {
            Element row = (Element) sheetRows.item(i);
            //c is for cells and not columns
            NodeList cells = row.getElementsByTagName("c");
            for (int j = 0; j < cells.getLength(); j++) {
                Element cell = (Element) cells.item(j);
                //r stands for reference which is the name of the cell, example: <c r="B1"><v>6</v></c>
                if (cell.getAttribute("r").equals(cellName)) {
                    Element value = (Element) cell.getElementsByTagName("v").item(0);
                    //when cell has attribute t, it's value is "s" which stands for string, therefore if t == true we return the value from the hashmap
                    if (cell.hasAttribute("t")) {
                        Integer key = Integer.valueOf(value.getTextContent());
                        return this.sharedStrings.get(key);
                    }
                    return value.getTextContent();

                }
            }

        }
        return null;
    }

    // Not sure if this method could be useful either
//    public Enumeration<? extends ZipEntry> GetEntries() {
//        return this.excelFile.entries();
//    }
// I don't think this method could be useful, but it's going to stay commented until I figure it out.
//    public void getSharedString() throws IOException {
//        ZipEntry entry = this.excelFile.getEntry("xl/sharedStrings.xml");
//        try {
//            //clearly a bad practice but needed for speed
//            TransformerFactory tf = TransformerFactory.newInstance();
//            Transformer transformer = tf.newTransformer();
//            StringWriter writer = new StringWriter();
//            transformer.transform(new DOMSource(this.documentBuilder.parse(this.excelFile.getInputStream(entry))), new StreamResult(writer));
//            String xmlString = writer.getBuffer().toString();
//            FileWriter fileWriter = new FileWriter("sharedStrings.xml");
//            fileWriter.write(xmlString);
//            fileWriter.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//
//
//    }
}

