package io.github.yorucyber.XLSXReader;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
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
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class XLSXReader {
    private DocumentBuilder documentBuilder = null;
    private Document currentSheet = null;
    private ZipFile excelFile = null;

    public XLSXReader(DocumentBuilder documentBuilder, ZipFile excelFile) {
        this.documentBuilder = documentBuilder;
        this.excelFile = excelFile;


    }

    //FOR DEBUG!!!
    public void createXMLFile() {
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


    public Document GetCurrentSheet() {
        return this.currentSheet;
    }

    public void SetCurrentSheet(String sheetName) throws IOException, SAXException {
        ZipEntry sheet = this.excelFile.getEntry(sheetName);
        InputStream sheetStream = this.excelFile.getInputStream(sheet);
        this.currentSheet = this.documentBuilder.parse(sheetStream);
    }

    public Node GetCell(String cellName) {
        Document doc = this.currentSheet;
        doc.getDocumentElement().normalize();
        NodeList sheetDataNodes = doc.getElementsByTagName("sheetData");
        Element sheetData = (Element) sheetDataNodes.item(0);
        NodeList sheetRows = sheetData.getElementsByTagName("row");
        for (int i = 0; i < sheetRows.getLength(); i++) {
            Element element = (Element) sheetRows.item(i);
            NodeList cells = element.getElementsByTagName("c");
            for (int j = 0; j < cells.getLength(); j++) {
                Element cell = (Element) cells.item(j);
               if (cell.getAttribute("r").equals(cellName)){
                   return cell.getElementsByTagName("v").item(0);
               }
            }

        }
        return null;
    }

    public void GetEntries() {
        Enumeration<? extends ZipEntry> entries = this.excelFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = entries.nextElement();
            String entryName = entry.getName();
            if (entryName.endsWith(".xml") && entryName.startsWith("xl/")) {
                InputStream inputStream = null;
                try {
                    inputStream = this.excelFile.getInputStream(entry);
                    Document document = this.documentBuilder.parse(inputStream);
                } catch (IOException | SAXException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }
}
