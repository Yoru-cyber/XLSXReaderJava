package io.github.yorucyber;

import io.github.yorucyber.XLSXReader.XLSXReader;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.zip.ZipFile;

public class Main {
    public static void main(String[] args) throws ParserConfigurationException {
        String xlsxFilePath = "./FinancialSample.xlsx";
        DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();
        try (ZipFile excelFile = new ZipFile(xlsxFilePath)) {
            XLSXReader xlsxReader = new XLSXReader(documentBuilder, excelFile);
            xlsxReader.SetCurrentSheet("xl/worksheets/sheet1.xml");
            Element cell = (Element) xlsxReader.GetCell("B1");
            System.out.println(cell.getTextContent());
//            xlsxReader.createXMLFile();

        } catch (IOException | SAXException e) {
            throw new RuntimeException(e);
        }


    }
}
