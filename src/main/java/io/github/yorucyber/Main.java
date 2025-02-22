package io.github.yorucyber;

import io.github.yorucyber.XLSXReader.XLSXReader;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

public class Main {
    public static void main(String[] args) throws ParserConfigurationException {
        String xlsxFilePath = "./FinancialSample.xlsx";
        DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder documentBuilder = documentBuilderFactory.newDocumentBuilder();
        XLSXReader xlsxReader = new XLSXReader(documentBuilder);
        xlsxReader.GetEntries(xlsxFilePath);
//        var dirs = XLSXReader.GetEntries(zipFilePath);
    }
}
