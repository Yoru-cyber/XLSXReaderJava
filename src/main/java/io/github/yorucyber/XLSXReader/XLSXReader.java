package io.github.yorucyber.XLSXReader;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import java.io.IOException;
import java.io.InputStream;
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

    public Document GetCurrentSheet() {
        return this.currentSheet;
    }

    public void SetCurrentSheet(String sheetName) throws IOException, SAXException {
        ZipEntry sheet = this.excelFile.getEntry(sheetName);
        InputStream sheetStream = this.excelFile.getInputStream(sheet);
        this.currentSheet = this.documentBuilder.parse(sheetStream);
    }

//    public void PrintCell() {
//
//    }

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
