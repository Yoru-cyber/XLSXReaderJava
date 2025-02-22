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
    private final DocumentBuilder documentBuilder;

    public XLSXReader(DocumentBuilder documentBuilder) {
        this.documentBuilder = documentBuilder;
    }

    public void GetEntries(String filePath) {
        try (ZipFile zipFile = new ZipFile(filePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String entryName = entry.getName();
                if (entryName.endsWith(".xml") && entryName.startsWith("xl/")) {
                    InputStream inputStream = zipFile.getInputStream(entry);
                    Document document = this.documentBuilder.parse(inputStream);
                    document.getDocumentElement().normalize();
                }
            }
        } catch (IOException | SAXException e) {
            throw new RuntimeException(e);
        }
    }
}
