# XLSX Reader 

This is a personal Java library made from scratch to read data from `.xlsx` files. 
This library implements a basic, custom parsing approach without relying on external libraries like Apache POI. Currently, 
it provides the fundamental functionality: parsing an XLSX file and creating a document object representing its structure.


## Features

* **XLSX Parsing (Custom):** Reads `.xlsx` files and converts them into an internal document object, without external dependencies.
## Getting Started

### Prerequisites

* Java Development Kit (JDK) 8 or later

### Installation

1.  **Clone the repository (or download the source code):**

    ```bash
    git clone https://github.com/Yoru-cyber/XLSXReaderJava
    ```

2.  **Build the library:**

    If you are using an IDE like IntelliJ IDEA or Eclipse, you can build the project directly from the IDE. If you are using the command line, navigate to the project directory and compile the Java files.

## Example

```java
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
    }
}
```
## Notes

* Currently, this is just a hobby project so it's very basic.

## Roadmap

- [x] Set and retrieve sheet.
- [ ] Set and retrieve document.
- [x] Get cell value.
- [ ] Set cell value.
- [ ] Change sheet's name.
- [ ] Core tests.
