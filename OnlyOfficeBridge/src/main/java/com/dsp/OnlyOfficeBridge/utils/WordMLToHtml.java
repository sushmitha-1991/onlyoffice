package com.dsp.OnlyOfficeBridge.utils;

import javax.xml.transform.*;
import javax.xml.transform.stream.*;
import java.io.*;

public class WordMLToHtml {

    public static String convert(String wordMLXml) throws Exception {

        // Load the XSL from resources
        InputStream xslStream = WordMLToHtml.class
                .getClassLoader()
                .getResourceAsStream("WordML2HTML.xsl");

        if (xslStream == null) {
            throw new RuntimeException("WordML2HTML.xsl not found in resources folder");
        }

        TransformerFactory factory = TransformerFactory.newInstance();

        Source xslt = new StreamSource(xslStream);
        Transformer transformer = factory.newTransformer(xslt);

        Source xmlSource = new StreamSource(new StringReader(wordMLXml));

        StringWriter htmlWriter = new StringWriter();
        StreamResult result = new StreamResult(htmlWriter);

        transformer.transform(xmlSource, result);

        return htmlWriter.toString();
    }


}

