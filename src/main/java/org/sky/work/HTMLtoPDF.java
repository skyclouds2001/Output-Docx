package org.sky.work;

import java.io.*;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.html2pdf.resolver.font.DefaultFontProvider;
import com.itextpdf.layout.font.FontProvider;

public class HTMLtoPDF {
    private static final String resourceDir = System.getProperty("user.dir") + "/src/main/resources";
    public void outputHTMLtoPDF () {
        String sourcePath = resourceDir + "/data.html";
        String outputPath = resourceDir + "/data.pdf";

        ConverterProperties converterProperties = new ConverterProperties();

        FontProvider fontProvider = new DefaultFontProvider();
        fontProvider.addFont("C:/Windows/Fonts");

        converterProperties.setFontProvider(fontProvider);

        try {
            InputStream inputStream = new FileInputStream(sourcePath);
            OutputStream outputStream = new FileOutputStream(outputPath);
            HtmlConverter.convertToPdf(inputStream, outputStream, converterProperties);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
