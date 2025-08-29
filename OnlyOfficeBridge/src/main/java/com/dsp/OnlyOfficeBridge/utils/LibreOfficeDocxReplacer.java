package com.dsp.OnlyOfficeBridge.utils;



import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeManager;
//import org.jodconverter.local.LocalConverter;
//import org.jodconverter.local.office.LocalOfficeManager;
import org.jodconverter.remote.RemoteConverter;
import org.jodconverter.remote.office.RemoteOfficeManager;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.nio.file.Files;
import java.util.Map;

public class LibreOfficeDocxReplacer {

    public static ByteArrayResource replaceWithLibreOffice(String templatePath, Map<String, String> replacements) throws Exception {
        ClassPathResource resource = new ClassPathResource(templatePath);
        File inputFile = resource.getFile();//new File(inputPath);
//        File inputFile = new File(templatePath);
        File tempFile = File.createTempFile("converted", ".fodt"); // Flat ODT

        // Start LibreOffice
//        OfficeManager officeManager = LocalOfficeManager.install();
        RemoteOfficeManager officeManager = RemoteOfficeManager.builder()
                // If your LibreOffice server runs in socket mode:
                //.urlConnection("socket://<REMOTE_SERVER_IP>:2002")

                // If your LibreOffice server exposes REST API:
                .urlConnection("http://10.247.128.22:2002:8080")
                .build();

        officeManager.start();

        try {
            // Step 1: Convert DOCX → FODT
//            LocalConverter.make(officeManager)
//                    .convert(inputFile)
//                    .to(tempFile)
//                    .execute();
            RemoteConverter.make(officeManager)
                    .convert(inputFile)
                    .to(tempFile)
                    .execute();
//            DocumentConverter convertter = RemoteConverterrter.make(officeManager);

            // Step 2: Replace placeholders in FODT
            String content = Files.readString(tempFile.toPath());
            for (Map.Entry<String, String> entry : replacements.entrySet()) {
                content = content.replace("{" + entry.getKey() + "}", entry.getValue());
            }
            Files.writeString(tempFile.toPath(), content);

            // Step 3: Convert FODT → DOCX
            File outputFile = File.createTempFile("generated-libre", ".docx");
            RemoteConverter.make(officeManager)
                    .convert(tempFile)
                    .to(outputFile)
                    .execute();
            String content1 = Files.readString(tempFile.toPath());
            for (Map.Entry<String, String> entry : replacements.entrySet()) {
                content1 = content1.replace("{" + entry.getKey() + "}", entry.getValue());
            }
            Files.writeString(tempFile.toPath(), content1);

            // Step 3: Convert FODT → DOCX
            RemoteConverter.make(officeManager)
                    .convert(tempFile)
                    .to(outputFile)
                    .execute();

            // Step 4: Read as bytes and return as resource
            byte[] fileBytes = Files.readAllBytes(outputFile.toPath());
            return new ByteArrayResource(fileBytes);
//            return outputFile;

        } finally {
            officeManager.stop();
        }
    }


}
