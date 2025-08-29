/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.controller;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import com.dsp.OnlyOfficeBridge.utils.LibreOfficeDocxReplacer;
import com.dsp.OnlyOfficeBridge.utils.Utils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.dsp.OnlyOfficeBridge.service.OnlyOfficeService;

/**
 * 
 */
@RestController
@RequestMapping()
public class OnlyOfficeController {

	@Autowired
	private OnlyOfficeService onlyOfficeService;

	@PostMapping("/config")
	public ResponseEntity<Map<String, Object>> getEditorConfig(@RequestBody Map<String, String> payload) {
		Map<String, Object> response = onlyOfficeService.getEditorConfig(payload);
		return ResponseEntity.ok(response);
	}

//	309610fc-dd27-4668-8920-2bf2a154fc01
	@GetMapping("/download/{key}/{templateId}/{refId}")
	public ResponseEntity<Resource> downloadByKey(@PathVariable String key,@PathVariable String templateId,@PathVariable String refId) {
		return onlyOfficeService.downloadByKey(key,templateId,refId);
	}

	@PostMapping("/callback")
	public ResponseEntity<Map<String, Object>> handleCallback(@RequestParam("dsp2Key") String dsp2Key,@RequestParam("applicationId") String applicationId,
			@RequestBody Map<String, Object> payload) {
		return onlyOfficeService.handleCallback(dsp2Key,applicationId, payload);
	}


    @GetMapping("/download-empty-docx/{fileName}")
    public ResponseEntity<byte[]> downloadEmptyDocx(@PathVariable String fileName) throws IOException {
        // Create an empty DOCX document
        XWPFDocument document = new XWPFDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        document.write(out);
        document.close();

        byte[] docBytes = out.toByteArray();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", fileName);

        return ResponseEntity.ok()
                .headers(headers)
                .body(docBytes);
    }


    @GetMapping("/download-empty-test/test")
    public ResponseEntity<byte[]> downloadEmptyDocxTest() throws Exception {
        // Create an empty DOCX document
        Map<String, String> replacements = Map.of(
                "range", "2024-25",
                "date", "29-08-2025",
                "myno", "123",
                "branch", "Research Branch",
                "subject", "Parliamentary Updates",
                "contain", "This is the dynamic content inserted into the template."
        );

        ByteArrayResource result = LibreOfficeDocxReplacer.replaceWithLibreOffice
                ("/doc-template/header.docx", replacements);
        ByteArrayResource result1 = LibreOfficeDocxReplacer.replaceWithLibreOffice
                ("/doc-template/contain.docx", replacements);
        ByteArrayResource result2 = LibreOfficeDocxReplacer.replaceWithLibreOffice
                ("/doc-template/footer.docx", replacements);
        byte[] bb = Utils.mergeDocxFiles(List.of(result.getByteArray(),result1.getByteArray(),result2.getByteArray()));

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=generated.docx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(bb.length)
                .body(bb);
    }

}
