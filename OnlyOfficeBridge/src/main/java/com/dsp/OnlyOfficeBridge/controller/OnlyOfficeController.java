/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.controller;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
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
	@GetMapping("/download/{key}")
	public ResponseEntity<Resource> downloadByKey(@PathVariable String key) {
		return onlyOfficeService.downloadByKey(key);
	}

	@PostMapping("/callback")
	public ResponseEntity<Map<String, Object>> handleCallback(@RequestParam("dsp2Key") String dsp2Key,
			@RequestBody Map<String, Object> payload) {
		return onlyOfficeService.handleCallback(dsp2Key, payload);
	}


    @GetMapping("/download-empty-docx")
    public ResponseEntity<byte[]> downloadEmptyDocx() throws IOException {
        // Create an empty DOCX document
        XWPFDocument document = new XWPFDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        document.write(out);
        document.close();

        byte[] docBytes = out.toByteArray();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "empty.docx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(docBytes);
    }




}
