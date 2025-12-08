/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.controller;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.dsp.OnlyOfficeBridge.response.GenericResponse;
import com.dsp.OnlyOfficeBridge.service.OnlyOfficeService;
import com.dsp.OnlyOfficeBridge.utils.LibreOfficeDocxReplacer;
import com.dsp.OnlyOfficeBridge.utils.ThirdPartyDocGen;
import com.dsp.OnlyOfficeBridge.utils.Utils;

/**
 * 
 */
@RestController
@RequestMapping()
public class OnlyOfficeController {

	@Autowired
	private OnlyOfficeService onlyOfficeService;
	@Autowired
	ThirdPartyDocGen thirdPartyDocGen;

	@PostMapping("/config")
	public ResponseEntity<Map<String, Object>> getEditorConfig(@RequestBody Map<String, String> payload) {
		Map<String, Object> response = onlyOfficeService.getEditorConfig(payload);
		return ResponseEntity.ok(response);
	}

	@GetMapping("/convertHtml")
	public ResponseEntity<String> convertHtml() {
		onlyOfficeService.convertHtml();
		return ResponseEntity.ok("response");
	}

	@GetMapping("/download/{key}/{templateId}/{refId}")
	public ResponseEntity<Resource> downloadByKey(@PathVariable String key, @PathVariable String templateId,
			@PathVariable String refId) {
		System.out.println(">>> here");
		return onlyOfficeService.downloadByKey(key, templateId, refId);
	}

//	only for pc html to doc conversion
	@PostMapping("/download/{key}/{templateId}/{refId}")
	public ResponseEntity<Resource> downloadByKey(@PathVariable String key, @PathVariable String templateId,
			@PathVariable String refId, @RequestBody String content) {
		return onlyOfficeService.downloadByKeyPc(key, templateId, refId, content);
	}

	@PostMapping("/callback")
	public ResponseEntity<Map<String, Object>> handleCallback(@RequestParam("dsp2Key") String dsp2Key,
			@RequestParam("applicationId") String applicationId, @RequestBody Map<String, Object> payload) {
		return onlyOfficeService.handleCallback(dsp2Key, applicationId, payload);
	}

	@GetMapping("/download-empty-docx/{fileName}/{dsp2Key}")
	public ResponseEntity<byte[]> downloadEmptyDocx(@PathVariable String fileName, @PathVariable String dsp2Key)
			throws Exception {
		// Create an empty DOCX document
		// XWPFDocument document = new XWPFDocument();
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        document.write(out);
//        document.close();
		Map<String, String> map = new HashMap<>();
		ByteArrayResource resource;
		if (fileName.equalsIgnoreCase("Translator-translated.docx")) {
			resource = Utils.replacePlaceholders("translation/blank/blank.docx", map);
		} else {
			resource = Utils.replacePlaceholders("blank/blank.docx", map);
		}
//        byte[] docBytes = out.toByteArray();
//        ByteArrayResource mergedResource = new ByteArrayResource(docBytes);
		thirdPartyDocGen.saveOpenKm(dsp2Key, resource);
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		headers.setContentDispositionFormData("attachment", fileName);

		return ResponseEntity.ok().headers(headers).body(resource.getByteArray());
	}

	@GetMapping("/download-empty-test/test")
	public ResponseEntity<byte[]> downloadEmptyDocxTest() throws Exception {
		// Create an empty DOCX document
		Map<String, String> replacements = Map.of("range", "2024-25", "date", "29-08-2025", "myno", "123", "branch",
				"Research Branch", "subject", "Parliamentary Updates", "contain",
				"This is the dynamic content inserted into the template.");

		ByteArrayResource result = LibreOfficeDocxReplacer.replaceWithLibreOffice("/doc-template/bck/header.docx",
				replacements);
		ByteArrayResource result1 = LibreOfficeDocxReplacer.replaceWithLibreOffice("/doc-template/bck/contain.docx",
				replacements);
		ByteArrayResource result2 = LibreOfficeDocxReplacer.replaceWithLibreOffice("/doc-template/bck/footer.docx",
				replacements);
		byte[] bb = Utils
				.mergeDocxFiles(List.of(result.getByteArray(), result1.getByteArray(), result2.getByteArray()));

		return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=generated.docx")
				.contentType(MediaType.APPLICATION_OCTET_STREAM).contentLength(bb.length).body(bb);
	}

	@GetMapping("/validate-content")
	public ResponseEntity<?> validateDocContent(@RequestParam String dsp2Key) {
		Boolean data = onlyOfficeService.convertDocxToText(dsp2Key);
		if (data == true) {
			GenericResponse<String> response = new GenericResponse<>(HttpStatus.OK.value(), true,
					"Data saved Successfully", null);
			return ResponseEntity.ok(response);
		} else {
			GenericResponse<String> response = new GenericResponse<>(HttpStatus.OK.value(), false, "Data not Saved",
					null);
			return ResponseEntity.ok(response);
		}
	}

	@PostMapping(value = "/convert-html-to-docx", produces = "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	public ResponseEntity<byte[]> convertHtmlToDocx(@RequestBody String htmlContent) {

		try {
			byte[] docBytes = onlyOfficeService.convertHtmlToDocx(htmlContent);

			return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted.docx")
					.body(docBytes);

		} catch (Exception e) {
			throw new RuntimeException("HTML → DOCX conversion failed: " + e.getMessage(), e);
		}
	}
}
