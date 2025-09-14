/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.utils;

import java.io.*;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlOptions;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
//import org.jodconverter.local.LocalConverter;
//import org.jodconverter.local.office.LocalOfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
//import org.jodconverter.remote.RemoteConverter;
//import org.jodconverter.remote.office.RemoteOfficeManager;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.RestTemplate;

import com.dsp.OnlyOfficeBridge.dto.UserDetailsDTO;
import com.dsp.OnlyOfficeBridge.entity.BaseUser;
import com.dsp.OnlyOfficeBridge.repository.MyUserRepository;

import jakarta.persistence.EntityManager;
import jakarta.persistence.Query;
import jakarta.servlet.http.HttpServletRequest;

@Component
public class Utils {

	@Autowired 
	private MyUserRepository myUserRepository;

	@Autowired
	private EntityManager entityManager;

	@Autowired
	private HttpServletRequest httpRequest;

	@Autowired
	private RestTemplate restTemplate;

	@Value("${download.url}")
	String downloadOkmUrl;
	
	@Value("${upload.url}")
	String uploadOkmUrl;
	
	/**
	 * Fetches the user ID of the currently authenticated user
	 * 
	 * @return userId
	 */
	public Long getUserIdFromAuthentication() {
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		if (authentication == null || !authentication.isAuthenticated()) {
			throw new IllegalStateException("No authenticated user found");
		}

		String username = authentication.getName();
		BaseUser baseUser = myUserRepository.findByUsername(username);
		if (baseUser == null) {
			throw new UsernameNotFoundException("Authenticated user not found in the database");
		}

		return baseUser.getId();
	}

	/**
	 * Fetches the user ID of the currently authenticated user
	 * 
	 * @return userId
	 */
	public String getUserFullNameFromAuthentication() {
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		if (authentication == null || !authentication.isAuthenticated()) {
			throw new IllegalStateException("No authenticated user found");
		}

		String username = authentication.getName();
		BaseUser baseUser = myUserRepository.findByUsername(username);
		if (baseUser == null) {
			throw new UsernameNotFoundException("Authenticated user not found in the database");
		}

		// Concatenate firstName, middleName, and lastName
		String firstName = baseUser.getFirstName();
		String middleName = baseUser.getMiddleName() != null ? baseUser.getMiddleName() : ""; // Handle null middleName
		String lastName = baseUser.getLastName();

		// Construct full name by trimming any extra spaces
		return String.join(" ", firstName, middleName, lastName).trim();
	}

	public Long getHouseIdFromAuthentication() {
		Long userId = getUserIdFromAuthentication(); // Get the authenticated user's ID

		// Define the native query
		String nativeQuery = "SELECT u.house_id FROM authentication.base_user u " + "WHERE u.id = :userId";

		// Execute the query
		Query query = entityManager.createNativeQuery(nativeQuery);
		query.setParameter("userId", userId);

		// Retrieve and return the house_id
		Object result = query.getSingleResult();
		if (result != null) {
			return ((Number) result).longValue(); // Cast result to Long
		} else {
			throw new IllegalStateException("House ID not found for the authenticated user.");
		}
	}

	/**
	 * Fetches the role ID of the authenticated user by joining user_roles and users
	 * table
	 * 
	 * @return roleId
	 */
	public Long getRoleIdByUserId() {
		Long userId = getUserIdFromAuthentication(); // Get the authenticated user's ID

		String nativeQuery = "SELECT ur.role_id FROM authentication.user_roles ur "
				+ "JOIN authentication.base_user u ON ur.user_id = u.id "
				+ "WHERE u.id = :userId AND ur.is_active = true AND ur.is_deleted = false";

		Query query = entityManager.createNativeQuery(nativeQuery);
		query.setParameter("userId", userId);

		try {
			// Fetch the result and ensure it's correctly cast to Long
			Object result = query.getSingleResult();

			if (result instanceof Integer) {
				// Convert Integer to Long
				return ((Integer) result).longValue();
			} else if (result instanceof Long) {

				// If it's already Long, return it directly
				return (Long) result;
			} else {
				throw new IllegalStateException("Unexpected data type for role_id: " + result.getClass());
			}
		} catch (Exception e) {
			throw new IllegalStateException("No active role found for the user with ID: " + userId, e);
		}
	}

	/**
	 * Fetches the role IDs of the authenticated user by joining user_roles and
	 * users table.
	 * 
	 * @return List<Long> - a list of role IDs
	 */
	public List<Long> getRoleIdsByUserId() {
		Long userId = getUserIdFromAuthentication(); // Get the authenticated user's ID

		String nativeQuery = "SELECT ur.role_id FROM authentication.user_roles ur "
				+ "JOIN authentication.base_user u ON ur.user_id = u.id "
				+ "WHERE u.id = :userId AND ur.is_active = true AND ur.is_deleted = false";

		Query query = entityManager.createNativeQuery(nativeQuery);
		query.setParameter("userId", userId);

		try {
			// Fetch the result list (which may contain multiple roles)
			List<Object> resultList = query.getResultList();

			// Convert the list of Object into a list of Longs
			List<Long> roleIds = new ArrayList<>();
			for (Object result : resultList) {
				if (result instanceof Integer) {
					// Convert Integer to Long
					roleIds.add(((Integer) result).longValue());
				} else if (result instanceof Long) {
					// Add Long values directly
					roleIds.add((Long) result);
				} else {
					throw new IllegalStateException("Unexpected data type for role_id: " + result.getClass());
				}
			}

			return roleIds;

		} catch (Exception e) {
			throw new IllegalStateException("No active roles found for the user with ID: " + userId, e);
		}
	}

	public Long getHouseIdFromHouseAbbrAuthentication(String houseAbbr) {
		// Define the native query
		String nativeQuery = "SELECT h.id FROM master.houses h " + "WHERE h.abbr = :houseAbbr";

		// Execute the query
		Query query = entityManager.createNativeQuery(nativeQuery);
		query.setParameter("houseAbbr", houseAbbr);

		// Retrieve and return the house_id
		Object result = query.getSingleResult();
		if (result != null) {
			return ((Number) result).longValue(); // Cast result to Long
		} else {
			throw new IllegalStateException("House ID not found for the authenticated user.");
		}
	}

	public String getOrdinalSuffix(String termStr) {
		try {
			int term = Integer.parseInt(termStr);
			if (term >= 11 && term <= 13) {
				return term + "th";
			}
			return switch (term % 10) {
			case 1 -> term + "st";
			case 2 -> term + "nd";
			case 3 -> term + "rd";
			default -> term + "th";
			};
		} catch (NumberFormatException e) {
			return termStr; // fallback if not a number
		}
	}

	public String getHouseAbbr() {
		Object principal = SecurityContextHolder.getContext().getAuthentication().getPrincipal();
		if (!(principal instanceof UserDetailsDTO udto)) {
			throw new RuntimeException("Authentication Problems");
		}
		return udto.getHouseAbbr();
	}

	public static String toCamelCase(String input) {

		String[] parts = input.split("\\s+");
		StringBuilder camelCaseString = new StringBuilder(parts[0].toLowerCase());
		for (int i = 1; i < parts.length; i++) {
			camelCaseString.append(parts[i].substring(0, 1).toUpperCase()).append(parts[i].substring(1).toLowerCase());
		}
		return camelCaseString.toString();
	}

	public ByteArrayResource downloadFileFromUUID(UUID uuid) {
		try {
			String downloadUrl = downloadOkmUrl + uuid;
			
			RestTemplate restTemplate = new RestTemplate();

			HttpHeaders headers = new HttpHeaders();
			HttpEntity<Void> entity = new HttpEntity<>(headers);

			ResponseEntity<byte[]> response = restTemplate.exchange(downloadUrl, HttpMethod.GET, entity, byte[].class);

			if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
				String filename = uuid.toString(); // Default filename

				List<String> contentDisposition = response.getHeaders().get(HttpHeaders.CONTENT_DISPOSITION);
				if (contentDisposition != null && !contentDisposition.isEmpty()) {
					// Example: Content-Disposition: attachment; filename="Document.docx"
					String disposition = contentDisposition.get(0);
					int start = disposition.indexOf("filename=");
					if (start != -1) {
						filename = disposition.substring(start + 9).replace("\"", "").trim();
					}
				}

				final String finalFilename = filename;
				return new ByteArrayResource(response.getBody()) {
					@Override
					public String getFilename() {
						return finalFilename;
					}
				};
			}
		} catch (Exception e) {
			e.printStackTrace(); // Use proper logging
		}

		return null;
	}

	public String uploadPdfToOpenKM(File pdfFile, String documentName, String folderPath, String masterId,
			String referenceId, String referenceType, String applicationNo) {
		try {
			System.out.println(">>> calling uploadPdfToOpenKM -");
			System.out.println(">>> documentName - " + documentName);
			System.out.println(">>> folderPath - " + folderPath);
			System.out.println(">>> masterId - " + masterId);
			System.out.println(">>> referenceId - " + referenceId);
			System.out.println(">>> referenceType - " + referenceType);
			System.out.println(">>> applicationNo - " + applicationNo);

			// The JWT token is automatically added by the RestTemplate interceptor
			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.MULTIPART_FORM_DATA);

			// Create a FileSystemResource from the file
			FileSystemResource fileResource = new FileSystemResource(pdfFile);

			// Prepare the body with all required fields
			MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
			body.add("folderPath", folderPath);
			body.add("documentName", documentName);
			body.add("file", fileResource);
			body.add("masterId", masterId);
			body.add("referenceId", referenceId);
			body.add("referenceType", referenceType);
			body.add("applicationNo", applicationNo);

			System.out.println(">>> inside");

			// Prepare the request entity with headers and body
			HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);
			System.out.println(requestEntity);
//			String uploadUrl = "https://dspstg.sansad.in/dms/openkm/dsp1/document/upload";
			System.out.println(">>> uploadUrl - " + uploadOkmUrl);
			ResponseEntity<String> response = restTemplate.exchange(uploadOkmUrl, HttpMethod.POST, requestEntity,
					String.class);

			// Return the response body as a string
			return response.getBody();
		} catch (HttpClientErrorException e) {
			e.printStackTrace();
			throw new RuntimeException("Error uploading document to OpenKM. HTTP Error: " + e.getMessage());
		} catch (Exception e) {
			throw new RuntimeException("Error uploading document to OpenKM.", e);
		}
	}
    public ByteArrayResource replacePlaceholders(ByteArrayResource inputResource,
                                                        Map<String, String> replacements)  {

        try (ByteArrayInputStream bis = new ByteArrayInputStream(inputResource.getByteArray());
             XWPFDocument document = new XWPFDocument(bis)) {

            // Replace in paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (XWPFRun run : runs) {
                        String text = run.getText(0);
                        if (text != null) {
                            for (Map.Entry<String, String> entry : replacements.entrySet()) {
                                if (text.contains(entry.getKey())) {
                                    text = text.replace(entry.getKey(), entry.getValue());
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }

            // Replace in tables also (if placeholders exist inside tables)
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            List<XWPFRun> runs = paragraph.getRuns();
                            if (runs != null) {
                                for (XWPFRun run : runs) {
                                    String text = run.getText(0);
                                    if (text != null) {
                                        for (Map.Entry<String, String> entry : replacements.entrySet()) {
                                            if (text.contains(entry.getKey())) {
                                                text = text.replace(entry.getKey(), entry.getValue());
                                                run.setText(text, 0);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Write to ByteArrayOutputStream
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                document.write(bos);
                return new ByteArrayResource(bos.toByteArray());
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public ByteArrayResource getFileAsResource(String filename) throws IOException {
        // Load file from resources/doc-template
        Resource resource = new ClassPathResource("doc-template/" + filename);

        // Convert to ByteArrayResource
        byte[] fileBytes = Files.readAllBytes(resource.getFile().toPath());

        return new ByteArrayResource(fileBytes) {
            @Override
            public String getFilename() {
                return filename; // preserve filename
            }
        };
    }
    public static ByteArrayResource replacePlaceholders(String resourcePath, Map<String, String> replacements) throws Exception {

        // Load file from resources
        ClassPathResource resource = new ClassPathResource("doc-template/" + resourcePath);

        // Copy resource to a temp file (LibreOffice requires a real File, not just InputStream)
        File inputFile = File.createTempFile("template", ".docx");
        try (InputStream in = resource.getInputStream();
             OutputStream out = new FileOutputStream(inputFile)) {
            in.transferTo(out);
        }

        File tempFile = File.createTempFile("converted", ".fodt");

        LocalOfficeManager officeManager = LocalOfficeManager.install();
        officeManager.start();
//        RemoteOfficeManager officeManager = RemoteOfficeManager.builder()
//                // If your LibreOffice server runs in socket mode:
//                //.urlConnection("socket://<REMOTE_SERVER_IP>:2002")
//
//                // If your LibreOffice server exposes REST API:
//                .urlConnection("http://10.247.128.22:2002")
//                .build();

        try {
            // Convert DOCX → FODT
            LocalConverter.builder()
                    .officeManager(officeManager)
                    .build()
                    .convert(inputFile)
                    .to(tempFile)
                    .as(DefaultDocumentFormatRegistry.FODT)
                    .execute();

            // Read + replace placeholders in FODT XML
            String xml = new String(java.nio.file.Files.readAllBytes(tempFile.toPath()), StandardCharsets.UTF_8);

            for (Map.Entry<String, String> entry : replacements.entrySet()) {
//                String placeholder = "\\{" + entry.getKey() + "\\}";
//                xml = xml.replaceAll(placeholder, entry.getValue());
                String placeholder = "{" + entry.getKey() + "}";
                String safeValue = StringEscapeUtils.escapeXml11(entry.getValue()); // escapes &, <, >, ", '
                xml = xml.replace(placeholder, safeValue);
            }

            // Write modified FODT back
            java.nio.file.Files.write(tempFile.toPath(), xml.getBytes(StandardCharsets.UTF_8));

            // Convert FODT → DOCX into memory
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            LocalConverter.builder()
                    .officeManager(officeManager)
                    .build()
                    .convert(tempFile)
                    .to(baos)
                    .as(DefaultDocumentFormatRegistry.DOCX)
                    .execute();

            return new ByteArrayResource(baos.toByteArray());

        } finally {
            officeManager.stop();
            inputFile.delete(); // cleanup temp files
            tempFile.delete();
        }
    }

//    public static ByteArrayResource replacePlaceholders(String resourcePath, Map<String, String> replacements) throws Exception {
//        ClassPathResource resource = new ClassPathResource("doc-template/" + resourcePath);
//
//        try (InputStream fis = resource.getInputStream();
//             XWPFDocument document = new XWPFDocument(fis)) {
//
//            for (XWPFParagraph paragraph : document.getParagraphs()) {
//                List<XWPFRun> runs = paragraph.getRuns();
//                if (runs != null) {
//                    for (XWPFRun run : runs) {
//                        String text = run.getText(0);
//                        if (text != null) {
//                            for (Map.Entry<String, String> entry : replacements.entrySet()) {
//                                if (text.contains(entry.getKey())) {
//                                    text = text.replace(entry.getKey(), entry.getValue());
//                                    run.setText(text, 0);
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            // Return as ByteArrayResource
//            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
//                document.write(bos);
//                return new ByteArrayResource(bos.toByteArray());
//            }
//        }
//    }

    public static byte[] mergeDocxFiles(List<byte[]> docxFiles) throws Exception {
        try (XWPFDocument mergedDoc = new XWPFDocument()) {
        	
            boolean first = true;

            for (int i = 0; i < docxFiles.size(); i++) {
                byte[] fileBytes = docxFiles.get(i);
                try (ByteArrayInputStream bais = new ByteArrayInputStream(fileBytes);
                     XWPFDocument doc = new XWPFDocument(bais)) {

                    CTBody srcBody = doc.getDocument().getBody();
                    CTBody destBody = mergedDoc.getDocument().getBody();

                    // Strip <w:sectPr> from all but the last document
                    CTBody cleaned = (i == docxFiles.size() - 1)
                            ? srcBody
                            : removeSectPr(srcBody);

                    if (first) {
                        destBody.set(cleaned);
                        first = false;
                    } else {
                        appendBody(destBody, cleaned);
                    }
                }
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            mergedDoc.write(baos);
            return baos.toByteArray();
        }
    }

    /**
     * Remove trailing <w:sectPr> which causes page breaks.
     */
    private static CTBody removeSectPr(CTBody body) throws Exception {
        String xml = body.xmlText();
        // Drop section properties
        xml = xml.replaceAll("(?s)<w:sectPr.*?</w:sectPr>", "");
        return CTBody.Factory.parse(xml);
    }

    /**
     * Append one document body into another.
     */
    private static void appendBody(CTBody src, CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);

        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String suffix = srcString.substring(srcString.lastIndexOf("<"));

        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));

        CTBody newBody = CTBody.Factory.parse(prefix + mainPart + addPart + suffix);
        src.set(newBody);
    }
    public String getTokenFromRequest() {
        String authorizationHeader = httpRequest.getHeader("Authorization");
        if (authorizationHeader != null && authorizationHeader.startsWith("Bearer ")) {
            return authorizationHeader.substring(7);
        }
        return null;
    }
     
	// 🔧 Helper: set default margins = 1 inch
    public void setDefaultMargins(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().isSetSectPr()
                ? doc.getDocument().getBody().getSectPr()
                : doc.getDocument().getBody().addNewSectPr();

        // --- Set page size (A4 = 210mm × 297mm) ---
        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(11906)); // A4 width in twips
        pageSize.setH(BigInteger.valueOf(16838)); // A4 height in twips
        
        CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();

        int topMargin = 1440; // twips
        int sideMargin = 1080; // twips
        pageMar.setTop(BigInteger.valueOf(topMargin));
        pageMar.setBottom(BigInteger.valueOf(topMargin));
        pageMar.setLeft(BigInteger.valueOf(sideMargin));
        pageMar.setRight(BigInteger.valueOf(sideMargin));
    }
    
}
