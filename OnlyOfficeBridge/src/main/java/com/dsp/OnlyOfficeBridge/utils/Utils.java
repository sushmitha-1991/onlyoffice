/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.FileSystemResource;
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
			String downloadUrl = "https://dspstg.sansad.in/dms/openkm/dsp1/downloads/" + uuid;
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
			String uploadUrl = "https://dspstg.sansad.in/dms/openkm/dsp1/document/upload";
			System.out.println(">>> uploadUrl - " + uploadUrl);
			ResponseEntity<String> response = restTemplate.exchange(uploadUrl, HttpMethod.POST, requestEntity,
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

}
