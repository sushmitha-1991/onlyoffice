/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.exception;

import java.util.HashMap;
import java.util.Map;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.FieldError;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;

import com.dsp.OnlyOfficeBridge.response.GenericResponse;

@ControllerAdvice
public class GlobalExceptionHandler {

	@ExceptionHandler(ResourceNotFoundException.class)
	public ResponseEntity<GenericResponse<Void>> handleResourceNotFoundException(ResourceNotFoundException ex) {
		GenericResponse<Void> response = new GenericResponse<>(HttpStatus.NOT_FOUND.value(), false, ex.getMessage(),
				null);
		return new ResponseEntity<>(response, HttpStatus.NOT_FOUND);
	}

	@ExceptionHandler(MethodArgumentNotValidException.class)
	public ResponseEntity<GenericResponse<Map<String, String>>> handleValidationExceptions(
			MethodArgumentNotValidException ex) {
		Map<String, String> errors = new HashMap<>();
		for (FieldError error : ex.getBindingResult().getFieldErrors()) {
			errors.put(error.getField(), error.getDefaultMessage());
		}

		GenericResponse<Map<String, String>> response = new GenericResponse<>(HttpStatus.BAD_REQUEST.value(), false,
				"Validation failed", errors);

		return new ResponseEntity<>(response, HttpStatus.BAD_REQUEST);
	}

	@ExceptionHandler(UserNotFoundException.class)
	public ResponseEntity<GenericResponse<?>> handleUserNotFoundException(UserNotFoundException ex) {
		// Create the generic response
		GenericResponse<?> response = new GenericResponse<>(HttpStatus.NOT_FOUND.value(), false, ex.getMessage(), null);
		return new ResponseEntity<>(response, HttpStatus.NOT_FOUND);
	}

	@ExceptionHandler(DatabaseException.class)
	public ResponseEntity<GenericResponse<Void>> handleDatabaseException(DatabaseException ex) {
		GenericResponse<Void> response = new GenericResponse<>(HttpStatus.INTERNAL_SERVER_ERROR.value(), false,
				"Database error: " + ex.getMessage(), null);
		return new ResponseEntity<>(response, HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@ExceptionHandler(DataProcessingException.class)
	public ResponseEntity<GenericResponse<Void>> handleDataProcessingException(DataProcessingException ex) {
		GenericResponse<Void> response = new GenericResponse<>(HttpStatus.INTERNAL_SERVER_ERROR.value(), false,
				"Data processing error: " + ex.getMessage(), null);
		return new ResponseEntity<>(response, HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@ExceptionHandler(Exception.class)
	public ResponseEntity<GenericResponse<Void>> handleGeneralException(Exception ex) {
		return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(new GenericResponse<>(
				HttpStatus.INTERNAL_SERVER_ERROR.value(), false, "An unexpected error occurred.", null));
	}

}
