/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.exception;

/**
 * 
 */
public class DuplicateActiveMemberHistoryException extends RuntimeException {
	public DuplicateActiveMemberHistoryException(Long memberId) {
		super("Only one active (reasonToVacate=null) record is allowed for memberId: " + memberId);
	}
}
