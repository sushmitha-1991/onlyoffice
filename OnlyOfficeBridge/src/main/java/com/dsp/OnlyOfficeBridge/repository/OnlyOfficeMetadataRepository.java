/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;

/**
 * 
 */
@Repository
public interface OnlyOfficeMetadataRepository extends JpaRepository<OnlyOfficeMetadata, Long> {

	/**
	 * @param documentKey
	 * @return
	 */
	Optional<OnlyOfficeMetadata> findByDocumentKey(String documentKey);

	/**
	 * @param key
	 * @return
	 */
	Optional<OnlyOfficeMetadata> findByDsp2Key(String key);

}
