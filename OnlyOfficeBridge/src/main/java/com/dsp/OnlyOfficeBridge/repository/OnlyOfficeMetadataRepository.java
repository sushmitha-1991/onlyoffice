/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
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
	
	@Query(value="select id FROM corrected_debate.history where allocation_input_id=:refId order by created_at desc limit 1",nativeQuery=true)
	Long getHistoryId(String refId);
	
	@Query(value="select lok_sabha_number from master.lok_sabha_masters where id=:id",nativeQuery=true)
	String getLokSabhaNo(Long id);
	
	@Query(value="select session_number from master.lok_sabha_session_masters where id=:id",nativeQuery=true)
	String getLokSabhaSessionNo(Long id);

}