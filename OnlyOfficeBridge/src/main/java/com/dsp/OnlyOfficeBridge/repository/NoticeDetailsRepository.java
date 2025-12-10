package com.dsp.OnlyOfficeBridge.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import com.dsp.OnlyOfficeBridge.entity.NoticeDetails;

@Repository
public interface NoticeDetailsRepository extends JpaRepository<NoticeDetails, Long> {

	@Query(value = "SELECT nd.html_content FROM notices.notice_details nd WHERE nd.id = :id", nativeQuery = true)
	Optional<String> findHtmlById(@Param("id") Long id);

}
