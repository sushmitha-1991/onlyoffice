package com.dsp.OnlyOfficeBridge.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import com.dsp.OnlyOfficeBridge.entity.StatementOfMatter377ZeroHour;


@Repository
public interface Statement377Repository extends JpaRepository<StatementOfMatter377ZeroHour, Long> {

    @Query(
        value = "SELECT html_content FROM notices.statement_of_matter_377_zerohour WHERE id = :id AND is_deleted = false AND is_active = true",
        nativeQuery = true
    )
    Optional<String> findHtmlById(@Param("id") Long id);
}

