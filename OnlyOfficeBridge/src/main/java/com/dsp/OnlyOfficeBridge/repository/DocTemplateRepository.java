package com.dsp.OnlyOfficeBridge.repository;

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface DocTemplateRepository extends JpaRepository<DocTemplate, Long> {
    List<DocTemplate> findByIdOrChieldIdOrderBySeq(Long id,Long chieldId);
}
