/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.repository;

import com.dsp.OnlyOfficeBridge.entity.Acronym;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;
import java.util.Set;

@Repository
public interface AcronymRepository extends JpaRepository<Acronym, Long> {

	Optional<Acronym> findByIdAndIsActive(Long id, boolean isActive);

	Optional<List<Acronym>> findByIsActive(boolean isActive);

    Optional<List<Acronym>> findByIsActiveAndHouseId(boolean b, Long houseId);

    @Query("SELECT a.abbr FROM Acronym a WHERE a.isActive = true")
    Optional<Set<String>> findActiveAbbreviations();
    Optional<Acronym> findByAbbrAndIsActive(String abbr, boolean isActive);

}
