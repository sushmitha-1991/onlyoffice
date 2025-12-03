/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.entity;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
@Table(name = "acronym_repo", schema = "verbatim")
public class Acronym extends Auditable {

	@Id
	@GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "acronym_seq")
	@SequenceGenerator(name = "acronym_seq", sequenceName = "acronym_seq", allocationSize = 1)
	private Long id;

	@Column(name = "abbr", nullable = false, length = 30)
	private String abbr;

	@Column(name = "full_text", nullable = false, length = 4098)
	private String fullText;

	@Column(name="house_id")
	private Long houseId;
}
