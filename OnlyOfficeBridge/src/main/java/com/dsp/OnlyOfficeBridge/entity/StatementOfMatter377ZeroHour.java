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
@Table(name = "statement_of_matter_377_zerohour", schema = "notices")
public class StatementOfMatter377ZeroHour {

    @Id
    private Long id;

    @Column(name = "html_content", columnDefinition = "text")
    private String htmlContent;

}