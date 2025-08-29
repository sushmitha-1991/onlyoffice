package com.dsp.OnlyOfficeBridge.entity;


import jakarta.persistence.*;
import lombok.Data;

@Entity
@Table(name = "doc_template", schema = "onlyoffice")
@Data
public class DocTemplate {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    @Column(name = "chield_id")
    private  Long chieldId;
    @Column(name = "type")
    private  String type;
    @Column(name = "place_holders")
    private  String placeHolders;
    @Column(name = "doc_uuid")
    private  String docUuid;
    @Column(name = "save_location")
    private  String saveLocation;
    @Column(name = "code")
    private  String code;
    @Column(name = "place_holders_type")
    private  String placeHoldersType;
    @Column(name = "seq")
    private  Long seq;


}
