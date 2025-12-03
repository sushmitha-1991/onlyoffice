package com.dsp.OnlyOfficeBridge.utils;



import com.dsp.OnlyOfficeBridge.dto.MargeFileDto;
import jakarta.xml.bind.JAXBElement;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.*;

import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;

import org.docx4j.XmlUtils;
import org.docx4j.wml.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.util.*;

/**
 * FINAL continuous DOCX merger for docx4j-core 11.5.4 (Jakarta)
 */
public class DocumentMergerContinuous {

    // Relationship URIs (Jakarta version – no RelationshipTypes class)
    private static final String RT_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    private static final String RT_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    private static final String RT_COMMENTS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    // =================================================================
    // PUBLIC MERGE METHOD
    // =================================================================
    public static byte[] mergeDocuments(List<MargeFileDto> docs) throws Exception {

        if (docs == null || docs.isEmpty())
            throw new IllegalArgumentException("No documents to merge");

        WordprocessingMLPackage target = WordprocessingMLPackage.load(
                new ByteArrayInputStream(docs.get(0).getFile())
        );

        ensureParts(target);

        for (int i = 1; i < docs.size(); i++) {

            WordprocessingMLPackage src = WordprocessingMLPackage.load(
                    new ByteArrayInputStream(docs.get(i).getFile())
            );

            ensureParts(src);

            mergeStyles(target, src);

            Map<BigInteger, BigInteger> absMap = new HashMap<>();
            Map<BigInteger, BigInteger> numMap = new HashMap<>();
            mergeNumbering(target, src, absMap, numMap);

            Map<BigInteger, BigInteger> footMap = new HashMap<>();
            mergeFootnotes(target, src, footMap);

            Map<BigInteger, BigInteger> endMap = new HashMap<>();
            mergeEndnotes(target, src, endMap);

            Map<BigInteger, BigInteger> commentMap = new HashMap<>();
            mergeComments(target, src, commentMap);

            mergeBody(target, src, absMap, numMap, footMap, endMap, commentMap);
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        target.save(baos);
        return baos.toByteArray();
    }


    // =================================================================
    // ENSURE REQUIRED PARTS EXIST on docx
    // =================================================================
    private static void ensureParts(WordprocessingMLPackage pkg) throws Exception {

        MainDocumentPart mdp = pkg.getMainDocumentPart();

        // Styles
        if (mdp.getStyleDefinitionsPart() == null) {
            StyleDefinitionsPart sdp = new StyleDefinitionsPart();
            mdp.addTargetPart(sdp);
            sdp.setContents(new Styles());
        }

        // Numbering
        if (mdp.getNumberingDefinitionsPart() == null) {
            NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
            mdp.addTargetPart(ndp);
            ndp.setJaxbElement(new Numbering());
        }

        // Footnotes
        if (getFootnotesPart(mdp) == null) {
            FootnotesPart fn = new FootnotesPart();
            mdp.addTargetPart(fn);
            fn.setContents(new CTFootnotes());
        }

        // Endnotes
        if (getEndnotesPart(mdp) == null) {
            EndnotesPart en = new EndnotesPart();
            mdp.addTargetPart(en);
            en.setContents(new CTEndnotes());
        }

        // Comments
        if (getCommentsPart(mdp) == null) {
            CommentsPart cp = new CommentsPart();
            mdp.addTargetPart(cp);
            cp.setContents(new Comments());
        }
    }


    // =================================================================
    // GET PARTS using RELATIONSHIP TYPE STRING URIs
    // =================================================================
    private static FootnotesPart getFootnotesPart(MainDocumentPart mdp) {
        RelationshipsPart rels = mdp.getRelationshipsPart();
        for (Relationship r : rels.getRelationships().getRelationship()) {
            if (RT_FOOTNOTES.equals(r.getType())) {
                return (FootnotesPart) rels.getPart(r);
            }
        }
        return null;
    }

    private static EndnotesPart getEndnotesPart(MainDocumentPart mdp) {
        RelationshipsPart rels = mdp.getRelationshipsPart();
        for (Relationship r : rels.getRelationships().getRelationship()) {
            if (RT_ENDNOTES.equals(r.getType())) {
                return (EndnotesPart) rels.getPart(r);
            }
        }
        return null;
    }

    private static CommentsPart getCommentsPart(MainDocumentPart mdp) {
        RelationshipsPart rels = mdp.getRelationshipsPart();
        for (Relationship r : rels.getRelationships().getRelationship()) {
            if (RT_COMMENTS.equals(r.getType())) {
                return (CommentsPart) rels.getPart(r);
            }
        }
        return null;
    }


    // =================================================================
    // STYLES MERGE
    // =================================================================
    private static void mergeStyles(WordprocessingMLPackage tgt,
                                    WordprocessingMLPackage src) throws Exception {

        StyleDefinitionsPart t = tgt.getMainDocumentPart().getStyleDefinitionsPart();
        StyleDefinitionsPart s = src.getMainDocumentPart().getStyleDefinitionsPart();
        if (s == null || s.getContents() == null) return;

        for (Style st : s.getContents().getStyle()) {
            if (st.getStyleId() != null && t.getStyleById(st.getStyleId()) == null) {
                t.getContents().getStyle().add(XmlUtils.deepCopy(st));
            }
        }
    }


    // =================================================================
    // NUMBERING MERGE
    // =================================================================
    private static void mergeNumbering(WordprocessingMLPackage tgt,
                                       WordprocessingMLPackage src,
                                       Map<BigInteger, BigInteger> absMap,
                                       Map<BigInteger, BigInteger> numMap)
            throws Exception {

        NumberingDefinitionsPart t = tgt.getMainDocumentPart().getNumberingDefinitionsPart();
        NumberingDefinitionsPart s = src.getMainDocumentPart().getNumberingDefinitionsPart();
        if (s == null) return;

        Numbering tn = t.getJaxbElement();
        Numbering sn = s.getJaxbElement();

        int maxAbs = tn.getAbstractNum().size();
        int maxNum = tn.getNum().size();

        for (Numbering.AbstractNum an : sn.getAbstractNum()) {
            Numbering.AbstractNum copy = XmlUtils.deepCopy(an);
            BigInteger newId = BigInteger.valueOf(++maxAbs);
            absMap.put(an.getAbstractNumId(), newId);
            copy.setAbstractNumId(newId);
            tn.getAbstractNum().add(copy);
        }

        for (Numbering.Num nm : sn.getNum()) {
            Numbering.Num copy = XmlUtils.deepCopy(nm);
            BigInteger newNumId = BigInteger.valueOf(++maxNum);
            numMap.put(nm.getNumId(), newNumId);
            copy.setNumId(newNumId);

            BigInteger oldAbs = nm.getAbstractNumId().getVal();
            if (absMap.containsKey(oldAbs)) {
                copy.getAbstractNumId().setVal(absMap.get(oldAbs));
            }

            tn.getNum().add(copy);
        }
    }


    // =================================================================
    // FOOTNOTES MERGE (using CTFootnotes)
    // =================================================================
    private static void mergeFootnotes(WordprocessingMLPackage tgt,
                                       WordprocessingMLPackage src,
                                       Map<BigInteger, BigInteger> map) throws Exception {

        FootnotesPart t = getFootnotesPart(tgt.getMainDocumentPart());
        FootnotesPart s = getFootnotesPart(src.getMainDocumentPart());
        if (s == null) return;

        CTFootnotes tf = t.getContents();
        CTFootnotes sf = s.getContents();

        int maxId = 0;
        for (CTFtnEdn f : tf.getFootnote()) {
            maxId = Math.max(maxId, f.getId().intValue());
        }

        for (CTFtnEdn f : sf.getFootnote()) {
            CTFtnEdn copy = XmlUtils.deepCopy(f);
            BigInteger newId = BigInteger.valueOf(++maxId);
            map.put(f.getId(), newId);
            copy.setId(newId);
            tf.getFootnote().add(copy);
        }
    }


    // =================================================================
    // ENDNOTES MERGE (using CTEndnotes)
    // =================================================================
    private static void mergeEndnotes(WordprocessingMLPackage tgt,
                                      WordprocessingMLPackage src,
                                      Map<BigInteger, BigInteger> map) throws Exception {

        EndnotesPart t = getEndnotesPart(tgt.getMainDocumentPart());
        EndnotesPart s = getEndnotesPart(src.getMainDocumentPart());
        if (s == null) return;

        CTEndnotes te = t.getContents();
        CTEndnotes se = s.getContents();

        int maxId = 0;
        for (CTFtnEdn e : te.getEndnote()) {
            maxId = Math.max(maxId, e.getId().intValue());
        }

        for (CTFtnEdn e : se.getEndnote()) {
            CTFtnEdn copy = XmlUtils.deepCopy(e);
            BigInteger newId = BigInteger.valueOf(++maxId);
            map.put(e.getId(), newId);
            copy.setId(newId);
            te.getEndnote().add(copy);
        }
    }


    // =================================================================
    // COMMENTS MERGE
    // =================================================================
    private static void mergeComments(WordprocessingMLPackage tgt,
                                      WordprocessingMLPackage src,
                                      Map<BigInteger, BigInteger> map) throws Exception {

        CommentsPart t = getCommentsPart(tgt.getMainDocumentPart());
        CommentsPart s = getCommentsPart(src.getMainDocumentPart());
        if (s == null) return;

        Comments tc = t.getContents();
        Comments sc = s.getContents();

        int maxId = 0;
        for (Comments.Comment c : tc.getComment()) {
            maxId = Math.max(maxId, c.getId().intValue());
        }

        for (Comments.Comment c : sc.getComment()) {
            Comments.Comment copy = XmlUtils.deepCopy(c);
            BigInteger newId = BigInteger.valueOf(++maxId);
            map.put(c.getId(), newId);
            copy.setId(newId);
            tc.getComment().add(copy);
        }
    }


    // =================================================================
    // BODY MERGE (continuous)
    // =================================================================
    private static void mergeBody(WordprocessingMLPackage tgt,
                                  WordprocessingMLPackage src,
                                  Map<BigInteger, BigInteger> absMap,
                                  Map<BigInteger, BigInteger> numMap,
                                  Map<BigInteger, BigInteger> footMap,
                                  Map<BigInteger, BigInteger> endMap,
                                  Map<BigInteger, BigInteger> commentMap)
            throws Exception {

        MainDocumentPart t = tgt.getMainDocumentPart();
        MainDocumentPart s = src.getMainDocumentPart();

        for (Object o : s.getContent()) {

            String xml = XmlUtils.marshaltoString(o, true, true);

            for (BigInteger oldId : numMap.keySet()) {
                xml = xml.replace("w:numId w:val=\"" + oldId + "\"",
                        "w:numId w:val=\"" + numMap.get(oldId) + "\"");
            }

            for (BigInteger oldId : absMap.keySet()) {
                xml = xml.replace("w:abstractNumId w:val=\"" + oldId + "\"",
                        "w:abstractNumId w:val=\"" + absMap.get(oldId) + "\"");
            }

            for (BigInteger oldId : footMap.keySet()) {
                xml = xml.replace("w:footnoteReference w:id=\"" + oldId + "\"",
                        "w:footnoteReference w:id=\"" + footMap.get(oldId) + "\"");
            }

            for (BigInteger oldId : endMap.keySet()) {
                xml = xml.replace("w:endnoteReference w:id=\"" + oldId + "\"",
                        "w:endnoteReference w:id=\"" + endMap.get(oldId) + "\"");
            }

            for (BigInteger oldId : commentMap.keySet()) {
                xml = xml.replace("w:commentReference w:id=\"" + oldId + "\"",
                        "w:commentReference w:id=\"" + commentMap.get(oldId) + "\"");
            }

            // remove section breaks for continuous merge
            xml = xml.replaceAll("(?s)<w:sectPr.*?</w:sectPr>", "");

            Object unmarshalled = XmlUtils.unmarshallFromTemplate(xml, new HashMap<>());
            t.addObject(unmarshalled);
        }
    }
}
