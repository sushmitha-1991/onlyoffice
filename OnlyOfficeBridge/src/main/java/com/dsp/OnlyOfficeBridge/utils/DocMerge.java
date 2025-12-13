package com.dsp.OnlyOfficeBridge.utils;

import com.dsp.OnlyOfficeBridge.dto.MargeFileDto;
import jakarta.xml.bind.JAXBElement;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class DocMerge {
    public static byte[] mergeDocxFiles(List<MargeFileDto> files) throws Exception {
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("No files to merge");
        }

        // Load first file as base
        WordprocessingMLPackage targetPkg = WordprocessingMLPackage
                .load(new ByteArrayInputStream(files.get(0).getFile()));
        MainDocumentPart targetMain = targetPkg.getMainDocumentPart();
        ensureNumberingPartExists(targetPkg);
        if (files.get(0).isBreak()) {
            targetMain.getContent().add(createPageBreak());
        }

        // Ensure footnotes, endnotes, comments exist
        FootnotesPart targetFootnotes = targetMain.getFootnotesPart();
        if (targetFootnotes == null) {
            targetFootnotes = new FootnotesPart();
            targetFootnotes.setContents(new CTFootnotes());
            targetMain.addTargetPart(targetFootnotes);
        }
        EndnotesPart targetEndnotes = getEndnotesPart(targetPkg);
        if (targetEndnotes == null) {
            targetEndnotes = new EndnotesPart();
            targetEndnotes.setContents(new CTEndnotes());
            targetMain.addTargetPart(targetEndnotes);
        }
        CommentsPart targetComments = targetMain.getCommentsPart();
        if (targetComments == null) {
            targetComments = new CommentsPart();
            targetComments.setContents(new Comments());
            targetMain.addTargetPart(targetComments);
        }

        // Track max IDs
        int maxFootnoteId = targetFootnotes.getContents().getFootnote().stream().map(fn -> fn.getId().intValue())
                .max(Integer::compareTo).orElse(100);
        int maxEndnoteId = targetEndnotes.getContents().getEndnote().stream().map(en -> en.getId().intValue())
                .max(Integer::compareTo).orElse(100);
        int maxCommentId = targetComments.getContents().getComment().stream().map(c -> c.getId().intValue())
                .max(Integer::compareTo).orElse(100);

        // For each source file
        for (int idx = 1; idx < files.size(); idx++) {
            WordprocessingMLPackage srcPkg = WordprocessingMLPackage
                    .load(new ByteArrayInputStream(files.get(idx).getFile()));
            try {
                if (!files.get(idx + 1).getOrientation().equalsIgnoreCase(files.get(idx).getOrientation())) {
                    files.get(idx).setBreak(true);
                }
            } catch (Exception ignored) {

            }
//            if ("LANDSCAPE".equalsIgnoreCase(files.get(idx).getOrientation())) {
//                convertToLandscape(srcPkg);
//            }
            MainDocumentPart srcMain = srcPkg.getMainDocumentPart();
            // 1) Copy images
            Map<String, String> relIdMapping = new HashMap<>();
            for (Part p : srcPkg.getParts().getParts().values()) {
                if (p instanceof BinaryPart && p.getPartName() != null
                        && p.getPartName().getName().startsWith("/word/media/")) {
                    BinaryPart srcBin = (BinaryPart) p;
                    ByteArrayOutputStream imgOut = new ByteArrayOutputStream();
                    srcBin.writeDataToOutputStream(imgOut);
                    BinaryPartAbstractImage newImgPart = BinaryPartAbstractImage.createImagePart(targetPkg,
                            imgOut.toByteArray());
                    Relationship newRel = newImgPart.getSourceRelationship();
                    if (newRel != null) {
                        for (Relationship srcRel : srcBin.getSourceRelationships()) {
                            relIdMapping.put(srcRel.getId(), newRel.getId());
                        }
                    }
                }
            }
// 2) Merge Numbering (and get mapping for numId)
            NumberingDefinitionsPart srcNumbering = srcMain.getNumberingDefinitionsPart();
            DocMerge.IdMaps numberingMaps = null;
            if (srcNumbering != null) {
                NumberingDefinitionsPart tgtNumbering = targetMain.getNumberingDefinitionsPart();
                numberingMaps = mergeNumbering(srcNumbering, tgtNumbering);
//                forceAbstractLevelStartOne(tgtNumbering);
                // normalizeNumbering(tgtNumbering);
            } else {

                numberingMaps = new DocMerge.IdMaps(); // empty maps
            }
            // 2) Merge Footnotes
            Map<BigInteger, BigInteger> footnoteIdMap = new LinkedHashMap<>();
            FootnotesPart srcFootnotes = srcMain.getFootnotesPart();
            if (srcFootnotes != null && srcFootnotes.getContents() != null) {
                for (CTFtnEdn fn : srcFootnotes.getContents().getFootnote()) {
                    BigInteger oldId = fn.getId();
                    BigInteger newId = BigInteger.valueOf(++maxFootnoteId);
                    fn.setId(newId);
                    footnoteIdMap.put(oldId, newId);
                }
                targetFootnotes.getContents().getFootnote().addAll(srcFootnotes.getContents().getFootnote());
            }

            // 3) Merge Endnotes
            Map<BigInteger, BigInteger> endnoteIdMap = new LinkedHashMap<>();
            EndnotesPart srcEndnotes = getEndnotesPart(srcPkg);
            if (srcEndnotes != null && srcEndnotes.getContents() != null) {
                for (CTFtnEdn en : srcEndnotes.getContents().getEndnote()) {
                    BigInteger oldId = en.getId();
                    BigInteger newId = BigInteger.valueOf(++maxEndnoteId);
                    en.setId(newId);
                    endnoteIdMap.put(oldId, newId);
                }
                targetEndnotes.getContents().getEndnote().addAll(srcEndnotes.getContents().getEndnote());
            }

            // 4) Merge Comments
            Map<BigInteger, BigInteger> commentIdMap = new LinkedHashMap<>();
            CommentsPart srcComments = srcMain.getCommentsPart();
            if (srcComments != null && srcComments.getContents() != null) {
                for (Comments.Comment c : srcComments.getContents().getComment()) {
                    BigInteger oldId = c.getId();
                    log.debug("comment old id {}", oldId);
                    BigInteger newId = BigInteger.valueOf(++maxCommentId);
                    c.setId(newId);
                    commentIdMap.put(oldId, newId);
                }
                targetComments.getContents().getComment().addAll(srcComments.getContents().getComment());
            }
            // 5) Append body content, fixing IDs
            log.info("relIdMapping map {}", relIdMapping);
            if (srcMain.getContent() != null) {
                for (Object o : srcMain.getContent()) {
                    Object clone = XmlUtils.deepCopy(o);
                    String xml = XmlUtils.marshaltoString(clone, true);

                    // Replace relIds

                    for (Map.Entry<String, String> e : relIdMapping.entrySet()) {
                        xml = xml.replace("r:embed=\"" + e.getKey() + "\"", "r:embed=\"" + e.getValue() + "\"");
                        xml = xml.replace("r:id=\"" + e.getKey() + "\"", "r:id=\"" + e.getValue() + "\"");
                    }
                    // Replace footnote IDs

                    List<Integer> currIdFootnote = getMapKeyList(xml, "w:footnoteReference[^>]*w:id=\"(\\d+)\"");
                    for (Map.Entry<BigInteger, BigInteger> e : footnoteIdMap.entrySet()) {
                        if (currIdFootnote.contains(e.getKey().intValue())) {

                            log.debug("Before replace footnote id {} >>> {}", e.getKey(), xml);

                            xml = xml.replaceAll("<w:footnoteReference([^>]*)w:id=\"" + e.getKey() + "\"",
                                    "<w:footnoteReference$1w:id=\"" + e.getValue() + "\"");

                            log.debug("After replace id {} with {} >>> {}", e.getKey(), e.getValue(), xml);

                            currIdFootnote.remove(Integer.valueOf(e.getKey().toString()));
                        }
                    }
                    // Replace endnote IDs
                    for (Map.Entry<BigInteger, BigInteger> e : endnoteIdMap.entrySet()) {
                        xml = xml.replace("w:endnoteReference w:id=\"" + e.getKey() + "\"",
                                "w:endnoteReference w:id=\"" + e.getValue() + "\"");
                    }
                    // Replace comment IDs
                    List<Integer> currId = getMapKeyList(xml, "w:commentReference[^>]*w:id=\"(\\d+)\"");

                    for (Map.Entry<BigInteger, BigInteger> e : commentIdMap.entrySet()) {

                        if (currId.contains(e.getKey().intValue())) {
                            xml = xml.replace("w:commentReference w:id=\"" + e.getKey() + "\"",
                                    "w:commentReference w:id=\"" + e.getValue() + "\"");
                            xml = xml.replace("w:commentRangeStart w:id=\"" + e.getKey() + "\"",
                                    "w:commentRangeStart w:id=\"" + e.getValue() + "\"");
                            xml = xml.replace("w:commentRangeEnd w:id=\"" + e.getKey() + "\"",
                                    "w:commentRangeEnd w:id=\"" + e.getValue() + "\"");
                            currId.remove(Integer.getInteger(e.getKey().toString()));
                        }

                    }

                    Object fixed = XmlUtils.unmarshalString(xml, Context.jc);
                    if (numberingMaps != null && numberingMaps.numIdMap != null) {
                        remapNumIdInContent(fixed, numberingMaps.numIdMap);
                    }
                    targetMain.getContent().add(fixed);
                }
            }

            // 6) Copy headers/footers (simple best-effort)
            List<SectionWrapper> sections = srcPkg.getDocumentModel().getSections();
            if (sections != null) {
                HeaderFooterPolicy srcPolicy = sections.get(0).getHeaderFooterPolicy();
                if (srcPolicy != null) {
                    if (srcPolicy.getDefaultHeader() != null) {
                        HeaderPart clonedHeader = new HeaderPart();
                        clonedHeader.setContents(XmlUtils.deepCopy(srcPolicy.getDefaultHeader().getContents()));
                        targetMain.addTargetPart(clonedHeader);
                    }
                    if (srcPolicy.getDefaultFooter() != null) {
                        FooterPart clonedFooter = new FooterPart();
                        clonedFooter.setContents(XmlUtils.deepCopy(srcPolicy.getDefaultFooter().getContents()));
                        targetMain.addTargetPart(clonedFooter);
                    }
                }
            }

            // Optional page break
            if (files.get(idx).isBreak()) {
//                targetMain.getContent().add(createPageBreak());
                insertSectionBreak(targetMain, srcMain);
            }
        }

        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            targetPkg.save(baos);
            return baos.toByteArray();
        }
    }
    private static void ensureNumberingPartExists(WordprocessingMLPackage pkg) throws Docx4JException {

        MainDocumentPart mdp = pkg.getMainDocumentPart();

        // Check if numbering part already exists
        NumberingDefinitionsPart numberingPart = mdp.getNumberingDefinitionsPart();

        if (numberingPart == null) {
            // Create new numbering part
            numberingPart = new NumberingDefinitionsPart();

            // Attach to the package
            mdp.addTargetPart(numberingPart);

            // Create empty numbering.xml content
            Numbering numbering = Context.getWmlObjectFactory().createNumbering();
            numberingPart.setContents(numbering);

            System.out.println("Created empty numbering part.");
        }
    }
    private static P createPageBreak() {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R run = factory.createR();
        Br breakObj = factory.createBr();
        breakObj.setType(STBrType.PAGE); // PAGE break
        run.getContent().add(breakObj);
        p.getContent().add(run);
        return p;
    }
    private static void insertSectionBreak(MainDocumentPart targetMain, MainDocumentPart srcMain) {

        Body srcBody = srcMain.getJaxbElement().getBody();

        if (srcBody == null || srcBody.getSectPr() == null) {
            return;
        }

        SectPr srcSectPr = XmlUtils.deepCopy(srcBody.getSectPr());

        P sectionBreakPara = Context.getWmlObjectFactory().createP();
        PPr ppr = Context.getWmlObjectFactory().createPPr();

        ppr.setSectPr(srcSectPr);
        sectionBreakPara.setPPr(ppr);

        targetMain.getContent().add(sectionBreakPara);
    }
    public static void convertToLandscape(WordprocessingMLPackage pkg) {

        MainDocumentPart mainPart = pkg.getMainDocumentPart();
        Body body = mainPart.getJaxbElement().getBody();

        if (body.getSectPr() == null) {
            body.setSectPr(Context.getWmlObjectFactory().createSectPr());
        }

        SectPr sectPr = body.getSectPr();

        if (sectPr.getPgSz() == null) {
            sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
        }

        SectPr.PgSz pgSz = sectPr.getPgSz();

        // Read current dimensions
        BigInteger width = pgSz.getW();
        BigInteger height = pgSz.getH();

        // Swap height and width
        pgSz.setW(height);
        pgSz.setH(width);

        // Set orientation to landscape
        pgSz.setOrient(STPageOrientation.LANDSCAPE);
    }
    // Helper to locate EndnotesPart in a package
    private static EndnotesPart getEndnotesPart(WordprocessingMLPackage pkg) {
        for (Part p : pkg.getParts().getParts().values()) {
            if (p instanceof EndnotesPart) {
                return (EndnotesPart) p;
            }
        }
        return null;
    }
    private static List<Integer> getMapKeyList(String xml, String rgx) {
        List<Integer> currId = new ArrayList<>();
        try {
            Pattern pr = Pattern.compile(rgx);
            Matcher mr = pr.matcher(xml);
            while (mr.find()) {
//                if(rgx.contains("commentReference")) {
//                    log.info("commentReference id {} xml {}", mr.group(1), xml);
//                }
                currId.add(Integer.parseInt(mr.group(1)));
            }
        } catch (Exception ignored) {
        }
        return currId;
    }
    private static class IdMaps {
        Map<BigInteger, BigInteger> abstractNumIdMap = new LinkedHashMap<>();
        Map<BigInteger, BigInteger> numIdMap = new LinkedHashMap<>();
    }

    /**
     * Merge numbering of srcPart into targetPart. Assumes targetPart is NOT null.
     */
    private static IdMaps mergeNumbering(NumberingDefinitionsPart srcPart, NumberingDefinitionsPart targetPart)
            throws Docx4JException {

        IdMaps maps = new IdMaps();
        if (targetPart == null || srcPart == null) return maps;

        Numbering targetContent = targetPart.getContents();
        Numbering srcContent = srcPart.getContents();

        int maxAbstractNumId = targetContent.getAbstractNum().stream()
                .map(an -> an.getAbstractNumId() != null ? an.getAbstractNumId().intValue() : 0)
                .max(Integer::compare).orElse(0);

        int maxNumId = targetContent.getNum().stream()
                .map(n -> n.getNumId() != null ? n.getNumId().intValue() : 0)
                .max(Integer::compare).orElse(0);

        // ---- COPY ABSTRACT NUMs ----
        if (srcContent.getAbstractNum() != null) {
            for (Numbering.AbstractNum srcAbs : srcContent.getAbstractNum()) {

                Numbering.AbstractNum copyAbs = XmlUtils.deepCopy(srcAbs);

                BigInteger oldAbstractId = srcAbs.getAbstractNumId();
                BigInteger newAbstractId = BigInteger.valueOf(++maxAbstractNumId);

                copyAbs.setAbstractNumId(newAbstractId);

                maps.abstractNumIdMap.put(oldAbstractId, newAbstractId);
                targetContent.getAbstractNum().add(copyAbs);
            }
        }

        // ---- COPY NUMs ----
        if (srcContent.getNum() != null) {
            for (Numbering.Num srcNum : srcContent.getNum()) {

                Numbering.Num copyNum = XmlUtils.deepCopy(srcNum);

                BigInteger oldNumId = srcNum.getNumId();
                BigInteger newNumId = BigInteger.valueOf(++maxNumId);

                copyNum.setNumId(newNumId);

                // Remap abstractNum reference
                if (copyNum.getAbstractNumId() != null) {
                    BigInteger oldAbs = copyNum.getAbstractNumId().getVal();
                    if (maps.abstractNumIdMap.containsKey(oldAbs)) {
                        copyNum.getAbstractNumId().setVal(maps.abstractNumIdMap.get(oldAbs));
                    }
                }

                maps.numIdMap.put(oldNumId, newNumId);
                targetContent.getNum().add(copyNum);
            }
        }

        return maps;
    }


    private static void remapNumIdInContent(Object obj, Map<BigInteger, BigInteger> numIdMap) {
        if (obj instanceof JAXBElement<?>) {
            obj = ((JAXBElement<?>) obj).getValue();
        }
        else if (obj instanceof P ) {
            P p = (P) obj;

            if (p.getPPr() != null && p.getPPr().getNumPr() != null && p.getPPr().getNumPr().getNumId() != null
                    && p.getPPr().getNumPr().getNumId().getVal() != null) {

                BigInteger oldId = p.getPPr().getNumPr().getNumId().getVal();

                if (numIdMap.containsKey(oldId)) {
                    p.getPPr().getNumPr().getNumId().setVal(numIdMap.get(oldId));
                }
            }
        }

        if (obj instanceof ContentAccessor) {
            for (Object child : ((ContentAccessor) obj).getContent()) {
                remapNumIdInContent(child, numIdMap);
            }
        }
    }
    public static String getOrientation(byte[] file) {
        String orientation = "PORTRAIT";
        try {
            WordprocessingMLPackage pkg = WordprocessingMLPackage.load(new ByteArrayInputStream(file));
            Body body = pkg.getMainDocumentPart().getJaxbElement().getBody();
            SectPr sectPr = body.getSectPr();

            // Default Word behavior: portrait

            if (sectPr != null && sectPr.getPgSz() != null) {
                SectPr.PgSz pgSz = sectPr.getPgSz();

                // Check explicit orientation attribute
                if (pgSz.getOrient() != null && pgSz.getOrient().value().equals("landscape")) {
                    orientation = "LANDSCAPE";
                } else {
                    // Fallback: check width vs height
                    if (pgSz.getW().intValue() > pgSz.getH().intValue()) {
                        orientation = "LANDSCAPE";
                    }
                }
            }
        } catch (Exception e) {
            orientation = "PORTRAIT";
        }

        return orientation;
    }
}
