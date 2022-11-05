package com.chetan.docx.docx_example;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.math.BigInteger;


public class UnNumberedBulletedList {

private static final String BULLET_TEMPLATE ="<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
        "<w:abstractNum w:abstractNumId=\"0\">" +
        "<w:nsid w:val=\"12D402B7\"/>" +
        "<w:multiLevelType w:val=\"hybridMultilevel\"/>" +
        "<w:tmpl w:val=\"AECAFC2E\"/>" +
        "<w:lvl w:ilvl=\"0\" w:tplc=\"04090001\">" +
        "<w:start w:val=\"1\"/>" +
        "<w:numFmt w:val=\"bullet\"/>" +
        "<w:lvlText w:val=\"\uF0B7\"/>" +
        "<w:lvlJc w:val=\"left\"/>" +
        "<w:pPr>" +
        "<w:ind w:left=\"360\" w:hanging=\"360\"/>" +
        "</w:pPr>" +
        "<w:rPr>" +
        "<w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/>" +
        "</w:rPr>" +
        "</w:lvl>" +
        "</w:abstractNum>"+
        "<w:num w:numId=\"1\">" +
        "<w:abstractNumId w:val=\"0\"/>" +
        "</w:num>" +
        "</w:numbering>";

public static void main(String[] args) throws Exception{

    WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
    createBulletedList(wordMLPackage);
    wordMLPackage.save(new File("/home/che/Downloads/out.docx"));
}

private static void createBulletedList(WordprocessingMLPackage wordMLPackage)throws Exception{
    NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
    wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
    ndp.setJaxbElement((Numbering) XmlUtils.unmarshalString(BULLET_TEMPLATE));
    wordMLPackage.getMainDocumentPart().addObject(createParagraph("India"));
    wordMLPackage.getMainDocumentPart().addObject(createParagraph("United Kingdom"));
    wordMLPackage.getMainDocumentPart().addObject(createParagraph("France"));

}
private static P createParagraph(String country) throws JAXBException {

    ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
    P p = factory.createP();
    String text =
            "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    " <w:rPr>" +
                    "<w:b/>" +
                    " <w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\"/><w:sz w:val=\"16\"/>" +
                    " </w:rPr>" +
                    "<w:t>" + country + "</w:t>" +
                    "</w:r>";

    R r = (R) XmlUtils.unmarshalString(text);
    p.getContent().add(r);
    
    org.docx4j.wml.PPr ppr = factory.createPPr();

    p.setPPr(ppr);
    // Create and add <w:numPr>
    PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
    ppr.setNumPr(numPr);


    // The <w:numId> element
    PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
    numPr.setNumId(numIdElement);
    numIdElement.setVal(BigInteger.valueOf(2));
    return p;
}


}