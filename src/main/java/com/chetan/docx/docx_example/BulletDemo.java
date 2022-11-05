package com.chetan.docx.docx_example;


import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPrBase;

import java.io.File;
import java.math.BigInteger;

public class BulletDemo {
   public static void main(String[] args) throws Docx4JException {
      
      WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
      ObjectFactory factory = new org.docx4j.wml.ObjectFactory();
      P p = factory.createP();
      org.docx4j.wml.Text t = factory.createText();
      t.setValue("yyyyyyy");
      org.docx4j.wml.R run = factory.createR();
      run.getContent().add(t);
      p.getContent().add(run);
      org.docx4j.wml.PPr ppr = factory.createPPr();
      p.setPPr(ppr);
      // Create and add <w:numPr>
      PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
      ppr.setNumPr(numPr);
      // The <w:ilvl> element
      PPrBase.NumPr.Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
      numPr.setIlvl(ilvlElement);
      ilvlElement.setVal(BigInteger.valueOf(0));
      // The <w:numId> element
      PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
      numPr.setNumId(numIdElement);
      numIdElement.setVal(BigInteger.valueOf(1));
      wordMLPackage.getMainDocumentPart().addObject(p);
      wordMLPackage.save(new File("/home/che/Downloads/testBullet1.docx"));
   }
}