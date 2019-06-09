package com.doc.demo.docx;

import java.io.File;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;

public class SampleReader {
	public static void main(String[] args) throws JAXBException, Docx4JException {
		File doc = new File("G:\\Work\\projects\\CHAPTER 03-3 SECURITY POLICY.docx");
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
		  .load(doc);
		MainDocumentPart mainDocumentPart = wordMLPackage
		  .getMainDocumentPart();
		String textNodesXPath = "//w:t";
		List<Object> textNodes= mainDocumentPart
		  .getJAXBNodesViaXPath(textNodesXPath, true);
		for (Object obj : textNodes) {
		    Text text = (Text) ((JAXBElement) obj).getValue();
		    String textValue = text.getValue();
		    System.out.println(textValue);
		}

	}
}
