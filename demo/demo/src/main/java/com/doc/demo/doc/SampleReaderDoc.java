package com.doc.demo.doc;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class SampleReaderDoc {
	
	public static void main(String[] args) {
		try {
			FileInputStream fis = new FileInputStream("G:\\Work\\projects\\CHAPTER 03-3 SECURITY POLICY.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
			List<XWPFParagraph> paragraphs= xdoc.getParagraphs();
			List<IBodyElement> elements= xdoc.getBodyElements();
					   
			/*
			 * int count =1;
			 * 
			 * for(IBodyElement element: elements) { System.out.print("element "+ count++ +
			 * ":"); System.out.println(element.getPart());
			 * //element.getBody().getParagraphs().stream().forEach( p->
			 * System.out.println(p.getParagraphText())); }
			 */
					   
			int count =1;
			   for(XWPFParagraph para: paragraphs) {
				   System.out.print("Paragraph "+ count++ + ":");
				   para.getRuns().forEach(r -> System.out.print("Bold :" +r.isBold()+ " Color: "+ r.getColor()+ "||"));
				   if(para.getStyle() != null && para.getStyle().contains("Heading"))
				   System.out.println(para.getStyle());
				   System.out.println();
				   System.out.println(para.getParagraphText());
				   para.getText();
				   
				   
				   
			   }
			   
			/*
			 * XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
			 * System.out.println(extractor.getText());
			 */
			} catch(Exception ex) {
			    ex.printStackTrace();
			}
	}
}
