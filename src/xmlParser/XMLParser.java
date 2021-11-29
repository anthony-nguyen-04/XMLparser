package xmlParser;

//https://mkyong.com/java/how-to-read-xml-file-in-java-dom-parser/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

public class XMLParser {

	private static final String fileName = "test.xml";
	
	private static final String[] attributes = {
			"Link to the Ovid Full Text or citation",
			"Authors",
			"Update Date",
			"Article Identifier",
			"Institution",
			"Conflict of Interest",
			"ISSN Linking",
			"Publication Type",
			"Entrez Date",
			"Journal Subset",
			"Publication Status",
			"Authors Full Name",
			"Status",
			"ISO Journal Abbreviation",
			"Keyword Heading",
			"Create Date",
			"Year of Publication",
			"Electronic Date of Publication",
			"Title Comment",
			"NLM Journal Name",
			"Publisher Item Identifier",
			"Record Owner",
			"ISSN Print",
			"Registry Number/Name of Substance",
			"ISSN Electronic",
			"Publication History Status",
			"Unique Identifier",
			"Grant Information",
			"Digital Object Identifier",
			"Other ID",
			"NLM Journal Code",
			"Abbreviated Source",
			"Entry Date",
			"Link to the External Link Resolver",
			"Publishing Model",
			"Database Code",
			"MeSH Subject Headings",
			"Source",
			"Revision Date",
			"MeSH Date",
			"Author NameID",
			"Language",
			"Indexing Method",
			"Version ID",
			"Date of Publication",
			"PMC Identifier",
			"Abstract",
			"Comments",
			"Title",
			"Cited References",
			"Country of Publication"
	}; 
	
	public static void main(String[] args) {
		
		// Instantiate the Factory
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();

		XPath xPath =  XPathFactory.newInstance().newXPath();
		
		//Preps workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
        XSSFSheet sheet = workbook.createSheet("Data");
		CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(false);
        ///
		
		try {
			
			
	        Row titleRow = sheet.createRow(0);
	        for (int i = 0; i < attributes.length; i++) {
	        	Cell cell = titleRow.createCell(i);
	            cell.setCellStyle(cellStyle);
	            cell.setCellValue(attributes[i]);
	        }
	        
	        
			 // parse XML file
	        DocumentBuilder db = dbf.newDocumentBuilder();
	
	        Document doc = db.parse(new File(fileName));
		
	        NodeList list = doc.getElementsByTagName("record");
	        
	        for (int temp = 0; temp < list.getLength(); temp++) {
	        	Node node = list.item(temp);
	        	
	        	if (node.getNodeType() == Node.ELEMENT_NODE) {

	        		Row row = sheet.createRow(temp + 1);
	        		
	                //Element element = (Element) node;
//	                System.out.println(element.getAttribute("index"));
//	                  
//	                String database = element.getElementsByTagName("F").item(1).getTextContent(); //.getAttributes().getNamedItem("L").getTextContent();
//	                System.out.println(database);
//	                  
//	                Node test = (Node) xPath.compile("F[@L = 'Authors']").evaluate(node, XPathConstants.NODE);
//	                System.out.println(test.getTextContent());
	                 
	                 for (int i = 0; i < attributes.length; i++) {
	                	 String data = attributes[i];
	                	 
	                	 String xpath = "F[@L = '" + data + "']";
	                	 System.out.println(xpath);
	                	 
	                	 String cellData = "";
	                	 
	                	 try {
		                	 Node test2 = (Node) xPath.compile(xpath).evaluate(node, XPathConstants.NODE);
		                	 System.out.println(test2.getTextContent());
		                	 cellData = test2.getTextContent();
	                	 }
	                	 catch (NullPointerException e) {
	                		 //System.out.println("NO DATA");
	                		 cellData = "NO DATA";
	                	 }
	                	 
	                	 Cell cell = row.createCell(i);
	                	 cell.setCellStyle(cellStyle);
	                	 cell.setCellValue(cellData);
	                	 
	                 }
	        	}
	        	
	        	
	        }
		}
		
		catch (Exception e){
			e.printStackTrace();
		}
		
		try {
			FileOutputStream out = new FileOutputStream(new File("data.xlsx"));
			workbook.write(out);
			
			workbook.close();
			out.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}

}
