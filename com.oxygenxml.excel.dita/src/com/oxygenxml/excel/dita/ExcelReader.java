package com.oxygenxml.excel.dita;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.lang.reflect.Constructor;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.DTDHandler;
import org.xml.sax.EntityResolver;
import org.xml.sax.ErrorHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.SAXNotRecognizedException;
import org.xml.sax.SAXNotSupportedException;
import org.xml.sax.XMLReader;

public class ExcelReader implements XMLReader {

	/**
	 * Entity resolver
	 */
	private EntityResolver resolver;
	/**
	 * Content Handler
	 */
	private ContentHandler handler;
	/**
	 * Error Handler.
	 */
	private ErrorHandler errorHandler;

	@Override
	public boolean getFeature(String name) throws SAXNotRecognizedException, SAXNotSupportedException {
		return false;
	}

	@Override
	public void setFeature(String name, boolean value) throws SAXNotRecognizedException, SAXNotSupportedException {
	}

	@Override
	public Object getProperty(String name) throws SAXNotRecognizedException, SAXNotSupportedException {
		return null;
	}

	@Override
	public void setProperty(String name, Object value) throws SAXNotRecognizedException, SAXNotSupportedException {
	}

	@Override
	public void setEntityResolver(EntityResolver resolver) {
		this.resolver = resolver;
	}

	@Override
	public EntityResolver getEntityResolver() {
		return resolver;
	}

	@Override
	public void setDTDHandler(DTDHandler handler) {
		//
	}

	@Override
	public DTDHandler getDTDHandler() {
		return null;
	}

	@Override
	public void setContentHandler(ContentHandler handler) {
		this.handler = handler;
	}

	@Override
	public ContentHandler getContentHandler() {
		return handler;
	}

	@Override
	public void setErrorHandler(ErrorHandler handler) {
		errorHandler = handler;
	}

	@Override
	public ErrorHandler getErrorHandler() {
		return errorHandler;
	}
	
	/**
	 * Create a workbook.
	 * @param extension The extension
	 * @param is The input stream.
	 * @return The workbook
	 * @throws IOException
	 */
	private static Workbook createWorkbook(String extension, InputStream is) throws IOException {
		Workbook wb = null;
		if("xlsx".equals(extension)){
			//New XML-type Excel files.
			try {
				Class<?> clazz = Class.forName("org.apache.poi.xssf.usermodel.XSSFWorkbook");
				Constructor<?> constructor = clazz.getConstructor(new Class[] {InputStream.class});
				wb = (Workbook) constructor.newInstance(new Object[] {is});
			} catch (Throwable e) {
				throw new IOException(e);
			}
		} else {
			POIFSFileSystem fs = new POIFSFileSystem(is);
			wb = new HSSFWorkbook(fs);
		}
		return wb;
	}

	@Override
	public void parse(InputSource input) throws IOException, SAXException {
		InputStream is = input.getByteStream();
		URL url = new URL(input.getSystemId());
		//Check out how many header rows we should have
		int headerRowsNo = 1;
		String query = url.getQuery();
		if(query != null) {
			String[] paramNameValue = query.split("&");
			for (int i = 0; i < paramNameValue.length; i++) {
				String nameValue = paramNameValue[i];
				String[] nameAndValue = nameValue.split("=");
				if(nameAndValue != null && nameAndValue.length == 2) {
					if("headerRowsNo".equals(nameAndValue[0])) {
						try {
							headerRowsNo = Integer.parseInt(nameAndValue[1]);
						} catch(Exception ex) {
							ex.printStackTrace();
						}
					}
				}
			}
		}
		if(is == null) {
			URL urlForConnect = url;
			if("file".equals(url.getProtocol())) {
				String urlStr = urlForConnect.toString();
				if(urlStr.contains("?")) {
					//Remove query part
					urlForConnect = new URL(urlStr.substring(0, urlStr.indexOf("?")));
				}
			}
			is = urlForConnect.openStream();
		}
		StringBuilder sb = new StringBuilder();
		String name = url.getPath();
		String extension = "";
		if(name.contains("/")) {
			name = name.substring(name.lastIndexOf("/") + 1, name.length());
			if(name.contains(".")) {
				extension = name.substring(name.lastIndexOf(".") + 1, name.length());
				name = name.substring(0, name.lastIndexOf("."));
			}
		}
		sb.append("<topic id='" + name + "' class='- topic/topic ' "
				+ "xmlns:ditaarch=\"http://dita.oasis-open.org/architecture/2005/\" ditaarch:DITAArchVersion=\"1.3\""
				+ " domains='(topic abbrev-d)                            a(props deliveryTarget)                            (topic equation-d)                            (topic hazard-d)                            (topic hi-d)                            (topic indexing-d)                            (topic markup-d)                            (topic mathml-d)                            (topic pr-d)                            (topic relmgmt-d)                            (topic sw-d)                            (topic svg-d)                            (topic ui-d)                            (topic ut-d)                            (topic markup-d xml-d)   '>");
		sb.append("<title class='- topic/title '>" + name + "</title>");
		Workbook workbook = createWorkbook(extension, is);
		int noSheets = workbook.getNumberOfSheets();
		sb.append("<body class='- topic/body '>");
		for (int i = 0; i < noSheets; i++) {
			Sheet datatypeSheet = workbook.getSheetAt(i);
			Iterator<Row> iterator = datatypeSheet.iterator();
			if(iterator.hasNext()) {
				sb.append("<table id='" + datatypeSheet.getSheetName().replace(' ', '_') + "' class='- topic/table '>");
				sb.append("<title class=\"- topic/title \">").append(datatypeSheet.getSheetName()).append("</title>");
				List<StringBuilder> rowsData = new ArrayList<>();
				//For each sheet we have a table
				int maxColCount = 0;
				while (iterator.hasNext()) {
					StringBuilder rowData = new StringBuilder();
					rowData.append("<row class=\"- topic/row \">");
					Row currentRow = iterator.next();
					Iterator<Cell> cellIterator = currentRow.iterator();
					int colCount = 0;
					while (cellIterator.hasNext()) {
						colCount++;
						Cell currentCell = cellIterator.next();
						rowData.append("<entry class='- topic/entry '>");
						rowData.append(getImportRepresentation(currentCell, true));
						rowData.append("</entry>");
					}
					if(colCount > maxColCount) {
						maxColCount = colCount;
					}
					rowData.append("</row>");
					rowsData.add(rowData);
				}
				sb.append("<tgroup cols='" + maxColCount + "' class=\"- topic/tgroup \">");
				if(headerRowsNo > 0) {
					sb.append("<thead class=\"- topic/thead \">");
					for (int j = 0; j < headerRowsNo && j < rowsData.size(); j++) {
						sb.append(rowsData.get(j));
					}
					sb.append("</thead>");
				}
				sb.append("<tbody class=\"- topic/tbody \">");
				for (int j = headerRowsNo; j < rowsData.size(); j++) {
					sb.append(rowsData.get(j));
				}
				sb.append("</tbody>");
				sb.append("</tgroup>");

				sb.append("</table>");
				
			}
		}
		sb.append("</body>");
		sb.append("</topic>");

		//Delegate to content handler.
		SAXResult result = new SAXResult(handler);
		try {
			TransformerFactory.newInstance().newTransformer().transform(new StreamSource(new StringReader(sb.toString()), url.toString()),
					result);
		} catch (Exception e) {
			throw new IOException(e);
		}
	}
	
	  /**
	   * Check if a cell contains a date Since dates are stored internally in Excel as double values we
	   * infer it is a date if it is formatted as such.
	   * 
	   * @param cell
	   *          Excel cell to check.
	   * @return true if cell is succeptible of showing a date.
	   * 
	   * @see #isInternalDateFormat(int)
	   */
	  private static boolean isCellDateFormatted(Cell cell) {
	    // Starting with POI 3.5 we rely on a new mechanism to discover if the number corresponds to a date or not.
	    if (cell == null) {
	      return false;
	    }
	    return DateUtil.isCellDateFormatted(cell);
	  }
	
	  /**
	   * Get the import string representation of the Excel cell.
	   * 
	   * @param cell
	   *          Excel cell to be represented.
	   * @param displayDataAsInExcel <code>true</code> if the data will be displayed as it appears in Excel,
	   *         <code>false</code> if it will display the row data.
	   * @return String representation of the Excel cell.
	   */
	  private static String getImportRepresentation(Cell cell, boolean displayDataAsInExcel) {
	    String importPresentationString = "";
	    DataFormatter df = null;
	    if (cell != null) {
	      df = new DataFormatter();
	      if (displayDataAsInExcel && (cell.getCellType() == Cell.CELL_TYPE_FORMULA)) {
	        // We are trying to evaluate the formula.
	        FormulaEvaluator fe = null;
	        try {
	          if(cell.getSheet().getWorkbook() instanceof HSSFWorkbook) {
	            //Older Excel
	            fe = new HSSFFormulaEvaluator((HSSFWorkbook) cell.getSheet().getWorkbook());
	          } else {
	            //Or newer 2007+
	            Class<?> newEvaluatorClazz = Class.forName("org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator");
	            Class<?> wbClazz = Class.forName("org.apache.poi.xssf.usermodel.XSSFWorkbook");
	            Constructor<?> constructor = newEvaluatorClazz.getConstructor(new Class[] {wbClazz});
	            fe = (FormulaEvaluator) constructor.newInstance(new Object[] {cell.getSheet().getWorkbook()});
	          }
	          fe.evaluateInCell(cell);
	          // After evaluation, the cell will change its type.
	        } catch (Exception e) {
	          //EXM-21075 If we can't evaluate formula we fall back to the old approach.
	        	System.err.println("Could not evaluate the cell formula: " + e.getMessage());
	        	e.printStackTrace();
	          // We were unable to evaluate the formula due an error, so we can't show data as it appears in Excel.
	          displayDataAsInExcel = false;
	        }
	      } 
	      switch (cell.getCellType()) {
	        case Cell.CELL_TYPE_NUMERIC:
	          // Numeric types are succeptible of being dates/time too.
	          boolean isDateCell = isCellDateFormatted(cell);
	          if (isDateCell) {
	            // For date, we prefer to have our own kind of representation.
	            // Default value for date objects
	            importPresentationString = getDescriptionOfDate(cell);
	          } else {
	            if (displayDataAsInExcel) {
	              importPresentationString = df.formatCellValue(cell);
	            } else {
	              double doubleValue = cell.getNumericCellValue();
	              int intValue = (int) doubleValue;
	              if (doubleValue == intValue) {
	                importPresentationString = Integer.toString(intValue);
	              } else {
	                importPresentationString = Double.toString(doubleValue);
	              }
	            }
	          }
	          break;
	        case Cell.CELL_TYPE_STRING:
	          importPresentationString =
	            displayDataAsInExcel ? df.formatCellValue(cell)
	                                 : cell.getRichStringCellValue().getString();
	          break;
	        case Cell.CELL_TYPE_BLANK:
	          importPresentationString = "";
	          break;
	        case Cell.CELL_TYPE_BOOLEAN:
	          importPresentationString =
	            displayDataAsInExcel ? df.formatCellValue(cell)
	                                 : Boolean.toString(cell.getBooleanCellValue());
	          break;
	        case Cell.CELL_TYPE_ERROR:
	          importPresentationString = "#ERROR: " + Byte.toString(cell.getErrorCellValue());
	          break;
	        case Cell.CELL_TYPE_FORMULA:
	          importPresentationString =
	            displayDataAsInExcel ? df.formatCellValue(cell)
	                                 : getDescriptionOfFormula(cell);
	           break;
	        default:
	        	System.err.println("unsuported cell type " + cell.getCellType());
	          break;
	      }
	    }
	    return importPresentationString;
	  }
	  
	  /**
	   * Get the description of a cell having a Excel formula.
	   * We assume that the cell was already checked and it is a formula cell.
	   * 
	   * @param cell Excel cell.
	   * @return Description of the cell having an Excel formula.
	   */
	  private static String getDescriptionOfFormula(Cell cell) {
	    String descr = "";
	    switch (cell.getCachedFormulaResultType()) {
	      case Cell.CELL_TYPE_NUMERIC:
	        // Numeric types are succeptible of being dates/time too.
	        boolean isDateCell = isCellDateFormatted(cell);
	        if (isDateCell) {
	          // Default value for date objects
	          descr = getDescriptionOfDate(cell);
	        } else {
	          double doubleValue = cell.getNumericCellValue();
	          int intValue = (int) doubleValue;
	          if (doubleValue == intValue) {
	            descr = Integer.toString(intValue);
	          } else {
	            descr = Double.toString(doubleValue);
	          }
	        }
	        break;
	      case Cell.CELL_TYPE_BOOLEAN:
	        descr = Boolean.toString(cell.getBooleanCellValue());
	        break;
	      case Cell.CELL_TYPE_ERROR:
	        // The formula cannot be evaluated ... 
	        // Byte.toString(cell.getErrorCellValue()) is not an interesting value .. 
	        descr = "#ERROR: " + cell.getCellFormula();
	        break;
	      case Cell.CELL_TYPE_STRING:
	        descr = cell.getRichStringCellValue().getString();
	        break;
	      default:
	        // Can't be of other type according to specs.
	        break;
	    }
	    return descr;
	  }
	  
	  /**
	   * Get the description of a cell representing a date cell.
	   * 
	   * @param cell Date cell (we assume that this fact is valid)
	   * @return Description of the date cell.s
	   */
	  private static String getDescriptionOfDate(Cell cell) {
		  // This is unformatted, that is Excel type.
		  DataFormatter formatter = new DataFormatter();
		  return formatter.formatCellValue(cell);
	  }

	@Override
	public void parse(String systemId) throws IOException, SAXException {
		parse(new InputSource(systemId));
	}
	
	public static void main(String[] args) throws MalformedURLException, IOException, SAXException {
		ExcelReader reader = new ExcelReader();
//		reader.setContentHandler(new ContentHandler() {
//			@Override
//			public void startPrefixMapping(String prefix, String uri) throws SAXException {
//			}
//			
//			@Override
//			public void startElement(String uri, String localName, String qName, Attributes atts) throws SAXException {
//				System.err.println("START " + localName);
//			}
//			
//			@Override
//			public void startDocument() throws SAXException {
//				System.err.println("START DOC ");
//			}
//			
//			@Override
//			public void skippedEntity(String name) throws SAXException {
//			}
//			
//			@Override
//			public void setDocumentLocator(Locator locator) {
//			}
//			
//			@Override
//			public void processingInstruction(String target, String data) throws SAXException {
//			}
//			
//			@Override
//			public void ignorableWhitespace(char[] ch, int start, int length) throws SAXException {
//			}
//			
//			@Override
//			public void endPrefixMapping(String prefix) throws SAXException {
//			}
//			
//			@Override
//			public void endElement(String uri, String localName, String qName) throws SAXException {
//			}
//			
//			@Override
//			public void endDocument() throws SAXException {
//			}
//			
//			@Override
//			public void characters(char[] ch, int start, int length) throws SAXException {
//			}
//		});
		reader.parse(new File("samples/Book1.xlsx").toURI().toURL().toString());
	}

}
