package com.joaodjunior.pdfunion.app;


import com.lowagie.text.Document;
import com.lowagie.text.pdf.PRAcroForm;
import com.lowagie.text.pdf.PdfCopy;
import com.lowagie.text.pdf.PdfImportedPage;
import com.lowagie.text.pdf.PdfReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Classe criada para mesclar arquivos em PDF e gerar um final
 * 
 * @author Joao Luiz
 * @category Utils
 * @version 0.0.1
 * @since 14/03/2018
 */

public class MergePDF {
	
	static ByteArrayOutputStream out = new ByteArrayOutputStream();
	
	static Document doc;
	
	static PdfCopy copia;
	
    /**
     * Metodo o qual converte um arquivo *.PDF de entrada em umByteArray 
     * @param stream
     * @return ByteArrayOutputStream
     * @throws IOException
     */
	public static ByteArrayOutputStream getBytesArray(InputStream stream) throws IOException {
    	byte[] buffer = new byte[8192];
    	
    	ByteArrayOutputStream ArrayDeSaida = new ByteArrayOutputStream();
    	
    	int bytesLidos;
    	while((bytesLidos = stream.read(buffer)) != -1 ) {
    		ArrayDeSaida.write(buffer, 0, bytesLidos);
    	}
    	
    	return ArrayDeSaida;
    	
    }
    /**
     * Metodo recebe um array de bytes e o converte em um arquivo *.PDF de saida
     * @param bytesToConvertFile
     * @throws IOException
     */
    public static void criaArquivo(byte[] bytesToConvertFile) throws IOException {
    	
        //below is the different part
        File someFile = new File("c:/users/14208/Desktop/java2.pdf");
        FileOutputStream fos = new FileOutputStream(someFile);
        fos.write(bytesToConvertFile);
        fos.flush();
        fos.close();
        
    }
    /**
     * Metodo o qual recebe uma "Lista" de array de bytes, e o concatena em um arquivo unico de saida
     * @param pdfBytes
     * @throws IOException 
     */
    public static void unirBytesDeArquivo(byte[]... pdfBytes) throws IOException {
    	
    	for (byte[] byteDeArquivo : pdfBytes) {
    		try {
    			
    			PdfReader reader = new PdfReader(byteDeArquivo);
    			
    			int numerodePaginas = reader.getNumberOfPages();
    			
    			if(doc == null) {
    				doc = new Document(reader.getPageSizeWithRotation(1));
    				
    				copia = new PdfCopy(doc, out );
    				
    				doc.open();
    				
    			}
    			
    			PdfImportedPage pagina;
    			
    			for(int i = 0 ; i< numerodePaginas;) {
    				i++;
    				pagina = copia.getImportedPage(reader, i);
    				copia.addPage(pagina);
    			}
    			
    			PRAcroForm form = reader.getAcroForm();
    			
    			if(form != null) {
    				copia.copyAcroForm(reader);
    			}
    		} catch (Exception e) {
    			throw new RuntimeException(e);
    		}	
    	}
    }
    
	public static byte[] getMergedPdfByteArray() {
		if (out != null) {
			return getOut().toByteArray();
		} else {
			return null;
		}
	}
	
	public static void fecharDocumentos() {
        try {
        	doc.close();
			out.close();
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}
	
    public static ByteArrayOutputStream getOut() {
		return out;
	}

	public void setOut(ByteArrayOutputStream out) {
		this.out = out;
	}
	
	  /*  public static void concatPDFs(List<InputStream> streamOfPDFFiles,
    OutputStream outputStream, boolean paginate) {

	Document document = new Document();
	try {
	    List<InputStream> pdfs = streamOfPDFFiles;
	    List<PdfReader> readers = new ArrayList<PdfReader>();
	    int totalPages = 0;
	    Iterator<InputStream> iteratorPDFs = pdfs.iterator();
	    // Create Readers for the pdfs.
	    while (iteratorPDFs.hasNext()) {
	        InputStream pdf = iteratorPDFs.next();
	        PdfReader pdfReader = new PdfReader(pdf);
	        
	        
	        
	        readers.add(pdfReader);
	        totalPages += pdfReader.getNumberOfPages();
	    }
	    // Create a writer for the outputstream
	    PdfWriter writer = PdfWriter.getInstance(document, outputStream);
	    document.open();
	    BaseFont bf = BaseFont.createFont(BaseFont.HELVETICA,
	            BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
	    PdfContentByte cb = writer.getDirectContent(); // Holds the PDF
	    // data
	    PdfImportedPage page;
	    int currentPageNumber = 0;
	    int pageOfCurrentReaderPDF = 0;
	    Iterator<PdfReader> iteratorPDFReader = readers.iterator();
	    // Loop through the PDF files and add to the output.
	    while (iteratorPDFReader.hasNext()) {
	        PdfReader pdfReader = iteratorPDFReader.next();
	        // Create a new page in the target for each source page.
	        while (pageOfCurrentReaderPDF < pdfReader.getNumberOfPages()) {
	            document.newPage();
	            pageOfCurrentReaderPDF++;
	            currentPageNumber++;
	            page = writer.getImportedPage(pdfReader,
	                    pageOfCurrentReaderPDF);
	            cb.addTemplate(page, 0, 0);
	            // Code for pagination.
	            if (paginate) {
	                cb.beginText();
	                cb.setFontAndSize(bf, 9);
	                cb.showTextAligned(PdfContentByte.ALIGN_CENTER, ""
	                        + currentPageNumber + " of " + totalPages, 520,
	                        5, 0);
	                cb.endText();
	            }
	        }
	        pageOfCurrentReaderPDF = 0;
	    }
	    outputStream.flush();
	    document.close();
	    outputStream.close();
	} catch (Exception e) {
	    e.printStackTrace();
	} finally {
	    if (document.isOpen())
	        document.close();
	    try {
	        if (outputStream != null)
	            outputStream.close();
	    } catch (IOException ioe) {
	        ioe.printStackTrace();
	    }
	}
	}*/



}
