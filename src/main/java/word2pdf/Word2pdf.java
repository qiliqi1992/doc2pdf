package word2pdf;


import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.jsoup.Jsoup; 
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.Pipeline;
import com.itextpdf.tool.xml.XMLWorker;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import com.itextpdf.tool.xml.html.CssAppliers;
import com.itextpdf.tool.xml.html.CssAppliersImpl;
import com.itextpdf.tool.xml.html.Tags;
import com.itextpdf.tool.xml.parser.XMLParser;
import com.itextpdf.tool.xml.pipeline.css.CSSResolver;
import com.itextpdf.tool.xml.pipeline.css.CssResolverPipeline;
import com.itextpdf.tool.xml.pipeline.end.PdfWriterPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipelineContext;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import java.io.*;
import java.util.Iterator;

public class Word2pdf {

    
    public static void convert(String wordPath, String pdfPath) throws UnsupportedEncodingException, FileNotFoundException, IOException, DocumentException, TransformerException, ParserConfigurationException{
        String html = convert2HTML(wordPath);
        convertHTML2Pdf(html, pdfPath);
    }
    
    public static void convertHTML2Pdf(String html, String outFile) throws UnsupportedEncodingException, FileNotFoundException, IOException, DocumentException{
        convertHTML2Pdf(html, new FileOutputStream(outFile));
    }
    
    public static void  convertHTML2Pdf(String html , OutputStream out) throws UnsupportedEncodingException, IOException, DocumentException{
        com.itextpdf.text.Document document = new com.itextpdf.text.Document();
        PdfWriter writer = PdfWriter.getInstance(document, out);
        document.open();
        MyFontsProvider fontProvider = new MyFontsProvider();
        fontProvider.addFontSubstitute("lowagie", "garamond");
        fontProvider.setUseUnicode(true);
        // 使用我们的字体提供器，并将其设置为unicode字体样式
        CssAppliers cssAppliers = new CssAppliersImpl(fontProvider);
        HtmlPipelineContext htmlContext = new HtmlPipelineContext(cssAppliers);
        htmlContext.setTagFactory(Tags.getHtmlTagProcessorFactory());
        CSSResolver cssResolver = XMLWorkerHelper.getInstance().getDefaultCssResolver(true);
        Pipeline<?> pipeline =
                new CssResolverPipeline(cssResolver, new HtmlPipeline(htmlContext,
                        new PdfWriterPipeline(document, writer)));
        XMLWorker worker = new XMLWorker(pipeline, true);
        XMLParser p = new XMLParser(worker);
        
        
        p.parse(new InputStreamReader(new ByteArrayInputStream(html.getBytes()), "UTF-8"));
        document.close(); 
    }
    
    public static String adjustMeta(String content){
      //将不标准的<meta> 转换为标准的<meta></meta>
        org.jsoup.nodes.Document doc = Jsoup.parse(content);
        Elements meta = doc.getElementsByTag("meta");
        Iterator<Element> iterator = meta.iterator();
        while(iterator.hasNext()){
            Element next = iterator.next();
            next.html(next.html()+" ");
        }
         content=doc.html();        
         return content;
    }
    
    public static String convert2HTML(InputStream in) throws FileNotFoundException, IOException, TransformerException, ParserConfigurationException{
        HWPFDocument wordDocument = new HWPFDocument(in);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        wordToHtmlConverter.processDocument(wordDocument); 
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(out);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "HTML");
        serializer.transform(domSource, streamResult);
        out.close();
        return adjustMeta(new String(out.toByteArray()));
    } 
    
    public static String convert2HTML(String wordPath) throws FileNotFoundException, IOException, TransformerException, ParserConfigurationException{
        return convert2HTML(new FileInputStream(wordPath));
    }

    public static void main(String argv[]) {
        
        try {
            convert("/Users/triompha/Documents/work/a.doc","/Users/triompha/Documents/work/1.pdf");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
