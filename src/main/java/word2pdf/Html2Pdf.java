package word2pdf;


import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontProvider;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.Pipeline;
import com.itextpdf.tool.xml.XMLWorker;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import com.itextpdf.tool.xml.exceptions.CssResolverException;
import com.itextpdf.tool.xml.html.CssAppliers;
import com.itextpdf.tool.xml.html.CssAppliersImpl;
import com.itextpdf.tool.xml.html.Tags;
import com.itextpdf.tool.xml.parser.XMLParser;
import com.itextpdf.tool.xml.pipeline.css.CSSResolver;
import com.itextpdf.tool.xml.pipeline.css.CssResolverPipeline;
import com.itextpdf.tool.xml.pipeline.end.PdfWriterPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipelineContext;

import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.nio.charset.Charset;

/**
 * Created by Carey on 15-2-2.
 */
public class Html2Pdf {

    public static class MyFontsProvider extends XMLWorkerFontProvider {
        @Override
        public Font getFont(final String fontname, final String encoding,
          final boolean embedded, final float size, final int style,
          final BaseColor color) {
         BaseFont bf = null;
         try {
          bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
            BaseFont.NOT_EMBEDDED);
         } catch (DocumentException e) {
          // TODO Auto-generated catch block
          e.printStackTrace();
         } catch (IOException e) {
          // TODO Auto-generated catch block
          e.printStackTrace();
         }
         Font font = new Font(bf, size, style, color);
         font.setColor(color);
         return font;
        }
    }


    public boolean convertHtmlToPdf(String inputFile, String outputFile) throws Exception {

        OutputStream os = new FileOutputStream(outputFile);
        ITextRenderer renderer = new ITextRenderer();
        String url = new File(inputFile).toURI().toURL().toString();
        renderer.setDocument(url);
        // 解决中文支持问题
        ITextFontResolver fontResolver = renderer.getFontResolver();

        // BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",BaseFont.NOT_EMBEDDED);

        fontResolver.addFont("/com/itextpdf/text/pdf/fonts/cmaps/STSong-Light", "UniGB-UCS2-H",
                BaseFont.NOT_EMBEDDED);
        // 解决图片的相对路径问题
        // renderer.getSharedContext().setBaseURL("file:/D:/test");
        renderer.layout();
        renderer.createPDF(os);
        os.flush();
        os.close();
        return true;
    }

    public static void convert2(String infile, String outfile) throws FileNotFoundException,
            IOException, DocumentException, CssResolverException {
        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outfile));
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
        File input = new File(infile);
        p.parse(new InputStreamReader(new FileInputStream(input), "UTF-8"));
        document.close();
    }


    public static void main(String[] args) throws DocumentException, IOException,
            CssResolverException {


        convert2("/Users/triompha/Documents/work/1.html", "/Users/triompha/Documents/work/1.pdf");

        //
        // Document document = new Document();
        // // step 2
        // PdfWriter writer = PdfWriter.getInstance(document, new
        // FileOutputStream("/Users/triompha/Documents/work/1.pdf"));
        // // step 3
        // document.open();
        // // step 4
        //
        // XMLWorkerHelper.getInstance().parseXHtml(writer, document,
        // new
        // FileInputStream("/Users/triompha/Documents/work/1.html"),null,Charset.forName("UTF-8"),FontProvider);
        // //step 5
        // document.close();
        //
        // System.out.println( "PDF Created!" );

        //
        // Html2Pdf html2Pdf =new Html2Pdf();
        // try {
        // html2Pdf.convertHtmlToPdf("/Users/triompha/Documents/work/1.html","/Users/triompha/Documents/work/1.pdf");
        // } catch (Exception e) {
        // e.printStackTrace();
        // }
    }
}
