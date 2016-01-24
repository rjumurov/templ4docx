package pl.jsolve.templ4docx.core;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.List;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import pl.jsolve.sweetener.io.Resources;
import pl.jsolve.templ4docx.cleaner.DocumentCleaner;
import pl.jsolve.templ4docx.exception.OpenDocxException;
import pl.jsolve.templ4docx.executor.DocumentExecutor;
import pl.jsolve.templ4docx.extractor.VariablesExtractor;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * The main class responsible for reading docx template, finding variables, saving completed docx template
 * @author Lukasz Stypka
 */
public class WordDocument implements Serializable {

    private static final long serialVersionUID = 1L;

    private VariablePattern variablePattern = new VariablePattern("${", "}");

    private XWPFDocument docx = null;

    private DocumentCleaner documentCleaner;

    private WordDocument(XWPFDocument docx) {
        this.documentCleaner = new DocumentCleaner();
        this.docx = docx;
    }

    public static WordDocument of(InputStream in) {
        XWPFDocument docx;
        try {
            docx = new XWPFDocument(in);
        } catch (IOException e) {
            throw new OpenDocxException(e.getMessage(), e.getCause());
        }
        return new WordDocument(docx);
    }

    public List<String> findVariables() {
        VariablesExtractor extractor = new VariablesExtractor();
        String content = readTextContent();
        return extractor.extract(content, variablePattern);
    }

    public String readTextContent() {
        XWPFWordExtractor extractor = null;
        try {
            extractor = new XWPFWordExtractor(docx);
            return extractor.getText();
        } finally {
            if (extractor != null) {
                Resources.closeStream(extractor);
            }
        }
    }

    public void fillTemplate(Variables variables) {
        documentCleaner.clean(this, variables, variablePattern);
        DocumentExecutor documentExecutor = new DocumentExecutor(variables);
        documentExecutor.execute(this);
    }

    public void save(String outputPath) {
        try {
            docx.write(new FileOutputStream(outputPath));
        } catch (Exception ex) {
            throw new OpenDocxException(ex.getMessage(), ex.getCause());
        }
    }

    public OutputStream toOutputStream() {
        OutputStream stream = new ByteArrayOutputStream();
        try {
            docx.write(stream);
        } catch (Exception ex) {
            throw new OpenDocxException(ex.getMessage(), ex.getCause());
        }
        return stream;
    }

    public void setVariablePattern(VariablePattern variablePattern) {
        this.variablePattern = variablePattern;
    }

    public XWPFDocument getXWPFDocument() {
        return docx;
    }

}
