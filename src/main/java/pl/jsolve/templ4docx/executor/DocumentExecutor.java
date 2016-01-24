package pl.jsolve.templ4docx.executor;

import java.util.List;

import pl.jsolve.templ4docx.core.WordDocument;
import pl.jsolve.templ4docx.extractor.VariableFinder;
import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * This class uses variableFinder to generate list of Insert, and then begins the process of filling a template.
 * @author Lukasz Stypka
 */
public class DocumentExecutor {

    private VariableFinder variableFinder;
    private Variables variables;

    public DocumentExecutor(Variables variables) {
        this.variableFinder = new VariableFinder(variables);
        this.variables = variables;
    }

    /**
     * This method replaces variables in wordDocument template. Note, that before executing this method you should invoke
     * DocumentCleaner.clean method
     * @param wordDocument WordDocument object with opened .wordDocument file
     */
    public void execute(WordDocument wordDocument) {
        List<Insert> inserts = variableFinder.find(wordDocument.getXWPFDocument(), variables);
        variableFinder.replace(inserts);
    }

}
