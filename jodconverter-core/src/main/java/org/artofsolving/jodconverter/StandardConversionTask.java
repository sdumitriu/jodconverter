//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.filter.OfficeDocumentFilter;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;

import com.sun.star.lang.XComponent;

public class StandardConversionTask extends AbstractConversionTask
{
    private final DocumentFormat outputFormat;

    private final List<OfficeDocumentFilter> filters = new ArrayList<OfficeDocumentFilter>();

    private Map<String, ?> defaultLoadProperties;

    private DocumentFormat inputFormat;

    public StandardConversionTask(File inputFile, File outputFile, DocumentFormat outputFormat)
    {
        super(inputFile, outputFile);
        this.outputFormat = outputFormat;
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties)
    {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setInputFormat(DocumentFormat inputFormat)
    {
        this.inputFormat = inputFormat;
    }

    public List<OfficeDocumentFilter> getFilters()
    {
        return this.filters;
    }

    @Override
    protected void modifyDocument(XComponent document, OfficeContext context) throws OfficeException
    {
        for (OfficeDocumentFilter filter : this.filters) {
            filter.filter(document, context);
        }
    }

    @Override
    protected Map<String, ?> getLoadProperties(File inputFile)
    {
        Map<String, Object> loadProperties = new HashMap<String, Object>();
        if (this.defaultLoadProperties != null) {
            loadProperties.putAll(this.defaultLoadProperties);
        }
        if (this.inputFormat != null && this.inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(this.inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    @Override
    protected Map<String, ?> getStoreProperties(File outputFile, XComponent document)
    {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        return this.outputFormat.getStoreProperties(family);
    }
}
