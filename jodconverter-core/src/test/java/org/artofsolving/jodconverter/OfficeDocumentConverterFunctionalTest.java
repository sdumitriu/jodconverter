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

import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.testng.annotations.Test;

@Test(groups = "functional")
public class OfficeDocumentConverterFunctionalTest
{
    public void runAllPossibleConversions() throws IOException
    {
        OfficeManager officeManager = new DefaultOfficeManagerConfiguration().buildOfficeManager();
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);

        officeManager.start();
        try {
            File dir = new File("src/test/resources/documents");
            File[] files = dir.listFiles(new FilenameFilter()
            {
                public boolean accept(File dir, String name)
                {
                    return !name.startsWith(".");
                }
            });
            List<OfficeException> failedConversions = new ArrayList<OfficeException>();
            for (File inputFile : files) {
                failedConversions.addAll(convert(converter, inputFile));
            }
            if (!failedConversions.isEmpty()) {
                throw failedConversions.get(0);
            }
        } finally {
            officeManager.stop();
        }
    }

    private List<OfficeException> convert(OfficeDocumentConverter converter, File inputFile) throws IOException
    {
        List<OfficeException> failedConversions = new ArrayList<OfficeException>();
        String inputExtension = FilenameUtils.getExtension(inputFile.getName());
        DocumentFormatRegistry formatRegistry = converter.getFormatRegistry();
        DocumentFormat inputFormat = formatRegistry.getFormatByExtension(inputExtension);
        assertNotNull(inputFormat, "unknown input format: " + inputExtension);
        Set<DocumentFormat> outputFormats = formatRegistry.getOutputFormats(inputFormat.getInputFamily());
        for (DocumentFormat outputFormat : outputFormats) {
            File outputFile = File.createTempFile("test", "." + outputFormat.getExtension());
            outputFile.deleteOnExit();
            System.out.printf("-- converting %s to %s... ", inputFormat.getExtension(), outputFormat.getExtension());
            try {
                converter.convert(inputFile, outputFile, outputFormat);
                System.out.printf("done.\n");
                assertTrue(outputFile.isFile() && outputFile.length() > 0);
                // TODO use file detection to make sure outputFile is in the expected format
            } catch (OfficeException e) {
                System.out.printf("failed.\n");
                failedConversions.add(e);
            }
        }
        return failedConversions;
    }
}
