/*
 * See the NOTICE file distributed with this work for additional
 * information regarding copyright ownership.
 *
 * This is free software; you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation; either version 2.1 of
 * the License, or (at your option) any later version.
 *
 * This software is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this software; if not, write to the Free
 * Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA
 * 02110-1301 USA, or see the FSF site: http://www.fsf.org.
 */
package org.artofsolving.jodconverter.filter;

import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeUtils;

import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XTextGraphicObjectsSupplier;

/**
 * Embeds external images.
 * 
 * @version $Id$
 * @since 3.1-xwiki
 */
public class ImageEmbedderFilter implements OfficeDocumentFilter
{
    @Override
    public void filter(XComponent document, OfficeContext context) throws OfficeException
    {
        if (OfficeUtils.cast(XServiceInfo.class, document).supportsService("com.sun.star.text.GenericTextDocument")) {
            embedWriterImages(document, context);
        }
    }

    private void embedWriterImages(XComponent document, OfficeContext context)
    {
        XIndexAccess indexAccess = OfficeUtils.cast(XIndexAccess.class,
            OfficeUtils.cast(XTextGraphicObjectsSupplier.class, document).getGraphicObjects());
        XGraphicProvider graphicProvider =
            OfficeUtils.cast(XGraphicProvider.class, context.getService("com.sun.star.graphic.GraphicProvider"));
        PropertyValue[] queryProperties = new PropertyValue[] {new PropertyValue()};
        queryProperties[0].Name = "URL";
        for (int i = 0; i < indexAccess.getCount(); i++) {
            try {
                XPropertySet graphicProperties = OfficeUtils.cast(XPropertySet.class, indexAccess.getByIndex(i));
                String graphicURL = (String) graphicProperties.getPropertyValue("GraphicURL");
                if (!graphicURL.contains("vnd.sun.star.GraphicObject")) {
                    queryProperties[0].Value = graphicURL;
                    // Before embedding the image, the "ActualSize" property holds the image size specified in the
                    // document content. If the width or height are not specified then their actual values will be 0.
                    Size specifiedSize = OfficeUtils.cast(Size.class, graphicProperties.getPropertyValue("ActualSize"));
                    graphicProperties.setPropertyValue("Graphic", graphicProvider.queryGraphic(queryProperties));
                    // Images are embedded as characters (see TextContentAnchorType.AS_CHARACTER) and their size is
                    // messed up if it's not explicitly specified (e.g. if the image height is not specified then it
                    // takes the line height).
                    adjustImageSize(graphicProperties, specifiedSize);
                }
            } catch (Exception e) {
                // Skip this graphic.
            }
        }
    }

    private void adjustImageSize(XPropertySet graphicProperties, Size specifiedSize)
    {
        try {
            // After embedding the image, the "ActualSize" property holds the actual image size.
            Size size = OfficeUtils.cast(Size.class, graphicProperties.getPropertyValue("ActualSize"));
            // Compute the width and height if not specified, preserving aspect ratio.
            if (specifiedSize.Width == 0 && specifiedSize.Height == 0) {
                specifiedSize.Width = size.Width;
                specifiedSize.Height = size.Height;
            } else if (specifiedSize.Width == 0) {
                specifiedSize.Width = specifiedSize.Height * size.Width / size.Height;
            } else if (specifiedSize.Height == 0) {
                specifiedSize.Height = specifiedSize.Width * size.Height / size.Width;
            }
            graphicProperties.setPropertyValue("Size", specifiedSize);
        } catch (Exception e) {
            // Ignore this image.
        }
    }
}
