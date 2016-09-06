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

import com.sun.star.lang.XComponent;

/**
 * Interface used to filter Office documents.
 * 
 * @version $Id$
 * @since 3.1-xwiki
 */
public interface OfficeDocumentFilter
{
    /**
     * Filter the given office document in the provided context.
     * 
     * @param document the office document to filter
     * @param context the office context, can be used to access office services
     * @throws OfficeException if filtering the office document fails
     */
    void filter(XComponent document, OfficeContext context) throws OfficeException;
}
