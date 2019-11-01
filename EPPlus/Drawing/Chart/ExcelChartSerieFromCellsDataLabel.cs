/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 * Garth Todd       Added       2019-07-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// This class is used for information about a single data label instance when using a cell range for the data label text
    /// </summary>
    class ExcelChartSerieFromCellsDataLabel : ExcelChartDataLabel
    {
        // may need a different base constructor so the XML is correctly generated
        internal ExcelChartSerieFromCellsDataLabel(XmlNamespaceManager ns, XmlNode node, ExcelScatterChartSerie scatterChartSerie, int idx, string id)
    : base(ns, node, idx, id)
        {
            
        }

    }
}
