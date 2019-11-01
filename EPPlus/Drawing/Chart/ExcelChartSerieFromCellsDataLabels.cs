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
    public sealed class ExcelChartSerieFromCellsDataLabels : ExcelChartDataLabel
    {
        internal ExcelScatterChartSerie _scatterChartSerie;

        internal ExcelChartSerieFromCellsDataLabels(XmlNamespaceManager ns, XmlNode node, ExcelScatterChartSerie scatterChartSerie)
    : base(ns, node, true)
        {
            _scatterChartSerie = scatterChartSerie;
            //CreateNode(positionPath);
            //Position = eLabelPosition.Center;
        }

        /// <summary>
        /// Magic string used in each dlBl ext instance
        /// </summary>
        const string dlBlExtString = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";
        /// <summary>
        /// Magic string used in the extLst with datalabelsRange
        /// </summary>
        const string extLstString = "{02D57815-91ED-43cb-92C2-25804820EDAC}";

        const string positionPath = "c:dLblPos/@val";
        /// <summary>
        /// Position of the labels
        /// </summary>
        public eLabelPosition Position
        {
            get
            {
                return GetPosEnum(GetXmlNodeString(positionPath));
            }
            set
            {
                SetXmlNodeString(positionPath, GetPosText(value));
            }
        }

        ExcelTextFont _font = null;
        /// <summary>
        /// Access font properties
        /// </summary>
        public new ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
                    {
                        CreateNode("c:txPr/a:bodyPr");
                        CreateNode("c:txPr/a:lstStyle");
                    }
                    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[] { "spPr", "txPr", "dLblPos", "showVal", "showCatName ", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
                }
                return _font;
            }
        }

        // when setting the range, delete/populate numCache node in xVal/yVal, delete/create the extLst node with datalabelsRange, and delete/create the dLbls node
        string _range;
        public string Range
        {
            get
            {
                return _range;
            }
            set
            {
                _range = value;
                var ws = _scatterChartSerie._chartSeries.Chart.WorkSheet;
                _scatterChartSerie.GenerateRef();
                _scatterChartSerie.GenerateDataLabelsRangeLst(_range, extLstString);
                Generate_dLbls();
            }
        }

        internal void Generate_dLbls()
        {

            ExcelAddress s = new ExcelAddress(Range);

            int count = s._toRow - s._fromRow + 1;

            List<string> dLblList = new List<string>();

            for (int i = 0; i < count; i++)
            {
                ExcelChartSerieFromCellsDataLabel label = new ExcelChartSerieFromCellsDataLabel(NameSpaceManager, TopNode, _scatterChartSerie, i, Guid.NewGuid().ToString().ToUpper());
            }

            TopNode.InnerXml += "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:dLblPos val=\"ctr\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/><c:showLeaderLines val=\"0\"/>";

            XmlElement extLst = TopNode.OwnerDocument.CreateElement("c:extLst", ExcelPackage.schemaChart);
            TopNode.AppendChild(extLst);
            extLst.InnerXml = "<c:ext uri=\"{CE6537A1-D6FC-4f65-9D91-7224C49458BB}\" xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\"><c15:layout/><c15:showDataLabelsRange val=\"1\"/><c15:showLeaderLines val=\"1\"/></c:ext>";

            string test = "asdf";
        }

    }
}
