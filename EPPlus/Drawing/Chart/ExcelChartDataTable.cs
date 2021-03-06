﻿/*******************************************************************************
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
 * Mark Kromis		Added		2017-01-07
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Datatable on chart level.
    /// </summary>
    public class ExcelChartDataTable : XmlHelper
    {
        private const string showHorzBorderPath = "c:showHorzBorder/@val";

        private const string showKeysPath = "c:showKeys/@val";

        private const string showOutlinePath = "c:showOutline/@val";

        private const string showVertBorderPath = "c:showVertBorder/@val";

        private ExcelDrawingBorder _border = null;

        private ExcelDrawingFill _fill = null;

        private ExcelTextFont _font = null;

        private string[] _paragraphSchemaOrder = new string[] { "spPr", "txPr", "dLblPos", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };

        internal ExcelChartDataTable(XmlNamespaceManager ns, XmlNode node)
                                                                           : base(ns,node)
       {
           XmlNode topNode = node.SelectSingleNode("c:dTable", NameSpaceManager);
           if (topNode == null)
           {
               topNode = node.OwnerDocument.CreateElement("c", "dTable", ExcelPackage.schemaChart);
               //node.InsertAfter(_topNode, node.SelectSingleNode("c:order", NameSpaceManager));
               InserAfter(node, "c:valAx,c:catAx", topNode);
               SchemaNodeOrder = new string[] { "dTable", "showHorzBorder", "showVertBorder", "showOutline", "showKeys", "spPr", "txPr" };
               topNode.InnerXml = "<c:showHorzBorder val=\"1\"/><c:showVertBorder val=\"1\"/><c:showOutline val=\"1\"/><c:showKeys val=\"1\"/>";
           }
           TopNode = topNode;
       }

        /// <summary>
        /// Access border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                }
                return _border;
            }
        }

        /// <summary>
        /// Access fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
                }
                return _fill;
            }
        }

        /// <summary>
        /// Access font properties
        /// </summary>
        public ExcelTextFont Font
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
                    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", _paragraphSchemaOrder);
                }
                return _font;
            }
        }

        /// <summary>
        /// The horizontal borders shall be shown in the data table
        /// </summary>
        public bool ShowHorizontalBorder
        {
           get
           {
               return GetXmlNodeBool(showHorzBorderPath);
           }
           set
           {
               SetXmlNodeString(showHorzBorderPath, value ? "1" : "0");
           }
       }

        /// <summary>
        /// The legend keys shall be shown in the data table
        /// </summary>
        public bool ShowKeys
        {
            get
            {
                return GetXmlNodeBool(showKeysPath);
            }
            set
            {
                SetXmlNodeString(showKeysPath, value ? "1" : "0");
            }
        }

        /// <summary>
        /// The outline shall be shown on the data table
        /// </summary>
        public bool ShowOutline
        {
            get
            {
                return GetXmlNodeBool(showOutlinePath);
            }
            set
            {
                SetXmlNodeString(showOutlinePath, value ? "1" : "0");
            }
        }

        /// <summary>
        /// The vertical borders shall be shown in the data table
        /// </summary>
        public bool ShowVerticalBorder
        {
            get
            {
                return GetXmlNodeBool(showVertBorderPath);
            }
            set
            {
                SetXmlNodeString(showVertBorderPath, value ? "1" : "0");
            }
        }

        protected eLabelPosition GetPosEnum(string pos)
        {
            switch (pos)
            {
                case "b":
                    return eLabelPosition.Bottom;

                case "ctr":
                    return eLabelPosition.Center;

                case "inBase":
                    return eLabelPosition.InBase;

                case "inEnd":
                    return eLabelPosition.InEnd;

                case "l":
                    return eLabelPosition.Left;

                case "r":
                    return eLabelPosition.Right;

                case "t":
                    return eLabelPosition.Top;

                case "outEnd":
                    return eLabelPosition.OutEnd;

                default:
                    return eLabelPosition.BestFit;
            }
        }

        protected string GetPosText(eLabelPosition pos)
        {
            switch (pos)
            {
                case eLabelPosition.Bottom:
                    return "b";

                case eLabelPosition.Center:
                    return "ctr";

                case eLabelPosition.InBase:
                    return "inBase";

                case eLabelPosition.InEnd:
                    return "inEnd";

                case eLabelPosition.Left:
                    return "l";

                case eLabelPosition.Right:
                    return "r";

                case eLabelPosition.Top:
                    return "t";

                case eLabelPosition.OutEnd:
                    return "outEnd";

                default:
                    return "bestFit";
            }
        }
    }
}