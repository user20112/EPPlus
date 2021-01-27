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
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/

using SkiaSharp;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFont : XmlHelper
    {
        private string _boldPath = "@b";
        private string _colorPath = "a:solidFill/a:srgbClr/@val";
        private string _fontCsPath = "a:cs/@typeface";
        private string _fontLatinPath = "a:latin/@typeface";
        private string _italicPath = "@i";
        private string _path;
        private XmlNode _rootNode;
        private string _sizePath = "@sz";

        private string _strikePath = "@strike";

        private string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";

        private string _underLinePath = "@u";

        internal ExcelTextFont(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
                                            : base(namespaceManager, rootNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _rootNode = rootNode;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    TopNode = node;
                }
            }
            _path = path;
        }

        public bool Bold
        {
            get
            {
                return GetXmlNodeBool(_boldPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_boldPath, value ? "1" : "0");
            }
        }

        public SKColor Color
        {
            get
            {
                string col = GetXmlNodeString(_colorPath);
                if (col == "")
                {
                    return SKColors.Empty;
                }
                else
                {
                    return new SKColor((uint.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier)));
                }
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_colorPath, ((uint)value).ToString("X").Substring(2, 6));
            }
        }

        public string ComplexFont
        {
            get
            {
                return GetXmlNodeString(_fontCsPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_fontCsPath, value);
            }
        }

        public bool Italic
        {
            get
            {
                return GetXmlNodeBool(_italicPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_italicPath, value ? "1" : "0");
            }
        }

        public string LatinFont
        {
            get
            {
                return GetXmlNodeString(_fontLatinPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_fontLatinPath, value);
            }
        }

        public float Size
        {
            get
            {
                return GetXmlNodeInt(_sizePath) / 100;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
            }
        }

        /// <summary>
        /// Set the font style from a font object
        /// </summary>
        /// <param name="Font"></param>
        public void SetFromFont(SKFont Font)
        {
            LatinFont = Font.Typeface.FamilyName;
            ComplexFont = Font.Typeface.FamilyName;
            Size = Font.Size;
            if (Font.Typeface.IsBold)
                Bold = Font.Typeface.IsBold;
            if (Font.Typeface.IsItalic) Italic = Font.Typeface.IsItalic;
        }

        protected internal void CreateTopNode()
        {
            if (_path != "" && TopNode == _rootNode)
            {
                CreateNode(_path);
                TopNode = _rootNode.SelectSingleNode(_path, NameSpaceManager);
            }
        }
    }
}