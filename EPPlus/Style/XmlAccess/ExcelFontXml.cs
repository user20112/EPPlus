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
using System;

using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for fonts
    /// </summary>
    public sealed class ExcelFontXml : StyleXmlHelper
    {
        private const string _colorPath = "d:color";

        private const string boldPath = "d:b";

        private const string familyPath = "d:family/@val";

        private const string italicPath = "d:i";

        private const string namePath = "d:name/@val";

        private const string schemePath = "d:scheme/@val";

        private const string sizePath = "d:sz/@val";

        private const string strikePath = "d:strike";

        private const string underLinedPath = "d:u";

        private const string verticalAlignPath = "d:vertAlign/@val";

        private bool _bold;

        private ExcelColorXml _color = null;

        private int _family;

        private bool _italic;

        private string _name;

        private string _scheme = "";

        private float _size;

        private ExcelUnderLineType _underlineType;

        private string _verticalAlign;

        internal ExcelFontXml(XmlNamespaceManager nameSpaceManager)
                                                                                                                                                                            : base(nameSpaceManager)
        {
            _name = "";
            _size = 0;
            _family = int.MinValue;
            _scheme = "";
            _color = _color = new ExcelColorXml(NameSpaceManager);
            _bold = false;
            _italic = false;
            _underlineType = ExcelUnderLineType.None;
            _verticalAlign = "";
        }

        internal ExcelFontXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _name = GetXmlNodeString(namePath);
            _size = (float)GetXmlNodeDecimal(sizePath);
            _family = GetXmlNodeIntNull(familyPath) ?? int.MinValue;
            _scheme = GetXmlNodeString(schemePath);
            _color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
            _bold = GetBoolValue(topNode, boldPath);
            _italic = GetBoolValue(topNode, italicPath);
            _verticalAlign = GetXmlNodeString(verticalAlignPath);
            if (topNode.SelectSingleNode(underLinedPath, NameSpaceManager) != null)
            {
                string ut = GetXmlNodeString(underLinedPath + "/@val");
                if (ut == "")
                {
                    _underlineType = ExcelUnderLineType.Single;
                }
                else
                {
                    _underlineType = (ExcelUnderLineType)Enum.Parse(typeof(ExcelUnderLineType), ut, true);
                }
            }
            else
            {
                _underlineType = ExcelUnderLineType.None;
            }
        }

        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return _bold;
            }
            set
            {
                _bold = value;
            }
        }

        /// <summary>
        /// Text color
        /// </summary>
        public ExcelColorXml Color
        {
            get
            {
                return _color;
            }
            internal set
            {
                _color = value;
            }
        }

        /// <summary>
        /// Font family
        /// </summary>
        public int Family
        {
            get
            {
                return (_family == int.MinValue ? 0 : _family); ;
            }
            set
            {
                _family = value;
            }
        }

        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return _italic;
            }
            set
            {
                _italic = value;
            }
        }

        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                Scheme = "";        //Reset schema to avoid corrupt file if unsupported font is selected.
                _name = value;
            }
        }

        /// <summary>
        /// Font Scheme
        /// </summary>
        public string Scheme
        {
            get
            {
                return _scheme;
            }
            private set
            {
                _scheme = value;
            }
        }

        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return _size;
            }
            set
            {
                _size = value;
            }
        }

        /// <summary>
        /// Vertical aligned
        /// </summary>
        public string VerticalAlign
        {
            get
            {
                return _verticalAlign;
            }
            set
            {
                _verticalAlign = value;
            }
        }

        internal override string Id
        {
            get
            {
                return Name + "|" + Size + "|" + Family + "|" + Color.Id + "|" + Scheme + "|" + Bold.ToString() + "|" + Italic.ToString() + "|" + false.ToString() + "|" + VerticalAlign + "|" + 0.ToString();
            }
        }

        public static float GetFontHeight(string name, float size)
        {
            name = name.StartsWith("@") ? name.Substring(1) : name;
            if (FontSize.FontHeights.ContainsKey(name))
            {
                return GetHeightByName(name, size);
            }
            else
            {
                return GetHeightByName("Calibri", size);
            }
        }

        public void SetFromFont(SKFont Font)
        {
            Name = Font.Typeface.FamilyName;
            Size = (int)Font.Size;
            Bold = Font.Typeface.IsBold;
            Italic = Font.Typeface.IsItalic;
        }

        internal ExcelFontXml Copy()
        {
            ExcelFontXml newFont = new ExcelFontXml(NameSpaceManager);
            newFont.Name = _name;
            newFont.Size = _size;
            newFont.Family = _family;
            newFont.Scheme = _scheme;
            newFont.Bold = _bold;
            newFont.Italic = _italic;
            newFont.VerticalAlign = _verticalAlign;
            newFont.Color = Color.Copy();
            return newFont;
        }

        internal override XmlNode CreateXmlNode(XmlNode topElement)
        {
            TopNode = topElement;

            if (_bold) CreateNode(boldPath); else DeleteAllNode(boldPath);
            if (_italic) CreateNode(italicPath); else DeleteAllNode(italicPath);
            if (false) CreateNode(strikePath); else DeleteAllNode(strikePath);

            if (_underlineType == ExcelUnderLineType.None)
            {
                DeleteAllNode(underLinedPath);
            }
            else if (_underlineType == ExcelUnderLineType.Single)
            {
                CreateNode(underLinedPath);
            }
            else
            {
                var v = _underlineType.ToString();
                SetXmlNodeString(underLinedPath + "/@val", v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1));
            }

            if (_verticalAlign != "") SetXmlNodeString(verticalAlignPath, _verticalAlign.ToString());
            if (_size > 0) SetXmlNodeString(sizePath, _size.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (_color.Exists)
            {
                CreateNode(_colorPath);
                TopNode.AppendChild(_color.CreateXmlNode(TopNode.SelectSingleNode(_colorPath, NameSpaceManager)));
            }
            if (!string.IsNullOrEmpty(_name)) SetXmlNodeString(namePath, _name);
            if (_family > int.MinValue) SetXmlNodeString(familyPath, _family.ToString());
            if (_scheme != "") SetXmlNodeString(schemePath, _scheme.ToString());

            return TopNode;
        }

        private static float GetHeightByName(string name, float size)
        {
            if (FontSize.FontHeights[name].ContainsKey(size))
            {
                return FontSize.FontHeights[name][size].Height;
            }
            else
            {
                float min = -1, max = float.MaxValue;
                foreach (var h in FontSize.FontHeights[name])
                {
                    if (min < h.Key && h.Key < size)
                    {
                        min = h.Key;
                    }
                    if (max > h.Key && h.Key > size)
                    {
                        max = h.Key;
                    }
                }
                if (min == max || max == float.MaxValue)
                {
                    return Convert.ToSingle(FontSize.FontHeights[name][min].Height);
                }
                else if (min == -1)
                {
                    return Convert.ToSingle(FontSize.FontHeights[name][max].Height);
                }
                else
                {
                    return Convert.ToSingle(FontSize.FontHeights[name][min].Height + (FontSize.FontHeights[name][max].Height - FontSize.FontHeights[name][min].Height) * ((size - min) / (max - min)));
                }
            }
        }
    }
}