// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace CKS_Converter
{
    public static class Helpers
    {
        #region Extensions
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null) return partXDocument;

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                        partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }
        #endregion

        public static MetaComment[] GetComments(WordprocessingDocument doc)
        {
            var xDoc = doc.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();

            var commentNodes = xDoc.Descendants().Elements(XNameExtension.XElem("comment"));
            var metaComments = commentNodes.ToList().Select(cn =>
            {
                int id = -1;
                int.TryParse(cn.Attribute(XNameExtension.Xattr("id")).Value, out id);

                var author = cn.Attribute(XNameExtension.Xattr("author")).Value;
                var date = cn.Attribute(XNameExtension.Xattr("date")).Value;
                var initials = cn.Attribute(XNameExtension.Xattr("initials")).Value;
                var content = cn.Descendants().Elements(XNameExtension.XElem("t")).FirstOrDefault().Value;

                Console.WriteLine(id);
                Console.WriteLine(author);
                Console.WriteLine(date);
                Console.WriteLine("*** *** ***");
                var mc = new MetaComment(id, author, date, initials, content);

                return mc;


                // commentNodes.Descendants().Elements(W.r).ToList().ForEach(Console.WriteLine);
                // Console.WriteLine("*** *** ***");
                // commentNodes.Descendants().Elements(W.t).ToList().ForEach(x => Console.WriteLine(x.Value));
            });

            return metaComments.ToArray();
        }
    }

    public class MetaComment
    {
        private int id;
        private string author;
        private string date;
        private string initials;

        private string value;

        public MetaComment(int id, string author, string date, string initials, string value)
        {
            this.id = id;
            this.author = author;
            this.date = date;
            this.initials = initials;
            this.value = value;
        }

        public object toHtml()
        {
            return new XElement(DOMElement.div, id, author, date);
        }
    }

    public static class XNameExtension
    {
        public static readonly XNamespace namSpec =
           "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XName Xattr(string attribute)
        {
            var a = namSpec + attribute;
            return XName.Get(a.ToString());
        }

        public static XName XElem(string name)
        {
            return namSpec + name;
        }
    }

    public static class DOMElement {
        public static readonly XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        public static readonly XName div = xhtml + "div";
    }
}
