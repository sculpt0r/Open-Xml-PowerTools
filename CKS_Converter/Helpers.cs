﻿using System;
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

        private static void GetComments(WordprocessingDocument doc)
        {
            var xDoc = doc.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
            var commentNodes = xDoc.Descendants().Elements(W.comment);
            commentNodes.ToList().ForEach(cn => {
                Console.WriteLine(cn);
                Console.WriteLine("*** *** ***");


                commentNodes.Descendants().Elements(W.r).ToList().ForEach(Console.WriteLine);
                Console.WriteLine("*** *** ***");
                commentNodes.Descendants().Elements(W.t).ToList().ForEach(x => Console.WriteLine(x.Value));
            });
            // string comments = null;

            // // Read the comments using a stream reader.  
            // using (StreamReader streamReader =
            //     new StreamReader(xDoc.GetStream()))
            // {
            //     comments = streamReader.ReadToEnd();
            // }
            // Console.WriteLine(comments);
        }
    }
}