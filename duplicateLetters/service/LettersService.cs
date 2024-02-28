using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using duplicateLetters.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace duplicateLetters.service
{
    public class LettersService
    {
        public LettersService() { }

        public void DuplicateBlock(string filePath, DataJson dataJson)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Dictionary<string, string> dictGeneral = new Dictionary<string, string>
                    {
                        {"@nomedapessoa", dataJson.Nomedapessoa},
                        {"@DATADEADMISSÃO", dataJson.Datadeadmissao},
                        {"@nomedogestor", dataJson.Gestor}
                    };
                ReplaceMarking(doc, dictGeneral);
                Body body = doc.MainDocumentPart.Document.Body;

                var startParagraph = body.Elements<Paragraph>().FirstOrDefault(para => para.InnerText.Contains("@start"));
                var endParagraph = body.Elements<Paragraph>().FirstOrDefault(para => para.InnerText.Contains("@end"));
                if (startParagraph != null && endParagraph != null)
                {

                    var elementsToDuplicate = body.Elements()
                        .SkipWhile(el => el != startParagraph)
                        .TakeWhile(el => el != endParagraph)
                        .Where(el => el != startParagraph && el != endParagraph)
                        .ToList();

                    var elementsToRemove = body.Elements()
                        .SkipWhile(el => el != startParagraph)
                        .TakeWhile(el => el != endParagraph)
                        .ToList();
                    // Remove the original block
                    foreach (var elementToRemove in elementsToRemove)
                    {
                        elementToRemove.Remove();
                    }

                    for (int i = 0; i < dataJson.Periodo.Count; i++)
                    {
                        Dictionary<string, string> dictPeriod = new Dictionary<string, string>
                    {
                        {"@period", $"Período {i +1}"},
                        {"@AREADAEMPRESA", dataJson.Periodo[i].Nome},
                        {"@CARGODAPESSOA", dataJson.Periodo[i].Data}
                    };

                        var clonedElements = elementsToDuplicate.Select(el => ReplaceAndCloneElement(el, dictPeriod)).ToList();

                        if(i > 0)
                        {
                            // Created page
                            Paragraph pageSeparator = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));

                            // Insert the page separator before the endParagraph
                            body.InsertBefore(pageSeparator, endParagraph);

                        }



                        foreach (var clonedElement in clonedElements)
                        {
                            body.InsertBefore(clonedElement, endParagraph);
                        }
                    }
                }
            }
        }

        static void ReplaceMarking(WordprocessingDocument documento, Dictionary<string, string> dictGeneral)
        {
            foreach (var texto in documento.MainDocumentPart.Document.Descendants<Text>())
            {
                foreach (var item in dictGeneral)
                {
                    if (texto.Text.Contains(item.Key))
                    {
                        texto.Text = texto.Text.Replace(item.Key, item.Value);
                    }

                }
            }
        }

        static OpenXmlElement ReplaceAndCloneElement(OpenXmlElement element, Dictionary<string, string> replacements)
        {
            var clonedElement = CloneElement(element);

            foreach (var run in clonedElement.Descendants<Run>())
            {
                foreach (var text in run.Descendants<Text>())
                {
                    foreach (var replacement in replacements)
                    {
                        text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                    }
                }
            }

            return clonedElement;
        }

        static OpenXmlElement CloneElement(OpenXmlElement element)
        {
            return element.CloneNode(true);
        }
    }
}
