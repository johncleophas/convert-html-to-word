using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace html2word
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start Test Word Doc");

            Console.Write("Enter the source path:");

            var sourcePath = Console.ReadLine();

            saveWordDoc(sourcePath);

            Console.WriteLine("End Test Word Doc");

            Console.WriteLine("Hello World!");
        }

        static void saveWordDoc(string sourcepath)
        {

            MemoryStream ms;
            MainDocumentPart mainPart;
            Body b;
            Document d;
            AlternativeFormatImportPart chunk;
            AltChunk altChunk;

            string altChunkID = "AltChunkId1";
            ms = new MemoryStream();


            using (var myDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                mainPart = myDoc.MainDocumentPart;

                if (mainPart == null)
                {
                    mainPart = myDoc.AddMainDocumentPart();
                    b = new Body();
                    d = new Document(b);
                    d.Save(mainPart);
                }

                chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, altChunkID);

                string file = @$"{sourcepath}\test-file.html";



                using (Stream chunkStream = chunk.GetStream(FileMode.Create, FileAccess.Write))
                {
                    var htmText = string.Empty;

                    if (File.Exists(file))
                    {
                        htmText = File.ReadAllText(file);
                        Console.WriteLine(htmText);
                    }

                    using (StreamWriter stringStream = new StreamWriter(chunkStream))
                    {
                        stringStream.Write(htmText);
                    }
                }

                altChunk = new AltChunk();
                altChunk.Id = altChunkID;
                mainPart.Document.Body.InsertAt(altChunk, 0);
                mainPart.Document.Save();

                myDoc.SaveAs(@$"{sourcepath}\test.docx");
            }

        }
    }
}
