using System;
using System.IO;
using Spire.Doc;

namespace SearchWordInDocx
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            try {
                Document document = new Document();
                string path;
                
                if (args.Length > 0)
                {
                    path = args[0];
                }
                else
                {
                    Console.Write("Enter the path of folder that contains the document: ");
                    path = Console.ReadLine();
                }
                path = path.Replace("\"", "");
                
                Console.Write("Enter the word to search: ");
                string word = Console.ReadLine();
                SaveDocumentToStream(document, path, word);
                
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }


        private static void SaveDocumentToStream(Document document, string path, string word)
        {
            MemoryStream finalStream = new MemoryStream();
            string[] files = Directory.GetFiles(path);
            foreach (string file in files)
            {
                document.LoadFromFile(file);
                
                using (MemoryStream stream = new MemoryStream())
                {
                    document.SaveToStream(stream, FileFormat.Txt);
                    stream.Position = 0;
                    stream.CopyTo(finalStream);
                }
            }

            finalStream.Position = 0;
            SearchWord(finalStream, word, files, document);
        }

        private static void SearchWord(MemoryStream stream, string word, string[] files, Document document)
        {
            stream.Position = 0;
            bool found = false;
            foreach (string file in files)
            {
                document.LoadFromFile(file);
                using (MemoryStream docStream = new MemoryStream())
                {
                    document.SaveToStream(docStream, FileFormat.Txt);
                    docStream.Position = 0;
                    StreamReader reader = new StreamReader(docStream);
                    string text = reader.ReadToEnd();
                    
                    if (text.ToLower().Contains(word.ToLower()))
                    {
                        Console.WriteLine("\"" + word + "\"" + " found in document " + Path.GetFileName(file));
                        found = true;
                    }
                }
            }
            
            if (!found)
            {
                Console.WriteLine("The word is not found in any document.");
            }
        }
    }
}
