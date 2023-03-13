using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO.DLS;

namespace ReadWordDocument
{
    class Program
    {
        private const string DIRECTORY_PATH = "C:\\Users\\kenan\\OneDrive\\Documents\\Softweare\\TemplateFolder Reduced";
        private static List<string> GetFileNames()
        {
            DirectoryInfo directory = new DirectoryInfo(DIRECTORY_PATH);
            FileInfo[] files = directory.GetFiles();

            List<string> fileNames = new List<string>();
            foreach (FileInfo f in files)
            {
                fileNames.Add(f.Name);
            }

            return fileNames;
        }

        private static SortedSet<string> GetFileTags(string fileName)
        {
            SortedSet<string> tags = new SortedSet<string>();

            using (WordDocument document = new WordDocument(DIRECTORY_PATH + "\\" + fileName))
            {
                //gets the word document text
                string text = document.GetText();

                for (int i = 0; i < text.Length; i++)
                {
                    if (text[i] == '{')
                    {
                        i++;
                        string tag = "";
                        while (i < text.Length && text[i] != '}')
                        {
                            tag += text[i];
                            i++;
                        }
                        tags.Add("\t" + tag);
                    }
                }
            }

            return tags;
        }
        static void Main(string[] args)
        {
            Console.WriteLine("Reading file names...");
            var fileNames = GetFileNames();
            List<string> filteredTags = new List<string>();

            Console.WriteLine("Reading tags...");
            int num = 0;

            foreach (string f in fileNames)
            {
                filteredTags.Add("* " + f);
                filteredTags.AddRange(GetFileTags(f));

                Console.WriteLine(num + ". file completed!");
                num++;
            }

            Console.WriteLine("Succesfull");
            Console.WriteLine("Number of tags: " + filteredTags.Count);

            Console.WriteLine("Writing in tags.txt");
            System.IO.File.WriteAllLines(DIRECTORY_PATH + "\\tags.txt", filteredTags);
            Console.WriteLine("Succesfully written!");
            Console.ReadLine();
        }
    }
}
