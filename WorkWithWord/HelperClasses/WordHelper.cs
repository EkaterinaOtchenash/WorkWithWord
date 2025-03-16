using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WorkWithWord.HelperClasses
{
    public class WordHelper
    {
        private FileInfo _fileinfo;

        public WordHelper(string filename)
        {
            if (File.Exists(filename))
            {
                _fileinfo = new FileInfo(filename);
            }
            else
            {
                throw new FileNotFoundException();
            }
        }

        public void Process(Dictionary<string, string> items, string path)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                object file = _fileinfo.FullName;
                object missing = Type.Missing;
                app.Documents.Open(file);

                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;
                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object replace = Word.WdReplace.wdReplaceAll;
                    find.Execute(
                        FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing,
                        Replace: replace);
                }

                app.ActiveDocument.SaveAs2(path);
                app.ActiveDocument.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
        }
    }
}