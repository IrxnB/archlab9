using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace archlabab9
{
    internal class TitlePageFormatter
    {
        private object _template = "C:\\Users\\Андрей Лузгин\\OneDrive\\Desktop\\учеба\\архитектура ИС\\9\\template.doc";




        internal void Format(string dest, TitlePageData data)
        {
            var word = new Word.Application();

            var doc = word.Documents.Add(ref _template, false, Word.WdNewDocumentType.wdNewBlankDocument, true);

            doc.Activate();

            ReplaceBookmark("workType", data.WorkType);
            Replace("{Номер}", data.WorkNumber);
            Replace("{Название}", data.Title);
            Replace("{Дисциплина}", data.Discilpline);
            Replace("{Преподаватель}", data.Teacher);
            Replace("{Год}", data.Year);

            TrySave();
            word.Quit(Word.WdSaveOptions.wdPromptToSaveChanges);


            void Replace(string toReplace, string replacement)
            {
                var range = doc.StoryRanges[Word.WdStoryType.wdMainTextStory];
                range.Find.ClearFormatting();

                range.Find.Execute(FindText: toReplace, ReplaceWith: replacement);
                TrySave();
            }

            void ReplaceBookmark(string bookmark, string replacement)
            {
                doc.Bookmarks[bookmark].Range.Text = replacement;
                TrySave();
            }

            void TrySave()
            {
                try
                {
                    doc.SaveAs(dest, Word.WdSaveFormat.wdFormatDocument);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        
    }


    
}
