using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using DevExpress.Utils.Menu;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Menu;
using DevExpress.XtraSpellChecker;

namespace RichEditPopupMenuExample {
    public partial class Form1 : Form {
        static List<char> wordSeparators = CreateWordSeparators();

        static List<char> CreateWordSeparators() {
            List<char> result = new List<char>();
            result.Add(' ');
            result.Add('\t');
            result.Add('\n');
            result.Add('\r');
            result.Add(',');
            result.Add('.');
            result.Add('-');
            result.Add('(');
            result.Add(')');
            result.Add('{');
            result.Add('}');
            result.Add('[');
            result.Add(']');
            result.Add('"');
            result.Add('\'');
            result.Add('<');
            result.Add('>');
            result.Add(':');
            result.Add(';');
            result.Add('\\');
            result.Add('/');
            return result;
        }

        public Form1() {
            InitializeComponent();
            this.richEditControl1.PreparePopupMenu += new PreparePopupMenuEventHandler(richEditControl1_PreparePopupMenu);
        }

        void richEditControl1_PreparePopupMenu(object sender, PreparePopupMenuEventArgs e) {
            if(spellChecker1.SpellCheckMode == SpellCheckMode.OnDemand) {
                DocumentPosition pos = this.richEditControl1.Document.CaretPosition;
                int wordEnd = GetWordEndIndex(pos);
                int wordStart = GetWordStartIndex(pos);
                if(wordEnd <= wordStart)
                    return;
                DocumentRange range = this.richEditControl1.Document.CreateRange(wordStart, wordEnd - wordStart);
                string word = this.richEditControl1.Document.GetText(range);
                if(this.spellChecker1.IsMisspelledWord(word, this.spellChecker1.Culture))
                    CreateMenuItems(e.Menu, range, word);
            }
        }
        void CreateMenuItems(RichEditPopupMenu menu, DocumentRange range, string word) {
            SuggestionCollection suggestions = this.spellChecker1.GetSuggestions(word);
            int count = suggestions.Count;
            if (count > 0) {
                int lastIndex = Math.Min(count - 1, 5);
                for (int i = lastIndex; i >= 0; i--) {
                    SuggestionBase suggestion = suggestions[i];
                    SuggestionMenuItem item = new SuggestionMenuItem(this.richEditControl1.Document, suggestion.Suggestion, range);
                    item.Image = RichEditPopupMenuExample.Properties.Resources.suggestion;
                    menu.Items.Insert(0, item);
                }
            }
            else {
                DXMenuItem emptyItem = new DXMenuItem("no spelling suggestions");
                emptyItem.Enabled = false;
                menu.Items.Insert(0, emptyItem);
            }
        }
        char GetCharacter(int position) {
            DocumentRange range = this.richEditControl1.Document.CreateRange(position, 1);
            return this.richEditControl1.Document.GetText(range)[0];
        }
        int GetWordEndIndex(DocumentPosition position) {
            int currentPosition = position.ToInt();
            //int result = currentPosition;
            int endPosition = this.richEditControl1.Document.Range.End.ToInt()-1;
            while (currentPosition <= endPosition && !wordSeparators.Contains(GetCharacter(currentPosition))) {
                currentPosition++;
                //result = currentPosition;
            }
            return currentPosition;
        }
        int GetWordStartIndex(DocumentPosition position) {
            int currentPosition = position.ToInt();
            int result = currentPosition;
            while (currentPosition >= 0 && !wordSeparators.Contains(GetCharacter(currentPosition))) {
                result = currentPosition;
                currentPosition--;
            }
            return result;
        }
        void Form1_Load(object sender, EventArgs e) {
            CultureInfo usCulture = new CultureInfo("en-US");
            SpellCheckerISpellDictionary dictionary1 = new SpellCheckerISpellDictionary(@"..\..\american.xlg", @"..\..\english.aff", usCulture);
            dictionary1.AlphabetPath = @"..\..\EnglishAlphabet.txt";
            dictionary1.Load();
            this.sharedDictionaryStorage1.Dictionaries.Add(dictionary1);

            this.richEditControl1.Text = @"Accordnig to an englnsih unviersitry sutdy the oredr of letetrs in a word dosen't mttaer, the olny thnig thta's imporantt is that the frsit and last ltteer of eevry word is in the crrecot psoition. The rset can be jmbueld and one is stlil able to read the txet withuot dificultfiy.";
            this.labelControl1.Text = "Right-click a misspelled word for a context menu containing a list of suggested replacements"; 
        }
    }

    public class SuggestionMenuItem : DXMenuItem {
        readonly Document document;
        readonly DocumentRange range;

        public SuggestionMenuItem(Document document, string suggestion, DocumentRange range)
            : base(suggestion) {
            this.document = document;
            this.range = range;
        }

        protected override void OnClick() {
            document.Replace(range, Caption);
        }
    }
}