<!-- default badges list -->
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E1758)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/RichEditPopupMenu/Form1.cs) (VB: [Form1.vb](./VB/RichEditPopupMenu/Form1.vb))
<!-- default file list end -->
# How to integrate SpellChecker menu items into the XtraRichEdit context menu


<p>This example demonstrates the technique used to customize the context menu of the RichEditControl.  <br />
Although recent versions of the RichEditControl are tightly integrated with our spell checking component, occasionally you may find it useful to add  spelling check menu items manually.</p><p>The <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditRichEditControl_PreparePopupMenutopic"><u>PreparePopupMenu</u></a> event of the RichEditControl is handled to insert new menu items, which are the suggested replacement words for the word where the caret is located. To obtain a list of suggestions, the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraSpellCheckerSpellChecker_GetSuggestionstopic"><u>GetSuggestions</u></a> method of the XtraSpellChecker is used. When menu item representing a suggested word is clicked, the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditAPINativeDocument_Replacetopic"><u>Replace</u></a> method substitutes the misspelled word for the selected suggestion.</p><p>Note that to implement spell checking functionality and display suggested words in a context menu, you may just drop the <strong>SpellChecker </strong>component on the form and set its <strong>SpellCheckMode </strong>to <strong>AsYouType </strong>value.</p>

<br/>


