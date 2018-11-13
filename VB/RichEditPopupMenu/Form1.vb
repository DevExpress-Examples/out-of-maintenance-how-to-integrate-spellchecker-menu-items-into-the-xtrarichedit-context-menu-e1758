Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Forms
#Region "#usings"
Imports DevExpress.Utils.Menu
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Menu
Imports DevExpress.XtraSpellChecker
#End Region ' #usings

Namespace RichEditPopupMenuExample
	Partial Public Class Form1
		Inherits Form
		Private Shared wordSeparators As List(Of Char) = CreateWordSeparators()

		Private Shared Function CreateWordSeparators() As List(Of Char)
			Dim result As New List(Of Char)()
			result.Add(" "c)
			result.Add(ControlChars.Tab)
			result.Add(ControlChars.Lf)
			result.Add(ControlChars.Cr)
			result.Add(","c)
			result.Add("."c)
			result.Add("-"c)
			result.Add("("c)
			result.Add(")"c)
			result.Add("{"c)
			result.Add("}"c)
			result.Add("["c)
			result.Add("]"c)
			result.Add(""""c)
			result.Add("'"c)
			result.Add("<"c)
			result.Add(">"c)
			result.Add(":"c)
			result.Add(";"c)
			result.Add("\"c)
			result.Add("/"c)
			Return result
		End Function

		Public Sub New()
			InitializeComponent()
			AddHandler richEditControl1.PopupMenuShowing, AddressOf richEditControl1_PopupMenuShowing
		End Sub

#Region "#popupmenushowing"
Private Sub richEditControl1_PopupMenuShowing(ByVal sender As Object,  _
ByVal e As DevExpress.XtraRichEdit.PopupMenuShowingEventArgs)
	If spellChecker1.SpellCheckMode = SpellCheckMode.OnDemand Then
		Dim pos As DocumentPosition = richEditControl1.Document.CaretPosition
		Dim wordEnd As Integer = GetWordEndIndex(pos)
		Dim wordStart As Integer = GetWordStartIndex(pos)
		If wordEnd <= wordStart Then
			Return
		End If
		Dim range As DocumentRange =  _
richEditControl1.Document.CreateRange(wordStart, wordEnd - wordStart)
		Dim word As String = richEditControl1.Document.GetText(range)
		If spellChecker1.IsMisspelledWord(word, spellChecker1.Culture) Then
			CreateMenuItems(e.Menu, range, word)
		End If
	End If
End Sub
Private Sub CreateMenuItems(ByVal menu As RichEditPopupMenu, ByVal range As DocumentRange, ByVal word As String)
	Dim suggestions As SuggestionCollection = Me.spellChecker1.GetSuggestions(word)
	Dim count As Integer = suggestions.Count
	If count > 0 Then
		Dim lastIndex As Integer = Math.Min(count - 1, 5)
		For i As Integer = lastIndex To 0 Step -1
		   Dim suggestion As SuggestionBase = suggestions(i)
		   Dim item As New SuggestionMenuItem(Me.richEditControl1.Document, suggestion.Suggestion, range)
		   item.Image = Resources.suggestion
		   menu.Items.Insert(0, item)
		Next i
	Else
		Dim emptyItem As New DXMenuItem("no spelling suggestions")
		emptyItem.Enabled = False
		menu.Items.Insert(0, emptyItem)
	End If
End Sub
#End Region ' #popupmenushowing

		Private Function GetCharacter(ByVal position As Integer) As Char
			Dim range As DocumentRange = Me.richEditControl1.Document.CreateRange(position, 1)
			Return Me.richEditControl1.Document.GetText(range)(0)
		End Function
		Private Function GetWordEndIndex(ByVal position As DocumentPosition) As Integer
			Dim currentPosition As Integer = position.ToInt()
			'int result = currentPosition;
			Dim endPosition As Integer = Me.richEditControl1.Document.Range.End.ToInt()-1
			Do While currentPosition <= endPosition _
 AndAlso Not wordSeparators.Contains(GetCharacter(currentPosition))
				currentPosition += 1
				'result = currentPosition;
			Loop
			Return currentPosition
		End Function
		Private Function GetWordStartIndex(ByVal position As DocumentPosition) As Integer
			Dim currentPosition As Integer = position.ToInt()
			Dim result As Integer = currentPosition
			Do While currentPosition >= 0 AndAlso Not wordSeparators.Contains(GetCharacter(currentPosition))
				result = currentPosition
				currentPosition -= 1
			Loop
			Return result
		End Function
		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			Dim usCulture As New CultureInfo("en-US")
			Dim dictionary1 As New SpellCheckerISpellDictionary("..\..\american.xlg", "..\..\english.aff", usCulture)
			dictionary1.AlphabetPath = "..\..\EnglishAlphabet.txt"
			dictionary1.Load()
			Me.spellChecker1.Dictionaries.Add(dictionary1)

			Me.richEditControl1.Text = "Accordnig to an englnsih unviersitry sutdy the oredr of letetrs in a word dosen't mttaer, the olny thnig thta's imporantt is that the frsit and last ltteer of eevry word is in the crrecot psoition. The rset can be jmbueld and one is stlil able to read the txet withuot dificultfiy."
			Me.labelControl1.Text = "Right-click a misspelled word for a context menu containing a list of suggested replacements"
		End Sub

	End Class
#Region "#menuitem"
Public Class SuggestionMenuItem
	Inherits DXMenuItem
	Private ReadOnly document As Document
	Private ReadOnly range As DocumentRange

	Public Sub New(ByVal document As Document, ByVal suggestion As String, ByVal range As DocumentRange)
		MyBase.New(suggestion)
		Me.document = document
		Me.range = range
	End Sub

	Protected Overrides Sub OnClick()
		document.Replace(range, Caption)
	End Sub
End Class
#End Region ' #menuitem
End Namespace
