Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Formatting
Imports Microsoft.VisualStudio.Text.Operations
Imports System.Windows
Imports System.Windows.Input


Public Class MouseProcessor
    Inherits MouseProcessorBase


    Private ReadOnly cgView As IWpfTextView
    Private ReadOnly cgNavigator As ITextStructureNavigator


    Public Sub New(
            view As IWpfTextView,
            navigator As ITextStructureNavigator
        )

        cgView = view
        cgNavigator = navigator
    End Sub


    Public Overrides Sub PreprocessMouseLeftButtonDown(e As MouseButtonEventArgs)
        If e.ChangedButton <> MouseButton.Left Then
            Return
        End If

        If e.ClickCount <> 2 Then
            Return
        End If

        ' Only do this special handling if no modifier keys are
        ' pressed. Ctrl-Click and Shift+Click have special behavior,
        ' so Alt+Click can be used to skip this special handling.
        If Keyboard.Modifiers <> ModifierKeys.None Then
            Return
        End If

        ' Try to select the hyphenated word. If we don't find
        ' a hyphenated word, then we won't handle the mouse
        ' click and let the default handling take care of it.
        If SelectHyphenatedSpan(e) Then
            e.Handled = True
        End If
    End Sub


    Private Function SelectHyphenatedSpan(e As MouseButtonEventArgs) As Boolean
        Dim snapshot As SnapshotPoint?


        snapshot = GetSnapshotAtCursor(e)

        If snapshot.HasValue Then
            Dim span As SnapshotSpan?


            span = GetHyphenatedSpan(snapshot.Value)

            If span.HasValue Then
                cgView.Selection.Select(span.Value, False)
                cgView.Caret.MoveTo(cgView.Selection.ActivePoint)
                Return True
            End If
        End If

        Return False
    End Function


    Private Function GetSnapshotAtCursor(e As MouseButtonEventArgs) As SnapshotPoint?
        Dim p As Point
        Dim line As ITextViewLine


        p = GetPositionInViewport(e)

        line = cgView.TextViewLines.GetTextViewLineContainingYCoordinate(p.Y)

        If line IsNot Nothing Then
            Return line.GetBufferPositionFromXCoordinate(p.X, True)
        Else
            Return Nothing
        End If
    End Function


    Private Function GetPositionInViewport(e As MouseButtonEventArgs) As Point
        Dim p As Point


        p = e.GetPosition(cgView.VisualElement)

        Return New Point(p.X + cgView.ViewportLeft, p.Y + cgView.ViewportTop)
    End Function


    Public Function GetHyphenatedSpan(snapshot As SnapshotPoint) As SnapshotSpan?
        Dim extent As TextExtent
        Dim lineExtent As SnapshotSpan
        Dim originalText As String


        lineExtent = snapshot.GetContainingLine().Extent

        ' Get the extent of the word at the cursor.
        extent = cgNavigator.GetExtentOfWord(snapshot)

        originalText = extent.Span.GetText()

        ' If there is any whitespace in the extent that would normally
        ' be selected then we won't do anything because we may end up
        ' expanding the whitespace to include other words.
        If Not originalText.Any(AddressOf Char.IsWhiteSpace) Then
            Dim startPoint As SnapshotPoint
            Dim endPoint As SnapshotPoint
            Dim foundHyphenAtStart As Boolean
            Dim foundHyphenAtEnd As Boolean


            ' From the word at the cursor, move to the bounds of the hyphenated word.
            startPoint = GetStartOfHyphenatedWord(extent.Span, lineExtent, foundHyphenAtStart)
            endPoint = GetEndOfHyphenatedWord(extent.Span, lineExtent, foundHyphenAtEnd)

            ' Only return the span of the word if a hyphen was found before or after the
            ' extent of the word at the cursor, or if the extent contained a hyphen. If
            ' the "word" at the cursor is a hyphen, then we want to select the words around
            ' it. If there isn't a hyphen anywhere in the span that we would select, then
            ' we will return null and let the default selection handling take care of it.
            If foundHyphenAtStart OrElse foundHyphenAtEnd OrElse (originalText.IndexOf("-"c) >= 0) Then
                Return New SnapshotSpan(startPoint, endPoint)
            End If
        End If

        Return Nothing
    End Function


    Private Function GetStartOfHyphenatedWord(
            span As SnapshotSpan,
            line As SnapshotSpan,
            ByRef foundHyphen As Boolean
        ) As SnapshotPoint

        Dim startOfWord As SnapshotPoint
        Dim point As SnapshotPoint


        startOfWord = span.Start
        point = startOfWord

        Do
            Dim ch As Char


            ' If the point is before the start of the
            ' line, move back to the previous character.
            If point > line.Start Then
                point = point.Subtract(1)
            Else
                Exit Do
            End If

            ch = point.GetChar()

            ' If the character is part of the word then move
            ' the start of the word back to that point.
            If IsCharacterPartOfWord(ch, foundHyphen) Then
                startOfWord = point
            Else
                Exit Do
            End If
        Loop

        Return startOfWord
    End Function


    Private Function GetEndOfHyphenatedWord(
            span As SnapshotSpan,
            line As SnapshotSpan,
            ByRef foundHyphen As Boolean
        ) As SnapshotPoint

        Dim endOfWord As SnapshotPoint
        Dim point As SnapshotPoint


        point = span.End

        Do
            Dim ch As Char


            ' The end of the span is the point after the end of the word (i.e.
            ' the character at that point is not included in the text of the
            ' span). So we always move to this point because if this character
            ' is not part of the word, then it will be the first character after
            ' the word, and that makes it the end point for the span of the word.
            endOfWord = point

            ch = point.GetChar()

            ' If the character is part of the word then move
            ' the start of the word back to that point.
            If IsCharacterPartOfWord(ch, foundHyphen) Then
                endOfWord = point
            Else
                Exit Do
            End If

            ' If the point is before the end of the
            ' line, move to the next character.
            If point < line.End Then
                point = point.Add(1)
            Else
                Exit Do
            End If
        Loop

        Return endOfWord
    End Function


    Private Function IsCharacterPartOfWord(
            ch As Char,
            ByRef foundHyphen As Boolean
        ) As Boolean

        Select Case ch
            Case "-"c
                foundHyphen = True
                Return True

            Case "_"c
                ' Underscores are treated as part of words by Visual Studio.
                Return True

            Case Else
                Return Char.IsLetterOrDigit(ch)

        End Select
    End Function

End Class
