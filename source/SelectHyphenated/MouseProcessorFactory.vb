Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Operations
Imports Microsoft.VisualStudio.Utilities
Imports System.ComponentModel.Composition


<ContentType("text")>
<Export(GetType(IMouseProcessorProvider))>
<Name("select-hyphenated")>
<TextViewRole(PredefinedTextViewRoles.Document)>
Public Class MouseProcessorFactory
    Implements IMouseProcessorProvider


    <Import()>
    Private cgTextStructureNavigatorService As ITextStructureNavigatorSelectorService


    Public Function GetAssociatedProcessor(wpfTextView As IWpfTextView) As IMouseProcessor _
        Implements IMouseProcessorProvider.GetAssociatedProcessor

        Return New MouseProcessor(
            wpfTextView,
            cgTextStructureNavigatorService.GetTextStructureNavigator(wpfTextView.TextBuffer)
        )
    End Function

End Class
