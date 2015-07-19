Attribute VB_Name = "modClipboard"
'http://kakaprogramming.blogspot.de/2010/01/copy-text-to-clipboard-in-vb6.html
Sub CopyTextToClipboard(ByVal TextToCopy As String)
    Clipboard.Clear
    Clipboard.SetText TextToCopy
End Sub

Function PasteTextFromClipboard() As String
    PasteTextFromClipboard = Clipboard.GetText
End Function
