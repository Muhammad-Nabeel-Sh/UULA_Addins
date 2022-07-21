Attribute VB_Name = "CustomTab"
Option Explicit
Sub CustomTab()
    Dim NumTabStop As Double
        NumTabStop = InputBox("Enter the tab stop value: ", "Tab Stops")
        ActiveWindow.Selection.TextRange2.ParagraphFormat.TabStops.Add msoTabStopLeft, 72 * NumTabStop
End Sub
