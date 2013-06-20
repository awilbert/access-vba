' Hide the Microsoft Access "Ribbon" toolbar.
' Also removes access to "Print Preview" ribbon, so make sure all printing needs are accommodated within the database itself.

Private Sub btnHideRibbon_Click()
DoCmd.ShowToolbar "Ribbon", acToolbarNo
End Sub


' Reverse the process to re-enable the Ribbon.

Private Sub btnShowRibbon_Click()
DoCmd.ShowToolbar "Ribbon", acToolbarYes
End Sub
