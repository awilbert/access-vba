Function dhAge(dtmBD As Date, Optional dtmDate As Date = 0) As Integer
   
    Dim intAge As Integer
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    intAge = DateDiff("yyyy", dtmBD, dtmDate)
    If dtmDate < DateSerial(Year(dtmDate), Month(dtmBD), _
     Day(dtmBD)) Then
        intAge = intAge - 1
    End If
    dhAge = intAge
End Function
