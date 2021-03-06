Public Sub CalendarGenerator(Target As Range)
    Dim val As Integer
    
    If IsNumeric(Target.Value) Then
        val = Target.Value
        
        If val > 2000 And val < 9999 Then
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            
            Dim colorBg As Long
            Dim colorDay As Long
            Dim colorWEnd As Long
            colorBg = RGB(22, 44, 51)
            colorDay = RGB(104, 156, 168)
            colorWEnd = RGB(21, 80, 100)
            Dim start As Range
            Dim cur As Range
            Dim nap As Date
            Dim lastM As Integer
            Dim curDay As VbDayOfWeek
            
            Set start = Target.Offset(2, 0)
            nap = DateSerial(val, 1, 1)
            Set cur = start
            
            Do While cur.Row < (start.Row + 12)
                curDay = vbMonday
            
                Do While Not Weekday(nap) = curDay
                    cur.Value = ""
                    cur.Interior.Color = colorBg
                    Set cur = cur.Offset(0, 1)
                    
                    If curDay < vbSaturday Then
                        curDay = curDay + 1
                    Else
                        curDay = vbSunday
                    End If
                Loop
                
                lastM = Month(nap)
                
                Do While cur.Column < (start.Column + 42)
                    If lastM = Month(nap) Then
                        cur.Value = Day(nap)
                        curDay = Weekday(nap)
                        
                        If curDay = vbSaturday Then
                            cur.Interior.Color = colorWEnd
                        ElseIf curDay = vbSunday Then
                            cur.Interior.Color = colorWEnd
                            Set cur = cur.Offset(0, 1)
                            cur.Interior.Color = colorBg
                        Else
                            cur.Interior.Color = colorDay
                        End If
                        
                        nap = DateAdd("d", 1, nap)
                    Else
                        cur.Value = ""
                        cur.Interior.Color = colorBg
                    End If
                    
                    Set cur = cur.Offset(0, 1)
                Loop
                
                Set cur = start.Offset(lastM, 0)
            Loop
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        End If
    End If
End Sub
