Sub test()
    
    json = "{foo:123, bar:""full"", tbl:[100,200,300]}"   ' string with JSON, from API

    Dim s As New ScriptControl: s.Language = "JScript"
    
    ' using more JavaScript
    s.ExecuteStatement "x = " & json
    Debug.Print s.Eval("x.foo")          ' 123
    Debug.Print s.Eval("typeof x.foo")   ' number
    Debug.Print s.Eval("x.tbl[1]+5")     ' 205
    
    s.AddCode "function fun(n) {return 'text ' + x.bar + ' (n)=' + x.tbl[n];} "
    Debug.Print s.Run("fun", 1)          ' text full (n)=200
    
    ' using more VBA
    Set tbl2 = s.Eval("x.tbl")
    Debug.Print CallByName(tbl2, "length", VbGet) ' 3
    Debug.Print CallByName(tbl2, "0", VbGet)      ' 100
    
    ' using even more VBA, no JavaScript variable
    Set j = s.Eval("(" & json & ")")
    Debug.Print j.foo; j.bar                       ' 123 full
    Debug.Print CallByName(j.tbl, "length", VbGet) ' 3
    Debug.Print CallByName(j.tbl, "0", VbGet)      ' 100
    Debug.Print CallByName(j.tbl, "1", VbGet)      ' 200
    
    Set s = Nothing
   
End Sub
