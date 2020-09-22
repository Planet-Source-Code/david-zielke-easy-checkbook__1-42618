Attribute VB_Name = "Module1"
Public Old6, FindPKey 'these MUST BE HERE FOR THE WHOLE APP.
Public Function MySubmit()
    Form1.Data1.Refresh
    With Form1.Data1.Recordset
        
        If FindPKey <> "" Then 'for empty db
            .FindLast "PrimaryKey = " & FindPKey
            .MovePrevious
        Else
            If Form1.MaskEdBox4.Text = "" Then 'Credit
                Form1.MaskEdBox6.Text = Form1.MaskEdBox5.Text
            Else
                Form1.MaskEdBox6.Text = "-" & Form1.MaskEdBox4.Text
            End If
        End If
        
        Old6 = Form1.MaskEdBox6.Text
        Do Until .EOF = True
            .MoveNext
            If Form1.MaskEdBox4.Text = "" Then 'Credit
                Form1.MaskEdBox6.Text = Old6 + Val(Form1.MaskEdBox5.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            Else 'debit
                Form1.MaskEdBox6.Text = Old6 - Val(Form1.MaskEdBox4.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            End If
        Loop
    End With
End Function
Public Function MyDelete()
    Form1.Data1.Refresh
    With Form1.Data1.Recordset
        .FindLast "PrimaryKey = " & FindPKey
        .Delete
        .MovePrevious
        Old6 = Form1.MaskEdBox6.Text
        Do Until .EOF = True
            .MoveNext
            If Form1.MaskEdBox4.Text = "" Then 'Credit
                Form1.MaskEdBox6.Text = Old6 + Val(Form1.MaskEdBox5.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            Else                        'debit
                Form1.MaskEdBox6.Text = Old6 - Val(Form1.MaskEdBox4.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            End If
        Loop
    End With
End Function

Public Function MyUpdate()
    Form1.Data1.Refresh
    With Form1.Data1.Recordset
        .FindLast "PrimaryKey = " & FindPKey
        .MovePrevious
        Old6 = Form1.MaskEdBox6.Text
        Do Until .EOF = True
            .MoveNext
            If Form1.MaskEdBox4.Text = "" Then 'Credit
                Form1.MaskEdBox6.Text = Old6 + Val(Form1.MaskEdBox5.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            Else 'debit
                Form1.MaskEdBox6.Text = Old6 - Val(Form1.MaskEdBox4.Text)
                Old6 = Val(Form1.MaskEdBox6.Text)
            End If
        Loop
    End With
End Function

