Module modError
    Public Sub ShowError(ByVal strText As String)
        frmError.lblError.Text = strText
        frmError.ShowDialog()
    End Sub

    Public Function BeEmpty(ByVal MyObject As Control) As Boolean
        BeEmpty = False
        If MyObject.Text = "" Then
            BeEmpty = True
            ShowError("یک عبارت وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
        End If
    End Function

    Public Function BeEmptyNotNumeric(ByVal MyObject As Control) As Boolean
        BeEmptyNotNumeric = False
        If MyObject.Text = "" Then
            BeEmptyNotNumeric = True
            ShowError("یک عبارت وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
        If Not IsNumeric(MyObject.Text) Then
            BeEmptyNotNumeric = True
            ShowError("یک عبارت عددي را وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
        If InStr(MyObject.Text, "/") <> 0 Then
            BeEmptyNotNumeric = True
            ShowError("یک عبارت عددي را وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
    End Function

    Public Function BeEmptyNotDate(ByVal MyObject As Control) As Boolean
        BeEmptyNotDate = False
        If MyObject.Text = "" Then
            BeEmptyNotDate = True
            ShowError("یک عبارت وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
        If Not BeDate(MyObject.Text) Then
            BeEmptyNotDate = True
            Call ShowError("يک تاريخ صحيح وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
    End Function

    Public Function BeNotNumeric(ByVal MyObject As Control) As Boolean
        BeNotNumeric = False
        If MyObject.Text = "" Then Exit Function
        If Not IsNumeric(MyObject.Text) Then
            BeNotNumeric = True
            ShowError("یک عبارت عددي را وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
    End Function

    Public Function BeNotDate(ByVal MyObject As Control) As Boolean
        BeNotDate = False
        If MyObject.Text = "" Then Exit Function
        If Not BeDate(MyObject.Text) Then
            BeNotDate = True
            Call ShowError("يک تاريخ صحيح وارد نماييد.")
            If TypeOf (MyObject.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject.Parent.Parent
                tbcControl.SelectedTab = MyObject.Parent
            End If
            MyObject.Select()
            Exit Function
        End If
    End Function

    Public Function BeNotBiggerDate(ByVal MyObject1 As Control, ByVal MyObject2 As Control) As Boolean
        BeNotBiggerDate = False
        If MyObject1.Text = "" Or MyObject2.Text = "" Then Exit Function
        If MyObject1.Text < MyObject2.Text Then
            BeNotBiggerDate = True
            Call ShowError("تاريخ دوم بايد بزرگتر از تاريخ اول باشد.")
            If TypeOf (MyObject1.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject1.Parent.Parent
                tbcControl.SelectedTab = MyObject1.Parent
            End If
            MyObject1.Select()
            Exit Function
        End If
    End Function

    Public Function BeNotBiggerNumeric(ByVal MyObject1 As Control, ByVal MyObject2 As Control) As Boolean
        BeNotBiggerNumeric = False
        If MyObject1.Text = "" Or MyObject2.Text = "" Then Exit Function
        If CDbl(MyObject1.Text) < CDbl(MyObject2.Text) Then
            BeNotBiggerNumeric = True
            Call ShowError("عدد دوم بايد بزرگتر از عدد اول باشد.")
            If TypeOf (MyObject1.Parent.Parent) Is TabControl Then
                Dim tbcControl As TabControl
                tbcControl = MyObject1.Parent.Parent
                tbcControl.SelectedTab = MyObject1.Parent
            End If
            MyObject1.Select()
            Exit Function
        End If
    End Function

End Module
