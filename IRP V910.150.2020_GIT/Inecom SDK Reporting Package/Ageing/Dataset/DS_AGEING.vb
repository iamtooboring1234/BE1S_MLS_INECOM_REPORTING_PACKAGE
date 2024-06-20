

Partial Public Class DS_AGEING
    Partial Public Class OCRDDataTable


    End Class

    Partial Public Class __NCM_AR_AGEINGDataTable
        Private Sub __NCM_AR_AGEINGDataTable_ColumnChanging(sender As Object, e As DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.U_IRPBPFIELD1Column.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class
End Class
