VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCapitalBudgets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ChampionNotes_AfterUpdate()
Call ManualUpdateBudgetNotes(Me.BudgetReference.Value, Me.TargetStartDate.Value, Me.ChampionNotes.Value, userName())

End Sub

Private Sub cmdClearFilter_Click()
Me.Filter = ""
Me.FilterOn = False

'Clear filter options
Me.cboFilter1 = Null
Me.cboFilter2 = Null
Me.cboFilter3 = Null

'Clear filter values
Me.cboFilterInput1 = Null
Me.cboFilterInput2 = Null
Me.cboFilterInput3 = Null
End Sub

Private Sub cmdFilter_Click()
If IsNull(Me.cboFilter1) = False And IsNull(Me.cboFilterInput1) = False Then
    If IsNull(Me.cboFilter2) = False And IsNull(Me.cboFilterInput2) = False Then
        If IsNull(Me.cboFilter3) = False And IsNull(Me.cboFilterInput3) = False Then
            Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value) _
                & " AND " & FilterFormula(Me.cboFilter2, Me.cboFilterInput2.Value) _
                & " AND " & FilterFormula(Me.cboFilter3, Me.cboFilterInput3.Value)
            Else
                Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value) _
                & " AND " & FilterFormula(Me.cboFilter2, Me.cboFilterInput2.Value)
            End If
    Else
        Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value)
    End If
Else
    Exit Sub
End If

Me.FilterOn = True
Me.cboFilter1.SetFocus

End Sub

Private Sub cmdSaveRecord_Click()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Form_AfterUpdate()
Call ManualUpdateBudgetNotes(Me.BudgetReference.Value, Me.TargetStartDate.Value, Me.ChampionNotes.Value, userName())
End Sub

Private Sub Form_Load()
Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value)
Me.FilterOn = True
End Sub
