[Forms]![Test Form]![FiscalYear]


[Forms]![Test Form]![Department]



create proc [dbo].[DepartmentBudgets_param_DeptandFY]  @FY Char(4), @dept char(20)
as
begin
--dept original budget
select  d.DeptID, DeptName, BudgetAmount 'Original Budget'
from Department d
where d.fiscalyear = @FY and DeptName = @dept
--where FiscalYear = '2021'
group by d.DeptID, DeptName, BudgetAmount
end




CREATE proc [dbo].[DepartmentBudgetsAll_param_FY] @FY Char(4)
as
begin
--dept original budget
select  d.DeptID, DeptName, BudgetAmount 'Original Budget'
from Department d
where d.fiscalyear = @FY
--where FiscalYear = '2021'
group by d.DeptID, DeptName, BudgetAmount
end


CREATE proc [dbo].[DepartmentBudgetsAll_param_FY] @FY Char(4)
as
begin
select  d.DeptID, DeptName, BudgetAmount 'Original Budget'
from Department d
where d.fiscalyear = @FY
group by d.DeptID, DeptName, BudgetAmount
end

[Forms]![Fiscal Year]![Fiscal_Year]


Like "*" & [Forms]![frm_DepartmentBudgetsAll(FY)]![txt_FiscalYearBox] & "*"

CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, ExpenseDate, ExpenseCat, ExpenseType, Team, MemberProject, Vendor, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.txt_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_ExpenseDate & "','" & Me.txt_ExpenseCat & "','" & Me.txt_ExpenseType & "','" & Me.txt_Team & "','" & Me.txt_MemberProject & "','" & Me.txt_Vendor & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"

frm_ExpensesActuals.Form.Requery

CurrentDb.Execute "INSERT INTO Table2(FirstName, LastName)" & "VALUES('" & Me.frst_Name_txt & "','" & Me.lst_Name_txt & "')"


 Me.txt_InitiatorCardHolder = ""
    Me.txt_AmountCharged = ""
    Me.txt_ExpenseDate = ""
    Me.txt_ExpenseType = ""
    Me.txt_Team = ""
    Me.txt_MemberProject = ""
    Me.txt_Vendor = ""
    Me.txt_Chartfield = ""
    Me.txt_Purpose = ""
    Me.txt_InvoiceNo = ""
    Me.txt_DocImage = ""


CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder) " & " VALUES(" & Me.txt_InitiatorCardHolder & "')"


CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, ExpenseDate, ExpenseCat, ExpenseType, Team, MemberProject, Vendor, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.txt_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_ExpenseDate & "','" & Me.txt_ExpenseCat & "','" & Me.txt_ExpenseType & "','" & Me.txt_Team & "','" & Me.txt_MemberProject & "','" & Me.txt_Vendor & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"

frm_ExpensesActualsSUB.Form.Requery



----------------------EDIT

CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, RequestDate, ExpenseCat, ExpenseType, Team, MemberProject, VendorReimbursee, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.cbo_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_RequestDate & "','" & Me.cbo_ExpenseCat & "','" & Me.cbo_ExpenseType & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"

CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, RequestDate, RequestFY, ExpenseCat, ExpenseType, Team, MemberProject, VendorReimbursee, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.cbo_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_RequestDate & "','" & Me.txt_RequestFY & "','" & Me.cbo_ExpenseCat & "','" & Me.cbo_ExpenseType & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"


frm_ExpensesActualsSUB.Form.Requery


------------------TO DO
Add New Record Pop Up of Update, Remove Subform 
Update Expense Actual Table Button
Button Red 
Figure out a way to add insert data into multiple tables
inserting records into multiple tables access 

    Bug 1: increase char length for purpose
    Bug 2: try/catch with pop up if purpose goes past 50 chars
    or
    if you can lock insert to 50 chars
    Add 1: remaining columns from insert add to right side of update form
    Add 2: date range for queries Team and Member
    worklist 1: get update function working




ExpenseCat = Classifications
Type = Authorized Method for Payments  
Alpha 0.1





---------UPDATE EDIT QUERY------------------------
CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, RequestDate, RequestFY, ExpenseCat, ExpenseType, Team, MemberProject, VendorReimbursee, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.cbo_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_RequestDate & "','" & Me.txt_RequestFY & "','" & Me.cbo_ExpenseCat & "','" & Me.cbo_ExpenseType & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"


CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VerificationDate] = " & CDbl(Me.txt_VerificationUpdate) & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [PaidFY] = " & (Me.txt_PaidFY1) & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VerificationDate] = " & CDbl(Me.txt_VerificationUpdate) & "" & ", [PaidFY] = " & (Me.txt_PaidFY1) & "" & ", [AmountCharged] = " & (Me.txt_AmountChargedUpdate) & "" & ", [Purpose] = " & (Me.txt_PurposeUpdate) & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

Current Version
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VerificationDate] = " & CDbl(Me.txt_VerificationUpdate) & "" & ", [PaidFY] = " & (Me.txt_PaidFY1) & "" & ", [AmountCharged] = " & Me.txt_AmountChargedUpdate & "" & ", [Purpose] = " & "'" & Me.txt_PurposeUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""



CurrentDb.Execute
Debug.Print

CurrentDb.Execute "UPDATE product " & " SET [product name] = '" & Me.txtName & "'" & ", [cost of product] = " & Me.txtCost & "" & ", [weight] = " & Me.txtWeight & "" & ", [group] = '" & Me.CmbGroup & "'" & ", [group ID] = '" & Me.txtGroupID & "'" & ", [Ordered] = " & Me.txtOrdered & "" & " WHERE [product name] = '" & Me.txtName.Tag & "'"






CurrentDB.Execute "UPDATE tbl_ExpenseActual " & " SET [PaidFY] = '" & Me.txt_PaidFY1 & "" & " WHERE [ExpenseTX] =" & Me.txt_ExpenseTX & "'"

CurrentDb.Execute "UPDATE product " & " SET [product name] = '" & Me.txtName & "" & " WHERE [product name] = '" & Me.txtName.Tag & "'"

CurrentDb.Execute " UPDATE damaged_card " & " SET Quantity='" & Me.txtquantity & "'" & " WHERE Quantity='" & Me.txtquantity.Tag & "'"



TO DO (5/15/20)
Queries: 
May 12 TeamAndMember
May 12 Team With OutMember
look at the attachment
Password 


Date range, or the 

'----------------------------------
'Expense Actuals Update Options'
'----------------------------------

'Update Version 3 (This one is the seperate IF Statements)
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VerificationDate] = " & CDbl(Me.txt_VerificationUpdate) & "" & ", [PaidFY] = " & (Me.txt_PaidFY1) & "" & ", [AmountCharged] = " & Me.txt_AmountChargedUpdate & "" & ", [Purpose] = " & "'" & Me.txt_PurposeUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

'Verification Column
If txt_VerificationUpdate.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_VerificationUpdate.Value) = False Then
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VerificationDate] = " & CDbl(Me.txt_VerificationUpdate) & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""
End If
End If

End Sub

'Paid FY Column
If txt_PaidFY1.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_VerificationUpdate.Value) = False Then
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [PaidFY] = " & Me.txt_PaidFY1 & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""
CurrentDb.Execute "UPDATE tbl_ExpenseFundSource " & " SET [PaidFY] = " & Me.txt_PaidFY1 & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

End If
End If

End Sub

'Amount Charged Column
If txt_PaidFY1.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_VerificationUpdate.Value) = False Then
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [AmountCharged] = " & & Me.txt_AmountChargedUpdate & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""
End If
End If

End Sub

'Purpose Column
If txt_PurposeUpdate.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_PurposeUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [Purpose] = " & Me.txt_PurposeUpdate & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""
        CurrentDb.Execute "UPDATE tbl_ExpenseFundSource " & " SET [Purpose] = " & Me.txt_PurposeUpdate & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""
End If
End If

End Sub


'----------------------------------
'Extra Update Options'
'----------------------------------

'Initiator Cardholder Column

If cbo_InitiatorUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_InitiatorUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [InitiatorCardHolder] = " & "'" & Me.cbo_InitiatorUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If


'Request Date
If txt_RequestUpdate.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_RequestUpdate.Value) = False Then
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [RequestDate] = " & CDbl(Me.txt_RequestUpdate) & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If

End Sub


'Paid Date
If txt_RequestUpdate.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_RequestUpdate.Value) = False Then
CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [PaidDate] = " & CDbl(Me.txt_RequestUpdate) & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If

End Sub


'Expense Category
If cbo_ExpenseCatUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_ExpenseCatUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [ExpenseCat] = " & "'" & Me.cbo_ExpenseCatUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If


'Expense Type
If cbo_ExpenseTypeUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_ExpenseTypeUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [ExpenseType] = " & "'" & Me.cbo_ExpenseTypeUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If


'Team
If cbo_TeamUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_TeamUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [ExpenseType] = " & "'" & Me.cbo_TeamUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
        CurrentDb.Execute "UPDATE tbl_ExpenseFundSource " & " SET [ExpenseType] = " & "'" & Me.cbo_TeamUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""

End If
End If


'Member Project
If cbo_MemberProjectUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_MemberProjectUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [MemberProject] = " & "'" & Me.cbo_MemberProjectUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
        CurrentDb.Execute "UPDATE tbl_ExpenseFundSource " & " SET [MemberProject] = " & "'" & Me.cbo_MemberProjectUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""

End If
End If


'Vendor Reimbursee
If cbo_VendorUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_VendorUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VendorReimbursee] = " & "'" & Me.cbo_VendorUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
        CurrentDb.Execute "UPDATE tbl_ExpenseFundSource " & " SET [VendorReimbursee] = " & "'" & Me.cbo_VendorUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""

End If
End If


'Chartfield
If txt_ChartfieldUpdate.Value = "" Then
    Exit Sub
Else
    If IsNull(txt_ChartfieldUpdate.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [Chartfield] = " & "'" & Me.txt_ChartfieldUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.cbo_ExpenseTX & ""
End If
End If

'----------------------------------
'Fund ID Insert Statement'
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(ExpenseTX, FundID, Amount)" & "VALUES('" & Me.cbo_ExpenseTX & "','" & Me.txt_FundID1 & "','" & Me.txt_Amount1 & "')"

Call cmd_AddFundID2
Call cmd_AddFundID3

End Sub



If txt_FundID2.Value = "" Then
    Exit Sub
Else
    If IsNull(txt_FundID2.Value) = False Then
        CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(ExpenseTX, FundID, Amount)" & "VALUES('" & Me.cbo_ExpenseTX & "','" & Me.txt_FundID2 & "','" & Me.txt_Amount2 & "')"
End If
End If

If txt_FundID3.Value = "" Then
    Exit Sub
Else
    If IsNull(txt_FundID3.Value) = False Then
        CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(ExpenseTX, FundID, Amount)" & "VALUES('" & Me.cbo_ExpenseTX & "','" & Me.txt_FundID3 & "','" & Me.txt_Amount3 & "')"
End If
End If

Current Team Balance: Sum([AmountCharged])-[tbl_TeamBudget]![CurrentFundedAmount]


'----------------------------------
'Initial Budget Department'
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_Department(DeptID, DeptName, BudgetAmount, FiscalYear)" & "VALUES('" & Me.cbo_DeptID & "','" & Me.cbo_DeptName & "','" & Me.txt_BudgetAmount & "','" & cbo_FiscalYear & "')"



CurrentDb.Execute "INSERT INTO tbl_Department(DeptID, DeptName, BudgetAmount, FiscalYear)" & "VALUES("& "'" & Me.cbo_DeptID & "'" & "','" & Me.cbo_DeptName & "','" & Me.txt_BudgetAmount & "','" & cbo_FiscalYear & "')"

SET [Chartfield] = " & "'" & Me.txt_ChartfieldUpdate & "'" & "" &


CurrentDb.Execute "INSERT INTO tbl_Department(DeptName, BudgetAmount, FiscalYear)" & "VALUES('" & Me.cbo_DeptName & "','" & Me.txt_BudgetAmount & "','" & cbo_FiscalYear & "')"


Debug.Print "INSERT INTO tbl_Department(DeptID, DeptName, BudgetAmount, FiscalYear)" & "VALUES('" & Me.cbo_DeptID & "','" & Me.cbo_DeptName & "','" & Me.txt_BudgetAmount & "','" & cbo_FiscalYear & "')"

INSERT INTO tbl_Department(DeptID, DeptName, BudgetAmount, FiscalYear)VALUES('0-0101-000','President's Office','333.33','2021')

CurrentDb.Execute "INSERT INTO tbl_ExpenseActual(InitiatorCardHolder, AmountCharged, RequestDate, RequestFY, ExpenseCat, ExpenseType, Team, MemberProject, VendorReimbursee, Chartfield, Purpose, InvoiceNo)" & "VALUES('" & Me.cbo_InitiatorCardHolder & "','" & Me.txt_AmountCharged & "','" & Me.txt_RequestDate & "','" & Me.txt_RequestFY & "','" & Me.cbo_ExpenseCat & "','" & Me.cbo_ExpenseType & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"



'----------------------------------
'Update Budget Department'
'----------------------------------

'CurrentDb.Execute "UPDATE tbl_Department " & " SET [BudgetAmount] = " & Me.txt_BudgetAmount & "" & " WHERE [DeptName] = " & "'" & Me.cbo_DeptName & "'" & " AND [FiscalYear] = " & Me.cbo_FiscalYear & ""



CurrentDb.Execute "INSERT INTO tbl_Department(DeptID, DeptName, BudgetAmount, FiscalYear)" & "VALUES('" & Me.cbo_DeptID & "','" & Me.cbo_DeptName & "','" & Me.txt_BudgetAmount & "','" & cbo_FiscalYear & "')"

CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [VendorReimbursee] = " & "'" & Me.cbo_VendorUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""

CurrentDb.Execute "UPDATE tbl_ExpenseActual " & " SET [ExpenseType] = " & "'" & Me.cbo_ExpenseTypeUpdate & "'" & "" & " WHERE [ExpenseTX] = " & Me.txt_ExpenseTX & ""



CurrentDb.Execute "UPDATE tbl_Department " & " SET [BudgetAmount] = " & Me.txt_BudgetAmount & "" & " WHERE [DeptName] = " & "'" & Me.cbo_DeptName & "'" & ""


CurrentDb.Execute "UPDATE loggingX " & " SET stdid=" & Me.txtID & ", stdname='" & Me.txtName & "'" & ", gender='" & Me.cboGender & "'" & ", phone='" & Me.txtPhone & "'" & ", address='" & Me.txtAddress & "'" & ", WHid='" & Me.txtWHid & "'" & " WHERE stdid=" & Me.txtID.Tag & " AND WHid=" & Me.txtWHid


CurrentDb.Execute "UPDATE tbl_Department " & " SET [BudgetAmount] = " & Me.txt_BudgetAmount & "" & " WHERE [DeptName] = " & "'" & Me.cbo_DeptName & "'" & " AND [FiscalYear] = " & Me.cbo_FiscalYear 




'----------------------------------
'Initial Budget Funding Source'
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_FundingSource(FundingSourceID, FiscalYear, FundingSourceName, Type, FundingAccountNumber, FundedAmount)" & "VALUES('" & Me.cbo_FundingSource & "','" & Me.cbo_FiscalYear & "','" & Me.cbo_FundingSourceName & "','" & Me.cbo_Type & "','" & Me.cbo_FundingAccountNumber & "','" & txt_FundedAmount & "')"

'----------------------------------
'Update Budget Funding Source'
'----------------------------------

CurrentDb.Execute "UPDATE tbl_FundingSource " & " SET [FundedAmount] = " & Me.txt_FundedAmount & "" & " WHERE [FundingSourceID] = " & Me.cbo_FundingSource & " AND [FiscalYear] = " & Me.cbo_FiscalYear & ""




'----------------------------------
'Initial Budget Team'
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_TeamBudget(TeamName, FiscalYear, OriginalFundedAmount, CurrentFundedAmount, LastUpdated)" & "VALUES('" & Me.cbo_TeamName & "','" & Me.cbo_FiscalYear & "','" & Me.cbo_OriginalFundedAmount & "','" & Me.cbo_CurrentFundedAmount & "','" & Me.txt_LastUpdated & "')"


'----------------------------------
'Update Budget Team'
'----------------------------------

CurrentDb.Execute "UPDATE tbl_TeamBudget " & " SET [CurrentFundedAmount] = " & Me.cbo_CurrentFundedAmount & "" & ", [LastUpdated] = " & CDbl(Me.txt_LastUpdated) & "" & " WHERE [TeamName] = " & "'" & Me.cbo_TeamName & "'" & " AND [FiscalYear] = " & Me.cbo_FiscalYear & ""



'----------------------------------
'Initial Budget Member/Project'
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_MemberProjectBudget(MemberProjectName, FiscalYear, OriginalFundedAmount, CurrentFundedAmount, LastUpdated)" & "VALUES('" & Me.cbo_TeamName & "','" & Me.cbo_FiscalYear & "','" & Me.cbo_OriginalFundedAmount & "','" & Me.cbo_CurrentFundedAmount & "','" & Me.txt_LastUpdated & "')"


'----------------------------------
'Update Budget Member/Project'
'----------------------------------

CurrentDb.Execute "UPDATE tbl_MemberProjectBudget " & " SET [CurrentFundedAmount] = " & Me.cbo_CurrentFundedAmount & "" & ", [LastUpdated] = " & CDbl(Me.txt_LastUpdated) & "" & " WHERE [MemberProjectName] = " & "'" & Me.cbo_MemberName & "'" & " AND [FiscalYear] = " & Me.cbo_FiscalYear & ""



new punch list for weekend:

'Add clear fields button for fundingID insert

'1) Add pop up the fundingID successfully entered

'2) Add a close button for each form.

'3) If possible, add a "hot button" for expense actual insert that returns last row inserted -
'select top 1 ExpenseTX from ExpenseActual (edited) 
'order by ExpenseTX Desc

'4) Fix Expense Actuals Update Statements Update all Button


CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(Team, MemberProject, VendorReimbursee, Purpose)" & "VALUES('" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Purpose & "','" & Me.cbo_ExpenseCat & "','" & Me.cbo_ExpenseType & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Chartfield & "','" & Me.txt_Purpose & "','" & txt_InvoiceNo & "')"


CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(ExpenseTX, Team, MemberProject, VendorReimbursee, Purpose)" & "VALUES('" & Me.txt_LatestExpenseTX2 & "','" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Purpose & "')"

CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(Team, MemberProject, VendorReimbursee, Purpose)" & "VALUES('" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Purpose & "','" & Me.txt_Chartfield & "')"


CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(Team, MemberProject, VendorReimbursee, Purpose)" & "VALUES('" & Me.cbo_Team & "','" & Me.cbo_MemberProject & "','" & Me.cbo_VendorReimbursee & "','" & Me.txt_Purpose & "')"



'----------------------------------
Punch List 6/23
'----------------------------------


1) 'also, need a start page. meaning have it open to the home form.
2) need department spend and roll up, just like team and member.
    -What is the (amount) column for, and how should it affect the QUERY
    -Does not have the same Reforecast/Forecast columns, come back and ask
3) also, we forgot to create a funding source rollup.
4) We forgot to add FY to the funding source for expense actuals.

CurrentDb.Execute "INSERT INTO tbl_ExpenseFundSource(ExpenseTX, PaidFY, FundID, Amount)" & "VALUES('" & Me.cbo_ExpenseTX & "','" & Me.cbo_PaidFY & "','" & Me.cbo_FundID1 & "','" & Me.txt_Amount1 & "')"


















