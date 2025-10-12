VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sinhalaRfixWindow 
   Caption         =   "UserForm1"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "userForm-for-replace-R-with-repaya-or-restore-default-R-0.3V-stable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sinhalaRfixWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //correcting the wrong "R" character changing/splitting error:

Private Sub ChangeToRwithRepaya_Click()
       Selection.Find.ClearFormatting
       Selection.Find.Replacement.ClearFormatting
       ' //change default "R" character to repaya:
       With Selection.Find
           .Text = ChrW(3515) & ChrW(3530) & "([" & ChrW(3461) & ChrW(3462) & ChrW(3463) & ChrW(3464) & ChrW(3465) & ChrW(3466) & ChrW(3467) & ChrW(3468) & ChrW(3469) & ChrW(3470) & ChrW(3471) & ChrW(3472) & ChrW(3473) & ChrW(3474) & ChrW(3475) & ChrW(3476) & ChrW(3477) & ChrW(3478) & ChrW(3482) & ChrW(3483) & ChrW(3484) & ChrW(3485) & ChrW(3486) & ChrW(3487) & ChrW(3488) & ChrW(3489) & ChrW(3490) & ChrW(3491) & ChrW(3492) & ChrW(3493) & ChrW(3494) & ChrW(3495) & ChrW(3496) & ChrW(3497) & ChrW(3498) & ChrW(3499) & ChrW(3500) & ChrW(3501) & ChrW(3502) & ChrW(3503) & ChrW(3504) & ChrW(3505) & ChrW(3507) & ChrW(3508) & ChrW(3509) & ChrW(3510) & ChrW(3511) & ChrW(3512) & ChrW(3513) & ChrW(3514) & ChrW(3517) & ChrW(3520) & ChrW(3521) & ChrW(3522) & ChrW(3523) & ChrW(3524) & ChrW(3525) & ChrW(3526) & ChrW(3558) & ChrW(3559) & ChrW(3560) & ChrW(3561) & ChrW(3562) & ChrW(3563) & ChrW(3564) & ChrW(3565) & ChrW(3566) & ChrW(3567) & ChrW(3572) _
& "]?)"
           .Replacement.Text = ChrW(3515) & ChrW(3530) & ChrW(8205) & ChrW(3515) & "\1"
           .Forward = True
           .Wrap = wdFindContinue
           .Format = True
           .MatchWildcards = True
           .Execute Replace:=wdReplaceAll
       End With
    exampleOfRwithRepaya.Visible = True
    exampleOfRestoredR.Visible = False
    MsgBox ("done!.. :-)")
End Sub


Private Sub restoreRepayaToDefaultR_Click()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    ' //change repaya to default "R" character:
    With Selection.Find
        .Text = ChrW(3515) & ChrW(3530) & ChrW(8205) & ChrW(3515) & "([" & ChrW(3461) & ChrW(3462) & ChrW(3463) & ChrW(3464) & ChrW(3465) & ChrW(3466) & ChrW(3467) & ChrW(3468) & ChrW(3469) & ChrW(3470) & ChrW(3471) & ChrW(3472) & ChrW(3473) & ChrW(3474) & ChrW(3475) & ChrW(3476) & ChrW(3477) & ChrW(3478) & ChrW(3482) & ChrW(3483) & ChrW(3484) & ChrW(3485) & ChrW(3486) & ChrW(3487) & ChrW(3488) & ChrW(3489) & ChrW(3490) & ChrW(3491) & ChrW(3492) & ChrW(3493) & ChrW(3494) & ChrW(3495) & ChrW(3496) & ChrW(3497) & ChrW(3498) & ChrW(3499) & ChrW(3500) & ChrW(3501) & ChrW(3502) & ChrW(3503) & ChrW(3504) & ChrW(3505) & ChrW(3507) & ChrW(3508) & ChrW(3509) & ChrW(3510) & ChrW(3511) & ChrW(3512) & ChrW(3513) & ChrW(3514) & ChrW(3517) & ChrW(3520) & ChrW(3521) & ChrW(3522) & ChrW(3523) & ChrW(3524) & ChrW(3525) & ChrW(3526) & ChrW(3558) & ChrW(3559) & ChrW(3560) & ChrW(3561) & ChrW(3562) & ChrW(3563) & ChrW(3564) & ChrW(3565) & ChrW(3566) & ChrW(3567) & ChrW(3572) & ChrW(8205) _
                & "]?)"
        .Replacement.Text = ChrW(3515) & ChrW(3530) & "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    exampleOfRwithRepaya.Visible = False
    exampleOfRestoredR.Visible = True
    MsgBox ("done!..")
End Sub
