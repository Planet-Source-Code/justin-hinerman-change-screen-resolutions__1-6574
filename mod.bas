Attribute VB_Name = "Module1"

Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, I As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    I = 0


    Do
        ReturnVal = EnumDisplaySettings(0&, I, DevM)
        I = I + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
    RetValue = MsgBox("Do You Wish To Keep Your Screen Resolution To " & Width & "x" & Height & " - " & BPP & " BPP?", vbQuestion + vbOKCancel, "Change Resolution Confirm:")


    If RetValue = vbCancel Then
        DevM.dmPelsWidth = OldWidth
        DevM.dmPelsHeight = OldHeight
        DevM.dmBitsPerPel = OldBPP
        Call ChangeDisplaySettings(DevM, 1)
        MsgBox "Old Resolution(" & OldWidth & " x " & OldHeight & ", " & OldBPP & " Bit) Successfully Restored!", vbInformation + vbOKOnly, "Resolution Confirm:"
        ChangeRes = 0
    Else
        ChangeRes = 1
    End If
    Exit Function
ERROR_HANDLER:
    ChangeRes = 0
End Function

