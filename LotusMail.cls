VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LotusMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class LotusMail
'Attributes

Private p_Recipient() As String
Private p_RecipientCC() As String
Private p_RecipientBCC() As String
'Private p_Subject As String
'Private p_Body As String
'Private p_Signature As Boolean
'Private p_Principal As Boolean
'Private p_PrincipalName As String
'Private p_PrincipalAddress As String
'Private p_Attachment As Boolean
Private p_AttachmentDetails() As String
'Private p_SaveToSent As Boolean
'Private p_LN_Server As String
'Private p_LN_DBFile As String

Public Property Let Subject(iStr As String)
    p_Subject = iStr
End Property
Public Property Let Body(iStr As String)
    p_Body = iStr
End Property
Public Property Let PrincipalName(iStr As String)
    p_PrincipalName = iStr
End Property
Public Property Let PrincipalAddress(iStr As String)
    p_PrincipalAddress = iStr
End Property
Public Property Let LN_Server(iStr As String)
    p_LN_Server = iStr
End Property
Public Property Let LN_DBFile(iStr As String)
    p_LN_DBFile = iStr
End Property
Public Property Let Signature(iBool As Boolean)
    p_Signature = iBool
End Property
Public Property Let Principal(iBool As Boolean)
    p_Principal = iBool
End Property
Public Property Let Attachment(iBool As Boolean)
    p_Attachment = iBool
End Property
Public Property Let SaveToSent(iBool As Boolean)
    p_SaveToSent = iBool
End Property

Private Sub LoadArray(iDes() As String, iSrc() As String)
    Dim iCnt As Integer: Dim iMax As Integer: iMax = UBound(iSrc) - LBound(iSrc) + 1
    ReDim iDes(iMax)
    
    For iCnt = 0 To iMax
        iDes(iCnt) = iSrc(LBound(iSrc) + iCnt)
    Next iCnt

End Sub

Public Property Let Recipient(iStr() As String)
    Call LoadArray(p_Recipient(), iStr())
    MsgBox "OK", vbinbfo
End Property

