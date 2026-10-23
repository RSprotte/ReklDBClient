Attribute VB_Name = "zumTesten"
Option Compare Database
Option Explicit

Private Declare PtrSafe Function tapiRequestMakeCall Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
' Possible error return values from tapiRequestMakeCall function
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&
Private Const TAPIERR_INVALPOINTER = -18&

Public Sub Dial(strPhoneNo As String, Optional strCalledParty As String, _
Optional strComment As String, Optional bError As Boolean)
Dim strErr      As String
Dim lngResult   As Long
Dim strCaption     As String

    ' Assign Default values to optional arguments
    If IsMissing(strCalledParty) Then strCalledParty = ""
    If IsMissing(strComment) Then strComment = ""
    If IsMissing(bError) Then bError = True

    lngResult = tapiRequestMakeCall(Trim(strPhoneNo), strCaption, strCalledParty, strComment)
    'Display error message unless requested otherwise
    If bError And lngResult <> 0 Then
        strErr = "Error dialing " + strPhoneNo + ": "
        Select Case lngResult
            Case TAPIERR_NOREQUESTRECIPIENT
                strErr = strErr & "Unable to start/use Dialing application."
            Case TAPIERR_REQUESTQUEUEFULL
                strErr = strErr & "Dialing requests queue is full."
            Case TAPIERR_INVALDESTADDRESS
                strErr = strErr & "Phone number is invalid."
            Case TAPIERR_INVALPOINTER
                strErr = strErr & "Pointer error."
                ' This error should not occur
            Case Else
                strErr = strErr & "Unknown error."
                ' The four errors listed are all of the current function
                ' error returns but who knows what the future will bring
        End Select
        MsgBox strErr, vbOKOnly + vbExclamation, "Dialing Error"
    End If
End Sub
