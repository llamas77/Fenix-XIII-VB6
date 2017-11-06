Attribute VB_Name = "modDirectSound"
Option Explicit

Public Sub IniciarDirectSound()

Err.Clear

On Error GoTo fin
    
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    
    Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
        
Exit Sub
fin:
    End
End Sub

Public Sub LiberarDirectSound()

Dim cloop As Integer

For cloop = 1 To NumSoundBuffers
    Set DSBuffers(cloop) = Nothing
Next cloop

Set DirectSound = Nothing

End Sub
