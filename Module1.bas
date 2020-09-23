Attribute VB_Name = "Module1"
Public r1
Public r2
Public r3
Public cindex1
Public cindex2
Public cindex3
Public counter1
Public counter2
Public counter3
Public flashimage
Public flashimage1
Public flashimage2
Public playingup As Boolean
Public neste As Boolean
Public won2000 As Boolean
Public lowwin As Boolean
Public bonus As Integer
Public holdbonus As Integer
Public xcount
Public xcount1
Public xcount2
Public scount
Public scount1
Public scount2
Public rounds1
Public rounds2
Public rounds3
Public flashcount
Public wincash As Integer
Public currentcash As Integer
Declare Function sndPlaySound Lib "winmm" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0 ' play synchronously (default)
Public Const SND_ASYNC = &H1 ' play asynchronously
Public Const SND_LOOP = &H8 ' loop the sound until next



Public Function getbonus()
Randomize
bonusnr = Int((Rnd * 40) + 1)
Select Case bonusnr
Case 1
bonus = 20
Case 2
bonus = 20
Case 3
bonus = 20
Case 4
bonus = 20
Case 5
bonus = 20
Case 6
bonus = 20
Case 7
bonus = 20
Case 8
bonus = 20
Case 9
bonus = 20
Case 10
bonus = 40
Case 11
bonus = 40
Case 12
bonus = 40
Case 13
bonus = 40
Case 14
bonus = 40
Case 15
bonus = 60
Case 16
bonus = 60
Case 17
bonus = 60
Case 18
bonus = 80
Case 19
bonus = 80
Case 20
bonus = 100
Case 21
bonus = 100
Case 22
If lowwin = True Then
bonus = 20
Else
bonus = 200
End If
Case 23
If lowwin = True Then
bonus = 20
Else
bonus = 300
End If
Case 24
If lowwin = True Then
bonus = 20
Else
bonus = 400
End If
Case 25
If lowwin = True Then
bonus = 60
Else
bonus = 500
End If
Case 26
If lowwin = True Then
bonus = 60
Else
bonus = 800
End If
Case 27
If lowwin = True Then
bonus = 80
Else
bonus = 1200
End If
Case 28
If lowwin = True Then
bonus = 100
Else
bonus = 1500
End If
Case 29
If lowwin = True Then
bonus = 100
Else
bonus = 1600
End If
Case 30
If lowwin = True Then
bonus = 200
Else
bonus = 2000
End If
Case 31
bonus = 20
Case 32
bonus = 20
Case 33
bonus = 20
Case 34
bonus = 40
Case 35
bonus = 40
Case 36
bonus = 40
Case 37
bonus = 60
Case 38
bonus = 60
Case 39
bonus = 20
Case 40
bonus = 20
End Select
getbonus = bonus
End Function
