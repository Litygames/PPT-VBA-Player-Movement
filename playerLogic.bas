Attribute VB_Name = "playerLogic"
 ' ---------------------------------------------
 ' Copyright (c) 2025 Litygames
 ' Licensed under the GNU General Public License v3.0
 ' https://www.gnu.org/licenses/gpl-3.0.txt
 ' ---------------------------------------------
 ' Gracias por utilizar PPTGameMaker - @litygames
 ' No olvides apoyar mi contenido ^^
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If
Dim initialSlide%, imagePath As String, idleShape As shape, movingShape As shape, startTime As Single, lastKey As String
Dim showDialogue As Boolean, currentTrigger As shape, movementBlocked As Boolean, dlgBox As shape, dialogSpeed As Double
Public Sub HoverOver()
 initialSlide = 2
 ActivePresentation.SlideShowWindow.View.PointerType = 3
 ActivePresentation.SlideShowWindow.View.GotoSlide initialSlide
 PlayerMovement
 End Sub
Private Sub PlayerMovement()
 Dim playerSpeed As Single, slideWidth As Integer, slideHeight As Integer, moveX As Single, moveY As Single
 Dim wallShape As shape, wallShapes As New Collection, doorShape As shape, doorShapes As New Collection, triggerShape As shape, triggerShapes As New Collection
 Dim shp As shape, allSlides As slide, keyLeft As Integer, keyUp As Integer, keyDown As Integer, keyRight As Integer, keyZ As Integer, resetInterval As Single
 Dim currentPosition As Integer
 currentPosition = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
 imagePath = ActivePresentation.Path & "\data\"
 playerSpeed = 3
 keyLeft = 65
 keyUp = 87
 keyDown = 83
 keyRight = 68
 keyZ = 69
 resetInterval = 0.5
 dialogSpeed = 0.02
 startTime = Timer
 With ActivePresentation.SlideShowWindow.View
  slideWidth = .slide.Master.Width: slideHeight = .slide.Master.Height
  Set idleShape = .slide.Shapes("playerIdle"): Set movingShape = .slide.Shapes("playerMoving")
    For Each shp In .slide.Shapes
      If shp.Name Like "wall*" Then wallShapes.Add shp
      If shp.Name Like "door*" Then doorShapes.Add shp
      If shp.Name Like "trigger*" Then triggerShapes.Add shp
    Next
  End With
Dim leftState As Boolean, upState As Boolean, rightState As Boolean, downState As Boolean
Dim newLeft As Boolean, newUp As Boolean, newRight As Boolean, newDown As Boolean, newZ As Boolean, lastZ As Boolean
movementBlocked = False: lastZ = False
Set dlgBox = Nothing
On Error Resume Next
Set dlgBox = ActivePresentation.SlideShowWindow.View.slide.Shapes("dialogueBox")
On Error GoTo 0
If Not dlgBox Is Nothing Then dlgBox.Visible = msoFalse
Dim dialogState As Integer, animEndTime As Double: dialogState = 0
Do While currentPosition > 0
  moveX = 0: moveY = 0: idleShape.Visible = True: movingShape.Visible = False
  newZ = (GetAsyncKeyState(keyZ) <> 0)
If Not dlgBox Is Nothing Then
If dialogState = 1 Then
  If Timer >= animEndTime Then dlgBox.AnimationSettings.Animate = False: dialogState = 2
End If
 If Not currentTrigger Is Nothing And newZ And Not lastZ Then
  Select Case dialogState
  Case 0
   movementBlocked = True: dialogSystem: dlgBox.Visible = msoTrue: dlgBox.ZOrder msoBringToFront
    animEndTime = Timer + Len(dlgBox.TextFrame.TextRange.text) * dialogSpeed: dialogState = 1
  Case 1
    dlgBox.AnimationSettings.Animate = False: dialogState = 2
  Case 2
    movementBlocked = False: dlgBox.Visible = msoFalse: dialogState = 0
  End Select
 End If
End If
 lastZ = newZ
If Not movementBlocked Then
  newLeft = (GetAsyncKeyState(keyLeft) <> 0): newUp = (GetAsyncKeyState(keyUp) <> 0): newRight = (GetAsyncKeyState(keyRight) <> 0): newDown = (GetAsyncKeyState(keyDown) <> 0)
  If newLeft And Not leftState Then lastKey = "left"
  If newUp And Not upState Then lastKey = "up"
  If newRight And Not rightState Then lastKey = "right"
  If newDown And Not downState Then lastKey = "down"
  Select Case lastKey
     Case "left": If Not newLeft Then lastKey = ""
     Case "up": If Not newUp Then lastKey = ""
     Case "right": If Not newRight Then lastKey = ""
     Case "down": If Not newDown Then lastKey = ""
  End Select
  If lastKey = "" Then
     If newLeft Then lastKey = "left"
     If newUp Then lastKey = "up"
     If newRight Then lastKey = "right"
     If newDown Then lastKey = "down"
  End If
  leftState = newLeft: upState = newUp: rightState = newRight: downState = newDown
 Select Case lastKey
      Case "left": moveX = -playerSpeed: SetShapeImages "idle_r.gif", "walk_r.gif", 180
      Case "up": moveY = -playerSpeed: SetShapeImages "idle_u.gif", "walk_u.gif"
      Case "right": moveX = playerSpeed: SetShapeImages "idle_r.gif", "walk_r.gif", 0
      Case "down": moveY = playerSpeed: SetShapeImages "idle_d.gif", "walk_d.gif"
  End Select
If showDialogue And newZ And Not lastZ Then dialogSystem
lastZ = newZ
  If moveX <> 0 Or moveY <> 0 Then
     ActivePresentation.Slides(1).Shapes("tiempo").TextFrame.TextRange.text = Time: idleShape.Visible = False: movingShape.Visible = True
     MoveShape idleShape, moveX, moveY, slideWidth, slideHeight, wallShapes, doorShapes, triggerShapes: MoveShape movingShape, moveX, moveY, slideWidth, slideHeight, wallShapes, doorShapes, triggerShapes
  End If
 If Timer - startTime >= resetInterval Then
     ActivePresentation.SlideShowWindow.View.GotoSlide currentPosition: startTime = Timer
 End If
End If
  DoEvents
Loop
End Sub
 Private Sub MoveShape(s As shape, moveX As Single, moveY As Single, slideWidth As Integer, slideHeight As Integer, wallShapes As Collection, doorShapes As Collection, triggerShapes As Collection)
   Dim wallShape As shape, doorShape As shape, triggerShape As shape, doorNumber As Integer, posX As Integer, posY As Integer
  s.Left = s.Left + moveX: s.Top = s.Top + moveY
   If s.Left < 0 Then s.Left = 0
   If s.Top < 0 Then s.Top = 0
   If s.Left + s.Width > slideWidth Then s.Left = slideWidth - s.Width
   If s.Top + s.Height > slideHeight Then s.Top = slideHeight - s.Height
    For Each wallShape In wallShapes
       If Colision(s, wallShape) Then
           s.Left = s.Left - moveX: s.Top = s.Top - moveY
       End If
   Next
For Each doorShape In doorShapes
    ExtractDoorData doorShape, doorNumber, posX, posY
 If Colision(s, doorShape) Then
  idleShape.Left = movingShape.Left: idleShape.Top = movingShape.Top
  idleShape.Visible = True: movingShape.Visible = False
 With ActivePresentation.Slides(doorNumber)
     Set idleShape = .Shapes("playerIdle"): Set movingShape = .Shapes("playerMoving")
     idleShape.Left = posX: idleShape.Top = posY: movingShape.Left = posX: movingShape.Top = posY
 End With
 Select Case lastKey
   Case "left": SetShapeImages "idle_r.gif", "walk_r.gif", 180, True
   Case "up": SetShapeImages "idle_u.gif", "walk_u.gif", 0, True
   Case "right": SetShapeImages "idle_r.gif", "walk_r.gif", 0, True
   Case "down": SetShapeImages "idle_d.gif", "walk_d.gif", 0, True
 End Select
 ActivePresentation.SlideShowWindow.View.GotoSlide doorNumber
 idleShape.Visible = True: movingShape.Visible = False
 Set idleShape = Nothing: Set movingShape = Nothing: Set wallShapes = Nothing: Set doorShapes = Nothing: PlayerMovement
  End If
Next
For Each triggerShape In triggerShapes
  If Colision(s, triggerShape) Then
      Set currentTrigger = triggerShape: showDialogue = True: Exit For
  Else
      showDialogue = False: Set currentTrigger = Nothing
  End If
Next
End Sub
Function Colision(a As shape, b As shape) As Boolean
  Colision = Not (a.Left + a.Width < b.Left Or a.Left > b.Left + b.Width Or a.Top + a.Height < b.Top Or a.Top > b.Top + b.Height)
End Function
Private Sub SetShapeImages(idleImage As String, movingImage As String, Optional rotationX As Integer = 0, Optional forceUpdate As Boolean = False)
  Static lastIdleImage As String, lastMovingImage As String
   If forceUpdate Or lastIdleImage <> idleImage Then
     If dir(imagePath & idleImage) <> "" Then idleShape.Fill.UserPicture imagePath & idleImage: lastIdleImage = idleImage
   End If
   If forceUpdate Or lastMovingImage <> movingImage Then
     If dir(imagePath & movingImage) <> "" Then movingShape.Fill.UserPicture imagePath & movingImage: lastMovingImage = movingImage
   End If
   idleShape.ThreeD.rotationX = rotationX: movingShape.ThreeD.rotationX = rotationX
End Sub
Private Sub ExtractDoorData(doorShape As shape, ByRef doorNumber As Integer, ByRef posX As Integer, ByRef posY As Integer)
  Dim shapeName As String, parts() As String
  shapeName = doorShape.Name: parts = Split(shapeName, "_")
  If UBound(parts) = 3 Then doorNumber = CInt(parts(1)): posX = CInt(parts(2)): posY = CInt(parts(3))
End Sub
Private Sub dialogSystem()
    If currentTrigger Is Nothing Then Exit Sub
    Dim partes() As String
    partes = Split(currentTrigger.Name, "_")
    Dim idx As String
    idx = partes(1)
    Dim contenido As String
    On Error Resume Next
   contenido = ActivePresentation.SlideShowWindow.View.slide.Shapes("txtSlideDialogs").TextFrame.TextRange.text
   On Error GoTo 0
   If contenido = "" Then Exit Sub
    contenido = Replace(contenido, vbCrLf, vbLf)
    contenido = Replace(contenido, vbCr, vbLf)
    Dim lineas() As String
    lineas = Split(contenido, vbLf)
    Dim l As Variant, mensaje As String
    For Each l In lineas
        l = Trim(l)
        If l <> "" And Left(l, Len(idx & "_")) = idx & "_" Then
            mensaje = Mid$(l, Len(idx & "_") + 1)
            Exit For
        End If
    Next
If Not dlgBox Is Nothing Then
  dlgBox.TextFrame.TextRange.text = mensaje
  With dlgBox.AnimationSettings
      .AdvanceMode = ppAdvanceOnTime: .AdvanceTime = 0: .EntryEffect = ppEffectFade
      .TextLevelEffect = ppAnimateByFirstLevel: .TextUnitEffect = ppAnimateByCharacter: .Animate = True
  End With
On Error Resume Next
 ActivePresentation.SlideShowWindow.View.slide.TimeLine.MainSequence.Item(1).Timing.Duration = dialogSpeed
On Error GoTo 0
End If
End Sub
