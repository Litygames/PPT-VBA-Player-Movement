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
Public Sub HoverOver()
 initialSlide = 2
 ActivePresentation.SlideShowWindow.View.PointerType = 3
 SlideShowWindows(1).View.GotoSlide initialSlide
 PlayerMovement
 End Sub
Private Sub PlayerMovement()
 Dim playerSpeed As Single, slideWidth As Integer, slideHeight As Integer, moveX As Single, moveY As Single, wallShape As shape, wallShapes As New Collection, doorShape As shape, doorShapes As New Collection
 Dim shp As shape, allSlides As slide, keyLeft As Integer, keyUp As Integer, keyDown As Integer, keyRight As Integer, resetInterval As Single
 Dim currentPosition As Integer
 currentPosition = SlideShowWindows(1).View.CurrentShowPosition
 imagePath = ActivePresentation.Path & "\data\"
 playerSpeed = 3
 keyLeft = 65
 keyUp = 87
 keyDown = 83
 keyRight = 68
 resetInterval = 0.5
 startTime = Timer
 With SlideShowWindows(1).View
  slideWidth = .slide.Master.Width: slideHeight = .slide.Master.Height
  Set idleShape = .slide.Shapes("playerIdle"): Set movingShape = .slide.Shapes("playerMoving")
    For Each shp In .slide.Shapes
      If shp.Name Like "wall*" Then wallShapes.Add shp
      If shp.Name Like "door*" Then doorShapes.Add shp
    Next
  End With
Dim leftState As Boolean, upState As Boolean, rightState As Boolean, downState As Boolean, newLeft As Boolean, newUp As Boolean, newRight As Boolean, newDown As Boolean
Do While currentPosition > 0
  moveX = 0: moveY = 0: idleShape.Visible = True: movingShape.Visible = False
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
  If moveX <> 0 Or moveY <> 0 Then
     ActivePresentation.Slides(1).Shapes("tiempo").TextFrame2.TextRange.Text = Time: idleShape.Visible = False: movingShape.Visible = True
     MoveShape idleShape, moveX, moveY, slideWidth, slideHeight, wallShapes, doorShapes: MoveShape movingShape, moveX, moveY, slideWidth, slideHeight, wallShapes, doorShapes
  End If
 If Timer - startTime >= resetInterval Then
     SlideShowWindows(1).View.GotoSlide currentPosition: startTime = Timer
  End If
  DoEvents
Loop
 End Sub
 Private Sub MoveShape(s As shape, moveX As Single, moveY As Single, slideWidth As Integer, slideHeight As Integer, wallShapes As Collection, doorShapes As Collection)
   Dim wallShape As shape, doorShape As shape, isCollision As Boolean, isOpen As Boolean, doorNumber As Integer, posX As Integer, posY As Integer
  s.Left = s.Left + moveX: s.Top = s.Top + moveY
   If s.Left < 0 Then s.Left = 0
   If s.Top < 0 Then s.Top = 0
   If s.Left + s.Width > slideWidth Then s.Left = slideWidth - s.Width
   If s.Top + s.Height > slideHeight Then s.Top = slideHeight - s.Height
    For Each wallShape In wallShapes
      If Not (s.Left + s.Width < wallShape.Left Or s.Left > wallShape.Left + wallShape.Width Or s.Top + s.Height < wallShape.Top Or s.Top > wallShape.Top + wallShape.Height) Then
           isCollision = True: Exit For
       End If
   Next
   If isCollision Then s.Left = s.Left - moveX: s.Top = s.Top - moveY
For Each doorShape In doorShapes
    ExtractDoorData doorShape, doorNumber, posX, posY
  If Not (s.Left + s.Width < doorShape.Left Or s.Left > doorShape.Left + doorShape.Width Or s.Top + s.Height < doorShape.Top Or s.Top > doorShape.Top + doorShape.Height) Then
    isOpen = True: Exit For
  End If
Next
If isOpen Then
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
 SlideShowWindows(1).View.GotoSlide doorNumber
 idleShape.Visible = True: movingShape.Visible = False
 Set idleShape = Nothing: Set movingShape = Nothing: Set wallShapes = Nothing: Set doorShapes = Nothing: PlayerMovement
End If
End Sub
Private Sub SetShapeImages(idleImage As String, movingImage As String, Optional rotationX As Integer = 0, Optional forceUpdate As Boolean = False)
  Static lastIdleImage As String, lastMovingImage As String
   If forceUpdate Or lastIdleImage <> idleImage Then idleShape.Fill.UserPicture imagePath & idleImage: lastIdleImage = idleImage
   If forceUpdate Or lastMovingImage <> movingImage Then movingShape.Fill.UserPicture imagePath & movingImage: lastMovingImage = movingImage
   idleShape.ThreeD.rotationX = rotationX: movingShape.ThreeD.rotationX = rotationX
End Sub
Private Sub ExtractDoorData(doorShape As shape, ByRef doorNumber As Integer, ByRef posX As Integer, ByRef posY As Integer)
  Dim shapeName As String, parts() As String
  shapeName = doorShape.Name: parts = Split(shapeName, "_")
  If UBound(parts) = 3 Then doorNumber = CInt(parts(1)): posX = CInt(parts(2)): posY = CInt(parts(3))
End Sub