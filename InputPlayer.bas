Attribute VB_Name = "InputPlayer"
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
Dim imagePath As String, idleShape As Shape, movingShape As Shape, startTime As Single
Public Sub PlayerMovement()
 Dim initialSlide%, playerSpeed As Single, slideWidth As Single, slideHeight As Single, moveX As Single, moveY As Single, wallShape As Shape, wallShapes As New Collection
 Dim shp As Shape, allSlides As Slide
 imagePath = ActivePresentation.Path & "\data\": initialSlide = 1
 playerSpeed = Val(ActivePresentation.Slides(initialSlide).Shapes("playerSpeedText").TextFrame2.TextRange.Text)
 ActivePresentation.SlideShowWindow.View.PointerType = 3
 startTime = Timer
 With SlideShowWindows(1).View

    .GotoSlide initialSlide: slideWidth = .Slide.Master.Width: slideHeight = .Slide.Master.Height
  Set idleShape = .Slide.Shapes("playerIdle"): Set movingShape = .Slide.Shapes("playerMoving")
 For Each wallShape In .Slide.Shapes
  If wallShape.Name Like "wall*" Then wallShapes.Add wallShape
 Next
  End With
  Do While SlideShowWindows(1).View.CurrentShowPosition > 0

       moveX = 0: moveY = 0: idleShape.Visible = 1: movingShape.Visible = 0
  On Error Resume Next
     If GetAsyncKeyState(vbKeyA) Then moveX = -playerSpeed: SetShapeImages "idle_r.gif", "walk_r.gif", 180
     If GetAsyncKeyState(vbKeyW) Then moveY = -playerSpeed: SetShapeImages "idle_u.gif", "walk_u.gif"
     If GetAsyncKeyState(vbKeyS) Then moveY = playerSpeed: SetShapeImages "idle_d.gif", "walk_d.gif"
     If GetAsyncKeyState(vbKeyD) Then moveX = playerSpeed: SetShapeImages "idle_r.gif", "walk_r.gif", 0
  On Error GoTo 0
        If moveX <> 0 Or moveY <> 0 Then
          idleShape.Visible = 0: movingShape.Visible = 1
          MoveShape idleShape, moveX, moveY, slideWidth, slideHeight, wallShapes: MoveShape movingShape, moveX, moveY, slideWidth, slideHeight, wallShapes
       End If
       ActivePresentation.Slides(initialSlide).Shapes("tiempo").TextFrame2.TextRange.Text = Time: DoEvents
   Loop
 End Sub
 Private Sub MoveShape(s As Shape, moveX As Single, moveY As Single, slideWidth As Single, slideHeight As Single, wallShapes As Collection)
   Dim wallShape As Shape, isCollision As Boolean
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
 End Sub
Private Sub SetShapeImages(idleImage As String, movingImage As String, Optional rotationX As Double = 0)
idleShape.Fill.UserPicture imagePath & idleImage: movingShape.Fill.UserPicture imagePath & movingImage
idleShape.ThreeD.rotationX = rotationX: movingShape.ThreeD.rotationX = rotationX
 End Sub