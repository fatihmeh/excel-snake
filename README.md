# Excel-Snake


![](/snake_process_flowchart.png)


```vba

Option Explicit
Dim arrSnake()
Dim i, k, j, g, ofsColumn, ofsRow, tempColumn, tempRow, lengthArr As Integer

'snake body and head color values
Const headPaint As Integer = 37
Const bodyPaint As Integer = 14

Dim change, yemBool, st As Boolean
Dim yem As Range
Dim direct As String

'reset button
Sub resetSnake(ByVal control As IRibbonControl)
st = False
direct = ""
change = True
yemBool = False
On Error Resume Next
yem.Value = ""
clearPaint
End Sub

'direction buttons
Private Function moveRight(ByVal control As IRibbonControl)
direct = "right"
premove True
End Function
Private Function moveDown(ByVal control As IRibbonControl)
direct = "down"
premove True
End Function
Private Function moveUp(ByVal control As IRibbonControl)
direct = "up"
premove True
End Function
Private Function moveLeft(ByVal control As IRibbonControl)
direct = "left"
premove True
End Function

'stop button
Sub stopSnake(ByVal control As IRibbonControl)
stopAction
End Sub

'start button
Sub startSnake(ByVal control As IRibbonControl)
gogogo
End Sub

Private Function stopAction()

st = False
direct = ""
change = True
yemBool = False

'if any button is pressed before creation of yem value, if will throw an error. so resume next is safe to use here
On Error Resume Next
yem.Value = ""

End Function

Private Sub gogogo()

yemBool = False
st = True
arrSnake = Array("J11", "K11", "L11")
direct = "right"
createYem
premove False

End Sub

Private Function premove(change)

Do While st = True
If change = True Then Exit Do
Delay 1
clearPaint
moveSnake
If st = True Then paintSnake
Loop

End Function

Private Function moveSnake()

    Select Case direct
    Case "right"
    ofsColumn = 1
    ofsRow = 0
    Case "left"
    ofsColumn = -1
    ofsRow = 0
    Case "up"
    ofsColumn = 0
    ofsRow = -1
    Case "down"
    ofsColumn = 0
    ofsRow = 1
    Case ""
    Exit Function
    Case Else
    End Select


If st = False Then Exit Function

    'update snake coordinations in array
    For i = LBound(arrSnake) To UBound(arrSnake)
            If i = UBound(arrSnake) Then
            'Debug.Print Sheet1.Range(arrSnake(i)).Offset(ofsRow, ofsColumn).Address
            arrSnake(i) = Sheet1.Range(arrSnake(i)).Offset(ofsRow, ofsColumn).Address
            Else
            'Debug.Print arrSnake(i)
            arrSnake(i) = arrSnake(i + 1)
            End If
    Next i
    
    crashControl
    yemControl

End Function

Private Function growSnake()

        If yemBool = True Then
            
                Select Case checkDirection
                Case "right"
                ofsColumn = 1
                ofsRow = 0
                Case "left"
                ofsColumn = -1
                ofsRow = 0
                Case "up"
                ofsColumn = 0
                ofsRow = -1
                Case "down"
                ofsColumn = 0
                ofsRow = 1
                Case "caresiz"
                Debug.Print ""
                Exit Function
                Case Else
                End Select
            
            'extend snake
            ReDim Preserve arrSnake(UBound(arrSnake) + 1)
            'reverse loop to update location values, because new addition tile adds up to bottom of the snake
            For k = UBound(arrSnake) To LBound(arrSnake) Step -1
                If k <> 0 Then
                arrSnake(k) = arrSnake(k - 1)
                End If
                If k = 0 Then
                arrSnake(k) = Sheet1.Range(arrSnake(k + 1)).Offset(ofsRow * -1, ofsColumn * -1).Address
                End If
            Next k
            yemBool = False
        End If

End Function


Private Function crashControl()

lengthArr = UBound(arrSnake)

    If (Sheet1.Range(arrSnake(lengthArr)).Parent.Name = Sheet1.Range("AJ1:BE22").Parent.Name) Then
        Dim ints As Range
        Set ints = Application.Intersect(Sheet1.Range(arrSnake(UBound(arrSnake))), Sheet1.Range("B2:U21"))
        If (ints Is Nothing) Then
            MsgBox "Aaa kaza!"
            stopAction
            Exit Function
        End If

    End If
    
End Function

Private Function yemControl()

Dim yemCrash As Range

If (Sheet1.Range(arrSnake(lengthArr)).Parent.Name = Sheet1.Range("AJ1:BE22").Parent.Name) Then
Set yemCrash = Application.Intersect(Sheet1.Range(arrSnake(UBound(arrSnake))), yem)
        If (yemCrash Is Nothing) Then
        yemBool = False
        Else
            'Debug.Print "yem eaten!"
            yem.Value = ""
            createYem
            yemBool = True
            growSnake
            Exit Function
        End If
End If

End Function

Private Function checkDirection() As String

'this functions checks snake's bottom tiles' direction
'and uses this in growSnake function to decide which direction to grow bottom tiles
'there was a problem when snake's direction changed via buttons, bottom tiles growed regardless of current placement of snake


checkDirection = ""

Dim firstRow, firstColumn, secondRow, secondColumn As String

firstRow = Range(arrSnake(LBound(arrSnake))).Row
firstColumn = Range(arrSnake(LBound(arrSnake))).Column
secondRow = Range(arrSnake(LBound(arrSnake) + 1)).Row
secondColumn = Range(arrSnake(LBound(arrSnake) + 1)).Column

'Debug.Print "firstRow = " & firstRow & ", firstColumn = " & firstColumn
'Debug.Print "secondRow = " & secondRow & ", secondColumn = " & secondColumn

If firstRow = secondRow And firstColumn < secondColumn Then checkDirection = "right"
If firstRow = secondRow And firstColumn > secondColumn Then checkDirection = "left"
If firstRow < secondRow And firstColumn = secondColumn Then checkDirection = "down"
If firstRow > secondRow And firstColumn = secondColumn Then checkDirection = "up"
If firstRow <> secondRow And firstColumn <> secondColumn Then checkDirection = "caresiz"

End Function

Private Function clearPaint()

Sheet1.Range("A1:CB50").Interior.ColorIndex = 0

End Function
Private Function paintSnake()

lengthArr = UBound(arrSnake)
For j = LBound(arrSnake) To UBound(arrSnake)
    Sheet1.Range(arrSnake(j)).Interior.ColorIndex = headPaint
Next j
    Sheet1.Range(arrSnake(lengthArr)).Interior.ColorIndex = bodyPaint
End Function

Private Function createYem()

Dim yemAddress As String
yemAddress = Sheet1.Cells(rndTest, rndTest).Address
'yemAddress = "N11"
'Debug.Print yemAddress
Set yem = Sheet1.Range(yemAddress)
yem.Value = "•"

End Function
Private Function rndTest() As Integer

Dim myRnd As Integer
myRnd = Int(2 + rnd * (21 - 2 + 1))
rndTest = myRnd

End Function

Private Function Delay(count As Long)

Dim start As Long
start = Timer
Do While Timer < start + 1
DoEvents
Loop

End Function

```