Public Class Form1
    Dim board(15, 15) As Space
    Dim r, c, z, mineCount, rightMines As Integer
    Dim lblTitle, lblTimer, lblMines As Label
    Dim started As Boolean = False
    Dim ended As Boolean = False
    Dim maybeWin As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Randomize()
        buildForm()
    End Sub
    Public Sub buildForm()
        Me.Size = New Size(716, 804)
        addLabels()
        buildBoard()
        setMines()
    End Sub
    Public Sub addLabels()
        lblTitle = New Label
        lblTitle.Text = "Minesweeper"
        lblTitle.Font = New Font("Engravers MT", 20)
        lblTitle.BorderStyle = BorderStyle.FixedSingle
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        lblTitle.Location = New Point(200, 50)
        lblTitle.Size = New Size(300, 40)
        Controls.Add(lblTitle)

        lblTimer = New Label
        lblTimer.Text = 0
        lblTimer.Font = New Font("Engravers MT", 20)
        lblTimer.BorderStyle = BorderStyle.FixedSingle
        lblTimer.TextAlign = ContentAlignment.MiddleRight
        lblTimer.Location = New Point(570, 50)
        lblTimer.Size = New Size(80, 40)
        Controls.Add(lblTimer)

        lblMines = New Label
        lblMines.Text = 30
        lblMines.Font = New Font("Engravers MT", 20)
        lblMines.BorderStyle = BorderStyle.FixedSingle
        lblMines.TextAlign = ContentAlignment.MiddleLeft
        lblMines.Location = New Point(50, 50)
        lblMines.Size = New Size(80, 40)
        Controls.Add(lblMines)
    End Sub
    Public Sub buildBoard()
        For Me.r = 0 To 14
            For Me.c = 0 To 14
                board(c, r) = New Space
                board(c, r).Row = r
                board(c, r).Column = c
                board(c, r).Location = New Point(c * 40 + 50, r * 40 + 120)
                Controls.Add(board(c, r))
                AddHandler board(c, r).MouseDown, AddressOf SearchBoard
            Next
        Next
    End Sub
    Public Sub setMines()
        For Me.z = 1 To 30
Here:
            r = Int(Rnd() * 14)
            c = Int(Rnd() * 14)
            If board(c, r).IsMine = True Then
                GoTo Here
            Else
                board(c, r).IsMine = True
            End If
        Next
        setNumbers()
    End Sub
    Public Sub SearchBoard(ByVal sender As Object, ByVal e As MouseEventArgs)
        If started = False And Not ended Then
            started = True
        End If
        If e.Button = MouseButtons.Left And Not ended Then
            If board(sender.column, sender.row).IsMine And board(sender.column, sender.row).Flagged = 0 Then
                endGame()
                rightMines = 0
                For Me.r = 0 To 4
                    For Me.c = 0 To 4
                        If board(c, r).IsMine And board(c, r).Flagged = 1 Then rightMines += 1
                    Next
                Next
                MsgBox("You lost!" & Environment.NewLine & Environment.NewLine & Environment.NewLine & "Press 'Menu' >> 'Reset Game' to play again.")
                Exit Sub
            End If

            For Me.r = 0 To 4
                For Me.c = 0 To 4
                    board(c, r).Checked = False
                Next
            Next
            If board(sender.column, sender.row).Flagged = 0 Then
                sender.search()
                sender.Checked = True
                If board(sender.column, sender.row).Number = 0 Then
                    If sender.column > 0 Then
                        board(sender.column - 1, sender.row).search()
                        If board(sender.column - 1, sender.row).Number = 0 Then SearchContinued(sender.column - 1, sender.row)
                    End If
                    If sender.column < 14 Then
                        board(sender.column + 1, sender.row).search()
                        If board(sender.column + 1, sender.row).Number = 0 Then SearchContinued(sender.column + 1, sender.row)
                    End If
                    If sender.row > 0 Then
                        board(sender.column, sender.row - 1).search()
                        If board(sender.column, sender.row - 1).Number = 0 Then SearchContinued(sender.column, sender.row - 1)
                    End If
                    If sender.row < 14 Then
                        board(sender.column, sender.row + 1).search()
                        If board(sender.column, sender.row + 1).Number = 0 Then SearchContinued(sender.column, sender.row + 1)
                    End If
                    If sender.column > 0 And sender.row > 0 Then
                        board(sender.column - 1, sender.row - 1).search()
                        If board(sender.column - 1, sender.row - 1).Number = 0 Then SearchContinued(sender.column - 1, sender.row - 1)
                    End If
                    If sender.column < 14 And sender.row < 14 Then
                        board(sender.column + 1, sender.row + 1).search()
                        If board(sender.column + 1, sender.row + 1).Number = 0 Then SearchContinued(sender.column + 1, sender.row + 1)
                    End If
                    If sender.column > 0 And sender.row < 14 Then
                        board(sender.column - 1, sender.row + 1).search()
                        If board(sender.column - 1, sender.row + 1).Number = 0 Then SearchContinued(sender.column - 1, sender.row + 1)
                    End If
                    If sender.column < 14 And sender.row > 0 Then
                        board(sender.column + 1, sender.row - 1).search()
                        If board(sender.column + 1, sender.row - 1).Number = 0 Then SearchContinued(sender.column + 1, sender.row - 1)
                    End If
                End If
            End If
        End If
        If e.Button = MouseButtons.Right And Not ended Then sender.flag()

        checkWin()
    End Sub
    Public Sub SearchContinued(ByVal c As Integer, ByVal r As Integer)
        board(c, r).search()
        board(c, r).Checked = True

        If board(c, r).Number = 0 Then
            If c > 0 Then
                board(c - 1, r).search()
                If board(c - 1, r).Number = 0 And board(c - 1, r).Checked = False Then SearchContinued(c - 1, r)
            End If
            If c < 14 Then
                board(c + 1, r).search()
                If board(c + 1, r).Number = 0 And board(c + 1, r).Checked = False Then SearchContinued(c + 1, r)
            End If
            If r > 0 Then
                board(c, r - 1).search()
                If board(c, r - 1).Number = 0 And board(c, r - 1).Checked = False Then SearchContinued(c, r - 1)
            End If
            If r < 14 Then
                board(c, r + 1).search()
                If board(c, r + 1).Number = 0 And board(c, r + 1).Checked = False Then SearchContinued(c, r + 1)
            End If
            If c > 0 And r > 0 Then
                board(c - 1, r - 1).search()
                If board(c - 1, r - 1).Number = 0 And board(c - 1, r - 1).Checked = False Then SearchContinued(c - 1, r - 1)
            End If
            If c < 14 And r < 14 Then
                board(c + 1, r + 1).search()
                If board(c + 1, r + 1).Number = 0 And board(c + 1, r + 1).Checked = False Then SearchContinued(c + 1, r + 1)
            End If
            If c > 0 And r < 14 Then
                board(c - 1, r + 1).search()
                If board(c - 1, r + 1).Number = 0 And board(c - 1, r + 1).Checked = False Then SearchContinued(c - 1, r + 1)
            End If
            If c < 14 And r > 0 Then
                board(c + 1, r - 1).search()
                If board(c + 1, r - 1).Number = 0 And board(c + 1, r - 1).Checked = False Then SearchContinued(c + 1, r - 1)
            End If
        End If
    End Sub
    Public Sub setNumbers()
        For Me.r = 0 To 14
            For Me.c = 0 To 14
                mineCount = 0
                If c > 0 Then
                    If board(c - 1, r).IsMine Then mineCount += 1
                End If
                If c < 14 Then
                    If board(c + 1, r).IsMine Then mineCount += 1
                End If
                If r > 0 Then
                    If board(c, r - 1).IsMine Then mineCount += 1
                End If
                If r < 14 Then
                    If board(c, r + 1).IsMine Then mineCount += 1
                End If
                If c > 0 And r > 0 Then
                    If board(c - 1, r - 1).IsMine Then mineCount += 1
                End If
                If c < 14 And r < 14 Then
                    If board(c + 1, r + 1).IsMine Then mineCount += 1
                End If
                If c > 0 And r < 14 Then
                    If board(c - 1, r + 1).IsMine Then mineCount += 1
                End If
                If c < 14 And r > 0 Then
                    If board(c + 1, r - 1).IsMine Then mineCount += 1
                End If
                board(c, r).Number = mineCount
            Next
        Next
    End Sub
    Public Sub checkWin()
        maybeWin = True
        For Me.r = 0 To 14
            For Me.c = 0 To 14
                If Not board(c, r).IsMine And Not board(c, r).Revealed Then
                    maybeWin = False
                End If
            Next
        Next
        If maybeWin Then
            endGame()
            MsgBox("You win!" & Environment.NewLine & Environment.NewLine & "Your time: " & lblTimer.Text & " seconds.")
        End If
    End Sub
    Public Sub endGame()
        started = False
        ended = True
        For Me.r = 0 To 14
            For Me.c = 0 To 14
                If board(c, r).Flagged = 1 And Not board(c, r).IsMine Then board(c, r).revealMine(True)
                If board(c, r).Flagged <> 1 And board(c, r).IsMine Then board(c, r).revealMine(False)
            Next
        Next
    End Sub
    Public Sub restartGame()
        lblTimer.Text = 0
        lblMines.Text = 30
        For Me.r = 0 To 14
            For Me.c = 0 To 14
                board(c, r).Dispose()
                board(c, r) = Nothing
            Next
        Next
        buildBoard()
        setMines()
        ended = False
    End Sub
    Public Sub deCount(ByVal flag As Boolean)
        If flag Then
            lblMines.Text -= 1
        Else
            lblMines.Text += 1
        End If
    End Sub

    Private Sub setRestart_Click(sender As Object, e As EventArgs) Handles setRestart.Click
        endGame()
        restartGame()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If started Then
            lblTimer.Text += 1
        End If
    End Sub
End Class
