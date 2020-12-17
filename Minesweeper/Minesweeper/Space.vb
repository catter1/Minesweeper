Public Class Space
    Inherits PictureBox
    Dim M As Boolean 'Is it hiding a mine?
    Dim S As Boolean 'Has it been revealed yet?
    Dim N As Integer 'How many mines are near it?
    Dim H As Boolean 'Did it get checked yet?
    Dim F As Integer '0 = Flag, 1 = Question, 2 = Nothing
    Dim r, c As Integer 'Row and column

    Public Sub New()
        Me.Size = New Size(40, 40)
        Me.SizeMode = PictureBoxSizeMode.StretchImage
        Me.Image = Image.FromFile("unrevealed.png")
        Me.Revealed = False
        Me.IsMine = False
    End Sub

    Public Property Revealed As Boolean
        Get
            Return S
        End Get
        Set(value As Boolean)
            S = value
        End Set
    End Property
    Public Property Checked As Boolean
        Get
            Return H
        End Get
        Set(value As Boolean)
            H = value
        End Set
    End Property
    Public Property IsMine As Boolean
        Get
            Return M
        End Get
        Set(value As Boolean)
            M = value
        End Set
    End Property
    Public Property Flagged As Integer
        Get
            Return F
        End Get
        Set(value As Integer)
            F = value
        End Set
    End Property
    Public Property Number As Integer
        Get
            Return N
        End Get
        Set(value As Integer)
            N = value
        End Set
    End Property
    Public Property Row As Integer
        Get
            Return r
        End Get
        Set(value As Integer)
            r = value
        End Set
    End Property
    Public Property Column As Integer
        Get
            Return c
        End Get
        Set(value As Integer)
            c = value
        End Set
    End Property

    Public Sub search()
        If Me.Revealed = False And Me.Flagged = 0 Then
            Me.Revealed = True
            If Me.IsMine = False Then
                If Me.Number = 0 Then Me.Image = Image.FromFile("revealed.png")
                If Me.Number = 1 Then Me.Image = Image.FromFile("one.png")
                If Me.Number = 2 Then Me.Image = Image.FromFile("two.png")
                If Me.Number = 3 Then Me.Image = Image.FromFile("three.png")
                If Me.Number = 4 Then Me.Image = Image.FromFile("four.png")
                If Me.Number = 5 Then Me.Image = Image.FromFile("five.png")
                If Me.Number = 6 Then Me.Image = Image.FromFile("six.png")
                If Me.Number = 7 Then Me.Image = Image.FromFile("seven.png")
                If Me.Number = 8 Then Me.Image = Image.FromFile("eight.png")
            End If
        End If
    End Sub
    Public Sub flag()
        If Me.Revealed = False Then
            If Me.Flagged = 0 Then
                Me.Image = Image.FromFile("flag.png")
                Me.Flagged = 1
                Form1.deCount(True)
            ElseIf Me.Flagged = 1 Then
                Me.Image = Image.FromFile("question.png")
                Me.Flagged = 2
            Else
                Me.Image = Image.FromFile("unrevealed.png")
                Me.Flagged = 0
                Form1.deCount(False)
            End If
        End If
    End Sub
    Public Sub revealMine(ByVal flagged As Boolean)
        If flagged = False Then Me.Image = Image.FromFile("mine.png")
        If flagged Then Me.Image = Image.FromFile("xmine.png")
    End Sub
End Class
