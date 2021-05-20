Public Class Form1
    Dim picture As New OpenFileDialog()
    Dim image, grey, sobel, buffer As Bitmap
    Dim i, j, value, r3, r4, valX, valY, gradient As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Dim GX(3, 3) As Integer
    Dim GY(3, 3) As Integer
    Dim color As Color
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        picture.Filter = "File Type|*.jpg;*.png;*.bmp"
        picture.Title = "Choose Picture"
        If picture.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
            Return
        End If
        PictureBox1.ImageLocation = picture.FileName
        Button3.Enabled = True
    End Sub

    Function makeGrey(ByVal bmp As Bitmap) As Bitmap

        For i = 0 To bmp.Height - 1
            For j = 0 To bmp.Width - 1
                color = bmp.GetPixel(j, i)
                r3 = color.G
                r4 = color.B
                value = (bmp.GetPixel(j, i).R + r3 + r4) / 3
                bmp.SetPixel(j, i, Color.FromArgb(bmp.GetPixel(j, i).A, value, value, value))
            Next
        Next
        Return bmp
    End Function

    Function applySobel(ByVal image As Bitmap) As Bitmap
        grey = makeGrey(image)
        buffer = New Bitmap(grey.Width, grey.Height)

        GX(0, 0) = -1
        GX(0, 1) = 0
        GX(0, 2) = 1

        GX(1, 0) = -2
        GX(1, 1) = 0
        GX(1, 2) = 2

        GX(2, 0) = -1
        GX(2, 1) = 0
        GX(2, 2) = 1

        GY(0, 0) = -1
        GY(0, 1) = -2
        GY(0, 2) = -1

        GY(1, 0) = 0
        GY(1, 1) = 0
        GY(1, 2) = 0

        GY(2, 0) = 1
        GY(2, 1) = 2
        GY(2, 2) = 1
        For i = 0 To grey.Height - 1
            For j = 0 To grey.Width - 1

                If i = 0 Or i = grey.Height - 1 Or j = 0 Or j = grey.Width - 1 Then
                    color = Color.FromArgb(255, 255, 255)
                    buffer.SetPixel(j, i, color)
                    valX = 0
                    valY = 0
                Else
                    valX = grey.GetPixel(j - 1, i - 1).R * GX(0, 0) +
                          grey.GetPixel(j, i - 1).R * GX(0, 1) +
                          grey.GetPixel(j + 1, i - 1).R * GX(0, 2) +
                          grey.GetPixel(j - 1, i).R * GX(1, 0) +
                          grey.GetPixel(j, i).R * GX(1, 1) +
                          grey.GetPixel(j + 1, i).R * GX(1, 2) +
                          grey.GetPixel(j - 1, i + 1).R * GX(2, 0) +
                          grey.GetPixel(j, i + 1).R * GX(2, 1) +
                          grey.GetPixel(j + 1, i + 1).R * GX(2, 2)

                    valY = grey.GetPixel(j - 1, i - 1).R * GY(0, 0) +
                         grey.GetPixel(j, i - 1).R * GY(0, 1) +
                         grey.GetPixel(j + 1, i - 1).R * GY(0, 2) +
                         grey.GetPixel(j - 1, i).R * GY(1, 0) +
                         grey.GetPixel(j, i).R * GY(1, 1) +
                         grey.GetPixel(j + 1, i).R * GY(1, 2) +
                         grey.GetPixel(j - 1, i + 1).R * GY(2, 0) +
                         grey.GetPixel(j, i + 1).R * GY(2, 1) +
                         grey.GetPixel(j + 1, i + 1).R * GY(2, 2)

                    gradient = Int(Math.Abs(valX) + Math.Abs(valY))
                    If gradient < 0 Then
                        gradient = 0
                    End If
                    If gradient > 255 Then
                        gradient = 255
                    End If

                    buffer.SetPixel(j, i, Color.FromArgb(gradient, gradient, gradient))

                End If
            Next
        Next
        Return buffer
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        image = New Bitmap(PictureBox1.Image)
        sobel = applySobel(image)
        PictureBox2.Image = sobel
    End Sub



End Class
