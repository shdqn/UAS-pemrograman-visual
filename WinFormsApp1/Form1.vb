Imports System.Reflection.Emit

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Meminta pengguna memilih file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Text files (.txt)| .txt"
        openFileDialog.Title = "Select a text file"
        If openFileDialog.ShowDialog() <> DialogResult.OK Then
            Return
        End If

        ' Membaca file dan menampilkan isinya pada label
        Dim fileContents As String = System.IO.File.ReadAllText(openFileDialog.FileName)
        Label1.Text = fileContents

        ' Memproses teks yang telah dibaca dari file
        Dim words() As String = fileContents.Split(" ")
        Dim wordCount As Integer = words.Length
        Dim longestWord As String = ""
        For Each word As String In words
            ' Mencari kata terpanjang
            If word.Length > longestWord.Length Then
                longestWord = word
            End If
        Next
        Label2.Text = "Jumlah kata: " & wordCount

        ' Menampilkan pesan berbeda tergantung pada panjang kata terpanjang
        If longestWord.Length > 10 Then
            MessageBox.Show("Kata terpanjang memiliki lebih dari 10 huruf.")
        Else
            MessageBox.Show("Kata terpanjang memiliki 10 huruf atau kurang.")
        End If

        ' Memanggil procedure untuk menghitung jumlah karakter pada file
        Dim characterCount As Integer = CountCharacters(fileContents)
        Label3.Text = "Jumlah karakter: " & characterCount
    End Sub

    Private Function CountCharacters(text As String) As Integer
        ' Menghitung jumlah karakter pada teks
        Dim characterCount As Integer = 0
        For Each c As Char In text
            characterCount += 1
        Next
        Return characterCount
    End Function

End Class
