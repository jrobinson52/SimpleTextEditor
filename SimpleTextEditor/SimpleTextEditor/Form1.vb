Imports System.IO

Public Class Form1
    ' Class level variables
    Private strFilename As String = String.Empty    ' Document filename
    Dim blnIsChanged As Boolean = False             ' file change flag

    Sub ClearDocument()
        ' Clear the contents of the text box
        txtDocument.Clear()

        ' Clear the document name
        strFilename = String.Empty

        ' Set isChanged to False
        blnIsChanged = False
    End Sub

    ' The OpenDocument procedure opens a file and loads it
    ' into the TextBox for editing

    Sub OpenDocument()
        Dim inputFile As StreamReader ' object variable

        If ofdOpenFile.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' Retrieve the selected filename
            strFilename = ofdOpenFile.FileName

            Try
                ' Open the file
                inputFile = File.OpenText(strFilename)

                ' Read the file's contents into the textbox
                txtDocument.Text = inputFile.ReadToEnd

                ' Close the file.
                inputFile.Close()

                ' Update the isChanged variable
                blnIsChanged = False
            Catch ex As Exception
                ' Error message for file open error
                MessageBox.Show("Error opening the file.")
            End Try
        End If
    End Sub

    ' the SaveDocument procedure saves the current document

    Sub SaveDocument()
        Dim outputFile As StreamWriter ' object variable

        Try
            ' create the file
            outputFile = File.CreateText(strFilename)

            ' Write the textbox to the file
            outputFile.Write(txtDocument.Text)

            ' close the file
            outputFile.Close()

            ' update the isChanged variable
            blnIsChanged = False
        Catch ex As Exception
            ' error message for file creation error
            MessageBox.Show("Error creating the file.")
        End Try
    End Sub

    Private Sub txtDocument_TextChanged(sender As Object, e As EventArgs) Handles txtDocument.TextChanged
        ' Indicate the text has changed
        blnIsChanged = True
    End Sub

    Private Sub mnuFileNew_Click(sender As Object, e As EventArgs) Handles mnuFileNew.Click
        ' has the current document changed
        If blnIsChanged = True Then
            ' confirm before clearing the document
            If MessageBox.Show("The current document is not saved. " &
                               "Are you sure?", "Confirm",
                               MessageBoxButtons.YesNo) =
                            System.Windows.Forms.DialogResult.Yes Then
                ClearDocument()
            End If
        Else
            ' Document has not changed, so clear it.
            ClearDocument()
        End If
    End Sub

    Private Sub mnuFileOpen_Click(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        ' Has the current document changed?
        If blnIsChanged = True Then
            ' confirm before clearing the document
            If MessageBox.Show("The current document is not saved." &
                               " Are you sure?",
                               "Confirm", MessageBoxButtons.YesNo) =
                            System.Windows.Forms.DialogResult.Yes Then
                ClearDocument()
                OpenDocument()
            End If
        Else
            ' Document has not changed, so replace it
            ClearDocument()
            OpenDocument()
        End If
    End Sub

    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click
        ' Does the current document have a filename?
        If strFilename = String.Empty Then
            ' The document has not been saved, so
            ' use Save As dialog box
            If sfdSaveFile.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                strFilename = sfdSaveFile.FileName
                SaveDocument()
            End If
        Else
            'Save the document with the current filename
            SaveDocument()
        End If
    End Sub

    Private Sub mnuFileSaveAs_Click(sender As Object, e As EventArgs) Handles mnuFileSaveAs.Click
        ' save the current document undere a new filename
        If sfdSaveFile.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            strFilename = sfdSaveFile.FileName
            SaveDocument()
        End If
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        ' close the form
        Close()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        ' display an about box
        MessageBox.Show("Simple Text Editor version 1.0" & ControlChars.CrLf & "Created by Joshua Robinson")
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' if the document has been modified, confirm
        ' before exiting
        If blnIsChanged = True Then
            If MessageBox.Show("The current document is not saved. " &
                               "Do you wish to discard your changes?",
                               "Confirm",
                               MessageBoxButtons.YesNo) =
                               System.Windows.Forms.DialogResult.Yes Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
        End If
    End Sub
End Class
