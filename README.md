# Remove Audio from PowerPoint Slides

This VBA script helps you remove audio from specific slides in a PowerPoint presentation.

## How to use

1. Open your PowerPoint presentation.

2. Press `Alt + F11` to open the VBA editor.

3. Click `Insert > Module` to create a new module.

4. Copy the following code and paste it into the module:

    ```vba
    Sub RemoveAllAudio()

        Dim oSlide As Slide
        Dim oShape As Shape
        Dim slideNums As String
        Dim slideRange As Variant
        Dim slideIndex As Variant
        Dim slideNumber As Integer
        Dim i As Integer
        
        ' Request user input for the slide numbers
        slideNums = InputBox("Enter slide numbers to remove audio (separated by comma, use dash for range):", "Slide Numbers")
        slideRange = Split(slideNums, ",")
        
        ' Loop through each slide number or range provided
        For Each slideIndex In slideRange
            ' Check if the slideIndex is a range
            If InStr(slideIndex, "-") > 0 Then
                ' Split the range into start and end
                Dim rangeParts As Variant
                rangeParts = Split(slideIndex, "-")
                
                ' Loop through each slide in the range
                For slideNumber = CInt(Trim(rangeParts(0))) To CInt(Trim(rangeParts(1)))
                    ' Call the function to remove audio from the slide
                    RemoveAudioFromSlide slideNumber
                Next slideNumber
            Else
                ' Remove audio from the individual slide
                RemoveAudioFromSlide CInt(Trim(slideIndex))
            End If
        Next slideIndex
        
        MsgBox "Audio has been removed from the specified slides.", vbInformation

    End Sub

    Sub RemoveAudioFromSlide(slideNumber As Integer)

        Dim oSlide As Slide
        Dim oShape As Shape
        Dim i As Integer
        
        ' Check if the slide number is valid
        If slideNumber > 0 And slideNumber <= ActivePresentation.Slides.Count Then
            Set oSlide = ActivePresentation.Slides(slideNumber)
            
            i = oSlide.Shapes.Count
            
            ' Loop through each shape on the slide
            For i = i To 1 Step -1
                Set oShape = oSlide.Shapes(i)
                
                ' Check if the shape is an audio file
                If oShape.Type = msoMedia Then
                    If oShape.MediaType = ppMediaTypeSound Then
                        ' Delete the audio file
                        oShape.Delete
                    End If
                End If
            Next i
        End If

    End Sub
    ```

5. Press `Ctrl + S` to save, and then close the VBA editor.

6. Run the script in PowerPoint by pressing `Alt + F8`, select `RemoveAllAudio`, and click `Run`.

7. Enter the slide numbers you wish to remove audio from when prompted.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)
