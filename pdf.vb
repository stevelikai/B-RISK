Option Strict Off
Option Explicit On
Imports iTextSharp.text.pdf
Imports iTextSharp.text
'Imports System.IO
' We need System.IO namespace to read and write to disc
Imports System.IO

Module Pdf


    Public Sub create_pdf()

        Try

            Dim filepath2 As String = RiskDataDirectory
            filepath2 = filepath2 & "temp_out.txt"
            Call view_output(filepath2)

            Dim FileContent As String = File.ReadAllText(filepath2)

            Dim filepath3 As String = basefile.Replace(".xml", "_" & itcounter & ".pdf")

            Dim fntTextFont As Font = FontFactory.GetFont("COURIER", 7, Font.NORMAL, BaseColor.BLACK)
            Dim fntHeaderFont As Font = FontFactory.GetFont("ARIAL", 8, Font.BOLD, BaseColor.BLACK)

            Dim pdfdoc As New Document()

            Dim pdfwriter As PdfWriter = PdfWriter.GetInstance(pdfdoc, New FileStream(filepath3, FileMode.Create))

            pdfdoc.Open()
            'pdfdoc.Add(New Paragraph("Hello World."))
            'pdfdoc.NewPage()
            pdfdoc.Add(New Paragraph("VERIFIED PDF CREATED BY B-RISK : " & Now, fntHeaderFont))
            pdfdoc.Add(New Paragraph(FileContent, fntTextFont))

            pdfdoc.Close()

            'add page numbering to the pdf
            Dim bytes As Byte() = File.ReadAllBytes(filepath3)
            Dim blackFont As Font = FontFactory.GetFont("Arial", 10, Font.NORMAL, BaseColor.BLACK)
            Using stream As New MemoryStream()
                Dim reader As New PdfReader(bytes)
                Using stamper As New PdfStamper(reader, stream)
                    Dim pages As Integer = reader.NumberOfPages
                    For i As Integer = 1 To pages
                        ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, New Phrase(i.ToString(), blackFont), 568.0F, 15.0F, 0)
                    Next
                End Using
                bytes = stream.ToArray()
            End Using
            File.WriteAllBytes(filepath3, bytes)


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in pdf.vb create_pdf")
        End Try

    End Sub



End Module
