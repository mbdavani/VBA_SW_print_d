Attribute VB_Name = "Print_JX_ENG_HP4000_D1"
Sub JX_ENG_HP4000win10D()                                                                       'Print D-Size to HP4000win10, landscape, monochrome
Dim TroubleShoot As Boolean: TroubleShoot = False                                               'Set "Troubleshoot" to True if you want to trouble shoot
Dim TBLShootTitle As String: TBLShootTitle = "USNR: Trouble Shoot HP4000 D-Size"
Dim swApp As Object
Dim PrinterPath As String: PrinterPath = "\\jx-vfs\JX-ENG-HP4000win10"                          'Intended printer
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim FeatureData As Object
Dim Feature As Object
Dim Component As Object
Dim Part As SldWorks.ModelDoc2
Dim Size As String
Dim PartExt As SldWorks.ModelDocExtension
Dim PageSettings As PageSetup
Dim CurSheet As SldWorks.Sheet
Dim PartName As String
Dim vPath As Variant
Dim c1 As Long
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc
Set PartExt = Part.Extension
Set PageSettings = PartExt.AppPageSetup
Dim ExistingPrinter As String
Dim TBLShoot As String
ExistingPrinter = Part.Printer                                                                  'Printer assigned to active doc, prior to running this script

    If TroubleShoot = True Then                                                                 'Trouble Shooting Message.  This describes the existing print set up
        TBLShoot = MsgBox("Printer:   " & ExistingPrinter _
        & Chr(10) & "Size:            " & PageSettings.PrinterPaperSize _
        & Chr(10) & "Orientation:    " & PageSettings.Orientation _
        & Chr(10) & Chr(10) & "Scale to Fit:   " & PageSettings.ScaleToFit, vbInformation, TBLShootTitle)
    ElseIf TroubleShoot = False Then
    End If
    
PartExt.UsePageSetup = swPageSetupInUse_Application
Part.Printer = PrinterPath                                                                      'Assign the intended printer to active doc
PageSettings.Orientation = 2                                                                    'Where 1 = Portrait, 2 = Landscape
PageSettings.LeftFooter = ""                                                                    'Sets footer if we ever want to use that, ex.. "Last Plot Date: &[Date] &[Time]"
PageSettings.ScaleToFit = True                                                                  'Turns scale to fit on
PageSettings.PrinterPaperLength = 34
PageSettings.PrinterPaperWidth = 22
PageSettings.DrawingColor = 3                                                                   'Prints drawing in black & white

    If TroubleShoot = True Then                                                                 'Trouble Shooting Message.  This describes the final print set up
        TBLShoot = MsgBox("Printer:   " & Part.Printer _
        & Chr(10) & "Size:            " & PageSettings.PrinterPaperSize _
        & Chr(10) & "Orientation:    " & PageSettings.Orientation _
        & Chr(10) & Chr(10) & "Scale to Fit:   " & PageSettings.ScaleToFit, vbInformation, TBLShootTitle)
    ElseIf TroubleShoot = False Then
    End If
                                                                                                'Insert an EXIT SUB here if in trouble shoot mode

PartExt.PrintOut 1, 0, 1, False, PrinterPath, Empty                                              'Send printjob to specified printer (1st page, all pages, 1 copy, collate, intended printer, no PDF)

End Sub
