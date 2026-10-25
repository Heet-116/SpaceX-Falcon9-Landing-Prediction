$pptxPath = "C:\Users\Admin\College files\Labs\Coursera\Data Science Capstone Project Report.pptx"
$pdfPath = "C:\Users\Admin\College files\Labs\Coursera\Data Science Capstone Project Report.pdf"

Write-Host "Opening PowerPoint Application..."
$pptApp = New-Object -ComObject PowerPoint.Application

try {
    Write-Host "Opening presentation: $pptxPath"
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $presentation = $pptApp.Presentations.Open($pptxPath, 2, 0, 0)
    
    Write-Host "Exporting to PDF: $pdfPath"
    # 32 is the enum value for ppSaveAsPDF
    $presentation.SaveAs($pdfPath, 32)
    
    $presentation.Close()
    Write-Host "PDF successfully created!"
} catch {
    Write-Host "Error during PDF export: $_"
} finally {
    $pptApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
}
