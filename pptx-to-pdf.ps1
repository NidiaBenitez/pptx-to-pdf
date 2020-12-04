# Batch convert all .pptx files in a folder
# based on scripts found at https://gist.github.com/mp4096/1a2279ec7b3dfec659f58e378ddd9aee and http://tech.franzone.blog/2019/11/19/how-to-create-a-word-to-pdf-converter-script-in-powershell/

#set directory to the folder which contains the files you wish to convert
Set-Location -Path C:\Users\User\Documents\pdf_files

# Create a PowerPoint object
$pptx_app = New-Object -ComObject PowerPoint.Application
$pptx_app.Visible = $True

# Get all objects of type .pptx? in $documents_path
echo "Processing $($documents_path)"
Get-ChildItem -Path $documents_path -Filter *.pptx? | ForEach-Object {
	echo $_.FullName
 	$document = $pptx_app.Presentations.Open($_.FullName)
	
	# Create a name for the PDF documents. If you wish them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
	$pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"

	# Save as PDF. The number 32 is the value for ppSaveAsPDF on Power Point 2007 (check more formats at https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb251061(v=office.12)?redirectedfrom=MSDN)
	$document.SaveAs($pdf_filename, 32)

	# Close PowerPoint files
 	$document.Close()
	}

# Exit and release PowerPoint object
$pptx_app.Quit()
