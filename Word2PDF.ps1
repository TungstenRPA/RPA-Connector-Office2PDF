# Convert Word Document to PDF
# This requires Microsoft Word to be installed on the Roboserver
# Make Sure Kofax RPA Roboserver Settings/Security/AllowCommandLineExecution is set to true
# This takes one parameter - the path to a Word Document.
# A pdf document is created in the same folder with the same filename.
# If the pdf already exists it will be replaced
$Word = New-Object -ComObject Word.Application
$Doc = $Word.Documents.Open($args[0])
$Name = ($Doc.FullName).replace('.docx','.pdf')
$Doc.SaveAs($Name,17)
$Doc.Close()