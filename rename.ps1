
param (
  $base_dir = [string](Get-Location),
  $sheet_name = "excel_sheet.xlsx",
  $doc_name = "word_doc.docx"
)

#gets child directories
$sub_dir = $(Get-ChildItem "$base_dir");

foreach($sub in $sub_dir) {
  if ($sub -is [System.IO.DirectoryInfo]){ #makes sure is a directory - gives each parent directory
    #get full path to bid folder inside of parent directory
    $bid = $(Get-ChildItem $sub.FullName -Filter bid | % { $_.FullName })

    #get sheet_name specific xlsx file inside of bid
    $sheet = $(Get-ChildItem $bid -Filter $sheet_name)
    #rename sheet to append parent directory if sheet not null (found a match for exactly $sheet_name)
    if ($sheet -ne $null){
      Rename-Item -Path $sheet.FullName -NewName ($sheet.BaseName + $sub.tostring() + $sheet.Extension)
    }

    #get doc_name specific docx file inside of bid
    $doc = $(Get-ChildItem $bid -Filter $doc_name)
    if ($doc -ne $null){
      Rename-Item -Path $doc.FullName -NewName ($doc.BaseName + $sub.tostring() + $doc.Extension)
    }
  }
}
