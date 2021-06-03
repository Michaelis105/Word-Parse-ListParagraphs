param ($requirementXml, $outputFilePath)

function Write-To-CSV ([array]$requirements, [string]$outputFilePath) {
    Remove-Item -Path $outputFilePath
    Add-Content -Path $outputFilePath -Value 'Requirement|'
    $requirements | foreach {
        Add-Content -Path $outputFilePath -Value "$_|"
    }
}

function Combine-Text {
    param($pNode)
    $string = ''
    foreach ($r in $pNode.r) {
        $text = ''

        try {
            if ($r.t."#text" -ne $null) {
                $text = $($r.t."#text")
            } elseif ($r.t."#significant-whitespace" -ne $null) {
                continue
            } elseif ($r.t -ne $null) {
                $text = $($r.t)
            } 
        
            if (![String]::IsNullOrWhiteSpace($text)) {
                $text = $($text.Trim())
                $string = "$($string) $($text)"
            }
        
            #Write-Host ($r.t | Format-Table | Out-String)

        } catch {
            echo "Error text was: $text" 
            Write-Host ($r.t | Format-Table | Out-String)
            echo ""
        }
        
    }
    return $string 
}


function Get-Requirements {
    param($node)
    $requirements = @()

    foreach ($p in $node.p) {
        $style = $p.pPr.pStyle
        if ($style.val -eq 'ListParagraph') {
            $words = Combine-Text $p
            if (![String]::IsNullOrWhiteSpace($words)) {
                $requirements += $words
            }
        }
    }
    return $requirements
}

function Get-XML-Body-Node {
    param($filePath)
     [xml]$xml = Get-Content $filePath -Encoding UTF8
     foreach ($part in $xml.package.part) {
        if($part.name -eq '/word/document.xml') {
            return $part.xmlData.document.body
        }
     }
     throw "Could not find the body"
}


$xmlBodyNode = Get-XML-Body-Node $requirementXml
$requirements = Get-Requirements($xmlBodyNode)
Write-To-CSV -requirements $requirements -outputFilePath $outputFilePath