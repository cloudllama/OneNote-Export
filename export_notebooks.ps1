<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER
- See https://norspire.atlassian.net/wiki/spaces/One/pages/101646374/Exporting+OneNote+pages#Parameters

.EXAMPLE

.INPUTS

.OUTPUTS

.NOTES
- Export_Notebooks does not validate the e_structure_ of the notes. It will faithfully reproduce the structure of the note into Markdown, such as starting with a second-level bullet without having a first-level bullet. This ig the garbage-in, garbage-out maxim.
- Export_Notebooks does not understand all HTML tags. It supports bold, italics, bold+italics, strikethrough, and anchor (links). All other tags are ignored.
- Export_Notebooks does not make use of all OneNote tags in the XML of a page nor their attributes. There is a loss of metadata when converting from OneNote pages into markdown.

.LINK
#>

# Validate parameters for the script itself
[CmdletBinding(DefaultParameterSetName = 'Set1')]
param(
    [Parameter(Mandatory=$False, ParameterSetName = 'Set1')]
    [switch] $NoExport,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set2')]
    [string] $ExportSelected,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set3')]
    [switch] $ExportAll,

    [Parameter(Mandatory=$False)]
    [string] $NotebookDir,

    # The following three options are mutually exclusive. $NoPrintPage overrides $PrintSnippet overrides $PrintPage. $PrintSnippet is the default.
    [Parameter(Mandatory=$False)]
    [switch] $NoPrintPage,
    [Parameter(Mandatory=$False)]
    [switch] $PrintSnippet,
    [Parameter(Mandatory=$False)]
    [switch] $PrintPage,

    # The following three options are mutually exclusive. $Markdown overrides $PlainText overrides $HTML. $Markdown is the default.
    [Parameter(Mandatory=$False)]
    [switch] $Markdown,
    [Parameter(Mandatory=$False)]
    [switch] $PlainText,
    [Parameter(Mandatory=$False)]
    [switch] $HTML,

    [Parameter(Mandatory=$False)]
    [switch] $PrintStructure,
    [Parameter(Mandatory=$False)]
    [switch] $PrintStyles,
    [Parameter(Mandatory=$False)]
    [switch] $PrintTags,
    [Parameter(Mandatory=$False)]
    [switch] $SuppressOneNoteLinks,
    [Parameter(Mandatory=$False)]
    [switch] $NoDirCreation,
    [Parameter(Mandatory=$False)]
    [string] $ExportDir,

    [Parameter(Mandatory=$False)]
    [switch] $v,
    [Parameter(Mandatory=$False)]
    [switch] $vv,
    [Parameter(Mandatory=$False)]
    [switch] $vvv,
    [Parameter(Mandatory=$False)]
    [switch] $vvvv

)

# Reference: Application interface (OneNote): https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote

# -----------------------------------------------------------------------------
# Constants

$ILLEGAL_CHARACTERS = "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars()))

# -----------------------------------------------------------------------------
# Reference values

$LogLevels = @("DEBUG", "INFO", "WARNING", "ERROR")
$LogLevel = ""

# -----------------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Level,
        [string]$Message
    )

    switch ($Level) {
        "DEBUG" {
            If ($Loglevel -eq "DEBUG") {
                Write-Debug "$Message" 
            }
        }
        "INFO" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "INFO")) {
                Write-Host "INFO: $Message" -ForegroundColor Green
            }
        }
        "WARNING" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "INFO") -or ($LogLevel -eq "WARNING")) {
                Write-Warning "$Message"
            }
        }
        "ERROR" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "INFO") -or ($LogLevel -eq "WARNING") -or ($LogLevel -eq "ERROR")) {
                Write-Error "$Message"
            }
        }
        default {
        }
    }
    
}

# -----------------------------------------------------------------------------

function Find-OneNoteOutline {
    param (
        $PageNode,     # Accept XmlElement or any XML node type
        [int]$Depth = 0
        )
        
        if (($PageNode.Name) -eq "one:Outline"){
            return $PageNode
        } else {
            # Recurse into child nodes (if any)
            foreach ($Child in $PageNode.ChildNodes) {
                $FoundNode = Find-OneNoteOutline -PageNode $Child -Depth ($Depth + 1)
                If ($FoundNode) {
                    return $FoundNode
                }
            }
        
        }
    }
    
# -----------------------------------------------------------------------------

function Get-Tags {
    param (
        $PageNode     # We're expecting the DocumentElement.
    )

    # Known page styles
    #   - To Do
    #   - Important
    #   - Question
    #   - Idea
    #   - Remember for later

    $FoundTags = @{}

    # We depend on the tags being identified by increasing index number, starting at zero. This lets
    # us add them sequentially to the array.
    # It's valid to have NO tags defined for a page.
    # Tags can also have a highlight color that applies to the entire paragraph. We ignore it.

    ForEach ($Child in $PageNode.ChildNodes){
        If ($Child.LocalName -eq "TagDef" ){
            If ($Child.GetAttribute("index")) {
                If ($Child.GetAttribute("name")){
                    # Write-Output "  + Page style name: $($Child.GetAttribute("name"))"
                    $FoundTags[$Child.GetAttribute("index")] = $Child.GetAttribute("name")
                }
            }
        }
    }

    return $FoundTags
}
# -----------------------------------------------------------------------------
function Get-PageStyles {
    param (
        $PageNode     # We're expecting the DocumentElement.
    )

    # Known page styles
    #   - PageTitle
    #   - h1, h2, h3, h4, h5, h6 (headings)
    #   - p (normal)
    #   - code
    #   - blockquote (quote)
    #   - cite (citation)

    $FoundStyles = @{}

    # We depend on the styles being identified by increasing index number, starting at zero. This
    # lets us add them sequentially to the array.

    ForEach ($Child in $PageNode.ChildNodes){
        If ($Child.LocalName -eq "QuickStyleDef" ){
            If ($Child.GetAttribute("index")) {
                If ($Child.GetAttribute("name")){
                    # Write-Output "  + Page style name: $($Child.GetAttribute("name"))"
                    $FoundStyles[$Child.GetAttribute("index")] = $Child.GetAttribute("name")
                }
            }
        }
    }

    return $FoundStyles
}

# -----------------------------------------------------------------------------

function FormatHTMLTo-Markdown {
    param(
        [string] $Text,
        [bool] $RemoveNewlines = $True
    )
    [string] $ReturnText = ""
    
    If ($RemoveNewlines){
        $Text = $Text -replace "(`r`n|`r|`n)", ""
    }

    $OpeningTagRegex = "^([^<]*?)(<[^/][^>]+?>)(.*)$"
    If ($Text -match $OpeningTagRegex){
        do {
            $OpeningText = $matches[1]
            $OpeningTag = $matches[2]
            $RemainingOpeningText = $matches[3]
            

            [string] $ConvertedText = ""
            $ConvertedText = FormatHTMLTo-Markdown -Text $RemainingOpeningText -RemoveNewlines $False
            
            $ClosingTagRegex = "^([^<]*?)(</[^>]+?>)(.*)$"
            If ($ConvertedText -match $ClosingTagRegex) {
                $BeforeClosingText = $matches[1]
                # $ClosingTag = $matches[2]
                $RemainingClosingText = $matches[3]
            }
            
            # Bold tags: <span style='font-weight:bold'> ... </span>
            # Italic tags: <span style='font-style:italic'> ... </span>
            # Links: <a href="URL">URL-name</a>
            # Strikethrough: <span style='text-decoration:line-through'> ... </span>
            #   - Strikethrough isn't universally supported. It's a markdown extension.

            $BoldRegex = "font-weight:[\s]*?bold"
            $ItalicRegex = "font-style:[\s]*?italic"
            $LinkRegex = "(<a[\s]*?href="")(.*)("">)"
            $StrikethroughRegex = "style='text-decoration:line-through'"

            If (($OpeningTag -match $BoldRegex) -and ($OpeningTag -match $ItalicRegex)) {
                $ReturnText += $OpeningText + "___" + $BeforeClosingText + "___"

            } ElseIf ($OpeningTag -match $BoldRegex) {
                $ReturnText += $OpeningText + "__" + $BeforeClosingText + "__"

            } ElseIf ($OpeningTag -match $ItalicRegex) {
                $ReturnText += $OpeningText + "_" + $BeforeClosingText + "_"

            } ElseIf ($OpeningTag -match $LinkRegex) {
                $ReturnText += $OpeningText + "[" + $BeforeClosingText + "]" + "(" + $matches[2] + ")"

            } ElseIf ($OpeningTag -match $StrikethroughRegex){
                $ReturnText += $OpeningText + "~~" + $BeforeClosingText + "~~"

            } Else {
                # We don't care about any other tags, so remove them
                $ReturnText += $OpeningText + $BeforeClosingText
            }

            $Text = $RemainingClosingText
        } while ($Text -match $OpeningTagRegex)

        $ReturnText += $Text

    } Else {
        $ReturnText = $Text
    }

    return $ReturnText
}

# -----------------------------------------------------------------------------

function Get-Email {
    param(
        [string]$emailFilePath,
        [string]$spacingLeader = ""
    )
    [string]$EmailMessage = ""

    # Create an Outlook application instance
    $OutlookApp = New-Object -ComObject Outlook.Application

    # Open the email file
    $email = $OutlookApp.CreateItemFromTemplate($emailFilePath)

    # Extract email details
    $subject = $email.Subject
    $body = $email.Body
    $sender = $email.SenderName

    # Output the email details
    $EmailMessage += $spacingLeader + "--BEGIN EMAIL MESSAGE-------------------`n"
    $EmailMessage += $spacingLeader + "Subject: $subject`n"
    $EmailMessage += $spacingLeader + "Sender: $sender`n"
    $EmailMessage += $spacingLeader + "Body: $body`n"

    # Check for attachments
    if ($email.Attachments.Count -gt 0) {
        foreach ($attachment in $email.Attachments) {
            $EmailMessage += $spacingLeader + "  - Attachment: `"$($attachment.FileName)`"`n"
        }
    }
    $EmailMessage += $spacingLeader + "--END EMAIL MESSAGE---------------------`n"

    # Clean up
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($email) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutlookApp) | Out-Null

    return $EmailMessage
}

# -----------------------------------------------------------------------------
# Count-Tags is a recursive function that will descende the XML tree and count every tag without
# processing it.

function Count-Tags {
    param (
        $PageNode,     # Accept any XML node type
        $TagCount = 0
    )

    $TagCount += 1

    ForEach ($Child in $PageNode.ChildNodes){
        $TagCount = Count-Tags -PageNode $Child -TagCount $TagCount
    }

    return $TagCount
}

# -----------------------------------------------------------------------------

# Convert-Page is a recursive function that will descend the XML tree of a OneNote page and convert 
# it to markdown.
function Convert-Page {
    param (
        $PageName,
        $PageID,           # Object ID of the page
        $PageNode,         # Accept any XML node type
        $PageStyles,       # Hashtable of instantiated page styles {index, style_name}
        $Tags,             # Hashtable of instantiated tags {index, tag_name}
        $LastObjectID,     # Object ID of the parent
        $LastStyleName,    # StyleName of parent
        $IndentLevel = -1, # We start at -1 because the top tag is already an OEChildren
        $BulletLevel = 0,
        $AllNodeCount = 0, # The total number of nodes in the XML document
        $SumNodeCount = 0, # The sum of all nodes traversed
        $ProgressCounterDelay = 0
        )
        [String]$Paragraph = ""
        [int]$Bullet = $BulletLevel
        $ObjectID = $LastObjectID
        
    If ($AllNodeCount -eq 0){
        $AllNodeCount = Count-Tags $PageNode
        Write-Log "DEBUG" "AllNodeCount: $AllNodeCount"
    }

    # If the ProgressCounterDelay is zero, then set it to 5% of the total number of nodes. 5% is a 
    # compromise between updating the progress bar too frequently and not frequently enough.
    If ($ProgressCounterDelay -eq 0){
        $ProgressCounterDelay = [int]($AllNodeCount * 0.05)
    }
        
    $SumNodeCount += 1
    Write-Log "DEBUG" "Node: $($PageNode.Name), SumNodeCount: $SumNodeCount"

    If (($Loglevel -eq "INFO") -or ($LogLevel -eq "DEBUG")){
        Write-Log "DEBUG" "`$SumNodeCount: $SumNodeCount, `$AllNodeCount: $AllNodeCount, `$ProgressCounterDelay: $ProgressCounterDelay"
        $ProgressCounterDelay += 1
        If ($ProgressCounterDelay % 25 -eq 0){
            Write-Progress -Activity "$($PageName)" -Status "Converting page" -PercentComplete (($SumNodeCount / $AllNodeCount) * 100)
        }
    }

    $StyleName = $LastStyleName
    If ($PageNode.quickStyleIndex){
        $StyleName = $PageStyles[$PageNode.quickStyleIndex]
    } 

    If ($PageNode.Name -eq "one:OEChildren"){
        $IndentLevel += 1

    } ElseIf ($PageNode.Name -eq "one:OE"){
        $ObjectID = $PageNode.objectID

    } ElseIf ($PageNode.Name -eq "one:Image"){

        # Only include the alt text of the image if it is available.
        If ($PageNode.alt){
            $Paragraph = "{Image: ""$($PageNode.alt)""}`n"
        }
        
    } ElseIf ($PageNode.Name -eq "one:InsertedFile" ) {

        # Only include the file name and the original location of the file if they are available.
        If ($PageNode.preferredName){
            $Paragraph = "{File: ""$($PageNode.preferredName)"""
            If ($PageNode.pathSource) {
                $Paragraph += ", originally located at ""$($PageNode.pathSource)"""
            }
            $Paragraph += "}`n"

            If ( [System.Io.Path]::GetExtension($PageNode.preferredName) -eq ".msg" ){
                # Email message
                If (Test-Path $PageNode.pathCache) {
                    $EmailMessage = Get-Email -emailFilePath $PageNode.pathCache -spacingLeader "    "
                    If ($EmailMessage) {
                        $Paragraph += $EmailMessage
                    }
                }
            }
        }

    } ElseIf ($PageNode.Name -eq "one:Bullet"){
        $Bullet = $PageNode.GetAttribute("bullet")

    } ElseIf ($PageNode -is [System.Xml.XmlCDataSection]){
        # Only actual text that is typed into OneNote appears in ![CDATA] sections.

        If ($PageNode.Value.trim() -ne "") {
            If ($ToMarkdown.IsPresent){
                # We'll replace bold, italics, strikethrough, and links.
                # All other HTML tags will be removed. This includes but isn't limited to:
                #   - Font name 
                #   - Font size
                #   - 

                $PageText = FormatHTMLTo-Markdown -Text $PageNode.Value.trim()

            } ElseIf ($WithHTML.IsPresent) {
                $PageText = $PageNode.Value.trim()

            } Else { 
                # Default to $PlainText=$True
                $PageText = $PageNode.Value.trim() -replace "<[^>]+>", ""
            }

            $Leader = ""
            Switch ($StyleName){
                "PageTitle"{
                    $Leader += ""
                }
                "h1" {
                    $Leader += "# "
                }
                "h2" {
                    $Leader += "## "
                }
                "h3" {
                    $Leader += "### "
                }
                "h4" {
                    $Leader += "#### "
                }
                "p" {
                    $Leader += ""
                }
                "blockquote" {
                    $Leader += "> "
                }
                "code" {
                    $Leader += "``` "
                }
                "cite" {
                    $Leader += ""
                }
                default {
                    $Leader += ""
                }
            }
            $Paragraph = "$($Leader)"
            If ($Bullet -gt 0){
                $Paragraph += " " * ((($Bullet) -1)* 2) + "+ "
                $Bullet = 0
            } Else {
                $Paragraph += " " * (($IndentLevel) * 2)
            }
            $Paragraph += "$($PageText) `n"

            If (-not $SuppressOneNoteLinks.IsPresent){
                If ($StyleName -eq "h1"){
                    $HyperlinkToObject = ""
                    $OneNoteApp.GetHyperLinkToObject( $PageID, $ObjectID, [ref]$HyperlinkToObject)
                    $PlainPageText = $PageNode.Value.trim() -replace "<[^>]+>", ""
                    $Paragraph += "[$($PlainPageText)]($($HyperlinkToObject))`n"
                    $ObjectID = ""
                }
            }
        }
    }

    # Recurse into child nodes (if any)
    foreach ($Child in $PageNode.ChildNodes) {
        $ConvertResult = Convert-Page -PageName $PageName -PageID $PageID -PageNode $Child -PageStyles $PageStyles $LastObjectID $ObjectID -LastStyleName $StyleName -IndentLevel $IndentLevel -BulletLevel $Bullet -AllNodeCount $AllNodeCount -SumNodeCount $SumNodeCount -ProgressCounterDelay $ProgressCounterDelay 
        $Paragraph += $ConvertResult.Paragraph
        $Bullet = $ConvertResult.Bullet
        $SumNodeCount = $ConvertResult.SumNodeCount
        $ProgressCounterDelay = $ConvertResult.ProgressCounterDelay
    }

    return [PSCustomObject]@{
        Paragraph = $Paragraph
        Bullet = $Bullet
        SumNodeCount = $SumNodeCount
        ProgressCounterDelay = $ProgressCounterDelay
    }
}

# -----------------------------------------------------------------------------
function Split-Pages {
    param (
        $PageMarkdown,
        $PageName,
        $AllH1Count = 0,
        $SumH1Count = 0,
        $ProgressCounterDelay = 0
    )

    $PageParagraphs = @{}

    If ($AllH1Count -eq 0){
        $AllH1Count = ($PageMarkdown -split "`r?`n" | Select-String "^# ").Count
        Write-Log "DEBUG" "AllH1Count: $AllH1Count"
    }

    # If the ProgressCounterDelay is zero, then set it to 5% of the total number of nodes. 5% is a 
    # compromise between updating the progress bar too frequently and not frequently enough.
    If ($ProgressCounterDelay -eq 0){
        $ProgressCounterDelay = [int]($AllNodeCount * 0.05)
    }
    
    $Lines = $PageMarkdown -split "`r?`n"
    $LastParagraphTitle = ""
    $LastParagraph = ""
    
    ForEach($Line in $Lines){
        If ($($Line) -match "^# ") {
            
            $SumH1Count += 1
            If (($Loglevel -eq "INFO") -or ($LogLevel -eq "DEBUG")){
                Write-Log "DEBUG" "`$SumH1Count: $SumH1Count, `$AllH1Count: $AllH1Count, `$ProgressCounterDelay: $ProgressCounterDelay"
                $ProgressCounterDelay += 1
                If ($ProgressCounterDelay % 2 -eq 0){
                    Write-Progress -Activity "$($PageName)" -Status "Splitting pages" -PercentComplete (($SumH1Count / $AllH1Count) * 100)
                }
            }

            If ($PageParagraphs.ContainsKey($LastParagraphTitle)){
                $PageParagraphs[$LastParagraphTitle] += "$($LastParagraph)"
                
            } Else {
                If ($LastParagraphTitle -eq ""){
                    If ($LastParagraph){
                        $LastParagraphTitle = $PageName
                        $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                    }
                } Else {
                    $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                }
            }
            
           $LastParagraph=""

            $TitleLine = $Line -match "^(# )([\s]*?[\*+-]*[\s]*)(.*)$"
            If ($matches[3]){
                $LastParagraphTitle = $Matches[3]
                
                # + Remove the leading day name, if any, from the expected
                #   format of "dddd, MMMM dd, yyyy". Occasionally, the wrong
                #   day is attached to the right date. 
                $CleanedDate = $LastParagraphTitle -replace "^[^,]+,\s*", ""
                
                # + If we can recognize the heading as a valid date in the form
                #   of "MMMM dd, yyyy", then reformat it to "yyyy-MM-dd", which
                #   is friendlier to a file system.
                # + Also add the full date ("dddd, MMMM dd, yyyy") to the cache
                #   so that it is available as text to the LLM.
                try{
                    If ((Get-Date $CleanedDate).ToString("yyyy-MM-dd")){
                        $LastParagraphTitle = (Get-Date $CleanedDate).ToString("yyyy-MM-dd")
                    } 
                }
                catch {}
            } Else {
                # + We know that the line is an h1 heading (starts with "^# ")
                #   but it doesn't match the regular expression above. So, just
                #   set the line value to whatever comes after the "^# ".
                # + Note that the "Untitled" section can only occur once, at the
                #   beginning of the page. After any h1 heading has been 
                #   encountered, a section will *always* have a name.
                $LastParagraphTitle = "$($Matches[2])$($Matches[3])"
            }

            If (!($PageParagraphs.ContainsKey($LastParagraphTitle))) {
                try {
                    $LastParagraph = "# " + (Get-Date $LastParagraphTitle).ToString("dddd, MMMM dd, yyyy") +"`n"
                } 
                catch {
                    $LastParagraph = "# $($LastParagraphTitle)" + "`n"
                }
            }

        } Else {
            # + The line is not an h1 heading. Just add it to the cache.
            $LastParagraph += "$($Line)`n"
        }
    }

    # Issue-3
    # + We only add lines when we encounter an h1 heading. However, if an h1 
    #   heading is *not* the last line, we'll leave lines in the $LastParagraph
    #   buffer without adding them. So, look in $LastParagraph and see if any
    #   lines need to be added.
    If ($LastParagraph) {
        If ($PageParagraphs.ContainsKey($LastParagraphTitle)){
            $PageParagraphs[$LastParagraphTitle] += "$($LastParagraph)"
        } Else {
            If ($LastParagraphTitle -eq ""){
                If ($LastParagraph){
                    $LastParagraphTitle = $PageName
                    $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                }
            } Else {
                $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
            }        }
    }

    return $PageParagraphs
}


# ==================================================================================================
# Main 

# If there is any logging level, then allow debug printing.
If ($v.IsPresent -or $vv.IsPresent -or $vvv.IsPresent -or $vvvv.IsPresent){
    $DebugPreference = "Continue"
}

If (($v.IsPresent) -or ($vv.IsPresent)){
    $WarningPreference = "Continue"
}

# Set logging level. In case multiple levels are specified, the highest level is used.
If ( $vvvv.IsPresent){
    $LogLevel = "DEBUG"
    Write-Log "INFO" "Log level set to DEBUG"
} ElseIf ( $vvv.IsPresent) {
    $LogLevel = "INFO"
    Write-Log "INFO" "Log level set to INFO"
} ElseIf ( $vv.IsPresent) {
    # If LogLevel is set to WARNING, we can't print an INFO message that it is set to WARNING.
    $LogLevel = "WARNING"
} ElseIf ( $v.IsPresent) {
    # If LogLevel is set to ERROR, we can't print an INFO message that it is set to ERROR.
    $LogLevel = "ERROR"
} Else {
    # Default to no logging.
    $LogLevel = "NONE"
}

Write-Log -Level "DEBUG" -Message "NoExport: $NoExport"
Write-Log -Level "DEBUG" -Message "ExportSelected: $ExportSelected"
Write-Log -Level "DEBUG" -Message "ExportAll: $ExportAll"
Write-Log -Level "DEBUG" -Message "NotebookDir: $NotebookDir"
Write-Log -Level "DEBUG" -Message "NoPrintPage: $NoPrintPage"
Write-Log -Level "DEBUG" -Message "PrintSnippet: $PrintSnippet"
Write-Log -Level "DEBUG" -Message "PrintPage: $PrintPage"
Write-Log -Level "DEBUG" -Message "Markdown: $Markdown"
Write-Log -Level "DEBUG" -Message "PlainText: $PlainText"
Write-Log -Level "DEBUG" -Message "HTML: $HTML"
Write-Log -Level "DEBUG" -Message "PrintStructure: $PrintStructure"
Write-Log -Level "DEBUG" -Message "PrintStyles: $PrintStyles"
Write-Log -Level "DEBUG" -Message "PrintTags: $PrintTags"
Write-Log -Level "DEBUG" -Message "SuppressOneNoteLinks: $SuppressOneNoteLinks"
Write-Log -Level "DEBUG" -Message "NoDirCreation: $NoDirCreation"
Write-Log -Level "DEBUG" -Message "ExportDir: $ExportDir"
Write-Log -Level "DEBUG" -Message "v: $v"
Write-Log -Level "DEBUG" -Message "vv: $vv"
Write-Log -Level "DEBUG" -Message "vvv: $vvv"
Write-Log -Level "DEBUG" -Message "vvvv: $vvvv"

# Advanced logic for parameters.
If ($PSCmdlet.ParameterSetName -eq "Set1") {
    $NoExport=$True
}

# If no print option is specified, then default to PrintSnippet.
If ((-not $NoPrintPage.IsPresent) -and (-not $PrintPage.IsPresent)){
    Write-Log("INFO", "No print option specified. Defaulting to PrintSnippet.")
    $PrintSnippet=$True
} Else {
    # If $PrintSnippet is specified, ignore the other print options.
    If ($PrintSnippet.IsPresent){
        $NoPrintPage=$false
        $PrintPage=$false
    } ElseIf ($NoPrintPage.IsPresent){
        # If $NoPrint is specified, ignore the $PrintPage option (whether it was specified or not).
        # Also ignore the $PrintSnippet option for good hygiene.
        $PrintSnippet=$false
        $PrintPage=$false
    } Else {
        # $PrintPage must be set.
        $PrintPage=$True
        $PrintSnippet=$false
        $NoPrintPage=$false
    }
}
Write-Log "DEBUG" "After parameter logic: NoPrintPage: $NoPrintPage"
Write-Log "DEBUG" "After parameter logic: PrintSnippet: $PrintSnippet"
Write-Log "DEBUG" "After parameter logic: PrintPage: $PrintPage"

If ((-not $HTML) -and (-not $PlainText)){
    Write-Log("INFO", "No output format specified. Defaulting to Markdown.")
    $Markdown=$True
} Else {
    # If $Markdown is specified, ignore the other output options.
    If ($Markdown){
        $HTML=$false
        $PlainText=$false
    } ElseIf ($PlainText){
        $Markdown=$false
        $HTML=$false
    } Else {
        # $HTML must be set.
        $HTML=$True
        $PlainText=$false
    }
}
Write-Log "DEBUG" "After parameter logic: Markdown: $Markdown"
Write-Log "DEBUG" "After parameter logic: PlainText: $PlainText"
Write-Log "DEBUG" "After parameter logic: HTML: $HTML"

# NoDirCreation overrides the ExportDir parameter, if ExportDir is specified.
If (!$NoDirCreation){
    If ($ExportDir){
        If (!(Test-Path -Path $ExportDir -PathType Container)) {
            Write-Log "ERROR" "The specified export directory does not exist."
            Exit 1
        }
    } Else {
        Write-Log("INFO", "No export directory specified. Defaulting to the current directory.")
        $ExportDir = "."
    }
}


# Ensure that the notebook directory is valid. A null value is acceptable.
If ($NotebookDir){
    If (!(Test-path $NotebookDir)){
        Write-Log "ERROR" "The specified notebook directory does not exist."
        Exit 1
    }
}

# Start the OneNote application COM object.

    # The following code can be used to manually load the assembly, but it *should* work without it.
    # $AssemblyFile = (get-childitem $env:windir\assembly -Recurse Microsoft.Office.Interop.OneNote.dll | Sort-Object Directory -Descending | Select-Object -first 1).FullName
    # Add-Type -Path $AssemblyFile -IgnoreWarnings

$OneNoteApp = New-Object -ComObject OneNote.Application

# Ask OneNote for the hiearchy all the way down to the individual pages, which is as low as you can
# go.
# The 2013 schema is the most recent schema available.
# If the notebook directory is not specified, then the default notebook is used.
Write-Log "DEBUG" "Getting the hierarchy of notebooks and pages."
[xml]$NotebooksXML = ""
$Scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages
$OneNoteVersion = [Microsoft.Office.Interop.OneNote.XMLSchema]::xs2013
try {
    $OneNoteApp.GetHierarchy($NotebookDir, $Scope, [ref]$NotebooksXML, $OneNoteVersion)
}
catch {
    Write-Log "ERROR" "An error occurred while getting the hierarchy of notebooks and pages."
    Exit 1
}
Write-Log "DEBUG" "Got the hierarchy of notebooks and pages."

ForEach($Notebook in $NotebooksXML.Notebooks.Notebook)
{
    If ($PrintStructure.IsPresent) {
        Write-Output "Notebook Name: ""$($Notebook.Name.trim())"""
    }

    $CleansedNotebookName = $Notebook.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
    If ($CleansedNotebookName -ne $Notebook.Name.trim()){
        Write-Log "INFO" "The notebook name contains illegal characters. It has been cleansed to: `"$($CleansedNotebookName)`""
    }
    $NotebookPath = Join-Path -Path $ExportDir -ChildPath "$($CleansedNotebookName) notebook"
    If ((!$NoExport) -and (!$NoDirCreation)) {
        If (!(Test-Path -Path $NotebookPath -PathType Container)) {
            Write-Log "INFO" "Creating the notebook directory: `"$($CleansedNotebookName) notebook`""
            New-Item -Path $NotebookPath -ItemType Directory | Out-Null
        }
    }

    ForEach($Section in $Notebook.Section)
    {
        If ($PrintStructure.IsPresent){
            Write-Output "- Section Name: ""$($Section.Name.trim())"""
        }

        $CleansedSectionName = $Section.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
        If ($CleansedSectionName -ne $Section.Name.trim()){
            Write-Log "INFO" "The section name contains illegal characters. It has been cleansed to: `"$($CleansedSectionName)`""
        }
        $SectionPath = Join-Path -Path $NotebookPath -ChildPath "$($CleansedSectionName) section"
        If ((!$NoExport) -and (!$NoDirCreation)) {
            If ( !(Test-Path -Path $SectionPath -PathType Container)) {
                Write-Log "INFO" "Creating the section directory: `"$($CleansedSectionName) section`""
                New-Item -Path $SectionPath -ItemType Directory | Out-Null
            }
        }

        ForEach($Page in $Section.Page) 
        {
            If ($PrintStructure.IsPresent) {
                Write-Output "  - Page Name: ""$($Page.Name)"""
            }

            $CleansedPageName = $Page.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
            If ($CleansedPageName -ne $Page.Name.trim()){
                Write-Log "INFO" "The page name contains illegal characters. It has been cleansed to: `"$($CleansedPageName)`""
            }
            $PagePath = Join-Path -Path $SectionPath -ChildPath "$($CleansedPageName) page"
            If ((!$NoExport) -and (!$NoDirCreation)) {
                If (!(Test-Path -Path $PagePath -PathType Container)){
                    Write-Log "INFO" "Creating the page directory: `"$($CleansedPageName) page`""
                    New-Item -Path $PagePath -ItemType Directory | Out-Null
                }
            }

            # This operation can potentially take a long time because it's fetching the entire 
            # contents of the page.
            Write-Log("DEBUG", "Getting the content of the page: `"$($Page.Name)`"")
            [xml]$PageXML = ""
            $OneNoteApp.GetPageContent($Page.ID, [ref]$PageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic, $OneNoteVersion)
            Write-Log "DEBUG" "Got the content of the page: `"$($Page.Name)`""
            
            $DelimiterPrinted = $False

            # It is possible that a page is empty, but even empty pages still have a defined style
            # for the page title. This is a good check for the validity of the page.

            # The styles can be uniquely defined for each page. Styles are
            # defined in the order that they are first used on the page. So,
            # h2 migh tbe style #2 or #7, depending on when it was first
            # used on that page.
            Write-Log("DEBUG", "Finding the styles of the page: `"$($Page.Name)`"")
            $PageStyles = @{}
            $PageStyles = Get-PageStyles $PageXML.DocumentElement
            Write-Log "DEBUG" "Found the styles of the page: `"$($Page.Name)`""

            If ( $PageStyles.Keys.Count -eq 0 ){
                # All pages must have at least the PageTitle style.
                Write-Log "ERROR" "Could not find any styles for the page $($Page.Name)"
                Exit 1
            }
            
            If ($PrintStyles){
                Write-Output " "
                Write-Output "--------"
                $DelimiterPrinted=$True
                Write-Output "  + Page Styles"
                ForEach( $Key in $PageStyles.Keys) {
                    Write-Output "    - $($Key): $($PageStyles[$Key])"
                }
                Write-Output " "
            }    

            # Start at the beginning of the content, which is the first Outline node. There
            # can be multiple Outline nodes, but we're only interested in the first one.
            Write-Log "DEBUG" "Finding the first outline node of the page: `"$($Page.Name)`""
            [System.Xml.XmlElement]$OneNoteOutline = Find-OneNoteOutline $PageXML.DocumentElement 6
            
            If (!($OneNoteOutline)){
                Write-Log "INFO" "Could not find the first outline node of the page: `"$($Page.Name)`". The page is considered empty and will be ignored."
            } Else {
                Write-Log "DEBUG" "Found the first outline node of the page: `"$($Page.Name)`""
                
                # Tags are optional. There might not be any on the page.
                # It's possible that a page is empty and it's not an error to check for tags, but an
                # empty page will never have tags so we don't check unless the page has some
                # content. This is a performance optimization.
                Write-Log "DEBUG" "Finding the tags of the page: `"$($Page.Name)`""
                $Tags = @{}
                $Tags = Get-Tags $PageXML.DocumentElement
                Write-Log "DEBUG" "Found the tags of the page: `"$($Page.Name)`""
                
                If ($PrintTags){
                    If ($Tags.Keys.Count -gt 0){
                        If (!($DelimiterPrinted)){
                            Write-Output " "
                            Write-Output "--------"
                            $DelimiterPrinted=$True
                        }
                        Write-Output "  + Tags"
                        ForEach( $Key in $Tags.Keys) {
                            Write-Output "    - $($Key): $($Tags[$Key])"
                        }
                        Write-Output " "
                    }
                }     
     
                
                If (($PrintSnippet) -or ($PrintPage)){
                    If (!($DelimiterPrinted)){
                        Write-Output " "
                        Write-Output "--------"
                        $DelimiterPrinted=$True
                    }
                }
                    
                # Convert the page to markdown
                Write-Log "DEBUG" "Converting the page from XML: `"$($Page.Name)`""
                $ConvertResult = Convert-Page $Page.Name $Page.id $OneNoteOutline $PageStyles ""
                Write-Log "DEBUG" "Converted the page from XML: `"$($Page.Name)`""

                # Split the page into individual paragraphs and then
                # write them to individual files.
                Write-Log "DEBUG" "Splitting the page into paragraphs: `"$($Page.Name)`""
                $PageParagraphs = Split-Pages -PageMarkdown $ConvertResult.Paragraph -PageName $Page.Name
                Write-Log "DEBUG" "Split the page into paragraphs: `"$($Page.Name)`""

                ForEach ($PageParagraph in $PageParagraphs.Keys){
                    If ($PrintStructure.IsPresent) {
                        Write-Output "    * Paragraph Name: ""$($PageParagraph)"""
                    }

                    If ($PrintPage){
                        If (!($DelimiterPrinted)){
                            Write-Output " "
                            Write-Output "--------"
                            $DelimiterPrinted=$True
                        }
                        Write-Output("$($PageParagraphs[$PageParagraph])")
                    } ElseIf ($PrintSnippet){
                        If (!($DelimiterPrinted)){
                            Write-Output " "
                            Write-Output "--------"
                            $DelimiterPrinted=$True
                        }

                        #Write-Output($($PageParagraphs[$PageParagraph]) -split "`n" | Select-Object -First 3)
                        $Snippet = $($PageParagraphs[$PageParagraph]).Substring(0, [Math]::Min($($PageParagraphs[$PageParagraph]).Length, 100))
                        If ($Snippet.Length -eq 100){
                            $Snippet += "..."
                        }
                        Write-Output($Snippet)
                    }

                    $CleansedPageParagraph = $PageParagraph -replace $ILLEGAL_CHARACTERS, "_"
                    If ($CleansedPageParagraph -ne $PageParagraph){
                        Write-Log "INFO" "The paragraph name `"$($PageParagraph)`" contains illegal characters. It has been changed to `"$($CleansedPageParagraph)`"."
                    }
                    
                    If (($ExportAll) -or (($ExportSelected) -and ($Page.Name -eq $ExportedSelected))){
                        # ONE-2
                        # Remove illegal characters from the paragraph name, which will then be the file name.
                        $PageParagraphFileName = Join-Path -Path $PagePath -ChildPath "$($CleansedPageParagraph.TrimEnd()).md"
                        $PageParagraphs[$PageParagraph].TrimEnd() | Out-File -FilePath $PageParagraphFileName -Encoding utf8
                    }
                }

                If ($DelimiterPrinted){
                    Write-Output "--------"
                    Write-Output " "
                }
            }
            
        }
    }
}

# Clean up after yourself.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNoteApp) | Out-Null