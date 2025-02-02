<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER

.EXAMPLE

.INPUTS

.OUTPUTS

.NOTES
- Export_Notebooks does not validate the _structure_ of the notes. It will faithfully reproduce the structure of the note into Markdown, such as starting with a second-level bullet without having a first-level bullet. This ig the garbage-in, garbage-out maxim.
- Export_Notebooks does not understand all HTML tags. It supports bold, italics, bold+italics, strikethrough, and anchor (links). All other tags are ignored.
- Export_Notebooks does not make use of all OneNote tags in the XML of a page nor their attributes. There is a loss of metadata when converting from OneNote pages into markdown.

.LINK
#>

# Validate parameters for the script itself
[CmdletBinding(DefaultParameterSetName = 'Set1')]
param(
    [Parameter(Mandatory=$False, ParameterSetName = 'Set1')]
    [switch] $NoPagePrint,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set2')]
    [switch] $PrintAllPages,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set3')]
    [string] $PageToPrint,

    # The following three options are mutually exclusive. $ToMarkdown overrides $WithHTML overrides $PlainText, which is the default
    [Parameter(Mandatory=$False)]
    [switch] $PlainText,
    [Parameter(Mandatory=$False)]
    [switch] $WithHTML,
    [Parameter(Mandatory=$False)]
    [switch] $ToMarkdown,

    [Parameter(Mandatory=$False)]
    [switch] $PrintStructure,
    [Parameter(Mandatory=$False)]
    [switch] $PrintStyles,
    [Parameter(Mandatory=$False)]
    [switch] $PrintTags,
    [Parameter(Mandatory=$False)]
    [switch] $SuppressOneNoteLinks,
    [Parameter(Mandatory=$False)]
    [string] $ExportDir,
    [Parameter(Mandatory=$False)]
    [switch] $DebugMessages
)

# Reference: Applicaiton interface (OneNote): https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote

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
    $EmailMessage += $spacingLeader + "Subject: $subject`n"
    $EmailMessage += $spacingLeader + "Sender: $sender`n"
    $EmailMessage += $spacingLeader + "Body: $body`n"

    # Check for attachments
    if ($email.Attachments.Count -gt 0) {
        $EmailMessage += "Attachments:`n"
        foreach ($attachment in $email.Attachments) {
            $EmailMessage += " - $($attachment.FileName)`n"
        }
    }

    # Clean up
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($email) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutlookApp) | Out-Null

    return $EmailMessage
}

# -----------------------------------------------------------------------------

# Convert-Page is a recursive function that will descend the XML tree of a OneNote page and convert 
# it to markdown.
function Convert-Page {
    param (
        $PageID,           # Object ID of the page
        $PageNode,         # Accept any XML node type
        $PageStyles,       # Hashtable of instantiated page styles {index, style_name}
        $Tags,             # Hashtable of instantiated tags {index, tag_name}
        $LastObjectID,     # Object ID of the parent
        $LastStyleName,    # StyleName of parent
        $IndentLevel = -1, # We start at -1 because the top tag is already an OEChildren
        $BulletLevel = 0
    )
    [String]$Paragraph = ""
    [int]$Bullet = $BulletLevel
    $ObjectID = $LastObjectID

    $StyleName = $LastStyleName
    If ($PageNode.quickStyleIndex){
        $StyleName = $PageStyles[$PageNode.quickStyleIndex]
    } 

    If ($PageNode.Name -eq "one:OEChildren"){
        $IndentLevel += 1

    } ElseIf ($PageNode.Name -eq "one:OE"){
        $ObjectID = $PageNode.objectID

    } ElseIf ($PageNode.Name -eq "one:Image"){

        $Paragraph = "{Image: ""$($PageNode.alt)""}`n"
        
    } ElseIf ($PageNode.Name -eq "one:InsertedFile" ) {

        $Paragraph = "{File: ""$($PageNode.preferredName)"""
        If ($PageNode.pathSource) {
            $Paragraph += ", originally located at ""$($PageNode.pathSource)"""
        }
        $Paragraph += "}`n"

        If ( [System.Io.Path]::GetExtension($PageNode.preferredName) -eq ".msg" ){
            # Email message
            If (Test-Path $PageNode.pathCache) {
                $EmailMessage = Get-Email -emailFilePath $PageNode.pathCache -spacingLeader "  "
                If ($EmailMessage) {
                    $Paragraph += $EmailMessage
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
        $ConvertResult = Convert-Page -PageID $PageID -PageNode $Child -PageStyles $PageStyles $LastObjectID $ObjectID -LastStyleName $StyleName -IndentLevel $IndentLevel -BulletLevel $Bullet
        $Paragraph += $ConvertResult.Paragraph
        $Bullet = $ConvertResult.Bullet
    }

    return [PSCustomObject]@{
        Paragraph = $Paragraph
        Bullet = $Bullet
    }
}

# -----------------------------------------------------------------------------
function Split-Pages {
    param (
        $PageMarkdown,
        $PageName
    )

    $PageParagraphs = @{}

    $Lines = $PageMarkdown -split "`r?`n"
    $LastParagraphTitle = ""
    $LastParagraph = ""

    ForEach($Line in $Lines){
        If ($($Line) -match "^# ") {

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
                    $LastParagraph = "# " + (Get-Date $LastParagraphTitle).ToString("dddd, MMMM dd, yyyy") +"`n`n"
                } 
                catch {
                    $LastParagraph = "# $($LastParagraphTitle)"
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

If ($DebugMessages.IsPresent){
    $DebugPreference = "Continue"
    Write-Debug("Debugging enabled")
}

# Advanced logic for parameters
If ($PSCmdlet.ParameterSetName -eq "Set1") {
    $NoPagePrint=$True
}
If ((-not $WithHTML) -and (-not $ToMarkdown)){
    $PlainText=$True
}
If ($ExportDir) {
    If (!(Test-Path -Path $ExportDir -PathType Container)) {
        New-Item -Path $ExportDir -ItemType Directory | Out-Null
    }
} Else {
    $ExportDir = "."
}

# Start the OneNote application COM object.
# $AssemblyFile = (get-childitem $env:windir\assembly -Recurse Microsoft.Office.Interop.OneNote.dll | Sort-Object Directory -Descending | Select-Object -first 1).FullName
# Add-Type -Path $AssemblyFile -IgnoreWarnings
$OneNoteApp = New-Object -ComObject OneNote.Application

# Get the hierarchy of notebooks and include all objects, down to the maximum level (pages).
<#
$NotebookPath = "C:\\Users\\kevin\\Documents\\coding\\export-onenote\\Sample OneNote"
$NotebookID = ""
$cftNone = [Microsoft.Office.Interop.OneNote.CreateFileType]::cftNone
$OneNoteApp.OpenHierarchy($NotebookPath, "", [ref]$NotebookID, $cftNone)
#>

# Ask OneNote for the hiearchy all the way down to the individual pages, which is as low as you can
# go.
# The 2013 schema is the most recent schema available.
[xml]$NotebooksXML = ""
$Scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages
$OneNoteVersion = [Microsoft.Office.Interop.OneNote.XMLSchema]::xs2013
$OneNoteApp.GetHierarchy("", $Scope, [ref]$NotebooksXML, $OneNoteVersion)

# TODO
# Technically, there an be multiple OneNote notebooks on a computer. We assume
# that there is only one notebook. This is a limitation of the script.

ForEach($Notebook in $NotebooksXML.Notebooks.Notebook)
{
    If ($PrintStructure.IsPresent) {
        Write-Output "Notebook Name: ""$($Notebook.Name.trim())"""
    }

    If (!$NoPagePrint){
        $NotebookPath = Join-Path -Path $ExportDir -ChildPath "$($Notebook.Name.trim()) notebook"
        If (!(Test-Path -Path $NotebookPath -PathType Container)) {
            New-Item -Path $NotebookPath -ItemType Directory | Out-Null
        }
    }

    ForEach($Section in $Notebook.Section)
    {
        If ($PrintStructure.IsPresent){
            Write-Output "- Section Name: ""$($Section.Name.trim())"""
        }

        If (!$NoPagePrint){
            $SectionPath = Join-Path -Path $NotebookPath -ChildPath "$($Section.Name.trim()) section"
            If ( !(Test-Path -Path $SectionPath -PathType Container)) {
                New-Item -Path $SectionPath -ItemType Directory | Out-Null
            }
        }

        ForEach($Page in $Section.Page) {
            If (!$NoPagePrint){
                $PagePath = Join-Path -Path $SectionPath -ChildPath "$($Page.Name.trim()) page"
                If (!(Test-Path -Path $PagePath -PathType Container)){
                    New-Item -Path $PagePath -ItemType Directory | Out-Null
                }
                [xml]$PageXML = ""
                $OneNoteApp.GetPageContent($Page.ID, [ref]$PageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic, $OneNoteVersion)
    
                # The styles can be uniquely defined for each page. Styles are
                # defined in the order that they are first used on the page. So,
                # h2 migh tbe style #2 or #7, depending on when it was first
                # used on that page.
                $PageStyles = @{}
                $PageStyles = Get-PageStyles $PageXML.DocumentElement
    
                # TODO
                # Do something with the tags. For now, we just have the code 
                # that collects them.

                $Tags = @{}
                $Tags = Get-Tags $PageXML.DocumentElement
    
                If ( $PageStyles.Keys.Count -eq 0 ){
                    # All pages have at least the PageTitle style
                    Throw "Could not find any styles for the page $($Page.Name)"
                    Exit 1
                }
    
                If (($PrintAllPages.IsPresent) -or ($PageToPrint)) {

                    # Start at the beginning of the content, which is the first Outline node. There
                    # can be multiple Outline nodes, but we're only interested in the first one.
                    [System.Xml.XmlElement]$OneNoteOutline = Find-OneNoteOutline $PageXML.DocumentElement 6
                    Write-Output "  - Page Name: ""$($Page.Name)"""
                    $DelimiterPrinted = $False
                    
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
                    
                    If ($PrintTags){
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
                    
                    If (($PrintAllPages.IsPresent) -or (($PageToPrint) -and (($Page.Name) -eq $PageToPrint))) {
                        If (!($DelimiterPrinted)){
                            Write-Output " "
                            Write-Output "--------"
                            $DelimiterPrinted=$True
                        }
                        
                        # Convert the page to markdown
                        $ConvertResult = Convert-Page $Page.id $OneNoteOutline $PageStyles ""

                        # Split the page into individual paragraphs and then
                        # write them to individual files.
                        $PageParagraphs = Split-Pages -PageMarkdown $ConvertResult.Paragraph -PageName $Page.Name

                        ForEach ($PageParagraph in $PageParagraphs.Keys){
                            $PageParagraphFileName = Join-Path -Path $PagePath -ChildPath "$($PageParagraph.TrimEnd()).md"
                            $PageParagraphs[$PageParagraph].TrimEnd() | Out-File -FilePath $PageParagraphFileName -Encoding utf8
                            Write-Output("$($PageParagraphs[$PageParagraph])")
                        }
                    }
    
                    If ($DelimiterPrinted){
                        Write-Output "--------"
                        Write-Output " "
                    }
                } 
                
                If ($PrintStructure.IsPresent) {
                    Write-Output "  - Page Name: ""$($Page.Name)"""
                }
            }
        }
    }
}

# Clean up after yourself.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNoteApp) | Out-Null