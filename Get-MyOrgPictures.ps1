<#
.SYNOPSIS
Get the pictures of all people specified, include all direct reports recursively.

.DESCRIPTION
This scripts enumerates all reports of the given people provided in the -Includes parameter, recursively, and retrieves their pictures, storing them in a folder.

.PARAMETER Includes
People to retrieve pictures for. If the given person is a manager, all its reports, direct and indirects, will be added as well.

.PARAMETER Excludes
List of people to not retrieve picture for.

.INPUTS
None.

.OUTPUTS
Pictures of people stored as files in a folder.

#>

param (
    [Parameter(Mandatory=$true)][string[]]$Includes,
    [string[]]$Excludes,
	[int]$Columns
	)

Connect-MgGraph

Write-Progress -Activity "Finding people" -PercentComplete 0

$todo = New-Object System.Collections.Queue
$excludePeople = @{}

foreach($person in $Includes)
{
	$personObject = Get-AzureADUser -SearchString $person
    if ($personObject -eq $null)
    {
        Write-Error "Unable to find person to include: $person"
        return
    }
    if ($personObject.Count -gt 1)
    {
        Write-Error "Searching for '$person' matched multiple people. Try using UPN or e-mail address of the person instead."
        return
    }

    $dummy = $todo.Enqueue($personObject)
}

if ($Excludes -ne $null)
{
    foreach($person in $Excludes)
    {
        $personObject = Get-AzureADUser -SearchString $person

        if ($personObject -eq $null)
        {
            Write-Error "Unable to find person to exclude: $person"
            return
        }
        if ($personObject.Count -gt 1)
        {
        Write-Error "Searching for '$person' matched multiple people. Try using UPN or e-mail address of the person instead."
            return
        }

        $excludePeople.Add($personObject.ObjectId, $null)
    }
}

$folder = Join-Path $([System.IO.Path]::GetTempPath()) "MyOrgPeople"

$people = New-Object System.Collections.ArrayList

$peopleFound = $todo.Count
$peopleCompleted = 0

while($todo.Count -gt 0)
{
    Write-Progress -Activity "Finding people" -PercentComplete $([Int]($peopleCompleted/$peopleFound*100))

    $person = $todo.Dequeue()
    if (-not $excludePeople.Contains($person.ObjectId))
    {
        $nameParts = $person.GivenName.Split(" ")
        if ($nameParts -eq $Null -or $nameParts.Count -eq 1)
        {
            $singleName = $person.GivenName
        }
        elseif ($nameParts[0][0] -cmatch "[a-zA-Z]")
        {
            $singleName = $nameParts[0]
        }
        else
        {
            $singleName = $nameParts[1]
        }

        $obj = [PSCustomObject]@{
            Name = $singleName
            Filename = $(Join-Path $folder $($person.ObjectId + ".jpg"))
            ObjectId = $person.ObjectId
        }
        $dummy = $people.Add($obj)
    }

	$reports = Get-AzureADUserDirectReport -ObjectId $person.ObjectId -All $true
    if ($reports -ne $null -and $reports.Length -gt 0)
    {
        foreach($report in $reports)
        {
            $dummy = $todo.Enqueue($report)
            $peopleFound++
        }
    }

    $peopleCompleted++
}
Write-Progress -Activity "Finding people" -Completed



$people =  $people | sort -Property Name

# Get folder ready, this folder is reused over multiple runs of the command so it can cache people pictures
if ($(Test-Path -Path $folder) -eq $False)
{
    New-Item -Type Directory -Path $folder
}

foreach($person in $people)
{
	if ($(Test-Path -Path $person.FileName) -eq $False)
	{
		$result = Get-MgUserPhotoContent -PassThru -UserId $person.ObjectId -OutFile $person.FileName -ErrorAction:SilentlyContinue
		if ($result -ne $null)
		{
			Write-Output "No pic for $personId"
		}
	}
}

Add-type -AssemblyName office
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $application.Presentations.add()
$slideType = “microsoft.office.interop.powerpoint.ppSlideLayout” -as [type]
$slide = $presentation.slides.Add(1,$slideType::ppLayoutTitleOnly)


$slideWidth = $presentation.PageSetup.SlideWidth
$slideHeight = $presentation.PageSetup.SlideHeight
$topHeader = 160
if ($Columns -gt 0)
{
	$columnCount = $Columns
	$rowCount = [math]::Ceiling($people.Count / $columnCount)
}
else
{
	$ratio = $slideWidth / ($slideHeight - $topHeader - 40)
	$picFitInSquare = [math]::Ceiling($people.Count / $ratio)
	$rowCount = [math]::Ceiling([math]::Sqrt($picFitInSquare))
	$columnCount = [math]::Ceiling($rowCount * $ratio)
}
$pictureSize = [math]::Floor($slideWidth / $columnCount)

$totalHeight = $rowCount * ($pictureSize + 20)
if ($totalHeight -gt ($slideHeight - $topHeader))
{
	$pictureSize = [math]::Floor(($slideHeight - $topHeader) / $rowCount) - 20
}



$currentColumn = 0
$vPos = $topHeader
foreach($person in $people)
{
	$hPos = $currentColumn * $pictureSize
	
	if ($(Test-Path -Path $person.FileName) -eq $True)
	{
		$pic = $slide.Shapes.AddPicture2($person.FileName, 0, 1, $hPos, $vPos)
	
		# make pic square using width as reference
		if ($pic.Height -gt $pic.Width)
		{
			$pic.PictureFormat.CropBottom = $pic.Height - $pic.Width
		}
		$pic.Width = [single]$pictureSize
	}
	else
	{
		$pic = $slide.Shapes.AddShape(1, $hPos, $vPos, $pictureSize, $pictureSize)
		$pic.TextFrame2.TextRange.Text = $person.Name[0].ToString()
		$pic.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # center aligned
		$pic.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x404040 # this should be grey color
		$pic.TextFrame2.TextRange.Font.AllCaps = 1
		$pic.TextFrame2.TextRange.Font.Size = 60
		$pic.Fill.ForeColor.RGB = 0x202020 # this should light grey color
		$pic.Line.Visible = 0
	}
	
	
	$label = $slide.Shapes.AddShape(1, $hPos, $vPos + $pictureSize, $pictureSize, 20)
	$label.TextFrame2.TextRange.Text = $person.Name
	$label.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # center aligned
	$label.TextFrame2.VerticalAnchor = 1 # top aligned
	$label.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF # this should be white
	$label.TextFrame2.TextRange.Font.Size = 9
	$label.Fill.ForeColor.RGB = 0 # this should be black
	$label.Line.Visible = 0

	
	$currentColumn++
	if ($currentColumn -ge $columnCount)
	{
		$currentColumn = 0
        $vPos = $vPos + $pictureSize + 20
	}
}

