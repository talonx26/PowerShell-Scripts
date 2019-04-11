
class wwStringTag 
{
    [string]$TagName
    [string]$Description
    [string]$IOServerComputerName
    [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    [string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [int]$AcquisitionRate 
    [int]$TimeDeadBand
    [string]$SamplesInAI
    [int]$MaxLength 
    [string]$Format = "ASCII"
    [string]$InitialValue 
    [int]$CurrentEditor
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp
}

class wwAnalogTag 
{
    [string]$TagName
    [string]$Description = " "
    [string]$IOServerComputerName
    [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    [string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [int]$AcquisitionRate 
    [int]$StorageRate = 1000
    [int]$TimeDeadBand
    [int]$SamplesInAI
    [string]$AIMode = "All"
    [string]$EngUnits = "None"
    [int]$MinEU
    [int]$MaxEU
    [int]$MinRaw
    [int]$MaxRaw
    [string]$Scaling = "None"
    [string]$RawType = "MSFloat"
    [int]$IntegerSize
    [string]$Sign
    [double]$ValueDeadband
    [int]$InitialValue 
    [int]$CurrentEditor
    [int]$RateDeadBand
    [ValidateSet("Linear", "Stair Step", "System Default" )][string]$InterpolationType = "System Default"
    [int]$RolloverValue
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp = "No"
    [string]$DeadBandType = "TimeValue"

}

class wwDescreteTag 
{
    [string]$TagName
    [string]$Description
    [string]$IOServerComputerName
    [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    [string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [int]$AcquisitionRate 
    [int]$TimeDeadBand
    [int]$SamplesInAI
    [string]$AIMode = "All" 
    [string]$Message0
    [string]$Message1
    [int]$InitialValue 
    [int]$CurrentEditor
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp = "No"

}


<#
.SYNOPSIS
Convert PCS7 Tag Export to WonderWare Format

.DESCRIPTION
Convert PCS7 Tag Export to WonderWare Format

.PARAMETER PCS7
Name of PCS7 System

.PARAMETER TopicName
WonderWare Topic Name

.PARAMETER ImportTagFile
Full path or Relative path to the import file from PCS7

.PARAMETER ExportTagFile
Full path or Relative path to the export file from WonderWare

.EXAMPLE
ConvertTo-WonderWareTags -PCS7 "WCPS7LJ005" -TopicName "TopicOne" -ImportTagFile .\Import.txt -ExportTagFile .\Export.txt

.NOTES
General notes
#>
function ConvertTo-WonderWareTags
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string]$PCS7,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1)]
        [string]$TopicName,
        # Param2 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 2)]
        [string]$ImportTagFile,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 3)]
        [string]$ExportTagFile
    )

    Begin
    {
        $wwAnalogTags = @()
        $wwDescreteTags = @()
        $importTags = import-csv $ImportTagFile -Delimiter "`t" -Encoding "UTF7"
    }
    Process
    {
        Foreach ($Tag in $importTags)
        {
            switch ($Tag.'Tag Type')
            {
                'Binary'
                {
                    #Create New object of Digital Wonderware Tag
                    $ww = new-object wwDescreteTag
                    $ww.TagName = "$($topicName)_" + $tag.'Process Tag'.Replace("/MonDigital.Out#Value", ".Out").Replace("/PID.PV_IN", ".PV").Replace("/PID.SP", ".SP").Replace("/PID.LMN", ".MV").Replace("/FQ.", ".").Replace("/PID.MV#Value", ".MV").Replace("/PID.PV_Out#Value", ".PV").Replace(".SP#Value", ".SP").Replace("/EM.", ".").Replace("/MonAnalog.PV_Out#Value", ".PV").replace(".Out#Value", ".Out").Replace("RC_Executer_1/Exe.", "DRC.").Replace("/MonDigital", "").Replace("SP_AO", "SP").Replace("PV_Out#Value", "PV")
           
                    $ww.Description = $Tag.Comment
                    $ww.IOServerComputerName = $PCS7
                    $ww.TopicName = "OPC_$TopicName"
                    $ww.ItemName = $tag.'Tag name'
                    switch ($tag.'Acquisition type')
                    {
                        # changed default to Delta. preferred WW storage.
                        "Cyclical, continuous" { $ww.StorageType = "Delta"}
                        "After every change" { $ww.StorageType = "Delta"}
                        
                        Default { $ww.StorageType = "Delta"}
                    }
            
                    #Add Tag to Descrete collection
                    $wwDescreteTags += $ww
                }

                'Analog'
                {
                    #Create New object of Analog Wonderware Tag
                    $ww = new-object wwAnalogTag
                    $ww.TagName = "$($topicName)_" + $tag.'Process Tag'.Replace("/PID.PV_IN", ".PV").Replace("/PID.SP", ".SP").Replace("/PROFIBUS.FQ_Out#Value",".Totalizer").Replace("/PID.LMN", ".MV").Replace("/FQ.", ".").Replace("/PID.MV#Value", ".MV").Replace("/PID.PV_Out#Value", ".PV").Replace(".SP#Value", ".SP").Replace("/EM.", ".").Replace("/MonAnalog.PV_Out#Value", ".PV").replace(".Out#Value", ".Out").Replace("RC_Executer_1/Exe.", "DRC.").Replace("/MonDigital", "").Replace("SP_AO", "SP").Replace("PV_Out#Value", "PV").Replace("/PROFIBUS","").Replace("Dy_Out","Density").Replace("TE_Out","Temperature").Replace("#Value","")
                    $ww.Description = $Tag.Comment
                    $ww.IOServerComputerName = $PCS7
                    $ww.TopicName = "OPC_$TopicName"
                    $ww.ItemName = $tag.'Tag name'
                    switch ($tag.'Acquisition type')
                    {
                        # changed default to Delta. preferred WW storage.
                        "Cyclical, continuous" { $ww.StorageType = "Delta"}
                        "After every change" { $ww.StorageType = "Delta"}
                        
                        Default { $ww.StorageType = "Delta"}
                    }
                    if ($tag.Unit.trim() -ne '' -and $tag.unit.Trim() -ne "Perry")
                    {$ww.EngUnits = $tag.Unit }
             

                    #Add Tag to Analog collection
                    $wwAnalogTags += $ww
                }


            }
     
        }
    }
    End
    {
        ":(Mode)update	" | out-file  $ExportTagFile -Force
        ":(AnalogTag)TagName	Description	IOServerComputerName	IOServerAppName	TopicName	ItemName	AcquisitionType	StorageType	AcquisitionRate	StorageRate	TimeDeadband	SamplesInAI	AIMode	EngUnits	MinEU	MaxEU	MinRaw	MaxRaw	Scaling	RawType	IntegerSize	Sign	ValueDeadband	InitialValue	CurrentEditor	RateDeadband	InterpolationType	RolloverValue	ServerTimeStamp	DeadbandType" | out-file  $ExportTagFile -Append   
        $wwAnalogTags | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | % { $_.replace("""", '')} | Select-object -skip 1 | Out-File $ExportTagFile -Append
        ":(DiscreteTag)TagName	Description	IOServerComputerName	IOServerAppName	TopicName	ItemName	AcquisitionType	StorageType	AcquisitionRate	TimeDeadband	SamplesInAI	AIMode	Message0	Message1	InitialValue	CurrentEditor	ServerTimeStamp" | out-file  $ExportTagFile -Append
        $wwDescreteTags | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | % { $_.replace("""", '')} |  Select-object -skip 1 | Out-File $ExportTagFile -Append
    }
}



$PCS7 = "wpcs7lj006ss2"
$TopicName = "M109H05"
$file = ".\WWTagsM109H05.txt"

$import = "C:\Users\nk23208\Documents\WonderWare\wpcs7lj006ss2 - Tag_Export_8-21-18.txt"


