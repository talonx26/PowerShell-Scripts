
class wwStringTag 
{
    [string]$TagName
    [string]$Description
    [string]$IOServerComputerName
    static [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    static[string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [string]$AcquisitionRate 
    [string]$TimeDeadBand
    [string]$SamplesInAI
    [string]$MaxLength 
    [string]$Format = "ASCII"
    [string]$InitialValue 
    [string]$CurrentEditor
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp
}

class wwAnalogTag 
{
    [string]$TagName
    [string]$Description
    [string]$IOServerComputerName
    static [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    static[string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [string]$AcquisitionRate 
    [string]$StorageRate
    [string]$TimeDeadBand
    [string]$SamplesInAI
    [string]$AIMode = "All"
    [string]$EngUnits
    [string]$MinEU
    [string]$MinRaw
    [string]$MaxRaw
    [string]$Scaling
    [string]$RawType
    [string]$IntegerSize
    [string]$Sign
    [double]$ValueDeadband
    [string]$InitialValue 
    [string]$CurrentEditor
    [string]$RateDeadBand
    [ValidateSet("Linear", "Stair Step", "System Default" )][string]$InterpolationType
    [string]$RolloverValue
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp = "No"
    [string]$DeadBandType = "TimeValue"

}

class wwDescreteTag 
{
    [string]$TagName
    [string]$Description
    [string]$IOServerComputerName
    static [string]$IOServerAppName = "FSGateway"
    [string]$TopicName
    [string]$ItemName
    static[string]$AcquisitionType = "IOServer"
    [ValidateSet("Cyclic", "Delta")][string]$StorageType
    [string]$AcquisitionRate 
    [string]$TimeDeadBand
    [string]$SamplesInAI
    [string]$AIMode = "All" 
    [string]$Message0
    [string]$Message1
    [string]$InitialValue 
    [string]$CurrentEditor
    [ValidateSet("Yes", "No")][string]$ServerTimeStamp = "No"

}
$PCS7System = "wpcs7lj005"
$TopicName = "M111H05"
$file = ".\Data\WPCS7LJ005 - Tag Export.txt"
$importTags = import-csv $file -Delimiter "`t" -Encoding "UTF7"
$wwAnalogTags = @()
$wwDescreteTags = @()
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
            $ww.IOServerComputerName = $PCS7System
            $ww.TopicName = "OPC_$TopicName"
            $ww.ItemName = $tag.'Tag name'
            switch ($tag.'Acquisition type')
            {
                "Cyclical, continuous" { $ww.StorageType = "Cyclic"}
                "After every change" { $ww.StorageType = "Delta"}
                Default { $ww.StorageType = "UNKNOWN"}
            }
            
            #Add Tag to Descrete collection
            $wwDescreteTags += $ww
        }

        'Analog'
        {
            #Create New object of Analog Wonderware Tag
            $ww = new-object wwAnalogTag
            $ww.TagName = "$($topicName)_" + $tag.'Process Tag'.Replace("/PID.PV_IN", ".PV").Replace("/PID.SP", ".SP").Replace("/PID.LMN", ".MV").Replace("/FQ.", ".").Replace("/PID.MV#Value", ".MV").Replace("/PID.PV_Out#Value", ".PV").Replace(".SP#Value", ".SP").Replace("/EM.", ".").Replace("/MonAnalog.PV_Out#Value", ".PV").replace(".Out#Value", ".Out").Replace("RC_Executer_1/Exe.", "DRC.").Replace("/MonDigital", "").Replace("SP_AO", "SP").Replace("PV_Out#Value", "PV")
            $ww.Description = $Tag.Comment
            $ww.IOServerComputerName = $PCS7System
            $ww.TopicName = "OPC_$TopicName"
            $ww.ItemName = $tag.'Tag name'
            switch ($tag.'Acquisition type')
            {
                "Cyclical, continuous" { $ww.StorageType = "Cyclic"}
                "After every change" { $ww.StorageType = "Delta"}
                Default { $ww.StorageType = "UNKNOWN"}
            }
            $ww.EngUnits = $tag.Unit

            #Add Tag to Analog collection
            $wwAnalogTags += $ww
        }


    }
     
}

$wwDescreteTags | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | % { $_.replace("""", '')} | Out-File .\Data\descrete-out.txt
$wwAnalogTags | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | % { $_.replace("""", '')} | Out-File .\Data\analog-out.txt