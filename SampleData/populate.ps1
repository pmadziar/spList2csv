$webUrl = "http://team/workgroups/SITECOLL/bigdata"
$bigListName = "Data"
$refListName = "Options"
$options = @("Option 1", "Option 2", "Option 3", "Option 4")

# First Name    : Sean
# Last Name     : Pennington
# Date of Birth : 29/01/2017
# NI            : 1678031092099
# Company       : Viverra Consulting
# Mobile        : 07398 234434
# Value         : 114296
# Dept          : Human Resources


function CreateRefList(){
    $listName = $refListName
    $w = Get-SPWeb $webUrl
    $l = $w.Lists.TryGetList($listName)
    if($l -eq $null){
        Write-Host "Adding List: $($listName)"
        $lId = $w.Lists.Add("Options", "Lookup Data", [Microsoft.Sharepoint.SPListTemplateType]::GenericList)
        $l = $w.Lists[$lId]
        foreach($option in $options){
            Write-Host "Adding option: $($option)"
            $i = $l.Items.Add()
            $i[[Microsoft.SharePoint.SPBuiltInFieldId]::Title] = $option;
            $i.Update()
        } 
    } else {
        Write-Host "List $($listName) already exists"
    }
    $w.Dispose()
}

function CreateDataList(){
    $listName = $bigListName
    $w = Get-SPWeb $webUrl
    $l = $w.Lists.TryGetList($listName)
    if($l -eq $null){
        Write-Host "Adding List: $($listName)"
        $lId = $w.Lists.Add($bigListName, "Data", [Microsoft.Sharepoint.SPListTemplateType]::GenericList)
        $l = $w.Lists[$lId]
        # add fields
        $l.Fields.AddFieldAsXml('<Field DisplayName="First Name" ID="86257217-42b1-666d-b869-611a32cccc01" Name="FirstNameX" StaticName="FirstNameX" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Last Name" ID="86257217-42b1-666d-b869-611a32cccc02" Name="LastNameX" StaticName="LastNameX" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Date of Birth" Format="DateOnly" ID="86257217-42b1-666d-b869-611a32cccc03" Name="BoD" StaticName="BoD" Type="DateTime" />')
        $l.Fields.AddFieldAsXml('<Field DisplayName="NI" ID="86257217-42b1-666d-b869-611a32cccc04" Name="NI" StaticName="NI" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Company" ID="86257217-42b1-666d-b869-611a32cccc05" Name="CompanyNameX" StaticName="CompanyNameX" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Mobile" ID="86257217-42b1-666d-b869-611a32cccc06" Name="MobileNameX" StaticName="MobileNameX" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field Decimals="0" DisplayName="Value" ID="86257217-42b1-666d-b869-611a32cccc07" Name="ValueX" StaticName="ValueX" Type="Number"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Dept" ID="86257217-42b1-666d-b869-611a32cccc08" Name="DeptNameX" StaticName="DeptNameX" Type="Text"/>')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Last Login" Format="DateTime" ID="86257217-42b1-666d-b869-611a32cccc09" Name="LLoginX" StaticName="LLoginX" Type="DateTime" />')
        $l.Fields.AddFieldAsXml('<Field DisplayName="Distribution" ID="86257217-42b1-666d-b869-611a32cccc10" Name="DistributionX" StaticName="DistributionX" Type="Choice">
        <CHOICES>
            <CHOICE>Choice 1</CHOICE>
            <CHOICE>Choice 2</CHOICE>
            <CHOICE>Choice 3</CHOICE>
            <CHOICE>Choice 4</CHOICE>
            <CHOICE>Choice 5</CHOICE>
        </CHOICES>
        <Default>Choice 1</Default>
        </Field>')

        # add Lookup field
        $wId = $w.Id.ToString('D')
        $rlid = $w.Lists[$refListName].Id.ToString('D')
        $xmlStr = '<Field ID ="86257217-42b1-666d-b869-611a32cccc11" Type="Lookup" Name="OptionX" StaticName="OptionX" DisplayName="Option" List="LISTID" WebId="WEBID" ShowField = "Title" />'
        $xmlStr = $xmlStr.Replace("LISTID", $rlid).Replace("WEBID", $wId)
        $l.Fields.AddFieldAsXml($xmlStr)
        $l.Update()

    } else {
        Write-Host "List $($listName) already exists"
    }
    $w.Dispose()
}

function PopulateDataList(){
    $recs = Import-Csv -Delimiter ";" -Path '.\dataMay-4-2017.csv' -Verbose
        $listName = $bigListName
    $w = Get-SPWeb $webUrl
    $l = $w.Lists[$listName]
    Write-Host "Starting " -NoNewLine
    $count = 0
    foreach($rec  in $recs){
        AddDataListItem $l $rec
        $count++
        if($count % 100 -eq 0){
            if($count % 1000 -eq 0){
                Write-Host "*" -NoNewLine
            } else {
                Write-Host "." -NoNewLine
            }

        }
    }
    $w.Dispose()

}

function AddDataListItem($l, $rec){
    $rnd = new-object -typename System.Random
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    $i = $l.Items.Add()
    $i[[Microsoft.SharePoint.SPBuiltInFieldId]::Title] = $rec."First Name" + " " + $rec."Last Name"

    $i["First Name"] = $rec."First Name"
    $i["Last Name"] = $rec."Last Name"
    $i["Date of Birth"] =  [datetime]::ParseExact($rec."Date of Birth", "dd/MM/yyyy", $ci)
    $i["NI"] = [string]$rec."NI"
    $i["Company"] = $rec."Company"
    $i["Mobile"] = $rec."Mobile"
    $i["Value"] = [decimal]($rec."Value"); $i.Update()
    $i["Dept"] = $rec."Dept"
    $i["Last Login"] = [datetime]::Now.AddHours($rnd.NextDouble()*(-400.0))
    $i["Distribution"] = "Choice " + $($rnd.Next(4) + 1).ToString()
    $oid = $rnd.Next(4) + 1;
    $ostr = "$($oid);#$($options[$oid - 1])"
    $i["Option"] = $ostr
    $i.Update()

}

# CreateRefList
# CreateDataList
# PopulateDataList
