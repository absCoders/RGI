# Import-Module AWS.Tools.SimpleNotificationService

$clientSettings = @{
    "RGI" = @{"emailTo"=@("maria@absolution.com","rick@absolution.com", "whr@absolution.com", "wjz@absolution.com", "ewz@absolution.com");
            "emailFrom"="abs@absolution.com";
            "SmtpServer"="mail.absolution.com";
            "PROD"="\\192.168.110.83\Shared\RGI";
            "QA"="";
            "ReportsDir"="C:\VS\VDI\Reports\";
            "Solution"="VDI"};
    "NYA" = @{"emailTo"=@("maria@absolution.com", "wjz@absolution.com", "ewz@absolution.com");
            "emailFrom"="abs@absolution.com";
            "SmtpServer"="mail.absolution.com";
            "PROD"="\\192.168.170.101\Share\NYA";
            "QA"="";
            "ReportsDir"="C:\VS\VDI\Reports\";
            "Solution"="VDI"};
    "VAN" = @{"emailTo"=@("rick@absolution.com", "whr@absolution.com", "wjz@absolution.com", "ewz@absolution.com");
            "emailFrom"="abs@absolution.com";
            "SmtpServer"="mail.absolution.com";
            "PROD"="\\192.168.180.34\G\VDI";
            "QA"="";
            "ReportsDir"="C:\VS\VDI\Reports\";
            "Solution"="VDI"};        
}

$assembliesList = "ABS","ABSCS","ABSX","AP","AR","AS","CC","EC","ED","GL","IC","PO","SA","SO","TA","TAC","WB","WH","WHC","WO"


function Create-Assemblies-Xml($deployToEnvironment, $client){

    $deployBaseDir=$clientSettings[$client][$deployToEnvironment]

    $assembliesFolder = $deployBaseDir + "\bin";

    $xmlSettings = New-Object System.Xml.XmlWriterSettings;
    $xmlSettings.Indent = $true;
    $xmlSettings.IndentChars = "  ";
    $xmlSettings.Encoding = New-Object System.Text.UTF8Encoding($false);

    $xmlW = [System.Xml.XmlWriter]::Create("$assembliesFolder\assemblies.xml", $xmlSettings);
    $xmlW.WriteStartDocument($true);
    $xmlW.WriteStartElement("Assemblies");

    $assembliesList | %{ 
          $assemblyFileName = $_ + $(If ($_ -eq "ABS"){".exe" } Else { ".dll" })
          $assemblyFullPath = Join-Path $assembliesFolder $assemblyFileName
          $assemblyVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($assemblyFullPath).FileVersion
      
          $xmlW.WriteStartElement("Assembly");
          $xmlW.WriteElementString("Name",$_);
          $xmlW.WriteElementString("Version",$assemblyVersion);
          $xmlW.WriteEndElement(); 
     }

    $xmlW.WriteEndElement();
    $xmlW.Close();
}


function Deploy-Assemblies([string[]]$deployToEnvironments,[string[]]$assembliesToDeploy, $client){
  
    $deployToEnvironments | %{

     $myReptsDir = $clientSettings[$client]["ReportsDir"];
     $myRepts = gci $myReptsDir
     $qaReptsDir = $clientSettings[$client][$_] + "\Reports";
     $qaRepts = gci $qaReptsDir; 
  
     [array]$reportsToDeploy = Compare-Object -ReferenceObject $qaRepts -DifferenceObject $myRepts -Property Name, LastWriteTime -PassThru | Group Name | %{
                                                                                                            New-Object PSObject -Property @{
                                                                                                                QA = $_.Group | ?{ $_.SideIndicator -eq '<=' }
                                                                                                                Mine = $_.Group | ?{ $_.SideIndicator -eq '=>' }
                                                                                                            }
                                                                                                        } | ?{ $_.QA.LastWriteTime -lt $_.Mine.LastWriteTime }
                                                                                                        
        
        $assembliesFolder = $clientSettings[$client][$_] + "\bin";

        if ($_ -eq "PROD"){
            Create-Release-Folder $client $assembliesToDeploy
        }


        $assembliesToDeploy | %{
            #copy to assembliesFolder
            $assemblyFileName = $_ + $(If ($_ -eq "ABS"){".exe" } Else { ".dll" })

            Copy-Item "C:\VS\$($clientSettings[$client]["Solution"])\$_\bin\x86\Release\$assemblyFileName" -Destination $assembliesFolder
        }

        $reportsToDeploy | %{
            Copy-Item $_.Mine.FullName -Destination $qaReptsDir
        }

        Create-Assemblies-Xml $_ $client;
    }

    
    $envs = $deployToEnvironments -join ', '

    $deployMessageText = "`nEnvironments: $envs"
    if ($assembliesToDeploy.Count -gt 0){
        $deployMessageText += "`n`nAssemblies:"
        $assembliesToDeploy | %{ $deployMessageText += "`n`t--$($_ + $(If ($_ -eq "ABS"){".exe" } Else { ".dll" }))" }
    }
    if($reportsToDeploy.Count -gt 0){
        $deployMessageText += "`n`nReports:"
        $reportsToDeploy | %{ $deployMessageText += "`n`t--$( $_.Mine.Name )" }
    }  

    Send-Deploy-Email "Deployment to $envs" $deployMessageText $client;
    #Send-Txt-Via-AWS "$envs Deployment" $deployMessageText;
}

function Create-Release-Folder($client, [string[]]$itemsForDeploy){

    $results = @{}

    $releaseYYYYMMDD = Get-Date -Format 'yyyyMMddHHmmss'

    $prodBaseDir=$clientSettings[$client]['PROD']


    $qaReptsDir = $clientSettings[$client]["ReportsDir"];
    $qaRepts = gci $qaReptsDir;

    $prodDir = $prodBaseDir + "\bin";
    $prodReptsDir = $prodBaseDir + "\Reports";
    $prodFiles = gci $prodDir;
    $prodRepts = gci $prodReptsDir;

    #compare the contents of the NY-QATS1 folder w/the prod folder and get the list of updated files
    #$itemsForDeploy = Compare-Object -ReferenceObject $prodFiles -DifferenceObject $qaFiles -Property Name, LastWriteTime -PassThru;
    
    $reportsForDeploy = Compare-Object -ReferenceObject $prodRepts -DifferenceObject $qaRepts -Property Name, LastWriteTime -PassThru | Group Name | %{
                                                                                                            New-Object PSObject -Property @{
                                                                                                                Prod = $_.Group | ?{ $_.SideIndicator -eq '<=' }
                                                                                                                QA = $_.Group | ?{ $_.SideIndicator -eq '=>' }
                                                                                                            }
                                                                                                        } | ?{ $_.Prod.LastWriteTime -lt $_.QA.LastWriteTime }

    
    #if($itemsForDeploy.Length -eq 0 -and $reportsForDeploy.Length -eq 0){
    #    $results = new-object psobject -property @{Success=$false; Message='No new files were found in QA -- nothing deployed'}
    #    return $results; 
    #}

    
    #$inc = 96;
    #$append = "";
    #while(Test-Path ($prodBaseDir + "\Releases\$releaseYYYYMMDD$append")){
    #    $inc += 1
    #    $append = ([char]$inc);      
    #}

    #$releaseDir = $prodBaseDir + "\Releases\$releaseYYYYMMDD$append";

    $releaseDir = $prodBaseDir + "\Releases\$releaseYYYYMMDD";
    

    
    #create prod/releases/releaseYYYYMMDD and Rollback folders
    New-Item ($releaseDir  + "\Rollback") -Type Directory -Force | Out-Null;
    if($reportsForDeploy) { New-Item ($releaseDir  + "\Reports") -Type Directory -Force | Out-Null; }
    if($reportsForDeploy | ?{ $_.Prod }) { New-Item ($releaseDir  + "\Rollback\Reports") -Type Directory -Force | Out-Null; }

    #copy existing prod items to Rollback folder inside release folder    
    $itemsForDeploy | %{ [System.IO.FileInfo]"$prodDir\$_$(If ($_ -eq "ABS"){".exe" } Else { ".dll" })" } | Copy-Item -Destination ($releaseDir  + "\Rollback");
    #copy QA items to release folder
    #$itemsForDeploy | ?{ $_.Directory.Fullname -eq $qaDir } | Copy-Item -Destination $releaseDir;
    $itemsForDeploy | %{
    #copy to assembliesFolder
    $assemblyFileName = $_ + $(If ($_ -eq "ABS"){".exe" } Else { ".dll" })
     Copy-Item "C:\VS\$($clientSettings[$client]["Solution"])\$_\bin\x86\Release\$assemblyFileName" -Destination $releaseDir
     }

    #copy prod items to Rollback folder [System.IO.FileInfo]$filename
    $reportsForDeploy | ?{ $_.Prod } | %{ $_.Prod } | Copy-Item -Destination ($releaseDir + "\Rollback\Reports")
    $reportsForDeploy | %{ $_.QA } | Copy-Item -Destination ($releaseDir + "\Reports")

    #gci ($releaseDir + "\Rollback") -Recurse | %{ $_.LastWriteTime = (Get-Date) }  #do a touch on the rollback items so that if we copy them back into the bin directory they will be pulled down

    $deployCount = $itemsForDeploy.Length + ($reportsForDeploy | Measure-Object).Count # ($itemsForDeploy | ?{ $_.Directory.Fullname -eq $qaDir } | Measure-Object).Count + ($reportsForDeploy | Measure-Object).Count

    $results = new-object psobject -property @{Success=$true; Message="$deployCount files were copied to $releaseDir"; Release=$releaseDir}
    return $results;
}

function Deploy-Release($releaseFolder, $client){


	$releasesDir = "$($clientSettings[$client]["PROD"])\Releases"
	$binDir = "$($clientSettings[$client]["PROD"])\bin"
	$reportsDir = "$($clientSettings[$client]["PROD"])\Reports"
	#get current release
	$currentRelease = Get-Content "$releasesDir\Release.txt"
	$currentDir = "$releasesDir\$releaseFolder"
	
    $deploySubject = ""
    $deployMessage = ""

	if($releaseFolder -lt $currentRelease){  #this is a rollback
		#a rollback needs to sequentially apply each release between current and target in reverse order
		$rollbackDirs =  gci $releasesDir -Directory | ?{ $_.Name -match '^\d{6}' -and $_.Name -gt $releaseFolder } | sort -Descending
		$rollbackDirs | %{
			gci "$($_.FullName)\Rollback" -File | Copy-Item -Destination $binDir
			if(Test-Path "$($_.FullName)\Rollback\Reports"){ gci "$($_.FullName)\Rollback\Reports" -File | Copy-Item -Destination $reportsDir }
		}
        $deploySubject = "PRODUCTION ROLLBACK TO RELEASE $releaseFolder"
        $deployMessage = "ROLLED BACK TO RELEASE $releaseFolder"

	}else{ #normal deploy
		#normal deploy simply copies the contents of $releaseFolder to the bin folder
        $filesToDeploy = gci $currentDir -File
		$filesToDeploy | Copy-Item -Destination $binDir

        $reportsToDeploy = @()
		if(Test-Path "$currentDir\Reports"){ 
            $reportsToDeploy = gci "$currentDir\Reports" -File 
            $reportsToDeploy | Copy-Item -Destination $reportsDir 
        }

        $deploySubject = "PROD Deployment"
        $deployMessage = "`nRelease $releaseFolder deployed to Production"
        if($filesToDeploy.Count -gt 0){ $deployMessage += "`n`nAssemblies:" }
        $filesToDeploy | %{ $deployMessage += "`n`t--$($_.Name)" }
        if($reportsToDeploy.Count -gt 0){ $deployMessage += "`n`nReports:" }
        $reportsToDeploy | %{ $deployMessage += "`n`t--$($_.Name)" }
	}
	
	$releaseFolder | Out-File -FilePath "$releasesDir\Release.txt"   #update release file to reflect the new current release
	
    Send-Deploy-Email $deploySubject $deployMessage $client | Out-Null;
    Send-Txt-Via-AWS $deploySubject $deployMessage  | Out-Null;
	
	return new-object psobject -property @{Success=$true; Message="Deployed release $releaseFolder to production"}
}

 
function Send-Deploy-Email($deploySubject,$deployText,$client){
    
    $deployEmailBody = "Deployment completed.

    $deployText"

    if ($clientSettings[$client]["SmtpServer"] -ne ''){
        Send-MailMessage -From $clientSettings[$client]["emailFrom"] -To $clientSettings[$client]["emailTo"] -Subject $deploySubject -Body $deployEmailBody -SmtpServer $clientSettings[$client]["SmtpServer"]
    }

}