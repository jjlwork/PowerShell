<#
    .SYNOPSIS
        Compares two the csv files and updates the a SharePoint List. This script requires
        the SharepointPnP module to be available on the host system.
    .DESCRIPTION
        Compares the csv exports that are located in the Kronos Server
        share-drive. The identity of an Open Shift is the ExceptionID. This
        script compares the ExceptionID, Start Time, End Time and Date in the
        csv outputs. 
        
        Where an item is unique to the previous days export, a query on 
        ExceptionID is excuted on the SharePoint list and the returned list item(s) 
        are deleted. 
        
        When an item is unique to to the newest list, the new item is created in
        the SharePoint List.
        
        Where an item is unique to both list but shares an Exception ID, this is
        the result of an updated shift exception. In this case as there is
        potential to have multiple records with the same ExceptionID, the current
        item(s) with the matching ExceptionID is deleted and the newest item(s)
        matching the exceptionID is added as a new list item(s)
#>

#SharePoint Site and List
$SiteURL = "http://mysharepoint/100/sandbox/"
$ListName = "Open Shift"
$logFile = ".\Update-OpenShifts.log"
$oldfile = Import-CSV -Delimiter ';' -Path .\oldfile.csv
$newfile = Import-CSV -Delimiter ';' -Path .\newfile.csv
$compareCSV = Compare-Object $newfile $oldfile -Property ExceptionID, "Start Time", "End Time","Date" -PassThru

Connect-PnPOnline -Url $SiteURL 

#If Exists in old file but not in new delete the item from SharePoint
ForEach ($item in $compareCSV) {
   
   
    if ($item.SideIndicator -eq "<=") {
       
        #Query the list for the item and delete it
        $Query = "<View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='ExceptionID' />
                                <Value Type='Number'>" + $item.ExceptionID + "
                                </Value>
                            </Eq>
                        </Where>
                    </Query>
                </View>"
        
        
        $deleteItems = Get-PnPListItem -List $ListName -Query $Query

        
        Try {
            ForEach ($deleteItem in $deleteItems){
                Remove-PnPListItem -List $ListName -Identity $deleteItem -Force
                $log = (Get-Date -Format "dd/MM/yyyy HH:mm:ss K") + " Deleted:" + $item
                Add-Content $logFile $log}
        }
        catch{
            $log = (Get-Date -Format "dd/MM/yyyy HH:mm:ss K") + " Error Deleting:" + $item
            Add-Content $logFile $log
        }
    }
}

#Add any items that were in the new file but not in the old file.
ForEach ($item in $compareCSV) {
    
    if ($item.SideIndicator -eq "=>") {
        
        Try {
            Add-PNPListItem -List $ListName -Values @{
                "Title" = $item.Facility;
                'User_x0020_ID' = $item.'User ID';
                "Unit" = $item.Unit;
                "ExceptionID" = $item.ExceptionID;
                "Occupation"=$item.Occupation;
                "Shift" = $item.Shift;
                "Start_x0020_Time" = [DateTime]::ParseExact($item.'Start Time',"dd/MM/yyyy HH:mm:ss",$null);
                "End_x0020_Time" = [DateTime]::ParseExact($item.'End Time',"dd/MM/yyyy HH:mm:ss",$null);
                "Status" = $item.Status;
                "Date" = $item.Date;
            }    
        }
        Catch{
            $log = (Get-Date -Format "dd/MM/yyyy HH:mm:ss K") + " Error Creating:" + $item
            Add-Content $logFile $log
        }

        $log = (Get-Date -Format "dd/MM/yyyy HH:mm:ss K") + " Created:" + $item
        Add-Content $logFile $log

    }
 }
