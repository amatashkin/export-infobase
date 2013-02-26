# Create a 1C 8.2 COM Connector with Powershell

$comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog
$apps = $comAdmin.GetCollection("Applications")
$apps.Populate();

$newComPackageName = "V82_COMCOnnector"

# Путь к dll. Надо добавить автоопределение либо передачу параметра
$dllPath = "C:\Program Files (x86)\1cv82\8.2.15.319\bin\comcntr.dll"

$appExistCheckApp = $apps | Where-Object {$_.Name -eq $newComPackageName}

if($appExistCheckApp)
{
    $appExistCheckAppName = $appExistCheckApp.Value("Name")
    "This COM+ Application already exists : $appExistCheckAppName"
}
Else
{
    $newApp1 = $apps.Add()
    $newApp1.Value(“Name”) = $newComPackageName
    # Security Tab, Authorization Panel, “Enforce access checks for this application 
    $newApp1.Value(“ApplicationAccessChecksEnabled”) = 0 


    # Optional (to set to a specific Identify) 
    # $newApp1.Value(“Identity”) = “MyDomain\myUserName”
    # $newApp1.Value(“Password”) = “myPassword”

    $saveChangesResult = $apps.SaveChanges()
    "Results of the SaveChanges operation : $saveChangesResult"
    $comAdmin.InstallComponent($newComPackageName, $dllPath, "", "")
}

# Full documentation of the properties here.
# See http://msdn.microsoft.com/en-us/library/ms686107(v=VS.85).aspx identity for full documentation.
