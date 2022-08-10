# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
. "$PSScriptRoot\Modules\Service\CreateService.ps1"
. "$PSScriptRoot\Modules\AutomationRule\NewAutomationRule.ps1"
. "$PSScriptRoot\functionTester.ps1"
Import-Module "$PSScriptRoot\Modules\Selenium\3.0.1\Selenium.psd1"
Import-Module "$PSScriptRoot\Modules\AnyBox\AnyBox.psd1"

# User prompts
$Environment = Show-AnyBox -Message 'Apply changes to:' -Buttons 'Sandbox', 'Production'
if ($Environment.Production) {
    $Domain = "https://langara.teamdynamix.com/"
} else {
    $Domain = "https://langara.teamdynamix.com/SB"
}

$PromptBox = New-Object AnyBox.AnyBox
$PromptBox.Title = 'Service Creator'
$PromptBox.Prompts = @(
    New-AnyBoxPrompt -Name "CatName" -InputType Text -Message 'Category name or category ID (case sensitive):' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "ShortName" -InputType Text -Message 'Short name of the category for general support. This will generate General -shortname- Support:' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "EvalOrder" -InputType Text -Message 'GTS Automation Rule evaluation order:' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "Staff" -Group "Audience:" -InputType Checkbox -Message "Staff"
    New-AnyBoxPrompt -Name "Students" -Group "Audience:" -InputType Checkbox -Message "Students"
    New-AnyBoxPrompt -Name "Faculty" -Group "Audience:" -InputType Checkbox -Message "Faculty"
)
$PromptBox.Buttons = @(
    New-AnyBoxButton -Text 'Submit' -IsDefault
    New-AnyBoxButton -Text 'Cancel' -IsCancel
)
$UserInput = $PromptBox | Show-AnyBox
$UserInput

if ($UserInput.Cancel -or $UserInput.Cancel) {
    Show-AnyBox -Message "Operation cancelled" -Buttons 'Ok'
    Exit
}

# Start browser
$Driver = Start-SeFirefox
$WindowSize = [System.Drawing.Size]::new(800, 700)
$driver.Manage().Window.Size = $WindowSize

# Category name and ID validation
if($UserInput.CatName -notmatch '^\d+$') { # If input is a service name
    Enter-SeUrl ("$Domain"+"TDClient/81/askit/Requests/ServiceCatalog") -Driver $Driver
    $CategoryURL = $Driver.FindElementByXPath("//a[text()='$($UserInput.CatName)']").getAttribute('href')
    $CategoryId = $CategoryURL.Substring($CategoryURL.IndexOf('?') + 12)
} else {
    $CategoryId = $CategoryName
}

# Create new service
New-Service $CategoryId $UserInput

# Create GTS Automation Rule
New-AutomationRule $ServiceId $ServiceName $EvalOrder

# Stop driver
Stop-SeDriver -Driver $Driver
