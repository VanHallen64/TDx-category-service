# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
. "$PSScriptRoot\Modules\Service\CreateService.ps1"
. "$PSScriptRoot\Modules\AutomationRule\NewAutomationRule.ps1"
. "$PSScriptRoot\functionTester.ps1"
Import-Module "$PSScriptRoot\Modules\Selenium\3.0.1\Selenium.psd1"
Import-Module "$PSScriptRoot\Modules\AnyBox\AnyBox.psd1"

# User prompts
$Environment = Show-AnyBox -Message 'Apply changes to:' -Buttons 'Sandbox', 'Production'
if ($ProdInput.Production) {
    $Domain = "https://langara.teamdynamix.com/"
} else {
    $Domain = "https://langara.teamdynamix.com/SB"
}

$PromptBox = New-Object AnyBox.AnyBox
$PromptBox.Prompts = @(
    New-AnyBoxPrompt -Name "CatName" -InputType Text -Message 'Category name or category ID (case sensitive):' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "ShortName" -InputType Text -Message 'Short name of the category for general support. This will generate General -shortname- Support:' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "EvalOrder" -InputType Text -Message 'GTS Automation Rule evaluation order:' -ValidateNotEmpty
    New-AnyBoxPrompt -InputType Text -Message "Audience:"
    New-AnyBoxPrompt -Name "Faculty" -InputType Checkbox -Message "Faculty"
    New-AnyBoxPrompt -Name "Staff" -InputType Checkbox -Message "Staff"
    New-AnyBoxPrompt -Name "Students" -InputType Checkbox -Message "Students"
)
$PromptBox.Buttons = @(
    New-AnyBoxButton -Text 'Submit' -IsDefault
    New-AnyBoxButton -Text 'Cancel' -IsCancel
)
$Input = $PromptBox | Show-AnyBox
$Input

if ($Input.CatName.Cancel -or $Input.ShortName.Cancel) {
    Show-AnyBox -Message "No category information provided" -Buttons 'Ok'
    Exit
}

# Start browser
$Driver = Start-SeFirefox
$WindowSize = [System.Drawing.Size]::new(800, 700)
$driver.Manage().Window.Size = $WindowSize

# Category name and ID validation
if($CategoryName -notmatch '^\d+$') { # If input is a service name
    Enter-SeUrl ("$Domain"+"TDClient/81/askit/Requests/ServiceCatalog") -Driver $Driver
    $CategoryURL = $Driver.FindElementByXPath("//a[text()='$CategoryName']").getAttribute('href')
    $CategoryId = $CategoryURL.Substring($CategoryURL.IndexOf('?') + 12)
} else {
    $CategoryId = $CategoryName
}
$ServiceName = "General $CategoryShortName Support"



# Create new service
New-Service $CategoryId

# Create GTS Automation Rule
# New-AutomationRule $ServiceId $ServiceName $EvalOrder

# # Stop driver
# Stop-SeDriver -Driver $Driver
