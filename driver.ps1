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
    New-AnyBoxPrompt -Name "ServiceName" -InputType Text -Message 'General support service name:' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "Responsible" -InputType Text -Message 'Responsible (do not include "Group" in the name):' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "EvalOrder" -InputType Text -Message 'GTS Automation Rule evaluation order:' -ValidateNotEmpty
    New-AnyBoxPrompt -Name "Staff" -Group "Audience (at least one required):" -InputType Checkbox -Message "Staff"
    New-AnyBoxPrompt -Name "Students" -Group "Audience (at least one required):" -InputType Checkbox -Message "Students"
    New-AnyBoxPrompt -Name "Faculty" -Group "Audience (at least one required):" -InputType Checkbox -Message "Faculty"
)
$PromptBox.Buttons = @(
    New-AnyBoxButton -Text 'Submit' -IsDefault
    New-AnyBoxButton -Text 'Cancel' -IsCancel
)
$UserInput = $PromptBox | Show-AnyBox

if ($UserInput.Cancel) {
    Show-AnyBox -Message "Operation cancelled" -Buttons 'Ok'
    Exit
}

# Start browser
$Driver = Start-SeFirefox
$WindowSize = [System.Drawing.Size]::new(800, 700)
$driver.Manage().Window.Size = $WindowSize

# Sign in and wait for categories div
$URL = "$Domain"+"TDClient/81/askit/Requests/ServiceCatalog"
Enter-SeUrl $URL -Driver $Driver
$SignInButton = Find-SeElement -Driver $Driver -XPath "//div[@title='Sign In']/a[contains(text(), 'Sign In')]"
Invoke-SeClick -Element $SignInButton
Find-SeElement -Wait -Timeout 60 -Driver $Driver -Id "divCats" | Out-null

# Category name and ID validation
if($UserInput.CatName -notmatch '^\d+$') { # If input is a service name
    $CategoryURL = $Driver.FindElementByXPath("//a[text()='$($UserInput.CatName)']").getAttribute('href')
    $CategoryId = $CategoryURL.Substring($CategoryURL.IndexOf('?') + 12)
} else { # If input is a service ID
    $CategoryId = $UserInput.CatName
    $CategoryURL = $Driver.FindElementByXPath("//a[contains(@href, $($UserInput.CatName))]")
    $UserInput.CatName = $CategoryURL.Text
}

# Create new service
$ServiceId = New-Service $CategoryId $UserInput

# Create GTS Automation Rule
New-AutomationRule $ServiceId $UserInput.ServiceName $UserInput.CatName $UserInput.Responsible $UserInput.EvalOrder

# Stop driver
Stop-SeDriver -Driver $Driver
Show-AnyBox -Message "Service created" -Button "Ok" -DefaultButton "Ok"
