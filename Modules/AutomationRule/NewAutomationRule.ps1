#------ Create General Service Offering ------#

function New-AutomationRule($ServiceId, $ServiceName, $CategoryName, $Responsible, $EvalOrder) {
    $TextInfo = (Get-Culture).TextInfo
    $Responsible = $TextInfo.toTitleCase($Responsible)

    # New rule
    Enter-SeUrl ("$Domain"+"TDAdmin/1cc3ff6f-33a6-4148-b145-f5581a4f32bd/82/AutomationRules/Index?Component=9") -Driver $Driver
    $NewBtn = Find-SeElement -Wait -Timeout 60 -Driver $Driver -XPath '//a[@class="btn btn-primary"]'
    Invoke-SeClick -Element $NewBtn

    # Name
    $CurrentField = Find-SeElement -Driver $Driver -Id "Name"
    $TextInfo = (Get-Culture).TextInfo
    Send-SeKeys -Element $CurrentField -Keys "GTS $CategoryName - Assign to $Responsible"

    # Order
    $CurrentField = Find-SeElement -Driver $Driver -Id "Order"
    $CurrentField.SendKeys([OpenQA.Selenium.Keys]::Backspace)
    Send-SeKeys -Element $CurrentField -Keys $EvalOrder

    # Stop on Match
    $Checkbox = Find-SeElement -Wait -Timeout 10 -Driver $Driver -Id "ShouldStopOnMatch"
    Invoke-SeClick -Element $Checkbox

    # Description
    $CurrentField = Find-SeElement -Driver $Driver -Id "Description"
    Send-SeKeys -Element $CurrentField -Keys "This rule assigns General Technical Support tickets created under the $CategoryName service to $Responsible."

    # Save
    $SaveBtn = Find-SeElement -Driver $Driver -XPath '//div[@id="divButtons"]//button//span[text()="Save"]'
    Invoke-SeClick -Element $SaveBtn

    # Edit
    $EditBtn = Find-SeElement -Wait -Timeout 60 -Driver $Driver -XPath '//button[@class="btn btn-primary"]'
    Invoke-SeClick -Element $EditBtn

    # Is Active
    $Checkbox = Find-SeElement -Wait -Timeout 10 -Driver $Driver -Id "Rule_IsActive"
    Invoke-SeClick -Element $Checkbox

    # Automation Conditions
    $Option = Find-SeElement -Driver $Driver -Id "filter_column_0"
    $SelectElement = [OpenQA.Selenium.Support.UI.SelectElement]::new($Option)
    $SelectElement.SelectByValue(789) # Condition is for Service
    $SearchBtn = Find-SeElement -Driver $Driver -XPath "//a[@data-textid='lu_text_0']"
    Invoke-SeClick -Element $SearchBtn
    $Windows = Get-SeWindow -Driver $Driver
    Switch-SeWindow -Driver $Driver -Window $Windows[1]
    $CurrentField = Find-SeElement -Wait -Timeout 3 -Driver $Driver -Id "ctl00_cphSearchRows_txtSearch"
    Send-SeKeys -Element $CurrentField -Keys $ServiceName
    $SearchBtn = Find-SeElement -Driver $Driver -Id "ctl00_btnSearch"
    Invoke-SeClick -Element $SearchBtn
    $ServiceCheckbox = Find-SeElement -Wait -Timeout 5 -Driver $Driver -Id "ctl00_cphGrid_grdItems_chkMass_$ServiceId"
    Invoke-SeClick -Element $ServiceCheckbox
    $InsertBtn = Find-SeElement -Driver $Driver -XPath "//div[@class='pull-left']//input[@value='Insert Checked']"
    Invoke-SeClick -Element $InsertBtn
    Switch-SeWindow -Driver $Driver -Window $Windows[0]

    # Automation Actions
    $CurrentField = Find-SeElement -Driver $Driver -Id "select2-chosen-7"
    Invoke-SeClick -Element $CurrentField
    $CurrentField = Find-SeElement -Driver $Driver -Id "s2id_autogen7_search"
    Send-SeKeys -Element $CurrentField -Keys $Responsible
    $Selection = Find-SeElement -Wait -Timeout 5 -Driver $Driver -XPath "//div[@class='select2-result-label']//div[contains(text(),'$Responsible')]"
    Invoke-SeClick -Element $Selection

    # Save edit
    $SaveBtn = Find-SeElement -Driver $Driver -XPath "//div[@id='divButtons']//button[@class='btn btn-primary']"
    Invoke-SeClick -Element $SaveBtn
}