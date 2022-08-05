#------ Create Main Service Offering ------#

function New-Service($CategoryId) {
    Enter-SeUrl ("$Domain"+"TDClient/81/askit/Requests/New?CategoryID=$CategoryId") -Driver $Driver
}