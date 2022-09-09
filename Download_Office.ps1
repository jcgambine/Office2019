<#
	.SYNOPSIS
	Download Office 2019, 2021, and 365

	.PARAMETER Branch
	Choose Office branch: 2019, 2021, and 365

	.PARAMETER Channel
	Choose Office channel: 2019, 2021, and 365

	.PARAMETER Components
	Choose Office components: Access, OneDrive, Outlook, Word, Excel, PowerPoint, Teams

	.EXAMPLE Download Office 2019 with the Word, Excel, PowerPoint components
	DownloadOffice -Branch ProPlus2019Retail -Channel Current -Components Word, Excel, PowerPoint

	.EXAMPLE Download Office 2021 with the Excel, Word components
	DownloadOffice -Branch ProPlus2021Volume -Channel PerpetualVL2021 -Components Excel, Word

	.EXAMPLE Download Office 365 with the Excel, Word, PowerPoint components
	DownloadOffice -Branch O365ProPlusRetail -Channel SemiAnnual -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word

	.LINK
	https://config.office.com/deploymentsettings

	.LINK
	https://docs.microsoft.com/en-us/deployoffice/vlactivation/gvlks
#>
function DownloadOffice
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateSet("ProPlus2019Retail")]
		[string]
		$Branch,

		

		[Parameter(Mandatory = $true)]
		[ValidateSet("Access", "OneDrive", "Outlook", "Word", "Excel", "PowerPoint", "Teams")]
		[string[]]
		$Components
	)

	if (-not (Test-Path -Path "$PSScriptRoot\Default.xml"))
	{
		Write-Warning -Message "Default.xml doesn't exist"
		exit
	}

	[xml]$Config = Get-Content -Path "$PSScriptRoot\Default.xml" -Encoding Default -Force

	switch ($Branch)
	{
		ProPlus2019Retail
		{
			($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "ProPlus2019Retail"
		}
		
	}

	switch ($Channel)
	{
		Current
		{
			($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "Current"
		}
		
	}

	foreach ($Component in $Components)
	{
		switch ($Component)
		{
			
			Excel
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Excel']")
				$Node.ParentNode.RemoveChild($Node)
			}
			
			Word
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Word']")
				$Node.ParentNode.RemoveChild($Node)
			}
			PowerPoint
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='PowerPoint']")
				$Node.ParentNode.RemoveChild($Node)
			}
			Teams
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Teams']")
				$Node.ParentNode.RemoveChild($Node)
			}
		}
	}

	$Config.Save("$PSScriptRoot\Config.xml")

	
	# Start downloading to the Office folder
	Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/download `"$PSScriptRoot\Config.xml`"" -Wait
}

# Download Offce. Firstly, download Office, then install it
DownloadOffice -Branch ProPlus2019Retail -Channel Current -Components Word

# Install
# Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$PSScriptRoot\Config.xml`"" -Wait
