# Export-LingeringLinkObjects

param([string]$problemAttributesFile, [string]$targetFile)

$linkTypeColumn = 2
$objectDnColumn = 0
$objectsAlreadyWritten = new-object 'System.Collections.Generic.List[string]'
$reader = new-object System.IO.StreamReader($problemAttributesFile)
$writer = new-object System.IO.StreamWriter($targetFile)

if ($reader -eq $null -or $writer -eq $null)
{
    return
}

while ($null -ne ($buffer = ($reader.ReadLine())))
{
	$split = $buffer.Split("`t")
	if ($split.Length -eq 4)
	{
		if ($split[$linkTypeColumn] -eq "LingeringLink")
		{
			if ($objectsAlreadyWritten.Contains($split[$objectDnColumn]))
			{
				# Do nothing, we already wrote this object to the file
			}
			else
			{
				# Save the group membership so we can import it later
				$entry = new-object System.DirectoryServices.DirectoryEntry("LDAP://" + $split[$objectDnColumn])
				if ($entry.Properties.Contains("member"))
				{
					if ($entry.Properties["member"].Count -gt 0)
					{
						$writer.WriteLine("dn: " + $split[$objectDnColumn])
						$writer.WriteLine("changetype: modify")
						$writer.WriteLine("replace: member")
						foreach ($memberEntry in $entry.Properties["member"])
						{
							$writer.WriteLine("member: " + $memberEntry)
						}
						$writer.WriteLine("-")
					}
				}
				
				if ($entry.Properties.Contains("mail"))
				{
					if ($entry.Properties["mail"].Count -gt 0)
					{
						$writer.WriteLine("replace: mail")
						$writer.WriteLine("mail: " + $entry.Properties["mail"][0].ToString())
						$writer.WriteLine("-")
					}
				}

				if ($entry.Properties.Contains("proxyAddresses"))
				{
					if ($entry.Properties["proxyAddresses"].Count -gt 0)
					{
						$writer.WriteLine("replace: proxyAddresses")
						foreach ($proxyEntry in $entry.Properties["proxyAddresses"])
						{
							$writer.WriteLine("proxyAddresses: " + $proxyEntry)
						}
						$writer.WriteLine("-")
					}
				}

				if ($entry.Properties.Contains("legacyExchangeDN"))
				{
					if ($entry.Properties["legacyExchangeDN"].Count -gt 0)
					{
						$writer.WriteLine("replace: legacyExchangeDN")
						$writer.WriteLine("legacyExchangeDN: " + $entry.Properties["legacyExchangeDN"][0].ToString())
						$writer.WriteLine("-")
					}
				}

				$writer.WriteLine("")
				
				$objectsAlreadyWritten.Add($split[$objectDnColumn])
			}
		}
	}
}

$writer.Close()
$reader.Close()