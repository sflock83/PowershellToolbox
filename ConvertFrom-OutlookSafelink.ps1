function ConvertFrom-OutlookSafelink {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipelineByPropertyName,
            ValueFromPipeline
        )]
        [string]$URL
    )
    process {
        try {
            $URL = [system.uri]::UnescapeDataString($URL)
            switch -regex ($URL) {
                '.safelinks.protection.outlook.com\/\?url=.+&data=' { $splitString = '&data='; break }
                '.safelinks.protection.outlook.com\/\?url=.+&amp;data=' { $splitString = '&amp;data='; break }
                Default { throw 'Invalid Safelinks URL.' }
            }
            $URL = $Matches[$Matches.Count - 1]
            $URL = (($URL -Split '\?url=')[1] -Split $splitString)[0]
            return $URL
        } catch {
            Write-Verbose 'Failed converting Safelink to original link.'
            Write-Verbose ('Safelink URL: ' + $URL)
            $_
        }
    }
}
