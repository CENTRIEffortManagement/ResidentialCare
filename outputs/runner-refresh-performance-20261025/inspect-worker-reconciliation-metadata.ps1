# Read only the authorized original workbook's connection/table metadata.
# Never open Excel, read M definitions/cell data, extract package files, or save.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$workbookPath = 'CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Worker Reconciliation.xlsx'
$outputPath = 'outputs/runner-refresh-performance-20261025/worker-reconciliation-metadata.json'
$before = Get-Item -LiteralPath $workbookPath
$stampBefore = "$($before.Length):$($before.LastWriteTimeUtc.Ticks)"
$archive = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $workbookPath).Path)
$readParts = [Collections.Generic.List[string]]::new()
try {
    function Open-MetadataReader {
        param([string] $Part)
        $entry = $archive.GetEntry($Part)
        if ($null -eq $entry) { return $null }
        $readParts.Add($Part)
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $settings.CloseInput = $true
        return [Xml.XmlReader]::Create($entry.Open(), $settings)
    }
    function Read-MetadataXml {
        param([string] $Part)
        $reader = Open-MetadataReader $Part
        if ($null -eq $reader) { return $null }
        try {
            $document = [Xml.XmlDocument]::new()
            $document.XmlResolver = $null
            $document.Load($reader)
            return ,$document
        }
        finally { $reader.Dispose() }
    }
    function Get-MetadataAttributes {
        param($Element)
        $attributes = [ordered] @{}
        foreach ($attribute in $Element.Attributes) { $attributes[$attribute.Name] = $attribute.Value }
        return $attributes
    }
    function Get-MetadataRelationships {
        param([string] $Part)
        $slash = $Part.LastIndexOf('/')
        $relsPart = $Part.Substring(0, $slash + 1) + '_rels/' + $Part.Substring($slash + 1) + '.rels'
        $document = Read-MetadataXml $relsPart
        if ($null -eq $document) { return }
        foreach ($relationship in $document.DocumentElement.ChildNodes) {
            if ($relationship.LocalName -ne 'Relationship' -or $relationship.GetAttribute('TargetMode') -eq 'External') { continue }
            $uri = [uri]::new([uri] ('http://metadata-package/' + $Part), $relationship.GetAttribute('Target'))
            if ($uri.Host -ne 'metadata-package') { throw 'Unexpected external metadata relationship.' }
            [pscustomobject] @{
                Id = $relationship.GetAttribute('Id')
                Type = ($relationship.GetAttribute('Type') -split '/')[-1]
                Part = [uri]::UnescapeDataString($uri.AbsolutePath.TrimStart('/'))
            }
        }
    }
    function Get-QueryRefreshMetadata {
        param($Query)
        $refresh = @($Query.ChildNodes | Where-Object LocalName -eq 'queryTableRefresh')
        if ($refresh.Count -eq 0) { return $null }
        $fields = @($refresh[0].ChildNodes | Where-Object LocalName -eq 'queryTableFields')
        return [pscustomobject] @{
            Attributes = (Get-MetadataAttributes $refresh[0])
            Fields = @($fields | ForEach-Object { $_.ChildNodes } | Where-Object NodeType -eq ([Xml.XmlNodeType]::Element) | ForEach-Object { Get-MetadataAttributes $_ })
        }
    }
    $relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $sheetRecords = [Collections.Generic.List[object]]::new()
    $reader = Open-MetadataReader 'xl/workbook.xml'
    try {
        while ($reader.Read()) {
            # Read sheet identity attributes only, not defined-name formulas.
            if ($reader.NodeType -eq [Xml.XmlNodeType]::Element -and $reader.LocalName -eq 'sheet') {
                $sheetRecords.Add([pscustomobject] @{
                    Name = $reader.GetAttribute('name')
                    Id = $reader.GetAttribute('id', $relationshipNamespace)
                    State = $reader.GetAttribute('state')
                })
            }
        }
    }
    finally { $reader.Dispose() }
    $sheetParts = @{}
    foreach ($relationship in @(Get-MetadataRelationships 'xl/workbook.xml')) {
        if ($relationship.Type -eq 'worksheet') { $sheetParts[$relationship.Id] = $relationship.Part }
    }
    $connectionDocument = Read-MetadataXml 'xl/connections.xml'
    $connections = @{}
    foreach ($connection in $connectionDocument.DocumentElement.ChildNodes) {
        if ($connection.LocalName -eq 'connection') { $connections[$connection.GetAttribute('id')] = $connection }
    }
    $outputs = [Collections.Generic.List[object]]::new()
    foreach ($sheet in $sheetRecords) {
        if (-not $sheetParts.ContainsKey($sheet.Id)) { continue }
        # Worksheet cell XML is never read. Relationships identify the query and
        # table metadata parts without examining saved rows or values.
        foreach ($relationship in @(Get-MetadataRelationships $sheetParts[$sheet.Id])) {
            if ($relationship.Type -eq 'queryTable') {
                $query = (Read-MetadataXml $relationship.Part).DocumentElement
                $outputs.Add([pscustomobject] @{
                    Sheet = $sheet.Name; SheetState = $sheet.State; Exposure = 'Worksheet.QueryTables'
                    Table = $null; Query = (Get-MetadataAttributes $query); Part = $relationship.Part
                    RefreshMetadata = (Get-QueryRefreshMetadata $query)
                })
            }
            elseif ($relationship.Type -eq 'table') {
                $table = (Read-MetadataXml $relationship.Part).DocumentElement
                foreach ($queryRelationship in @(Get-MetadataRelationships $relationship.Part)) {
                    if ($queryRelationship.Type -ne 'queryTable') { continue }
                    $query = (Read-MetadataXml $queryRelationship.Part).DocumentElement
                    $outputs.Add([pscustomobject] @{
                        Sheet = $sheet.Name; SheetState = $sheet.State; Exposure = 'ListObject.QueryTable'
                        Table = (Get-MetadataAttributes $table); Query = (Get-MetadataAttributes $query); Part = $queryRelationship.Part
                        RefreshMetadata = (Get-QueryRefreshMetadata $query)
                    })
                }
            }
        }
    }
    $targetOutputs = @($outputs | Where-Object { $_.Sheet -in @('Employees-ALL (2)', 'Employees-ALL (3)') -or $_.Exposure -eq 'Worksheet.QueryTables' })
    $selectedConnections = [ordered] @{}
    foreach ($id in @($targetOutputs | ForEach-Object { $_.Query['connectionId'] } | Select-Object -Unique)) {
        if (-not $connections.ContainsKey($id)) { continue }
        $connection = $connections[$id]
        $properties = [Collections.Generic.List[object]]::new()
        foreach ($child in $connection.ChildNodes) {
            if ($child.NodeType -ne [Xml.XmlNodeType]::Element) { continue }
            if ($child.LocalName -notin @('dbPr', 'olapPr', 'webPr', 'textPr')) { continue }
            $attributes = Get-MetadataAttributes $child
            if ($attributes.Contains('connection')) {
                $attributes['connection'] = [regex]::Replace($attributes['connection'], '(?i)\b(Password|Pwd|User ID|UID|AccessToken|Token)\s*=\s*("[^"]*"|''[^'']*''|[^;]*)', '$1=<redacted>')
            }
            $properties.Add([pscustomobject] @{ Type = $child.LocalName; Attributes = $attributes })
        }
        $selectedConnections[$id] = [pscustomobject] @{
            Attributes = (Get-MetadataAttributes $connection)
            Properties = @($properties)
        }
    }
    $result = [ordered] @{
        Workbook = $workbookPath
        Scope = 'Saved connection/table metadata only; no worksheet cell data or M definitions read.'
        TargetOutputs = $targetOutputs
        Connections = $selectedConnections
        OtherOutputsOnSameConnections = @($outputs | Where-Object {
            $_.Sheet -notin @('Employees-ALL (2)', 'Employees-ALL (3)') -and $selectedConnections.Contains($_.Query['connectionId'])
        })
        ConnectionCount = $connections.Count
        EmployeeQueryConnections = @($connections.Values | Where-Object {
            @($_.ChildNodes | Where-Object { $_.LocalName -eq 'dbPr' -and $_.GetAttribute('command') -eq 'SELECT * FROM [Employees-ALL]' }).Count -gt 0
        } | ForEach-Object { Get-MetadataAttributes $_ })
        MetadataOutputs = @($outputs | ForEach-Object {
            [pscustomobject] @{ Sheet = $_.Sheet; Exposure = $_.Exposure; TableName = $(if ($null -ne $_.Table) { $_.Table['name'] } else { $null }); QueryName = $_.Query['name']; ConnectionId = $_.Query['connectionId']; Part = $_.Part }
        })
        PartsRead = @($readParts | Select-Object -Unique)
    }
}
finally { $archive.Dispose() }
$after = Get-Item -LiteralPath $workbookPath
$result['MetadataUnchanged'] = $stampBefore -eq "$($after.Length):$($after.LastWriteTimeUtc.Ticks)"
if (-not $result['MetadataUnchanged']) { throw 'Workbook metadata changed during read-only inspection.' }
$result | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $outputPath
$result | ConvertTo-Json -Depth 12
