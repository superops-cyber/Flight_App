$xml = [xml](Get-Content "WMTSCapabilities.xml")
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("wmts", "http://www.opengis.net/wmts/1.0")
$ns.AddNamespace("ows", "http://www.opengis.net/ows/1.1")

$layers = @("GOES-East_ABI_GeoColor", "GOES-East_ABI_Band13")

foreach ($layerId in $layers) {
    Write-Host "Layer: $layerId"
    $layerNode = $xml.SelectSingleNode("//wmts:Layer[ows:Identifier=`'$layerId`']", $ns)
    if ($layerNode) {
        # Select the first ResourceURL if multiple exist
        $resourceUrl = $layerNode.ResourceURL[0].template
        if (-not $resourceUrl) { $resourceUrl = $layerNode.ResourceURL.template }
        
        $tileMatrixSet = $layerNode.TileMatrixSetLink.TileMatrixSet
        Write-Host "  Template: $resourceUrl"
        Write-Host "  TileMatrixSet: $tileMatrixSet"
        
        $sampleUrl = $resourceUrl.Replace("{TileMatrixSet}", $tileMatrixSet)
        $sampleUrl = $sampleUrl.Replace("{TileMatrix}", "0")
        $sampleUrl = $sampleUrl.Replace("{TileRow}", "0")
        $sampleUrl = $sampleUrl.Replace("{TileCol}", "0")
        $sampleUrl = $sampleUrl.Replace("{Time}", "2026-04-22T18:00:00Z")
        
        Write-Host "  Sample URL: $sampleUrl"
        try {
            $resp = Invoke-WebRequest -Method Head -Uri $sampleUrl -ErrorAction Stop
            Write-Host "  Status Code: $($resp.StatusCode)"
        } catch {
            if ($_.Exception.Response) {
                Write-Host "  Status Code: $($_.Exception.Response.StatusCode.Value__)"
            } else {
                Write-Host "  Error: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  Layer not found."
    }
}
