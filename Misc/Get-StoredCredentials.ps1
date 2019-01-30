# the path to stored credential
$credPath = "H:\Secrets\Cred_${env:USERNAME}_${env:COMPUTERNAME}.xml"
# check for stored credential
if ( Test-Path $credPath ) {
    #crendetial is stored, load it 
    $cred = Import-CliXml -Path $credPath
} else {
    # no stored credential: create store, get credential and save it
    $parent = split-path $credpath -parent
    if ( -not test-Path $parent) {
        New-Item -ItemType Directory -Force -Path $parent
    }
    $cred = get-credential
    $cred | Export-CliXml -Path $credPath
}