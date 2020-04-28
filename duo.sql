Select Top 1000000 tsysOS.Image As icon,
  tblAssets.AssetID,
  tblAssets.AssetName,
  tblAssets.Domain,
  tblAssets.Username,
  tblADusers.Displayname,
  tblAssets.Userdomain,
  tblAssets.IPAddress,
  tblAssets.Firstseen,
  tblAssets.Lastseen,
  tblAssets.Lasttried
From tblAssets
  Inner Join tblAssetCustom On tblAssets.AssetID = tblAssetCustom.AssetID
  Inner Join tsysOS On tsysOS.OScode = tblAssets.OScode
  Left Join tblADusers On tblADusers.Username = tblAssets.Username And
    tblADusers.Userdomain = tblAssets.Userdomain
Where tblAssets.AssetID Not In (Select Top 1000000 tblSoftware.AssetID
      From tblSoftware Inner Join tblSoftwareUni On tblSoftwareUni.SoftID =
          tblSoftware.softID
      Where tblSoftwareUni.softwareName Like '%Duo%') And tblAssetCustom.State =
  1
Order By tblAssets.Domain,
  tblAssets.AssetName
