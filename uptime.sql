Select Top 1000000 tsysAssetTypes.AssetTypeIcon10 As icon,
  tblAssets.AssetID,
  tblAssets.AssetName,
  Convert(Decimal(12,0),tblAssets.Uptime / 86400) As UptimeDays,
  Case
    When Convert(Decimal(12,0),tblAssets.Uptime / 86400) > 30 Then 'red'
  End As foregroundcolor,
  tsysAssetTypes.AssetTypename As AssetType,
  tblAssets.Username,
  'mailto:' + tblADusers.email + '&subject=' + tblAssets.AssetName As
  hyperlink_NotifyUser,
  '' + tblADusers.Displayname As hyperlink_name_NotifyUser,
  tblAssets.IPAddress,
  tblAssetCustom.Manufacturer,
  tblAssetCustom.Model,
  tblAssetCustom.Serialnumber,
  tblAssets.Lastseen
From tblAssets
  Inner Join tblAssetCustom On tblAssets.AssetID = tblAssetCustom.AssetID
  Inner Join tsysAssetTypes On tsysAssetTypes.AssetType = tblAssets.Assettype
  Inner Join tblADusers On tblADusers.Username = tblAssets.Username
Where Convert(Decimal(12,0),tblAssets.Uptime / 86400) Is Not Null And
  Convert(Decimal(12,0),tblAssets.Uptime / 86400) > 10 And
  tblAssetCustom.State = 1 And tsysAssetTypes.AssetType != 16
Order By UptimeDays Desc
