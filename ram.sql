Select Top 1000000 tsysOS.Image As Icon,
  tblAssets.AssetID,
  tblAssets.AssetName,
  Ceiling(tblPhysicalMemoryArray.MaxCapacity / 1024) As MaxCapacity,
  CorrectMemory.Memory,
  Case
    When CorrectMemory.Memory < 2048 Then '#afafaf'
    When CorrectMemory.Memory >= 2048 And CorrectMemory.Memory < 3072 Then
      '#ff4747'
    When CorrectMemory.Memory >= 3072 And CorrectMemory.Memory < 4096 Then
      '#ffda47'
    When CorrectMemory.Memory >= 4096 And CorrectMemory.Memory < 8192 Then
      '#7cff89'
    When CorrectMemory.Memory >= 8192 Then '#00ff3a'
    Else '#d4f4be'
  End As backgroundcolor,
  Cast(CorrectMemory.Used As numeric) As [Slots  used],
  tblPhysicalMemoryArray.MemoryDevices As [Slots available],
  tblPhysicalMemoryArray.MemoryDevices - CorrectMemory.Used As [Slots free],
  tblAssets.Domain,
  tblAssets.IPAddress,
  tblAssets.Description,
  tblAssetCustom.Model,
  tblAssetCustom.Location,
  tsysIPLocations.IPLocation,
  tsysOS.OSname As OS,
  tblAssets.SP As SP,
  tblAssets.Lastseen,
  tblWarrantyDetails.WarrantyStartDate,
  tblWarrantyDetails.WarrantyEndDate
From tblAssets
  Inner Join tblPhysicalMemoryArray On
    tblAssets.AssetID = tblPhysicalMemoryArray.AssetID
  Inner Join (Select tblAssets.AssetID,
        Sum(Ceiling(tblPhysicalMemory.Capacity / 1024 / 1024)) As Memory,
        Count(tblPhysicalMemory.Win32_PhysicalMemoryid) As Used
      From tblAssets
        Left Outer Join (TsysMemorytypes
        Right Outer Join tblPhysicalMemory On TsysMemorytypes.Memorytype =
          tblPhysicalMemory.MemoryType) On tblAssets.AssetID =
          tblPhysicalMemory.AssetID
      Group By tblAssets.AssetID,
        tblPhysicalMemory.MemoryType
      Having tblPhysicalMemory.MemoryType <> 11) CorrectMemory On
    CorrectMemory.AssetID = tblAssets.AssetID And
    Ceiling(tblPhysicalMemoryArray.MaxCapacity / 1024) > CorrectMemory.Memory
    And Ceiling(tblPhysicalMemoryArray.MaxCapacity / 1024) >
    CorrectMemory.Memory
  Inner Join tblAssetCustom On tblAssets.AssetID = tblAssetCustom.AssetID
  Inner Join tsysOS On tblAssets.OScode = tsysOS.OScode
  Left Join tsysIPLocations On tsysIPLocations.StartIP <= tblAssets.IPNumeric
    And tsysIPLocations.EndIP >= tblAssets.IPNumeric
  Inner Join tblWarranty On tblAssets.AssetID = tblWarranty.AssetId
  Inner Join tblWarrantyDetails On tblWarranty.WarrantyId =
    tblWarrantyDetails.WarrantyId,
  tblADusers
Where tblPhysicalMemoryArray.MemoryDevices - CorrectMemory.Used >= 0 And
  tblPhysicalMemoryArray.[Use] = 3 And tblAssetCustom.State = 1
Order By CorrectMemory.Memory,
  [Slots free] Desc,
  tblAssets.AssetName
