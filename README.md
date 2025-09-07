# This is suitable to develop on windows machine

## Steps from scratch
- Create the Add-in File:
- Open Excel → New Blank Workbook.
- Press Alt+F11 to open the VBA Editor.
- Write your macros, functions, or event handlers.\
    * For a Ribbon button, excel expects a specific signature in VBA, so add the `control As IRibbonControl` as a para:
        ```
        ' Must be in a standard module, not ThisWorkbook or Sheet
        Public Sub UpdateCoverSheet(control As IRibbonControl)
            ' your logic
        End Sub
        ```
        In manual VBA:
        ```
        Public Sub UpdateCoverSheet()
            ' no parameters
        End Sub
        ```
        Or wrap main function for Ribbon:
        ```
        ' Ribbon entry point (must be Public and take IRibbonControl)
        Public Sub DoUpdateCoverAndValidate(control As IRibbonControl)
            Call UpdateCoverAndValidate   ' just a wrapper
        End Sub

        ' Actual logic (no arguments, can be run manually)
        Public Sub DoUpdateCoverAndValidate()
            ' ... all your existing long code here ...
        End Sub
        ```
- Save As → Excel Add-In (*.xlam).
    `*.xlam` file saved in `Users\AppData\Roaming\Microsoft\AddIns`
- Use **officeribbonXEditor** to embed XML in Add-in for a ribbon button.
    1. Open `.xlam` file and choose it, click `Insert` -> `Office 2010+ Custom UI Part`.
    2. Add a new property named customUI and paste the XML content inside it.
        ```
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
            <ribbon>
                <tabs>
                <tab id="CustomTab" label="My Custom Macros">
                    <group id="MyGroup" label="Macro Tools">
                    <button id="btnUpdateCover" 
                            label="Update Cover Sheet" 
                            imageMso="FileSave" 
                            size="large"
                            onAction="UpdateCoverSheet" />
                    </group>
                </tab>
                </tabs>
            </ribbon>
        </customUI>
        ```
        **\* Ensure your VBA function in the add-in matches the onAction attribute in the XML:**

    3. Create a PowerShell Script `ps1` to Install the Add-in: The following PowerShell script installs your add-in and enables it automatically.
        ```
        $addinPath = "$env:APPDATA\Microsoft\AddIns\MyMacro.xlam"
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false

        # Copy the add-in to Excel's default AddIns folder
        Copy-Item "MyMacro.xlam" -Destination $addinPath -Force

        # Enable the add-in in Excel registry
        $excelKey = "HKCU:\Softwa re\Microsoft\Office\" + $xl.Version + "\Excel\Options"
        Set-ItemProperty -Path $excelKey -Name "OPEN" -Value $addinPath

        $xl.Quit()
        Write-Host "Excel add-in installed successfully."
        ```
    4. Create a `*.wxs` file.
    5. Convert the script to `msi`. \
        a. Download the WiX Toolset from https://wixtoolset.org/ and install it. \
        b. Create a new project in WiX or use a `.wxs` file template. \
        c. Add the PowerShell script as part of the installation process. You can include a CustomAction in your .wxs file to execute the PowerShell script. \
        d. Compile the `.wxs` File using WiX compiler:
        ```
        candle.exe installer.wxs
        light.exe installer.wixobj -o InstallAddin.msi
        ```

    \* **Recommended: Keep a master editable .xlsm file.**

## How to update the `msi` file
MSI relies on component GUIDs and file hashes to decide whether to replace a file.

### How to make sure the updated `.xlam` is installed
1. Ensure the `.xlam` file actually changed
* Open the `.xlam` in Excel, update the XML, save it, and make sure the modified timestamp updates.
* If you edit the Ribbon XML externally, make sure you re-import it into the .xlam before building the MSI.
2. Change the Component GUID
* Easiest way to force MSI to treat it as a new file is to change the Guid of the AddInFile component:
```
<Component Id="AddInFile" Guid="{NEW-GUID-HERE}">
    <File Id="UpdateCoverSheet.xlam" Source="UpdateCoverSheet.xlam" KeyPath="yes" Vital="yes" />
</Component>
```
\* You’ll still want to rebuild the MSI after updating the .xlam.


### UpgradeCode generation
```
Powershell: [guid]::NewGuid().ToString()
Powershell: [guid]::NewGuid()	
```

