# Office Activator

**Office Activator** is an Office activation tool.

![MainWindow](https://github.com/bb107/OfficeActivator/blob/master/screenshots/main.png?raw=true)
![ServiceManage](https://github.com/bb107/OfficeActivator/blob/service/screenshots/sc.png?raw=true)

Just like other activation tools,
>**Office Activator is:**
> - intended to help people who lost activation of their legally-owned licenses, e.g. due to a change of hardware (motherboard, CPU, ...)
>
>**Office Activator is not:**
> - a one-click activation or crack tool
> - intended to activate illegal copies of software (Office, Project, Visio)

## How It Works
In simple terms, Office Activator makes the license status of Office always true by hooking some functions.
Therefore, the patched Office has no activation time limit (difference from KMS). Once Office is updated, the patch may become invalid and you need to run the patch again.

## Supported Version
<table>
    <tr>
        <th/>
        <th>Office 2010</th>
        <th>Office 2013</th>
        <th>Office 2016</th>
        <th>Office 2019</th>
        <th>Office 2021</th>
    </tr>
    <tr>
        <th>Windows 7</th>
        <td>✔️(No LicenseManagement)</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
    </tr>
    <tr>
        <th>Windows 8</th>
        <td>✖️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
    </tr>
    <tr>
        <th>Windows 10</th>
        <td>✖️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
    </tr>
    <tr>
        <th>Windows 11</th>
        <td>✖️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
        <td>✔️</td>
    </tr>
</table>

## How To Use
1. Download Office Activator and extract to a directory.
2. Determine the architecture of the installed Office, and then run the corresponding ```OfficeActivator[86/64].exe``` as administrator.
3. Use ```ospp.vbs``` or LicenseManagement function to uninstall Office product keys for all Grace channels.
4. Use ```ospp.vbs``` or LicenseManagement function to install the retail or KMS product key corresponding to the Office version. If necessary, install the license file first. You can select a pre-built key in the Product Key drop-down box.
5. Click the "Check patch status" button to check the patch status. If the tool does not automatically recognize the Office installation directory and the path of ```MSO.DLL```, you need to locate these paths first.
6. Check patch status. Typically, "SppcHook Patch State" and "MSO Patch State" are "SppcHook.dll not found." and "MSO.DLL is not patched." respectively. If this is not the case, and the indicator color is not green, you may need to reinstall Office.
7. At this point, just click the "Apply patch" button. If successful, the patch status indicator will turn green. If you need to restore the patch, just click the "Restore" button.
8. Enjoy!

## Run OfficeActivator As Service
OfficeActivator can now run as a Windows service, which will monitor the MSO.DLL for changes in the background, patching it as soon as Office is updated.
> 1. Click the "Open Service Manager" button.
> 2. Click the "Create Service" button in the pop-up window.
> 3. Click the "Start Service" button.

If you need to delete the service, just click the "Stop Service" button, and then click the "Delete Service" button.

### Notes on running as a service:
1. Service mode will only patch MSO.DLL, so you need to do it in window mode first.
2. Once the service is created, please do not move the OfficeActivator folder to another location.
3. OfficeActivator uses directory monitoring API to monitor files for changes, so it doesn't take too much system resources.
