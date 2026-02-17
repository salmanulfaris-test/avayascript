[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [int]$sec
)

while($true)
{

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @'
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
    int a(); int o();
    int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int f();
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

public static string GetDefault (int direction) {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(direction, 1, out dev));
    string id = null;
    Marshal.ThrowExceptionForHR(dev.GetId(out id));
    return id;
}
'@ -name audio -Namespace system

function Get-AudioDeviceFriendlyName($id) {
    $reg = "HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\MMDEVAPI\$id"
    return (get-ItemProperty $reg).FriendlyName
}

$defaultMic = ""
$defaultMicGuid = ""
$defaultSpeaker = ""
$defaultSpeakerGuid = ""
    
for ($i = 1; $i -le 5; $i++) {
    # Runs 5 times
    $defaultSpeakerGuid = $([audio]::GetDefault(0))
    $defaultSpeaker = $(Get-AudioDeviceFriendlyName $defaultSpeakerGuid)
    $defaultMicGuid = $([audio]::GetDefault(1))
    $defaultMic = $(Get-AudioDeviceFriendlyName $defaultMicGuid)
    
    #Write-Host "Default Speaker: $defaultSpeaker" 
    #Write-Host "Default Microphone  : $defaultMic"
    
    if ($defaultMic -like "*Sanas*" -and $defaultSpeaker -like "*Sanas*") {
        break
    }
    Start-Sleep -Seconds 3
}
    
if ($defaultMic -like "*Sanas*" -or $defaultSpeaker -like "*Sanas*") {

    #Find Avaya Input Device
    $AvayaInputDevice = (Get-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "ActiveRealRecordingDevice").ActiveRealRecordingDevice

    #Find Avaya Output Device
    $AvayaOutputDevice = (Get-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "ActivePlaybackDevice").ActiveRealRecordingDevice

    #Write-Host "Avaya Speaker: $AvayaOutputDevice" 
    #Write-Host "Avaya Microphone  : $AvayaInputDevice"
    
    if ($AvayaInputDevice -notlike "*Sanas*")
    {
    #Make registry changes to set Sanas as default
    Set-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "ActiveRealRecordingDevice" -Value ""
    Set-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "PreferredWaveInDeviceName" -Value ""

    Set-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "ActivePlaybackDevice" -Value ""
    Set-ItemProperty -Path "HKCU:Software\AVAYA\Avaya one-X Agent\Audio" -Name "PreferredWaveOutDeviceName" -Value ""

    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.WindowState = 'Normal'
    $form.StartPosition = 'CenterScreen'
    $form.ShowInTaskbar = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.Text = "Sanas-Avaya Notification"
    $form.Size = New-Object System.Drawing.Size(420,200)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Label (centered, multiline)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Please re-launch Avaya to use Sanas"
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(380,70)
    $label.Location = New-Object System.Drawing.Point(20,25)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($label)

    $form.Controls.Add($okButton)

    # Ok button (skip)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Ok"
    $okButton.Size = New-Object System.Drawing.Size(90,30)
    $okButton.Location = New-Object System.Drawing.Point(230,110)
    $okButton.Add_Click({
        $form.Close()
    })
    $form.Controls.Add($okButton)

    # Show dialog
    $form.ShowDialog() | Out-Null
 }

 }

 else{
#Write-Host "Default Mic is not Sanas"
#Sanas Device Not Found Windows Pop-up
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.WindowState = 'Normal'
$form.StartPosition = 'CenterScreen'
$form.ShowInTaskbar = $true
$form.FormBorderStyle = 'FixedDialog'
$form.Text = 'Sanas Not Detected'

[System.Windows.Forms.MessageBox]::Show(
    $form,
    "Sanas is not found as the default device.`nPlease launch Sanas.",
    "Sanas Alert",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)
 }
 Start-Sleep -Seconds $sec
 }
