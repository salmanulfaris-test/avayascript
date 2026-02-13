##############################################################################
# SanasAvayaCheck.ps1
# Purpose : Ensures Sanas is the default audio device and configures
#           Avaya one-X Agent to use it. Shows user popup if action needed.
# Runs    : Every hour via Windows Scheduled Task (deployed by Intune)
# Context : Logged-on user (HKCU access + GUI popups required)
##############################################################################

# -------------------------
# Log setup (create early)
# -------------------------
$LogFolder = "C:\ProgramData\CustomScripts"
$LogPath   = Join-Path $LogFolder "SanasAvayaCheck.log"
if (-not (Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "$ts - $msg"
}

Write-Log "=== Script execution started ==="

try {
    # Assemblies (GUI)
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop

    # -------------------------
    # Audio default device helper
    # -------------------------
    Add-Type @'
using System;
using System.Runtime.InteropServices;

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

public static class AudioHelper {
    public static string GetDefault(int direction) {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(direction, 1, out dev));
        string id = null;
        Marshal.ThrowExceptionForHR(dev.GetId(out id));
        return id;
    }
}
'@ -Name AudioNS -Namespace SanasAudio -ErrorAction SilentlyContinue

    function Get-AudioDeviceFriendlyName($id) {
        $reg = "HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\MMDEVAPI\$id"
        try { return (Get-ItemProperty $reg -ErrorAction Stop).FriendlyName } catch { return "" }
    }

    # -----------------------------------------
    # Detect default audio devices (retry loop)
    # -----------------------------------------
    $defaultMic     = ""
    $defaultSpeaker = ""

    for ($i = 1; $i -le 5; $i++) {
        try {
            $defaultSpeakerGuid = [SanasAudio.AudioHelper]::GetDefault(0)  # speakers
            $defaultSpeaker     = Get-AudioDeviceFriendlyName $defaultSpeakerGuid

            $defaultMicGuid     = [SanasAudio.AudioHelper]::GetDefault(1)  # mic
            $defaultMic         = Get-AudioDeviceFriendlyName $defaultMicGuid
        }
        catch {
            Write-Log "Attempt $i - Error detecting audio devices: $($_.Exception.Message)"
        }

        if ($defaultMic -like "*Sanas*" -and $defaultSpeaker -like "*Sanas*") {
            Write-Log "Sanas detected as default mic and speaker on attempt $i"
            break
        }
        Start-Sleep -Seconds 3
    }

    Write-Log "Default Speaker    : $defaultSpeaker"
    Write-Log "Default Microphone : $defaultMic"

    # -------------------------
    # Main logic
    # -------------------------
    if ($defaultMic -like "*Sanas*" -or $defaultSpeaker -like "*Sanas*") {

        # Guard: Avaya HKCU key may not exist until Avaya is launched/logged in
        $avayaKey = "HKCU:\Software\AVAYA\Avaya one-X Agent\Audio"
        if (-not (Test-Path $avayaKey)) {
            Write-Log "Avaya HKCU audio registry path not found ($avayaKey). Avaya may not be initialized yet. Exiting."
            Write-Log "=== Script execution completed ==="
            exit 0
        }

        # Read current Avaya audio settings
        $AvayaInputDevice  = (Get-ItemProperty -Path $avayaKey -Name "ActiveRealRecordingDevice" -ErrorAction SilentlyContinue).ActiveRealRecordingDevice
        $AvayaOutputDevice = (Get-ItemProperty -Path $avayaKey -Name "ActivePlaybackDevice"      -ErrorAction SilentlyContinue).ActivePlaybackDevice

        Write-Log "Avaya Input Device : $AvayaInputDevice"
        Write-Log "Avaya Output Device: $AvayaOutputDevice"

        if ($AvayaInputDevice -notlike "*Sanas*") {
            Write-Log "Avaya NOT using Sanas — updating registry and notifying user"

            # Clear Avaya audio device registry entries so it picks up the system default (Sanas)
            Set-ItemProperty -Path $avayaKey -Name "ActiveRealRecordingDevice"  -Value "" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $avayaKey -Name "PreferredWaveInDeviceName"  -Value "" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $avayaKey -Name "ActivePlaybackDevice"       -Value "" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $avayaKey -Name "PreferredWaveOutDeviceName" -Value "" -ErrorAction SilentlyContinue

            Write-Log "Registry updated — cleared Avaya audio device entries"

            # Show notification to user
            $form = New-Object System.Windows.Forms.Form
            $form.TopMost         = $true
            $form.StartPosition   = 'CenterScreen'
            $form.FormBorderStyle = 'FixedDialog'
            $form.Text            = "Sanas-Avaya Notification"
            $form.Size            = New-Object System.Drawing.Size(420, 200)
            $form.MaximizeBox     = $false
            $form.MinimizeBox     = $false
            $form.ShowInTaskbar   = $true

            $label           = New-Object System.Windows.Forms.Label
            $label.Text      = "Please re-launch Avaya to use Sanas"
            $label.AutoSize  = $false
            $label.Size      = New-Object System.Drawing.Size(380, 70)
            $label.Location  = New-Object System.Drawing.Point(20, 25)
            $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $label.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
            $form.Controls.Add($label)

            $okButton          = New-Object System.Windows.Forms.Button
            $okButton.Text     = "Ok"
            $okButton.Size     = New-Object System.Drawing.Size(90, 30)
            $okButton.Location = New-Object System.Drawing.Point(160, 110)
            $okButton.Add_Click({ $form.Close() })
            $form.Controls.Add($okButton)

            $form.ShowDialog() | Out-Null
            Write-Log "User acknowledged the relaunch notification"
        }
        else {
            Write-Log "Avaya is already using Sanas — no action needed"
        }
    }
    else {
        Write-Log "Sanas NOT detected as default device — alerting user"

        $form         = New-Object System.Windows.Forms.Form
        $form.TopMost = $true
        $form.StartPosition   = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.Text            = 'Sanas Not Detected'

        [System.Windows.Forms.MessageBox]::Show(
            $form,
            "Sanas is not found as the default device.`nPlease launch Sanas.",
            "Sanas Alert",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        Write-Log "User alerted that Sanas is not detected"
    }

    Write-Log "=== Script execution completed ==="
    exit 0
}
catch {
    # Never let the scheduled task return 1 without logging
    Write-Log "FATAL ERROR: $($_.Exception.Message)"
    Write-Log "=== Script execution completed (with error) ==="
    exit 0
}

#EOF
