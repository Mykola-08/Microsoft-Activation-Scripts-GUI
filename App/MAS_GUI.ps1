<#
.SYNOPSIS
    Win11 Style GUI for Microsoft Activation Scripts (MAS)
.DESCRIPTION
    A modern, Fluent-inspired interface for MAS with auto-detection of system status and dynamic theming.
#>

# --- Self-Elevation ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    Exit
}

# --- Assemblies ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Helper: Get System Info ---
function Get-SystemInfo {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $comp = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        
        $edition = $os.Caption -replace "Microsoft ", ""
        $version = $os.Version
        $build = $os.BuildNumber
        
        # Simple Activation Check (Partial)
        try {
            # 1=Licensed
            $license = Get-CimInstance SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction SilentlyContinue | Select-Object -First 1
            $status = "Unknown / Check manually"

            if ($null -ne $license) {
                if ($license.LicenseStatus -eq 1) { 
                    $status = "Permanently Activated"
                    if ($license.GracePeriodRemaining -gt 0 -and $license.GracePeriodRemaining -lt 40000) { 
                       # KMS usually has grace period in minutes ~ 259200 (180 days)
                       $status = "Volume/KMS Activated" 
                    }
                } else { 
                    $status = "Not Activated" 
                }
            }
        } catch {
            $status = "Detection Failed"
        }

        return @{
            Edition = $edition
            Version = "Build $build"
            Status  = $status
            PCName  = $comp.Name
        }
    } catch {
        return @{ Edition = "Unknown"; Version = "Unknown"; Status = "Unknown"; PCName = "Unknown" }
    }
}

$sysInfo = Get-SystemInfo

# --- Theme Definitions ---
$DarkTheme = @{
    "AppBackground"  = "#202020"
    "NavBackground"  = "#191919"
    "CardBackground" = "#272727"
    "CardBorder"     = "#353535"
    "CardHover"      = "#323232"
    "TextPrimary"    = "#FFFFFF"
    "TextSecondary"  = "#A0A0A0"
    "AccentColor"    = "#60CDFF"
}

$LightTheme = @{
    "AppBackground"  = "#F3F3F3"
    "NavBackground"  = "#EBEBEB"
    "CardBackground" = "#FFFFFF"
    "CardBorder"     = "#E0E0E0"
    "CardHover"      = "#F9F9F9"
    "TextPrimary"    = "#000000"
    "TextSecondary"  = "#666666"
    "AccentColor"    = "#0078D4"
}

# --- XAML ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Activation Scripts" Height="700" Width="1050"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        Background="{DynamicResource AppBackground}" Foreground="{DynamicResource TextPrimary}" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    
    <Window.Resources>
        <!-- Win11 Palette (Placeholders, will be overwritten by PS) -->
        <SolidColorBrush x:Key="AppBackground" Color="#202020"/>
        <SolidColorBrush x:Key="NavBackground" Color="#191919"/>
        <SolidColorBrush x:Key="CardBackground" Color="#272727"/>
        <SolidColorBrush x:Key="CardBorder" Color="#353535"/>
        <SolidColorBrush x:Key="CardHover" Color="#323232"/>
        <SolidColorBrush x:Key="AccentColor" Color="#60CDFF"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#A0A0A0"/>
        
        <!-- Win11 Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource CardBackground}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource CardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource CardHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Navigation RadioButton Style -->
        <Style x:Key="NavButtonStyle" TargetType="RadioButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Margin" Value="5,2"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="4"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Rectangle x:Name="indicator" Grid.Column="0" Fill="{DynamicResource AccentColor}" Height="16" RadiusX="2" RadiusY="2" Visibility="Collapsed" Margin="0,0,0,0"/>
                                <ContentPresenter Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="12,0,0,0"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource CardHover}"/> <!-- Use CardHover for nav hover -->
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource CardHover}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                                <Setter TargetName="indicator" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Info Card Style -->
        <Style x:Key="InfoCard" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource CardBackground}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource CardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="24"/>
            <Setter Property="Margin" Value="0,0,0,15"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="10" ShadowDepth="2" Direction="270" Color="Black" Opacity="0.1"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Action Card Style -->
        <Style x:Key="ActionCard" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource CardBackground}"/>
            <Setter Property="Height" Value="Auto"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Margin" Value="0,0,0,15"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource CardBorder}" BorderThickness="1" CornerRadius="8" Padding="24">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter Grid.Column="0" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Column="1" Text="&#xE76C;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="18" Foreground="{DynamicResource AccentColor}" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                         <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource CardHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid Background="{DynamicResource NavBackground}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="260"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar -->
        <StackPanel Grid.Column="0" Margin="10,20,10,20">
            <!-- App Title -->
            <StackPanel Orientation="Horizontal" Margin="20,0,0,30">
                <Border Width="32" Height="32" CornerRadius="8" Background="{DynamicResource AccentColor}" Margin="0,0,12,0">
                     <TextBlock Text="&#xE770;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="Bold"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="MAS" FontSize="16" FontWeight="Bold" Foreground="{DynamicResource TextPrimary}"/>
                    <TextBlock Text="Activator" FontSize="12" Foreground="{DynamicResource TextSecondary}"/>
                </StackPanel>
            </StackPanel>

            <!-- Nav Items -->
            <RadioButton x:Name="navHome" Content="System Overview" Style="{StaticResource NavButtonStyle}" IsChecked="True"/>
            <RadioButton x:Name="navActivators" Content="Activators" Style="{StaticResource NavButtonStyle}"/>
            <RadioButton x:Name="navExtras" Content="Troubleshoot &amp; Extras" Style="{StaticResource NavButtonStyle}"/>
            
            <!-- Theme Toggle in Sidebar Bottom -->
            <Button x:Name="btnToggleTheme" Content="Switch to Light Theme" Margin="20,50,20,0" FontSize="12" Foreground="{DynamicResource TextSecondary}" Background="Transparent" BorderThickness="0" HorizontalAlignment="Left"/>

        </StackPanel>

        <!-- Content Area -->
        <Border Grid.Column="1" Background="{DynamicResource AppBackground}" CornerRadius="8,0,0,0" Padding="40,30" BorderBrush="{DynamicResource CardBorder}" BorderThickness="1,1,0,0">
            <Grid>
                <!-- Home View -->
                <StackPanel x:Name="viewHome" Visibility="Visible">
                    <TextBlock Text="System Overview" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20" Foreground="{DynamicResource TextPrimary}"/>
                    
                    <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                            <TextBlock Text="Windows Information" FontSize="14" FontWeight="SemiBold" Foreground="{DynamicResource AccentColor}" Margin="0,0,0,10"/>
                            <Grid Margin="0,5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="&#xE7F8;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="32" VerticalAlignment="Center" Margin="0,0,20,0" Foreground="{DynamicResource TextSecondary}"/>
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="$($sysInfo.Edition)" FontSize="20" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                                    <TextBlock Text="$($sysInfo.Version)" FontSize="14" Foreground="{DynamicResource TextSecondary}"/>
                                    <TextBlock Text="$($sysInfo.PCName)" FontSize="12" Foreground="{DynamicResource TextSecondary}" Opacity="0.7"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </Border>

                     <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                            <TextBlock Text="Activation Status" FontSize="14" FontWeight="SemiBold" Foreground="{DynamicResource AccentColor}" Margin="0,0,0,10"/>
                            <Grid Margin="0,5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="&#xE775;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="32" VerticalAlignment="Center" Margin="0,0,20,0" Foreground="{DynamicResource TextSecondary}"/>
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="$($sysInfo.Status)" FontSize="20" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                                    <TextBlock Text="Automatic system detection" FontSize="14" Foreground="{DynamicResource TextSecondary}"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <TextBlock Text="Recommended Action" FontSize="16" FontWeight="SemiBold" Margin="0,20,0,10" Foreground="{DynamicResource TextPrimary}"/>
                    <Button x:Name="btnHomeHWID" Style="{StaticResource ActionCard}" Margin="0,5">
                        <StackPanel>
                            <TextBlock Text="Activate Windows" FontSize="17" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                            <TextBlock Text="Uses the 'HWID' method to generate a permanent digital license." FontSize="14" Foreground="{DynamicResource TextSecondary}" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Button>
                </StackPanel>

                <!-- Activators View -->
                <StackPanel x:Name="viewActivators" Visibility="Collapsed">
                    <TextBlock Text="Activation Methods" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20" Foreground="{DynamicResource TextPrimary}"/>
                    
                    <TextBlock Text="Windows 10 / 11" FontSize="14" FontWeight="SemiBold" Foreground="{DynamicResource TextSecondary}" Margin="0,10,0,5" Opacity="0.8"/>
                    <Button x:Name="btnHWID" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="HWID Activation" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                            <TextBlock Text="Digital License. Permanent. Best for personal PCs." FontSize="13" Foreground="{DynamicResource TextSecondary}"/>
                        </StackPanel>
                    </Button>

                    <TextBlock Text="Microsoft Office" FontSize="14" FontWeight="SemiBold" Foreground="{DynamicResource TextSecondary}" Margin="0,10,0,5" Opacity="0.8"/>
                        <Button x:Name="btnOhook" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="Ohook Activation" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                            <TextBlock Text="Permanent activation for Office 2013-2024 &amp; 365." FontSize="13" Foreground="{DynamicResource TextSecondary}"/>
                        </StackPanel>
                    </Button>

                    <TextBlock Text="Enterprise / Server" FontSize="14" FontWeight="SemiBold" Foreground="{DynamicResource TextSecondary}" Margin="0,10,0,5" Opacity="0.8"/>
                        <Button x:Name="btnKMS" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="Online KMS38" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                            <TextBlock Text="Activate until 2038 or 180-days (KMS). Good for Servers." FontSize="13" Foreground="{DynamicResource TextSecondary}"/>
                        </StackPanel>
                    </Button>
                </StackPanel>

                <!-- Extras View -->
                <StackPanel x:Name="viewExtras" Visibility="Collapsed">
                    <TextBlock Text="Extras" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20" Foreground="{DynamicResource TextPrimary}"/>
                    
                    <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                           <TextBlock Text="Troubleshooting" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,15"/>
                           <WrapPanel>
                                <Button x:Name="btnTroubleshoot" Content="Run Troubleshooter" Width="200" Margin="0,0,10,10"/>
                                <Button x:Name="btnCheckStatus" Content="Check Activation Status" Width="200" Margin="0,0,10,10"/>
                            </WrapPanel>
                        </StackPanel>
                    </Border>
                    
                    <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                           <TextBlock Text="Edition Management" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}" Margin="0,0,0,15"/>
                           <WrapPanel>
                                <Button x:Name="btnChangeWin" Content="Change Windows Edition" Width="200" Margin="0,0,10,10"/>
                                <Button x:Name="btnChangeOffice" Content="Change Office Edition" Width="200" Margin="0,0,10,10"/>
                                <Button x:Name="btnOEM" Content="Extract OEM Folder" Width="200" Margin="0,0,10,10"/>
                            </WrapPanel>
                        </StackPanel>
                    </Border>
                    
                    <Border Background="{DynamicResource CardBackground}" Padding="15" CornerRadius="4" Margin="0,20,0,0" BorderBrush="{DynamicResource CardBorder}" BorderThickness="1">
                        <StackPanel>
                             <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="&#xE946;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="{DynamicResource AccentColor}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                <TextBlock Text="How it works" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimary}"/>
                             </StackPanel>
                             <TextBlock Text="This GUI acts as a launcher for the official MAS scripts. When you select an option, it will open a secure terminal window to perform the activation processes transparently. No hidden background tasks." TextWrapping="Wrap" Foreground="{DynamicResource TextSecondary}" FontSize="13" LineHeight="20"/>
                        </StackPanel>
                    </Border>
                </StackPanel>

            </Grid>
        </Border>
    </Grid>
</Window>
"@

# --- Load & Parse ---
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Failed to load XAML: $_"
    Exit
}

# --- Apply Theme Logic ---
$isDarkTheme = $true

function Apply-Theme {
    param($Theme)
    foreach ($key in $Theme.Keys) {
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($Theme[$key])
        $brush = New-Object System.Windows.Media.SolidColorBrush($color)
        if ($window.Resources.Contains($key)) {
            $window.Resources[$key].Color = $color
        }
    }
}

function Toggle-Theme {
    $global:isDarkTheme = -not $global:isDarkTheme
    if ($global:isDarkTheme) {
        Apply-Theme $DarkTheme
        $btnToggleTheme.Content = "Switch to Light Theme"
    } else {
        Apply-Theme $LightTheme
        $btnToggleTheme.Content = "Switch to Dark Theme"
    }
}

# Apply Initial Theme
Apply-Theme $DarkTheme

# --- Find Controls ---
$navHome = $window.FindName("navHome")
$navActivators = $window.FindName("navActivators")
$navExtras = $window.FindName("navExtras")
$btnToggleTheme = $window.FindName("btnToggleTheme")

$viewHome = $window.FindName("viewHome")
$viewActivators = $window.FindName("viewActivators")
$viewExtras = $window.FindName("viewExtras")

$btnHomeHWID = $window.FindName("btnHomeHWID")
$btnHWID = $window.FindName("btnHWID")
$btnOhook = $window.FindName("btnOhook")
$btnKMS = $window.FindName("btnKMS")

$btnChangeWin = $window.FindName("btnChangeWin")
$btnChangeOffice = $window.FindName("btnChangeOffice")
$btnCheckStatus = $window.FindName("btnCheckStatus")
$btnTroubleshoot = $window.FindName("btnTroubleshoot")
$btnOEM = $window.FindName("btnOEM")

# --- Events ---
$btnToggleTheme.Add_Click({ Toggle-Theme })

function Switch-View {
    param($view)
    $viewHome.Visibility = "Collapsed"
    $viewActivators.Visibility = "Collapsed"
    $viewExtras.Visibility = "Collapsed"
    $view.Visibility = "Visible"
}

$navHome.Add_Checked({ Switch-View $viewHome })
$navActivators.Add_Checked({ Switch-View $viewActivators })
$navExtras.Add_Checked({ Switch-View $viewExtras })

# --- Logic: Execution ---
function Invoke-MASScript {
    param([string]$Path)
    $fullPath = Join-Path $PSScriptRoot $Path
    if (Test-Path $fullPath) {
        Start-Process "cmd.exe" -ArgumentList "/c `"$fullPath`""
    } else {
        [System.Windows.MessageBox]::Show("Script not found at:`n$fullPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

$btnHomeHWID.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Activators\HWID_Activation.cmd" })
$btnHWID.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Activators\HWID_Activation.cmd" })
$btnOhook.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Activators\Ohook_Activation_AIO.cmd" })
$btnKMS.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Activators\Online_KMS_Activation.cmd" })

$btnChangeWin.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Change_Windows_Edition.cmd" })
$btnChangeOffice.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Change_Office_Edition.cmd" })
$btnCheckStatus.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Check_Activation_Status.cmd" })
$btnTroubleshoot.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Troubleshoot.cmd" })
$btnOEM.Add_Click({ Invoke-MASScript "MAS\Separate-Files-Version\Extract_OEM_Folder.cmd" })

# --- Run ---
$window.ShowDialog() | Out-Null
