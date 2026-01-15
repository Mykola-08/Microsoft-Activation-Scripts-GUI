<#
.SYNOPSIS
    Win11 Style GUI for Microsoft Activation Scripts (MAS)
.DESCRIPTION
    A modern, Fluent-inspired interface for MAS with auto-detection of system status.
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
        $os = Get-CimInstance Win32_OperatingSystem
        $comp = Get-CimInstance Win32_ComputerSystem
        
        $edition = $os.Caption -replace "Microsoft ", ""
        $version = $os.Version
        $build = $os.BuildNumber
        
        # Simple Activation Check (Partial)
        # 1=Licensed
        $license = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like "Windows*" } | Select-Object -First 1
        $status = if ($license.LicenseStatus -eq 1) { "Permanently Activated" } else { "Not Activated / check details" }
        if ($license.LicenseStatus -eq 1 -and $license.GracePeriodRemaining -gt 0) { $status = "Volume/KMS Activated" }

        return @{
            Edition = $edition
            Version = "Build $build"
            Status  = $status
            PCName  = $comp.Name
        }
    } catch {
        return @{ Edition = "Unknown"; Version = ""; Status = "Unknown"; PCName = "Unknown" }
    }
}

$sysInfo = Get-SystemInfo

# --- XAML ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Activation Scripts" Height="600" Width="950"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        Background="#202020" Foreground="#FFFFFF" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    
    <Window.Resources>
        <!-- Win11 Palette -->
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
            <Setter Property="Background" Value="{StaticResource CardBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource CardBorder}"/>
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
                                <Setter TargetName="border" Property="Background" Value="{StaticResource CardHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#303030"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#505050"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Navigation RadioButton Style -->
        <!-- We use RadioButtons for tabs to simulate NavigationView properly -->
        <Style x:Key="NavButtonStyle" TargetType="RadioButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#D0D0D0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="5,2"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="4"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <!-- Active Indicator -->
                                <Rectangle x:Name="indicator" Grid.Column="0" Fill="{StaticResource AccentColor}" Height="16" RadiusX="2" RadiusY="2" Visibility="Collapsed" Margin="0,0,0,0"/>
                                
                                <ContentPresenter Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,0,0,0"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2D2D2D"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#303030"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter TargetName="indicator" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Info Card Style -->
        <Style x:Key="InfoCard" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource CardBackground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource CardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="Margin" Value="0,0,0,15"/>
        </Style>

        <!-- Action Card Style -->
        <Style x:Key="ActionCard" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource CardBackground}"/>
            <Setter Property="Height" Value="Auto"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{StaticResource CardBorder}" BorderThickness="1" CornerRadius="8" Padding="20">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter Grid.Column="0" HorizontalAlignment="Left"/>
                                <TextBlock Grid.Column="1" Text="&#xE76C;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="16" Foreground="{StaticResource AccentColor}" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                         <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource CardHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid Background="{StaticResource NavBackground}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar -->
        <StackPanel Grid.Column="0" Margin="10,20,10,20">
            <!-- App Title -->
            <StackPanel Orientation="Horizontal" Margin="15,0,0,30">
                <Border Width="24" Height="24" CornerRadius="12" Background="{StaticResource AccentColor}" Margin="0,0,10,0">
                    <TextBlock Text="M" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="Bold" Foreground="Black"/>
                </Border>
                <TextBlock Text="MAS" FontSize="18" FontWeight="SemiBold" VerticalAlignment="Center"/>
            </StackPanel>

            <!-- Nav Items -->
            <RadioButton x:Name="navHome" Content="System Overview" Style="{StaticResource NavButtonStyle}" IsChecked="True"/>
            <RadioButton x:Name="navActivators" Content="Activators" Style="{StaticResource NavButtonStyle}"/>
            <RadioButton x:Name="navExtras" Content="Troubleshoot &amp; Extras" Style="{StaticResource NavButtonStyle}"/>
            
        </StackPanel>

        <!-- Content Area -->
        <Border Grid.Column="1" Background="{StaticResource AppBackground}" CornerRadius="10,0,0,0" Padding="40,30">
            <Grid>
                <!-- Home View -->
                <StackPanel x:Name="viewHome" Visibility="Visible">
                    <TextBlock Text="System Overview" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20"/>
                    
                    <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                            <TextBlock Text="Windows Information" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource AccentColor}" Margin="0,0,0,10"/>
                            <Grid Margin="0,5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="&#xE7F8;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="24" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{StaticResource TextSecondary}"/>
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="$($sysInfo.Edition)" FontSize="18" FontWeight="SemiBold"/>
                                    <TextBlock Text="$($sysInfo.Version)" FontSize="14" Foreground="{StaticResource TextSecondary}"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </Border>

                     <Border Style="{StaticResource InfoCard}">
                        <StackPanel>
                            <TextBlock Text="Activation Status" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource AccentColor}" Margin="0,0,0,10"/>
                            <Grid Margin="0,5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="&#xE775;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="24" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{StaticResource TextSecondary}"/>
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="$($sysInfo.Status)" FontSize="18" FontWeight="SemiBold"/>
                                    <TextBlock Text="This information is detected automatically." FontSize="14" Foreground="#666"/>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <TextBlock Text="Quick Actions" FontSize="16" FontWeight="SemiBold" Margin="0,20,0,10"/>
                    <Button x:Name="btnHomeHWID" Style="{StaticResource ActionCard}" Margin="0,5">
                        <StackPanel>
                            <TextBlock Text="Activate Windows" FontSize="16" FontWeight="SemiBold"/>
                            <TextBlock Text="Use HWID method (Recommended)" FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                        </StackPanel>
                    </Button>
                </StackPanel>

                <!-- Activators View -->
                <StackPanel x:Name="viewActivators" Visibility="Collapsed">
                    <TextBlock Text="Activation Methods" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20"/>
                    <ScrollViewer VerticalScrollBarVisibility="Hidden">
                        <StackPanel>
                             <!-- HWID -->
                             <TextBlock Text="Windows 10 / 11" FontSize="14" Foreground="{StaticResource TextSecondary}" Margin="0,10,0,5"/>
                             <Button x:Name="btnHWID" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="HWID Activation" FontSize="16" FontWeight="SemiBold"/>
                                    <TextBlock Text="Permanent digital license. Does not require renewal." FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                                </StackPanel>
                            </Button>

                            <!-- Ohook -->
                            <TextBlock Text="Office" FontSize="14" Foreground="{StaticResource TextSecondary}" Margin="0,10,0,5"/>
                             <Button x:Name="btnOhook" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Ohook Activation" FontSize="16" FontWeight="SemiBold"/>
                                    <TextBlock Text="Permanent Office activation (2013-2024, 365)." FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                                </StackPanel>
                            </Button>

                            <!-- KMS -->
                             <TextBlock Text="Volume / Enterprise" FontSize="14" Foreground="{StaticResource TextSecondary}" Margin="0,10,0,5"/>
                             <Button x:Name="btnKMS" Style="{StaticResource ActionCard}" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Online KMS38" FontSize="16" FontWeight="SemiBold"/>
                                    <TextBlock Text="Activates until 2038 or 180-days (Server/Enterprise)." FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                                </StackPanel>
                            </Button>
                        </StackPanel>
                    </ScrollViewer>
                </StackPanel>

                <!-- Extras View -->
                <StackPanel x:Name="viewExtras" Visibility="Collapsed">
                    <TextBlock Text="Extras" FontSize="26" FontWeight="SemiBold" Margin="0,0,0,20"/>
                    
                    <WrapPanel>
                        <Button x:Name="btnChangeWin" Content="Change Windows Edition" Width="200" Margin="0,0,10,10"/>
                        <Button x:Name="btnChangeOffice" Content="Change Office Edition" Width="200" Margin="0,0,10,10"/>
                        <Button x:Name="btnCheckStatus" Content="Check Activation Status" Width="200" Margin="0,0,10,10"/>
                        <Button x:Name="btnTroubleshoot" Content="Troubleshoot" Width="200" Margin="0,0,10,10"/>
                        <Button x:Name="btnOEM" Content="Extract OEM Folder" Width="200" Margin="0,0,10,10"/>
                    </WrapPanel>
                    
                    <TextBlock Text="Logs &amp; Info" FontSize="16" FontWeight="SemiBold" Margin="0,20,0,10"/>
                    <Border Background="#252525" Padding="15" CornerRadius="4">
                        <TextBlock Text="This GUI wraps the official MAS scripts. When you click an action, a terminal window will open to perform the operation safely and transparently." TextWrapping="Wrap" Foreground="{StaticResource TextSecondary}"/>
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

# --- Find Controls ---
$navHome = $window.FindName("navHome")
$navActivators = $window.FindName("navActivators")
$navExtras = $window.FindName("navExtras")

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

# --- Logic: Navigation ---
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
