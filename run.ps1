# AppCleaner Memory edition. Copy the full paste-ready text into Windows PowerShell 5.1.
# Opening the UI does not install/extract a package or write an application log.
# Destructive operations still require explicit confirmation in the UI.
# Windows, PowerShell, .NET and official uninstallers may keep normal system records.
# Windows PowerShell 5.1 / WPF. No cleanup occurs before explicit confirmation.
[CmdletBinding()]
param([switch]$Preview, [switch]$NoShow)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Use Windows PowerShell 5.1: open powershell.exe -NoProfile -STA, then paste this script.'
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
# This small embedded library contains only the two documented Windows helpers below:
# window sizing and resolving the current user's Downloads/Pictures locations.
# It is loaded from bytes. It is not downloaded or extracted to a DLL on disk.
if (-not ('AppCleaner.MemoryNativeFrame' -as [type])) {
    $interopBytes=[Convert]::FromBase64String('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ZGxlZABpZAB0b2tlbgBwYXRoAFN5c3RlbS5SdW50aW1lLkludGVyb3BTZXJ2aWNlcwBPdXRBdHRyaWJ1dGUALmN0b3IAbmFtZQBTeXN0ZW0uUnVudGltZS5Db21waWxlclNlcnZpY2VzAENvbXBpbGF0aW9uUmVsYXhhdGlvbnNBdHRyaWJ1dGUAUnVudGltZUNvbXBhdGliaWxpdHlBdHRyaWJ1dGUATWVtb3J5SW50ZXJvcC0wNGY2NzRlOWQ4OWE0ODlhYTQ2YTZlZmJkZGI3MDdkMgBEbGxJbXBvcnRBdHRyaWJ1dGUAdXNlcjMyLmRsbAA8PmNfX0Rpc3BsYXlDbGFzczEARXZlbnRBcmdzADxBdHRhY2g+Yl9fMABwYXJhbTAAcGFyYW0xAFN5c3RlbS5XaW5kb3dzLkludGVyb3AAV2luZG93SW50ZXJvcEhlbHBlcgBnZXRfSGFuZGxlAFByZXNlbnRhdGlvbkNvcmUASHduZFNvdXJjZQBGcm9tSHduZABId25kU291cmNlSG9vawBBZGRIb29rAEV2ZW50SGFuZGxlcgBhZGRfU291cmNlSW5pdGlhbGl6ZWQASW50UHRyAFplcm8AVHlwZQBSdW50aW1lVHlwZUhhbmRsZQBHZXRUeXBlRnJvbUhhbmRsZQBNYXJzaGFsAFNpemVPZgBvcF9FcXVhbGl0eQBQdHJUb1N0cnVjdHVyZQBTdHJ1Y3R1cmVUb1B0cgBTdHJ1Y3RMYXlvdXRBdHRyaWJ1dGUATGF5b3V0S2luZABDb21waWxlckdlbmVyYXRlZEF0dHJpYnV0ZQBzaGVsbDMyLmRsbABTdHJpbmcAQXJndW1lbnRFeGNlcHRpb24AVGhyb3dFeGNlcHRpb25Gb3JIUgBQdHJUb1N0cmluZ1VuaQBvcF9JbmVxdWFsaXR5AEZyZWVDb1Rhc2tNZW0AABNEAG8AdwBuAGwAbwBhAGQAcwAASTMANwA0AEQARQAyADkAMAAtADEAMgAzAEYALQA0ADUANgA1AC0AOQAxADYANAAtADMAOQBDADQAOQAyADUARQA0ADYANwBCAAERUABpAGMAdAB1AHIAZQBzAABJMwAzAEUAMgA4ADEAMwAwAC0ANABFADEARQAtADQANgA3ADYALQA4ADMANQBBAC0AOQA4ADMAOQA1AEMAMwBCAEMA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')
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $actualHash=([BitConverter]::ToString($sha.ComputeHash($interopBytes))).Replace('-','') } finally { $sha.Dispose() }
    if ($actualHash -ne '6D24AA5073FF188DADBE5A691D12FAC98C802E96EFEAC1187E28EAD5E2FB00BC') { throw 'The embedded Windows helper is damaged. Copy the complete script again.' }
    [void][Reflection.Assembly]::Load($interopBytes)
}
<# Full C# source of the embedded helper, included for inspection:
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace AppCleaner
{
    // Keep a borderless window inside the current monitor's work area when maximized.
    public static class MemoryNativeFrame
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)]
        private struct Rect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)]
        private struct MinMaxInfo
        {
            public Point Reserved, MaxSize, MaxPosition, MinTrackSize, MaxTrackSize;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo { public int Size; public Rect Monitor, Work; public uint Flags; }
        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr window, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

        public static void Attach(Window window)
        {
            window.SourceInitialized += delegate
            {
                var source = HwndSource.FromHwnd(new WindowInteropHelper(window).Handle);
                if (source != null) source.AddHook(HandleMessage);
            };
        }

        private static IntPtr HandleMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (message != 0x0024) return IntPtr.Zero; // WM_GETMINMAXINFO
            var monitor = MonitorFromWindow(window, 2);
            var info = new MonitorInfo { Size = Marshal.SizeOf(typeof(MonitorInfo)) };
            if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref info)) return IntPtr.Zero;
            var bounds = (MinMaxInfo)Marshal.PtrToStructure(lParam, typeof(MinMaxInfo));
            bounds.MaxPosition.X = info.Work.Left - info.Monitor.Left;
            bounds.MaxPosition.Y = info.Work.Top - info.Monitor.Top;
            bounds.MaxSize.X = info.Work.Right - info.Work.Left;
            bounds.MaxSize.Y = info.Work.Bottom - info.Work.Top;
            Marshal.StructureToPtr(bounds, lParam, false);
            handled = true;
            return IntPtr.Zero;
        }
    }
}

namespace AppCleaner
{
    public static class MemoryKnownFolders
    {
        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(ref Guid id, uint flags, IntPtr token, out IntPtr path);

        public static string Get(string name)
        {
            Guid id;
            if (name == "Downloads") id = new Guid("374DE290-123F-4565-9164-39C4925E467B");
            else if (name == "Pictures") id = new Guid("33E28130-4E1E-4676-835A-98395C3BC3BB");
            else throw new ArgumentException("Unsupported folder.");
            IntPtr result = IntPtr.Zero;
            try
            {
                // Resolve the current user's real folder, including Windows redirection.
                // DONT_VERIFY avoids creating missing directories or touching their contents.
                int status = SHGetKnownFolderPath(ref id, 0x4000, IntPtr.Zero, out result);
                if (status < 0) Marshal.ThrowExceptionForHR(status);
                return Marshal.PtrToStringUni(result);
            }
            finally { if (result != IntPtr.Zero) Marshal.FreeCoTaskMem(result); }
        }
    }
}

#>
$script:busy = $false
$script:isPreview = [bool]$Preview
$script:dialogAction = $null
$script:previousFocus = $null
$script:logFile = $null
$script:checkedLocations = New-Object System.Collections.ArrayList
$script:previewRemoved = @{}
$script:cleanupSucceeded = $false

[xml]$markup = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="AppCleaner - Memory" Width="1180" Height="850" MinWidth="1000" MinHeight="740" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="CanResizeWithGrip" Background="#0B1020" Foreground="#F1F5FF" FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True">
 <WindowChrome.WindowChrome>
  <WindowChrome CaptionHeight="0" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" NonClientFrameEdges="None" UseAeroCaptionButtons="False"/>
 </WindowChrome.WindowChrome>
 <Window.Resources>
  <SolidColorBrush x:Key="Muted" Color="#909DB9"/>
  <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
  <Style x:Key="Action" TargetType="Button">
   <Setter Property="Foreground" Value="#DDE6FA"/><Setter Property="Background" Value="#1B2740"/><Setter Property="BorderBrush" Value="#334361"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="18,11"/><Setter Property="Cursor" Value="Hand"/><Setter Property="HorizontalContentAlignment" Value="Center"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
    <Border x:Name="Surface" CornerRadius="10" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" RenderTransformOrigin=".5,.5"><Border.RenderTransform><TranslateTransform/></Border.RenderTransform><ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/></Border>
    <ControlTemplate.Triggers>
     <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Surface" Property="BorderBrush" Value="#7CE8DC"/><Trigger.EnterActions><BeginStoryboard><Storyboard><DoubleAnimation Storyboard.TargetName="Surface" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)" To="-2" Duration="0:0:0.14"/></Storyboard></BeginStoryboard></Trigger.EnterActions><Trigger.ExitActions><BeginStoryboard><Storyboard><DoubleAnimation Storyboard.TargetName="Surface" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)" To="0" Duration="0:0:0.18"/></Storyboard></BeginStoryboard></Trigger.ExitActions></Trigger>
     <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Surface" Property="BorderBrush" Value="#A7FFF1"/></Trigger>
     <Trigger Property="IsPressed" Value="True"><Setter TargetName="Surface" Property="Opacity" Value=".65"/></Trigger>
     <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Surface" Property="Opacity" Value=".35"/></Trigger>
    </ControlTemplate.Triggers>
   </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="Button" BasedOn="{StaticResource Action}"/>
  <Style x:Key="AppCard" TargetType="CheckBox">
   <Setter Property="Cursor" Value="Hand"/><Setter Property="Foreground" Value="#F1F5FF"/><Setter Property="Margin" Value="0,0,12,12"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="CheckBox">
    <Border x:Name="Card" Background="#141E32" BorderBrush="#293650" BorderThickness="1" CornerRadius="16" Padding="16" RenderTransformOrigin=".5,.5">
     <Border.RenderTransform><TranslateTransform/></Border.RenderTransform>
     <Grid><ContentPresenter Margin="0,0,28,0"/><Border x:Name="TickBox" Width="22" Height="22" CornerRadius="7" BorderBrush="#4E607E" BorderThickness="1" HorizontalAlignment="Right" VerticalAlignment="Top"><TextBlock x:Name="Tick" Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" Foreground="#092426" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/></Border></Grid>
    </Border>
    <ControlTemplate.Triggers>
     <Trigger Property="IsChecked" Value="True"><Setter TargetName="Card" Property="Background" Value="#142C39"/><Setter TargetName="Card" Property="BorderBrush" Value="#51D7C5"/><Setter TargetName="TickBox" Property="Background" Value="#70E4D1"/><Setter TargetName="TickBox" Property="BorderBrush" Value="#70E4D1"/><Setter TargetName="Tick" Property="Visibility" Value="Visible"/></Trigger>
     <Trigger Property="IsMouseOver" Value="True"><Trigger.EnterActions><BeginStoryboard><Storyboard><DoubleAnimation Storyboard.TargetName="Card" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)" To="-4" Duration="0:0:0.20"><DoubleAnimation.EasingFunction><CubicEase EasingMode="EaseOut"/></DoubleAnimation.EasingFunction></DoubleAnimation></Storyboard></BeginStoryboard></Trigger.EnterActions><Trigger.ExitActions><BeginStoryboard><Storyboard><DoubleAnimation Storyboard.TargetName="Card" Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)" To="0" Duration="0:0:0.20"/></Storyboard></BeginStoryboard></Trigger.ExitActions></Trigger>
     <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Card" Property="BorderBrush" Value="White"/><Setter TargetName="Card" Property="BorderThickness" Value="2"/></Trigger>
     <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Card" Property="Opacity" Value=".48"/></Trigger>
    </ControlTemplate.Triggers>
   </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="CheckBox"><Setter Property="Foreground" Value="#E1E9FA"/><Setter Property="VerticalContentAlignment" Value="Center"/><Setter Property="Cursor" Value="Hand"/></Style>
  <Style x:Key="ScrollThumb" TargetType="Thumb"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Thumb"><Border x:Name="ThumbSurface" Background="#344761" CornerRadius="4" Margin="2,3"/><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ThumbSurface" Property="Background" Value="#6A9DAB"/></Trigger><Trigger Property="IsDragging" Value="True"><Setter TargetName="ThumbSurface" Property="Background" Value="#78D9CC"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
  <Style x:Key="ScrollTrackButton" TargetType="RepeatButton"><Setter Property="Focusable" Value="False"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate></Setter.Value></Setter></Style>
  <Style TargetType="ScrollBar"><Setter Property="Width" Value="10"/><Setter Property="Background" Value="Transparent"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ScrollBar"><Grid Background="{TemplateBinding Background}"><Track x:Name="PART_Track" IsDirectionReversed="True"><Track.DecreaseRepeatButton><RepeatButton Style="{StaticResource ScrollTrackButton}" Command="ScrollBar.PageUpCommand"/></Track.DecreaseRepeatButton><Track.Thumb><Thumb Style="{StaticResource ScrollThumb}"/></Track.Thumb><Track.IncreaseRepeatButton><RepeatButton Style="{StaticResource ScrollTrackButton}" Command="ScrollBar.PageDownCommand"/></Track.IncreaseRepeatButton></Track></Grid></ControlTemplate></Setter.Value></Setter></Style>
 </Window.Resources>
 <Border BorderBrush="#2A3650" BorderThickness="1">
 <Grid>
  <Grid.RowDefinitions><RowDefinition Height="54"/><RowDefinition Height="*"/><RowDefinition Height="33"/></Grid.RowDefinitions>
  <Grid x:Name="TitleBar" Background="#0D1425">
   <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="22,0"><Border Background="#67DECF" Width="25" Height="25" CornerRadius="7"><TextBlock Text="&#xE80F;" FontFamily="Segoe MDL2 Assets" Foreground="#0B282C" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="13"/></Border><TextBlock Text="APPCLEANER" FontWeight="SemiBold" FontSize="12" Margin="11,0,0,0" VerticalAlignment="Center"/><TextBlock Text="/  WINDOWS" Foreground="#667791" FontSize="10" Margin="14,0" VerticalAlignment="Center"/></StackPanel>
   <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,8,0"><Button x:Name="MinimizeButton" Content="&#xE921;" FontFamily="Segoe MDL2 Assets" Width="42" Background="Transparent" BorderThickness="0" Padding="10" AutomationProperties.Name="Minimize"/><Button x:Name="MaximizeButton" Content="&#xE922;" FontFamily="Segoe MDL2 Assets" Width="42" Background="Transparent" BorderThickness="0" Padding="10" AutomationProperties.Name="Maximize"/><Button x:Name="CloseButton" Content="&#xE8BB;" FontFamily="Segoe MDL2 Assets" Width="42" Background="Transparent" BorderThickness="0" Padding="10" AutomationProperties.Name="Close"/></StackPanel>
  </Grid>
  <Grid x:Name="Workspace" Grid.Row="1">
   <Grid.ColumnDefinitions><ColumnDefinition Width="204"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <Border BorderBrush="#1C273D" BorderThickness="0,0,1,0" Background="#0D1425">
    <Grid Margin="18,30,18,20"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
     <StackPanel><TextBlock Text="YOUR WORKSPACE" FontSize="10" Foreground="#657795" FontWeight="SemiBold" Margin="13,0,0,20"/>
      <Button x:Name="CleanNav" HorizontalContentAlignment="Left" Background="#1C3443" BorderBrush="#34586A" Padding="14,13" Margin="0,0,0,10"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE80F;" FontFamily="Segoe MDL2 Assets" Foreground="#78E9D8" Margin="0,2,12,0"/><TextBlock Text="Clean apps" FontWeight="SemiBold"/></StackPanel></Button>
      <Button x:Name="FoldersNav" HorizontalContentAlignment="Left" Background="Transparent" BorderThickness="0" Padding="14,13" Margin="0,0,0,8"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE8B7;" FontFamily="Segoe MDL2 Assets" Margin="0,2,12,0"/><TextBlock Text="Folder cleanup"/></StackPanel></Button>
      <Button x:Name="ActivityNav" HorizontalContentAlignment="Left" Background="Transparent" BorderThickness="0" Padding="14,13"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE9D9;" FontFamily="Segoe MDL2 Assets" Margin="0,2,12,0"/><TextBlock Text="Activity log"/></StackPanel></Button>
     </StackPanel>
     <StackPanel Grid.Row="2"><Border Background="#131F32" CornerRadius="13" Padding="16" Margin="0,0,0,16"><StackPanel><TextBlock Text="&#xE753;" FontFamily="Segoe MDL2 Assets" Foreground="#82B7FF" FontSize="22" Margin="0,0,0,12"/><TextBlock Text="Local cleanup" FontWeight="SemiBold"/><TextBlock Text="Cloud accounts stay yours. Review local data before you delete." FontSize="11" Foreground="{StaticResource Muted}" LineHeight="18" Margin="0,7,0,0"/></StackPanel></Border><Button x:Name="HelpButton" Content="How it works  &#x2197;" Background="Transparent" BorderThickness="0" Foreground="#98ACC9" FontSize="12"/><TextBlock Text="APPCLEANER  /  MEMORY" Foreground="#526581" FontSize="9" HorizontalAlignment="Center" Margin="0,15,0,0"/></StackPanel>
    </Grid>
   </Border>
   <Grid Grid.Column="1" Margin="32,26,32,20">
    <ScrollViewer x:Name="CleanPage" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,8,0">
    <StackPanel x:Name="PageContent" Opacity="0"><StackPanel.RenderTransform><TranslateTransform Y="14"/></StackPanel.RenderTransform>
     <Grid Margin="0,0,0,18"><StackPanel><TextBlock Text="A little less clutter." FontSize="29" FontWeight="SemiBold"/><TextBlock Text="Choose your apps. Make room for a fresh start." Foreground="{StaticResource Muted}" Margin="0,7,0,0" FontSize="13"/></StackPanel><Border HorizontalAlignment="Right" VerticalAlignment="Top" Background="#172B33" CornerRadius="12" Padding="12,6" Margin="0,7,0,0"><StackPanel Orientation="Horizontal"><Ellipse Width="6" Height="6" Fill="#6FE2C9" Margin="0,0,7,0"/><TextBlock x:Name="ModeLabel" Text="MEMORY SESSION" Foreground="#8CE6D6" FontSize="9" FontWeight="SemiBold"/></StackPanel></Border></Grid>
     <Border CornerRadius="18" Padding="25,20" Margin="0,0,0,24" ClipToBounds="True">
      <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="1,1"><GradientStop Color="#203B5C" Offset="0"/><GradientStop Color="#172B42" Offset=".6"/><GradientStop Color="#173E46" Offset="1"/></LinearGradientBrush></Border.Background>
      <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="155"/></Grid.ColumnDefinitions><StackPanel VerticalAlignment="Center"><TextBlock Text="A CLEANER SPACE" Foreground="#82EBDD" FontSize="10" FontWeight="Bold"/><TextBlock Text="Everything, in one flow." FontSize="20" FontWeight="SemiBold" Margin="0,9,0,8"/><TextBlock x:Name="AllCleanupScope" Text="4 checked apps + Downloads contents + Pictures images.&#x0a;Scan and review everything before confirming." Foreground="#AEC2D9" FontSize="12" LineHeight="19"/><Button x:Name="AllCleanupButton" AutomationProperties.Name="Clean everything" HorizontalAlignment="Left" Background="#78E8D4" BorderBrush="#78E8D4" Foreground="#102C34" Padding="18,10" Margin="0,12,0,0"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE945;" FontFamily="Segoe MDL2 Assets" Margin="0,3,10,0"/><TextBlock Text="&#x0E25;&#x0E49;&#x0E32;&#x0E07;&#x0E17;&#x0E31;&#x0E49;&#x0E07;&#x0E2B;&#x0E21;&#x0E14;" FontSize="15" FontWeight="Bold"/><TextBlock Text="&#x2192;" Margin="18,0,0,0" FontSize="17"/></StackPanel></Button></StackPanel>
       <Grid Grid.Column="1" Width="125" Height="122" VerticalAlignment="Center"><Ellipse Width="112" Height="112" Stroke="#365F73" StrokeThickness="1"/><Ellipse Width="86" Height="86" Stroke="#497D8B" StrokeThickness="1"/><Border Width="60" Height="60" CornerRadius="20" Background="#72E5D1" RenderTransformOrigin=".5,.5"><Border.RenderTransform><RotateTransform Angle="-10"/></Border.RenderTransform><TextBlock Text="&#xE777;" FontFamily="Segoe MDL2 Assets" Foreground="#143B48" FontSize="29" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><Canvas x:Name="Orbit" Width="112" Height="112" RenderTransformOrigin=".5,.5"><Canvas.RenderTransform><RotateTransform/></Canvas.RenderTransform><Ellipse Width="9" Height="9" Fill="#79E8D6" Canvas.Left="51" Canvas.Top="-4"/><Ellipse Width="5" Height="5" Fill="#8BAAFF" Canvas.Left="10" Canvas.Top="91"/></Canvas></Grid>
      </Grid>
     </Border>
     <Grid Margin="0,0,0,14"><StackPanel Orientation="Horizontal"><TextBlock Text="Select apps" FontSize="16" FontWeight="SemiBold"/><TextBlock x:Name="SelectionCount" Text="4 selected" FontSize="11" Foreground="#83A3BF" Margin="12,4,0,0"/></StackPanel><Button x:Name="SelectAllButton" Content="Deselect all" HorizontalAlignment="Right" Padding="4,0" Background="Transparent" BorderThickness="0" Foreground="#80DCD6" FontSize="11"/></Grid>
     <UniformGrid x:Name="AppGrid" Columns="2" Margin="0,0,-12,6">
      <CheckBox x:Name="ChromeCard" IsChecked="True" Style="{StaticResource AppCard}" AutomationProperties.Name="Google Chrome">
       <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="56"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Grid Width="42" Height="42" HorizontalAlignment="Left" VerticalAlignment="Center"><Ellipse Fill="#F0C558"/><Path Data="M21,21 L0,21 A21,21 0 0 1 39,10 Z" Fill="#F57770"/><Path Data="M21,21 L39,10 A21,21 0 0 1 21,42 L0,21 Z" Fill="#63C997"/><Ellipse Width="20" Height="20" Fill="#70B7F2" Stroke="#D9F1FF" StrokeThickness="3"/></Grid>
        <StackPanel Grid.Column="1"><TextBlock Text="Google Chrome" FontSize="14" FontWeight="SemiBold"/><TextBlock Text="Browser &amp; profile data" Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,0,9"/><TextBlock x:Name="ChromeStatus" Text="Checking installation..." Foreground="#79C7BA" FontSize="10"/></StackPanel>
       </Grid>
      </CheckBox>
      <CheckBox x:Name="LineCard" IsChecked="True" Style="{StaticResource AppCard}" AutomationProperties.Name="LINE">
       <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="56"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Width="42" Height="42" Background="#29C96C" CornerRadius="12" HorizontalAlignment="Left" VerticalAlignment="Center"><Grid><Border Background="White" CornerRadius="10" Margin="7,9,7,10"><TextBlock Text="LINE" Foreground="#24B860" FontSize="9" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><Path Data="M16,30 L16,35 L22,29 Z" Fill="White"/></Grid></Border>
        <StackPanel Grid.Column="1"><TextBlock Text="LINE" FontSize="14" FontWeight="SemiBold"/><TextBlock Text="Messages &amp; local settings" Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,0,9"/><TextBlock x:Name="LineStatus" Text="Checking installation..." Foreground="#79C7BA" FontSize="10"/></StackPanel>
       </Grid>
      </CheckBox>
      <CheckBox x:Name="TelegramCard" IsChecked="True" Style="{StaticResource AppCard}" AutomationProperties.Name="Telegram Desktop">
       <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="56"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Width="42" Height="42" Background="#45A9E0" CornerRadius="14" HorizontalAlignment="Left" VerticalAlignment="Center"><Path Data="M8,20 L34,10 L28,33 L20,26 L15,30 L16,23 L29,15 L19,24 Z" Fill="White"/></Border>
        <StackPanel Grid.Column="1"><TextBlock Text="Telegram Desktop" FontSize="14" FontWeight="SemiBold"/><TextBlock Text="Sessions &amp; local chats" Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,0,9"/><TextBlock x:Name="TelegramStatus" Text="Checking installation..." Foreground="#79C7BA" FontSize="10"/></StackPanel>
       </Grid>
      </CheckBox>
      <CheckBox x:Name="DiscordCard" IsChecked="True" Style="{StaticResource AppCard}" AutomationProperties.Name="Discord">
       <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="56"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Width="42" Height="42" Background="#5865F2" CornerRadius="14" HorizontalAlignment="Left" VerticalAlignment="Center"><Grid><Path Data="M11,13 C14,11 16,11 17,11 L18,14 C20,13 22,13 24,14 L25,11 C28,11 30,12 31,13 C34,18 35,23 34,28 C31,30 28,31 26,31 L24,28 C22,29 20,29 18,28 L16,31 C13,31 10,30 8,28 C7,23 8,18 11,13 Z" Fill="White"/><Ellipse Width="4" Height="6" Fill="#5865F2" Margin="15,19,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/><Ellipse Width="4" Height="6" Fill="#5865F2" Margin="24,19,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/></Grid></Border>
        <StackPanel Grid.Column="1"><TextBlock Text="Discord" FontSize="14" FontWeight="SemiBold"/><TextBlock Text="Cache &amp; local settings" Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,0,9"/><TextBlock x:Name="DiscordStatus" Text="Checking installation..." Foreground="#79C7BA" FontSize="10"/></StackPanel>
       </Grid>
      </CheckBox>
     </UniformGrid>
     <Border Background="#2A252B" BorderBrush="#504039" BorderThickness="1" CornerRadius="10" Padding="14,11" Margin="0,0,0,18"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="25"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><TextBlock Text="&#xE7BA;" FontFamily="Segoe MDL2 Assets" Foreground="#E8BE86" Margin="0,3,0,0"/><TextBlock Grid.Column="1" Text="Local profiles, saved sign-ins, bookmarks and local chats may be permanently deleted. Back up what matters. Deleted data skips the Recycle Bin." Foreground="#CEBBA8" FontSize="11" LineHeight="17"/></Grid></Border>
     <Grid Margin="0,0,0,20"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="225"/></Grid.ColumnDefinitions><StackPanel VerticalAlignment="Center"><TextBlock x:Name="ReadyTitle" Text="Ready when you are." FontSize="15" FontWeight="SemiBold"/><TextBlock x:Name="ReadyDetail" Text="Review the selected apps before confirming." Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,12,0"/></StackPanel><Button x:Name="QuickButton" Grid.Column="1" Background="#78E8D4" BorderBrush="#78E8D4" Foreground="#102C34" Padding="20,14"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE945;" FontFamily="Segoe MDL2 Assets" Margin="0,2,10,0"/><TextBlock Text="Quick clean" FontWeight="Bold" FontSize="14"/><TextBlock Text="&#x2192;" FontSize="17" Margin="18,-2,0,0"/></StackPanel></Button></Grid>
     <Border Background="#141E30" CornerRadius="12" Padding="17,14"><StackPanel><TextBlock Text="Prefer to go step by step?" Foreground="#91A4C0" FontSize="11" Margin="0,0,0,11"/><UniformGrid Columns="3"><Button x:Name="UninstallButton" Content="1  Uninstall apps" Margin="0,0,8,0" FontSize="11" Padding="10,10"/><Button x:Name="ScanButton" Content="2  Scan leftovers" Margin="0,0,8,0" FontSize="11" Padding="10,10"/><Button x:Name="DeleteButton" Content="3  Delete checked" FontSize="11" Padding="10,10" IsEnabled="False"/></UniformGrid></StackPanel></Border>
     <Border x:Name="ResultsPanel" Visibility="Collapsed" Background="#141E30" CornerRadius="12" Padding="18" Margin="0,15,0,0"><StackPanel><TextBlock x:Name="ResultsTitle" Text="Review leftover locations" FontWeight="SemiBold" Margin="0,0,0,6"/><TextBlock Text="Select only the locations you want to delete. Nothing is selected automatically." FontSize="11" Foreground="{StaticResource Muted}" Margin="0,0,0,12"/><ItemsControl x:Name="ResultsList"/></StackPanel></Border>
    </StackPanel></ScrollViewer>
    <Grid x:Name="FoldersPage" Visibility="Collapsed">
     <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
     <StackPanel Margin="0,0,0,22"><TextBlock Text="Your files. Your call." FontSize="29" FontWeight="SemiBold"/><TextBlock Text="Scan a folder, review its contents, then choose to delete." Foreground="{StaticResource Muted}" Margin="0,8,0,0"/></StackPanel>
     <UniformGrid Grid.Row="1" Columns="2" Margin="0,0,-14,18">
      <Border Background="#172B40" CornerRadius="16" BorderBrush="#2F506D" BorderThickness="1" Padding="20" Margin="0,0,14,0"><StackPanel><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE896;" FontFamily="Segoe MDL2 Assets" FontSize="24" Foreground="#8DC8FF" Margin="0,0,12,0"/><TextBlock Text="Downloads" FontSize="20" FontWeight="SemiBold"/></StackPanel><TextBlock Text="All files &amp; subfolders" Foreground="#A7C0D8" Margin="0,12,0,5"/><TextBlock Text="Keeps the Downloads folder. No AppCleaner installation is needed." Foreground="{StaticResource Muted}" FontSize="11" MinHeight="32"/><Button x:Name="DownloadsScanButton" Content="Scan Downloads  &#x2192;" Margin="0,16,0,0" HorizontalAlignment="Left" Padding="16,10"/></StackPanel></Border>
      <Border Background="#24233D" CornerRadius="16" BorderBrush="#494364" BorderThickness="1" Padding="20" Margin="0,0,14,0"><StackPanel><StackPanel Orientation="Horizontal"><TextBlock Text="&#xEB9F;" FontFamily="Segoe MDL2 Assets" FontSize="24" Foreground="#C6B7FF" Margin="0,0,12,0"/><TextBlock Text="Pictures" FontSize="20" FontWeight="SemiBold"/></StackPanel><TextBlock Text="Image files only, including subfolders" Foreground="#C4BCDE" Margin="0,12,0,5"/><TextBlock Text="JPG, PNG, HEIC, RAW and other image extensions. Keeps videos, documents and folders." Foreground="{StaticResource Muted}" FontSize="11" MinHeight="32"/><Button x:Name="PicturesScanButton" Content="Scan pictures  &#x2192;" Margin="0,16,0,0" HorizontalAlignment="Left" Padding="16,10"/></StackPanel></Border>
     </UniformGrid>
     <Border Grid.Row="2" Background="#31242D" BorderBrush="#65414D" BorderThickness="1" CornerRadius="11" Padding="15,12" Margin="0,0,0,18"><StackPanel><TextBlock Text="PERMANENT DELETION  /  NO RECYCLE BIN" Foreground="#F2B3BA" FontSize="11" FontWeight="SemiBold"/><TextBlock Text="Back up files you want to keep. Deletions in synced folders (such as OneDrive) may also sync to the cloud and other devices." Foreground="#D1ADB6" FontSize="11" Margin="0,6,0,0"/></StackPanel></Border>
     <StackPanel Grid.Row="3" Margin="0,0,0,12"><TextBlock x:Name="FolderSummary" Text="Start with a scan." FontSize="15" FontWeight="SemiBold"/><TextBlock x:Name="FolderPath" Text="No folder selected" Foreground="#8DB5CA" FontSize="11" Margin="0,5,0,0" TextTrimming="CharacterEllipsis" TextWrapping="NoWrap"/><TextBlock x:Name="FolderDetail" Text="Nothing is deleted by scanning. Linked paths are skipped." Foreground="{StaticResource Muted}" FontSize="11" Margin="0,5,0,0"/></StackPanel>
     <Border Grid.Row="4" Background="#101A2A" CornerRadius="12" BorderBrush="#28374E" BorderThickness="1" Padding="12">
      <ListView x:Name="FolderItems" Background="Transparent" BorderThickness="0" Foreground="#C9D7E9" ScrollViewer.HorizontalScrollBarVisibility="Disabled" ScrollViewer.VerticalScrollBarVisibility="Auto" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling" FontSize="11">
       <ListView.Resources><Style TargetType="GridViewColumnHeader"><Setter Property="Foreground" Value="#A5BAD3"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="GridViewColumnHeader"><Border Background="#1D2C41" Padding="8" BorderBrush="#34455D" BorderThickness="0,0,0,1"><ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/></Border></ControlTemplate></Setter.Value></Setter></Style></ListView.Resources>
       <ListView.ItemContainerStyle><Style TargetType="ListViewItem"><Setter Property="Padding" Value="0,6"/><Setter Property="Foreground" Value="#C9D7E9"/><Setter Property="HorizontalContentAlignment" Value="Stretch"/><Style.Triggers><DataTrigger Binding="{Binding Type}" Value="Skipped"><Setter Property="Foreground" Value="#E6BC82"/></DataTrigger></Style.Triggers></Style></ListView.ItemContainerStyle>
       <ListView.View><GridView AllowsColumnReorder="False"><GridViewColumn Header="File / folder" Width="280"><GridViewColumn.CellTemplate><DataTemplate><TextBlock Text="{Binding RelativePath}" ToolTip="{Binding Path}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn><GridViewColumn Header="Type" Width="75" DisplayMemberBinding="{Binding Type}"/><GridViewColumn Header="Size" Width="110" DisplayMemberBinding="{Binding Size}"/><GridViewColumn Header="Action" Width="170"><GridViewColumn.CellTemplate><DataTemplate><TextBlock Text="{Binding Action}" ToolTip="{Binding Action}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn></GridView></ListView.View>
      </ListView>
     </Border>
     <Grid Grid.Row="5" Margin="0,18,0,0"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Only the reviewed scan is deleted.&#x0a;New or changed files are kept for another scan." FontSize="11" Foreground="{StaticResource Muted}" VerticalAlignment="Center" Margin="0,0,15,0"/><Button x:Name="FolderDeleteButton" Grid.Column="1" Content="Review permanent deletion  &#x2192;" Background="#AD5265" BorderBrush="#CB7484" Foreground="White" FontWeight="SemiBold" Padding="18,14" IsEnabled="False"/></Grid>
    </Grid>
    <Grid x:Name="ActivityPage" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><StackPanel Margin="0,0,0,22"><TextBlock Text="Every step, in view." FontSize="29" FontWeight="SemiBold"/><TextBlock Text="Follow the current session and review any skipped or failed operations." Foreground="{StaticResource Muted}" Margin="0,8,0,0"/></StackPanel><TextBlock x:Name="LogPathLabel" Grid.Row="1" Foreground="#7494B6" FontSize="10" Margin="0,0,0,14"/><Border Grid.Row="2" Background="#090F1D" CornerRadius="14" BorderBrush="#273550" BorderThickness="1" Padding="18"><TextBox x:Name="LogBox" IsReadOnly="True" Background="Transparent" Foreground="#B9CEDE" BorderThickness="0" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/></Border><Button x:Name="BackButton" Grid.Row="3" Content="&#x2190;  Back to cleanup" HorizontalAlignment="Left" Margin="0,18,0,0"/></Grid>
   </Grid>
  </Grid>
  <Border Grid.Row="2" Background="#0D1425" BorderBrush="#202D43" BorderThickness="0,1,0,0"><Grid Margin="20,0"><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="140"/></Grid.ColumnDefinitions><Ellipse x:Name="StatusDot" Fill="#74DEC8" Width="6" Height="6" Margin="0,0,8,0"/><TextBlock x:Name="LatestEvent" Grid.Column="1" Text="Ready" Foreground="#8598B3" FontSize="10" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" TextWrapping="NoWrap"/><ProgressBar x:Name="BusyProgress" Grid.Column="2" Height="3" IsIndeterminate="True" Foreground="#76E8D5" Background="#213248" Visibility="Collapsed"/></Grid></Border>
  <Grid KeyboardNavigation.TabNavigation="Cycle" x:Name="DialogOverlay" Grid.RowSpan="3" Background="#C5080D18" Visibility="Collapsed">
   <Border x:Name="DialogCard" Width="580" MaxHeight="680" VerticalAlignment="Center" HorizontalAlignment="Center" Background="#162237" BorderBrush="#405574" BorderThickness="1" CornerRadius="20" Padding="30"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><TextBlock x:Name="DialogEyebrow" Text="ONE LAST CHECK" Foreground="#7DE9D7" FontSize="10" FontWeight="Bold"/><TextBlock x:Name="DialogTitle" Grid.Row="1" Text="Review your cleanup" FontSize="25" FontWeight="SemiBold" Margin="0,12,0,15"/><StackPanel Grid.Row="2"><ScrollViewer x:Name="DialogBodyScroll" MaxHeight="390" VerticalScrollBarVisibility="Auto"><TextBlock x:Name="DialogBody" Foreground="#B4C3D9" FontSize="13" LineHeight="22"/></ScrollViewer><Border x:Name="DialogReviewPanel" Visibility="Collapsed" Background="#0E192A" BorderBrush="#344A62" BorderThickness="1" CornerRadius="10" Padding="10" Margin="0,12,0,0"><StackPanel><TextBlock Text="PERMANENT DELETE / NO RECYCLE BIN / REVIEW ALL TARGETS" FontSize="10" Foreground="#F2B3BA" Margin="0,0,0,8"/><ListBox x:Name="DialogReviewItems" Height="155" Background="Transparent" BorderThickness="0" Foreground="#B4C3D9" ScrollViewer.HorizontalScrollBarVisibility="Disabled" ScrollViewer.VerticalScrollBarVisibility="Auto" VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling"><ListBox.ItemContainerStyle><Style TargetType="ListBoxItem"><Setter Property="HorizontalContentAlignment" Value="Stretch"/><Setter Property="Padding" Value="4,6"/></Style></ListBox.ItemContainerStyle><ListBox.ItemTemplate><DataTemplate><StackPanel><TextBlock Text="{Binding Scope}" Foreground="#79E8D6" FontSize="10"/><TextBlock Text="{Binding Path}" ToolTip="{Binding Path}" Foreground="#D6E0EE" FontSize="11" TextWrapping="Wrap" Margin="0,3,0,0"/><TextBlock Text="{Binding Action}" Foreground="#A9B9CF" FontSize="10" TextWrapping="Wrap" Margin="0,2,0,0"/></StackPanel></DataTemplate></ListBox.ItemTemplate></ListBox></StackPanel></Border></StackPanel><StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0"><Button x:Name="DialogCancel" Content="Cancel" MinWidth="100" Margin="0,0,10,0"/><Button x:Name="DialogConfirm" Content="Confirm cleanup" Background="#78E8D4" BorderBrush="#78E8D4" Foreground="#102C34" FontWeight="SemiBold" MinWidth="140"/></StackPanel></Grid></Border>
  </Grid>
 </Grid>
 </Border>
</Window>






'@
$reader = New-Object System.Xml.XmlNodeReader $markup
$script:window = [Windows.Markup.XamlReader]::Load($reader)
[AppCleaner.MemoryNativeFrame]::Attach($script:window)
$reader.Close()
$namespaces = New-Object System.Xml.XmlNamespaceManager $markup.NameTable
$namespaces.AddNamespace('x','http://schemas.microsoft.com/winfx/2006/xaml')
foreach ($node in $markup.SelectNodes('//*[@x:Name]', $namespaces)) {
    Set-Variable -Name $node.GetAttribute('Name','http://schemas.microsoft.com/winfx/2006/xaml') -Value $window.FindName($node.GetAttribute('Name','http://schemas.microsoft.com/winfx/2006/xaml')) -Scope Script
}
$script:apps = @(
    @{ Name='Google Chrome'; Pattern='^Google Chrome$'; Processes=@('chrome','chrome_proxy'); Paths=@((Join-Path $env:LOCALAPPDATA 'Google\Chrome'), (Join-Path $env:ProgramFiles 'Google\Chrome')); Keys=@('HKCU:\Software\Google\Chrome') },
    @{ Name='LINE'; Pattern='^LINE$'; Processes=@('LINE','LineLauncher','LineUpdater'); Paths=@((Join-Path $env:LOCALAPPDATA 'LINE'), (Join-Path $env:APPDATA 'LINE')); Keys=@() },
    @{ Name='Telegram Desktop'; Pattern='^Telegram Desktop(?: version .+)?$'; Processes=@('Telegram','Updater'); Paths=@((Join-Path $env:APPDATA 'Telegram Desktop'), (Join-Path $env:LOCALAPPDATA 'Telegram Desktop'), (Join-Path $env:LOCALAPPDATA 'Programs\Telegram Desktop')); Keys=@('HKCU:\Software\Telegram Desktop') },
    @{ Name='Discord'; Pattern='^Discord$'; Processes=@('Discord','Update'); Paths=@((Join-Path $env:LOCALAPPDATA 'Discord'), (Join-Path $env:APPDATA 'discord')); Keys=@() }
)
if (${env:ProgramFiles(x86)}) { $script:apps[0].Paths += Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome' }
function Write-Log([string]$Message) {
    $line = '[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Message
    $script:LogBox.AppendText($line + "`r`n"); $script:LogBox.ScrollToEnd(); $script:LatestEvent.Text = $Message
}
function Get-Installs($App) {
    foreach ($root in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        if (Test-Path -LiteralPath $root) {
            foreach ($key in Get-ChildItem -LiteralPath $root) {
                $item = Get-ItemProperty -LiteralPath $key.PSPath
                if ($item.DisplayName -match $App.Pattern) {
                    [pscustomobject]@{ Name=$item.DisplayName; Command=$item.UninstallString; QuietCommand=$item.QuietUninstallString; Machine=$root.StartsWith('HKLM:') }
                }
            }
        }
    }
}
function Test-AppProcessPath($App, [string]$ExecutablePath) {
    if (-not $ExecutablePath -or -not [IO.Path]::IsPathRooted($ExecutablePath)) { return $false }
    $full = [IO.Path]::GetFullPath($ExecutablePath)
    foreach ($root in $App.Paths) {
        $prefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
        if ($full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}
function Get-AppProcesses($App) {
    $session = [Diagnostics.Process]::GetCurrentProcess().SessionId
    foreach ($name in $App.Processes) {
        foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            if ($process.SessionId -ne $session) { continue }
            # Update/Updater are shared names: never act without a matching app path.
            $path = $null
            try { $path = $process.Path } catch { }
            if (Test-AppProcessPath $App $path) { $process }
            elseif ($name -notin @('Updater','Update')) { throw ('Cannot verify process location: ' + $name + '. Cleanup stopped for ' + $App.Name + '.') }
        }
    }
}
function Assert-Stopped($App) {
    if (@(Get-AppProcesses $App).Count) { throw ('Please close ' + $App.Name + ' first, or use QUICK CLEAN to close it automatically.') }
}
function Stop-AppProcesses($App) {
    # Recheck briefly for app children which were already launching. Never kill by name.
    for ($attempt=0; $attempt -lt 3; $attempt++) {
        $running = @(Get-AppProcesses $App)
        if (-not $running.Count) { return }
        foreach ($process in $running) {
            if ($process.HasExited) { continue }
            Write-Log ('Closing ' + $App.Name + ': PID ' + $process.Id)
            try { $process.Kill() } catch { if (-not $process.HasExited) { throw } }
            if (-not $process.WaitForExit(3000)) { throw ('Process did not stop: ' + $process.Id) }
        }
        Invoke-UiPump
    }
    Assert-Stopped $App
}
function Assert-NoLinks([string]$Path) {
    $node = Get-Item -LiteralPath $Path -Force
    $parent = $node
    while ($null -ne $parent) {
        if ($parent.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked folder refused: $($parent.FullName)" }
        $parent = $parent.Parent
    }
    if ($node.PSIsContainer) {
        foreach ($child in Get-ChildItem -LiteralPath $Path -Force) {
            if ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked item refused: $($child.FullName)" }
            if ($child.PSIsContainer) { Assert-NoLinks $child.FullName }
        }
    }
}
function Start-Uninstall($Entry, [switch]$Quiet) {
    $rawCommand = $Entry.Command
    if ($Quiet -and $Entry.QuietCommand) {
        $rawCommand = $Entry.QuietCommand
        Write-Log ('Using registered quiet uninstaller: ' + $Entry.Name)
    } elseif ($Quiet) { Write-Log ($Entry.Name + ': no registered quiet command; follow the official uninstaller if prompted.') }
    $command = [Environment]::ExpandEnvironmentVariables([string]$rawCommand).Trim()
    if ($command -match '^"([^"\r\n]+\.exe)"\s*(.*)$') { $exe=$Matches[1]; $arguments=$Matches[2] }
    elseif ($command -match '^(.+?\.exe)(?:\s+(.*))?$') { $exe=$Matches[1]; $arguments=$Matches[2] }
    else { throw 'Unsupported uninstall command. Use Windows Settings > Apps for this installation.' }
    if ([IO.Path]::GetFileName($exe) -ieq 'msiexec.exe') {
        $exe = Join-Path $env:WINDIR 'System32\msiexec.exe'
        $arguments = $arguments -replace '(?i)/I(?=\s*\{)', '/X'
    }
    if (-not [IO.Path]::IsPathRooted($exe) -or -not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw 'Uninstaller executable was not found.' }
    $info = New-Object System.Diagnostics.ProcessStartInfo
    $info.FileName=$exe; $info.Arguments=$arguments; $info.UseShellExecute=$true
    if ($Entry.Machine) { $info.Verb='runas' }
    $process = [Diagnostics.Process]::Start($info)
    if ($null -ne $process) {
        while (-not $process.HasExited) { Invoke-UiPump; Start-Sleep -Milliseconds 150 }
        $exitCode = $process.ExitCode
        Write-Log ('Uninstaller exited: ' + $Entry.Name + ' (code ' + $exitCode + ')')
        $process.Dispose()
        if ($exitCode -in @(1641,3010)) { throw 'Restart required. Automatic cleanup was skipped.' }
        if ($exitCode -ne 0) { throw ('Uninstaller did not report success; cleanup skipped. Exit code: ' + $exitCode) }
    }
}
function Remove-AppLeftover($App, [string]$Kind, [string]$Path) {
    if ($Kind -notin @('Paths','Keys') -or $Path -notin $App[$Kind]) { throw 'Path is outside the exact allowlist.' }
    if (Test-Path -LiteralPath $Path) {
        if ($Kind -eq 'Paths') { Assert-NoLinks $Path }
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $Path) { throw 'Some data remains.' }
    }
    Write-Log ('Deleted: ' + $Path)
}
function Invoke-FastCleanup($App) {
    $entries = @(Get-Installs $App)
    if (-not $entries.Count) {
        Write-Log ($App.Name + ': no standard installer found. Use Scan leftovers to review existing data manually.')
        return
    }
    Stop-AppProcesses $App
    foreach ($entry in $entries) { Start-Uninstall $entry -Quiet }
    if (@(Get-Installs $App).Count) { throw 'Still registered as installed. Automatic cleanup skipped.' }
    Stop-AppProcesses $App
    foreach ($kind in @('Paths','Keys')) {
        foreach ($path in ($App[$kind] | Select-Object -Unique)) {
            Assert-Stopped $App
            Remove-AppLeftover $App $kind $path
        }
    }
    Write-Log ($App.Name + ': quick cleanup completed for known locations.')
}



function Invoke-UiPump {
    $frame = New-Object Windows.Threading.DispatcherFrame
    [void]$window.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false }.GetNewClosure())
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}
function Wait-PreviewStep {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.ElapsedMilliseconds -lt 450) { Invoke-UiPump; Start-Sleep -Milliseconds 25 }
}
# Preview replaces every destructive boundary. It never kills, uninstalls or deletes.
if ($Preview) {
    function Get-Installs($App) {
        if (-not $script:previewRemoved.ContainsKey($App.Name)) {
            [pscustomobject]@{ Name=$App.Name; Command='preview'; QuietCommand='preview'; Machine=$false }
        }
    }
    function Get-AppProcesses($App) { }
    function Stop-AppProcesses($App) { Write-Log ('[PREVIEW] Close ' + $App.Name); Wait-PreviewStep }
    function Start-Uninstall($Entry, [switch]$Quiet) {
        Write-Log ('[PREVIEW] Uninstall ' + $Entry.Name); Wait-PreviewStep
        $script:previewRemoved[$Entry.Name] = $true
    }
    function Remove-AppLeftover($App, [string]$Kind, [string]$Path) {
        Write-Log ('[PREVIEW] Would remove: ' + $Path); Wait-PreviewStep
    }
    $ModeLabel.Text = 'PREVIEW / NO CHANGES'
    $window.Title = 'AppCleaner - Preview'
}
$script:cards = @($ChromeCard, $LineCard, $TelegramCard, $DiscordCard)
$script:appStatuses = @($ChromeStatus, $LineStatus, $TelegramStatus, $DiscordStatus)
$script:workControls = $script:cards + @($QuickButton,$AllCleanupButton,$UninstallButton,$ScanButton,$DeleteButton,$SelectAllButton,$ResultsList,$DownloadsScanButton,$PicturesScanButton,$FolderDeleteButton)

function Get-SelectedApps {
    for ($i=0; $i -lt $script:cards.Count; $i++) {
        if ($script:cards[$i].IsChecked) { $script:apps[$i] }
    }
}
function Update-Selection {
    $count = @(Get-SelectedApps).Count
    $SelectionCount.Text = "$count selected"
    $AllCleanupScope.Text = $(if ($count) { "$count checked apps + Downloads contents + Pictures images.`nScan and review everything before confirming." } else { "Downloads contents + Pictures images only.`nNo apps selected. Scan and review before confirming." })
    $SelectAllButton.Content = $(if ($count -eq $script:cards.Count) { 'Deselect all' } else { 'Select all' })
    if (-not $script:busy) {
        $QuickButton.IsEnabled = $count -gt 0
        $UninstallButton.IsEnabled = $count -gt 0
        $ScanButton.IsEnabled = $count -gt 0
        $DeleteButton.IsEnabled = @($script:checkedLocations | Where-Object { $_.IsChecked }).Count -gt 0
        if ($count -eq 0) {
            $ReadyTitle.Text = 'Choose an app to begin.'
            $ReadyDetail.Text = 'Select one or more cards above.'
        } elseif ($ReadyTitle.Text -eq 'Choose an app to begin.') {
            $ReadyTitle.Text = 'Ready when you are.'
            $ReadyDetail.Text = 'Review the selected apps before confirming.'
        }
    }
}
function Update-InstallStatus {
    for ($i=0; $i -lt $script:apps.Count; $i++) {
        try {
            $found = @(Get-Installs $script:apps[$i]).Count -gt 0
            $script:appStatuses[$i].Text = $(if ($found) { 'Desktop installation found' } else { 'No standard installer found' })
        } catch { $script:appStatuses[$i].Text = 'Installation check unavailable' }
    }
}
function Start-Reveal($Element) {
    $fade = New-Object Windows.Media.Animation.DoubleAnimation
    $fade.From=0; $fade.To=1; $fade.Duration=[TimeSpan]::FromMilliseconds(420)
    $Element.BeginAnimation([Windows.UIElement]::OpacityProperty,$fade)
    if ($Element.RenderTransform -is [Windows.Media.TranslateTransform]) {
        $slide = New-Object Windows.Media.Animation.DoubleAnimation
        $slide.From=12; $slide.To=0; $slide.Duration=[TimeSpan]::FromMilliseconds(440)
        $ease = New-Object Windows.Media.Animation.CubicEase
        $ease.EasingMode='EaseOut'; $slide.EasingFunction=$ease
        $Element.RenderTransform.BeginAnimation([Windows.Media.TranslateTransform]::YProperty,$slide)
    }
}
function Set-Page([string]$Name) {
    $CleanPage.Visibility = $(if ($Name -eq 'Clean') {'Visible'} else {'Collapsed'})
    $FoldersPage.Visibility = $(if ($Name -eq 'Folders') {'Visible'} else {'Collapsed'})
    $ActivityPage.Visibility = $(if ($Name -eq 'Activity') {'Visible'} else {'Collapsed'})
    $CleanNav.Background = $(if ($Name -eq 'Clean') {'#1C3443'} else {'Transparent'})
    $CleanNav.BorderBrush = $(if ($Name -eq 'Clean') {'#34586A'} else {'Transparent'})
    $FoldersNav.Background = $(if ($Name -eq 'Folders') {'#1C3443'} else {'Transparent'})
    $ActivityNav.Background = $(if ($Name -eq 'Activity') {'#1C3443'} else {'Transparent'})
    if ($Name -eq 'Activity') { Start-Reveal $ActivityPage; $LogBox.ScrollToEnd() }
    elseif ($Name -eq 'Folders') { Start-Reveal $FoldersPage }
    else { Start-Reveal $PageContent }
}
function Hide-Review {
    $DialogOverlay.Visibility='Collapsed'
    $script:allReviewPlan=$null
    $DialogReviewPanel.Visibility='Collapsed'; $DialogReviewItems.ItemsSource=$null
    $DialogCard.Width=580; $DialogBodyScroll.MaxHeight=390
    $Workspace.IsEnabled=$true
    $script:dialogAction=$null
    if ($null -ne $script:previousFocus) { [void]$script:previousFocus.Focus() }
}
function Show-Review([string]$Title, [string]$Body, [string]$Confirm, [scriptblock]$Action) {
    if ($script:busy) { return }
    $DialogReviewPanel.Visibility='Collapsed'; $DialogReviewItems.ItemsSource=$null
    $DialogCard.Width=580; $DialogBodyScroll.MaxHeight=390
    $script:previousFocus=[Windows.Input.Keyboard]::FocusedElement
    $DialogTitle.Text=$Title; $DialogBody.Text=$Body; $DialogConfirm.Content=$Confirm
    $DialogConfirm.IsEnabled=$true
    $script:dialogAction=$Action
    $DialogEyebrow.Text=$(if ($null -eq $Action) {'A QUICK GUIDE'} elseif ($script:isPreview) {'PREVIEW / NO REAL CHANGES'} else {'ONE LAST CHECK'})
    $DialogCancel.Visibility=$(if ($null -eq $Action) {'Collapsed'} else {'Visible'})
    $Workspace.IsEnabled=$false
    $DialogOverlay.Visibility='Visible'
    Start-Reveal $DialogCard
    if ($null -eq $Action) { [void]$DialogConfirm.Focus() } else { [void]$DialogCancel.Focus() }
}
function Clear-Results {
    $ResultsList.Items.Clear(); $script:checkedLocations.Clear()
    $ResultsPanel.Visibility='Collapsed'; $DeleteButton.IsEnabled=$false
}
function Invoke-Work([scriptblock]$Action) {
    if ($script:busy) { return }
    $script:busy=$true
    $script:cleanupSucceeded=$false
    foreach ($control in $script:workControls) { $control.IsEnabled=$false }
    $CloseButton.IsEnabled=$false
    $BusyProgress.Visibility='Visible'
    $ReadyTitle.Text='Working through your selection...'
    $ReadyDetail.Text='Follow the Activity log for progress.'
    Invoke-UiPump
    try { & $Action }
    catch {
        $script:cleanupSucceeded=$false
        Write-Log ('ERROR: ' + $_.Exception.Message)
        $ReadyTitle.Text='Something needs your attention.'
        $ReadyDetail.Text=$_.Exception.Message
    } finally {
        $script:busy=$false
        foreach ($control in $script:workControls) { $control.IsEnabled=$true }
        $CloseButton.IsEnabled=$true; $BusyProgress.Visibility='Collapsed'
        Update-InstallStatus; Update-Selection; Update-FolderActions
    }
}
function Set-Outcome([int]$Done, [int]$Skipped, [int]$Failed) {
    $script:cleanupSucceeded=($Done -gt 0 -and $Skipped -eq 0 -and $Failed -eq 0)
    $ReadyTitle.Text=$(if ($Failed) {'Finished with issues to review.'} elseif ($Skipped) {'Finished. Some apps were skipped.'} else {'All done. A little more breathing room.'})
    if ($script:isPreview) { $ReadyTitle.Text='Preview complete. No apps were changed.' }
    $ReadyDetail.Text="$Done completed / $Skipped skipped / $Failed failed. See Activity log for details."
    Write-Log $ReadyDetail.Text
}
function Invoke-SelectedCleanup {
    Invoke-Work {
        Clear-Results
        $done=0; $skipped=0; $failed=0
        foreach ($app in @(Get-SelectedApps)) {
            try {
                if (-not @(Get-Installs $app).Count) {
                    Write-Log ($app.Name + ': no standard installer found. Use Scan leftovers to review existing data manually.'); $skipped++; continue
                }
                Write-Log ('Starting quick cleanup: ' + $app.Name)
                Invoke-FastCleanup $app; $done++
            } catch { $failed++; Write-Log ($app.Name + ': ' + $_.Exception.Message) }
        }
        Set-Outcome $done $skipped $failed
    }
}
function Invoke-SelectedUninstall {
    Invoke-Work {
        Clear-Results
        $done=0; $skipped=0; $failed=0
        foreach ($app in @(Get-SelectedApps)) {
            try {
                Assert-Stopped $app
                $entries=@(Get-Installs $app)
                if (-not $entries.Count) { Write-Log ($app.Name + ': no standard installer registration. Store/portable installs require manual removal.'); $skipped++; continue }
                foreach ($entry in $entries) { Write-Log ('Starting official uninstaller: ' + $entry.Name); Start-Uninstall $entry }
                if (@(Get-Installs $app).Count) { throw 'Still registered. Finish uninstallation or restart Windows before cleanup.' }
                Write-Log ($app.Name + ': uninstalled. You can now scan leftovers.'); $done++
            } catch { $failed++; Write-Log ($app.Name + ': ' + $_.Exception.Message) }
        }
        Set-Outcome $done $skipped $failed
    }
}
function Invoke-Scan {
    Invoke-Work {
        Clear-Results
        $skipped=0; $failed=0
        foreach ($app in @(Get-SelectedApps)) {
            try {
                if (@(Get-Installs $app).Count) { Write-Log ($app.Name + ': uninstall first; skipping cleanup scan.'); $skipped++; continue }
                foreach ($kind in @('Paths','Keys')) {
                    foreach ($path in ($app[$kind] | Select-Object -Unique)) {
                        if ($script:isPreview -or (Test-Path -LiteralPath $path)) {
                            $item=New-Object Windows.Controls.CheckBox
                            $item.Margin='0,7,0,7'; $item.Padding='7,0,0,0'
                            $item.Tag=@{ App=$app; Path=$path; Kind=$kind }
                            $label=New-Object Windows.Controls.StackPanel
                            $nameText=New-Object Windows.Controls.TextBlock
                            $nameText.Text=$app.Name + ' / ' + $(if ($kind -eq 'Paths') {'Folder'} else {'Registry'})
                            $nameText.Foreground='#E1E9FA'; $nameText.FontSize=12
                            $pathText=New-Object Windows.Controls.TextBlock
                            $pathText.Text=$path; $pathText.Foreground='#8BA4C2'; $pathText.FontSize=10
                            $pathText.TextWrapping='Wrap'; $pathText.Margin='0,4,0,0'; $pathText.MaxWidth=620
                            [void]$label.Children.Add($nameText); [void]$label.Children.Add($pathText)
                            $item.Content=$label
                            [Windows.Automation.AutomationProperties]::SetName($item,$app.Name + ': ' + $path)
                            $item.Add_Checked({ Update-Selection }); $item.Add_Unchecked({ Update-Selection })
                            [void]$ResultsList.Items.Add($item); [void]$script:checkedLocations.Add($item)
                        }
                    }
                }
            } catch { $failed++; Write-Log ($app.Name + ': scan failed: ' + $_.Exception.Message) }
        }
        $count=$ResultsList.Items.Count
        $ResultsTitle.Text="$count known leftover locations"
        $ResultsPanel.Visibility='Visible'; Start-Reveal $ResultsPanel
        $ReadyTitle.Text=$(if ($count) {'Review the locations below.'} else {'No leftover locations found to review.'})
        $ReadyDetail.Text="$count found / $skipped apps skipped / $failed failed. See Activity log for details."
        Write-Log $ReadyDetail.Text
        $CleanPage.ScrollToBottom()
    }
}
function Invoke-CheckedDelete {
    Invoke-Work {
        $done=0; $failed=0
        foreach ($item in @($script:checkedLocations | Where-Object {$_.IsChecked})) {
            $target=$item.Tag
            try {
                Assert-Stopped $target.App
                if (@(Get-Installs $target.App).Count) { throw 'App is still installed; cleanup refused.' }
                Remove-AppLeftover $target.App $target.Kind $target.Path
                $ResultsList.Items.Remove($item); $script:checkedLocations.Remove($item); $done++
            } catch { $failed++; Write-Log ('FAILED: ' + $target.Path + ' - ' + $_.Exception.Message) }
        }
        $ResultsTitle.Text=[string]$ResultsList.Items.Count + ' known leftover locations remaining'
        Set-Outcome $done 0 $failed
    }
}

$QuickButton.Add_Click({
    $selected=@(Get-SelectedApps)
    if (-not $selected.Count -or $script:busy) { return }
    $names=($selected | ForEach-Object {$_.Name}) -join ', '
    $paths=($selected | ForEach-Object {$_.Paths; $_.Keys} | Select-Object -Unique) -join "`r`n"
    Show-Review 'Ready for a fresh start?' ("Selected apps: $names`r`n`r`nThis will force-close the selected apps, uninstall them, and permanently delete their known local data. Unsaved work will be lost.`r`n`r`nProfiles, saved sign-ins, bookmarks and local chats may be deleted. No Recycle Bin. Back up anything you need.`r`n`r`nLocations to clean:`r`n$paths") 'Confirm cleanup' { Invoke-SelectedCleanup }
})
$UninstallButton.Add_Click({
    $selected=@(Get-SelectedApps)
    if (-not $selected.Count -or $script:busy) { return }
    $names=($selected | ForEach-Object {$_.Name}) -join ', '
    Show-Review 'Uninstall selected apps?' ("Selected apps: $names`r`n`r`nSave your work and close these apps first. Their registered uninstallers will open; follow any prompts they show.`r`n`r`nAfter uninstalling, scan to review leftover data separately.") 'Uninstall apps' { Invoke-SelectedUninstall }
})
$ScanButton.Add_Click({ Invoke-Scan })
$DeleteButton.Add_Click({
    $items=@($script:checkedLocations | Where-Object {$_.IsChecked})
    if (-not $items.Count -or $script:busy) { return }
    $paths=($items | ForEach-Object {$_.Tag.Path}) -join "`r`n"
    Show-Review 'Permanently delete this data?' ("The checked locations and their contents will be permanently deleted. This does not use the Recycle Bin.`r`n`r`n$paths") 'Delete permanently' { Invoke-CheckedDelete }
})
$DialogCancel.Add_Click({ Hide-Review })
$DialogConfirm.Add_Click({
    $action=$script:dialogAction
    $approvedAllPlan=$script:allReviewPlan
    Hide-Review
    $script:cleanupSucceeded=$false
    if ($null -ne $action) { & $action $approvedAllPlan }
})
$window.Add_PreviewKeyDown({
    if ($_.Key -eq [Windows.Input.Key]::Escape -and $DialogOverlay.Visibility -eq 'Visible') { Hide-Review; $_.Handled=$true }
})
$SelectAllButton.Add_Click({ $check=@(Get-SelectedApps).Count -lt $script:cards.Count; foreach ($card in $script:cards) { $card.IsChecked=$check }; Update-Selection })
foreach ($card in $script:cards) { $card.Add_Checked({ Update-Selection }); $card.Add_Unchecked({ Update-Selection }) }
$CleanNav.Add_Click({ Set-Page 'Clean' }); $BackButton.Add_Click({ Set-Page 'Clean' })
$ActivityNav.Add_Click({ Set-Page 'Activity' })
$HelpButton.Add_Click({
    Show-Review 'A clean start, directly from PowerShell.' "MEMORY SESSION`r`nNo package installation or app log file is needed. Closing this window ends the UI session. Windows and uninstallers can still record their normal activity.`r`n`r`nQUICK CLEAN`r`nSelect apps, review the confirmation, then close + uninstall + clean known locations in one flow.`r`n`r`nSTEP BY STEP`r`n1. Save your work, close the apps and uninstall.`r`n2. Scan known leftover locations.`r`n3. Check specific locations, then confirm deletion.`r`n`r`nOnly standard desktop installations are detected. Store, portable or custom installations may need Windows Settings.`r`n`r`nQuick clean handles selected apps only. The combined Clean everything button reviews checked apps, Downloads contents and Pictures image files together, including existing leftovers for checked apps. With no apps checked, only the folders are included. Cancel keeps everything. Individual folder actions remain available. Folder deletion is permanent and may sync to cloud storage. This is not secure disk erasure.`r`n`r`nA restart request, failed uninstall or remaining installer registration stops automatic cleanup." 'Got it' $null
})
$TitleBar.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) { $window.WindowState=$(if ($window.WindowState -eq 'Maximized') {'Normal'} else {'Maximized'}) }
    elseif ($_.ButtonState -eq 'Pressed') { $window.DragMove() }
})
$MinimizeButton.Add_Click({ $window.WindowState='Minimized' })
$MaximizeButton.Add_Click({ $window.WindowState=$(if ($window.WindowState -eq 'Maximized') {'Normal'} else {'Maximized'}) })
$CloseButton.Add_Click({ $window.Close() })
$window.Add_Closing({ if ($script:busy) { $_.Cancel=$true } })
$window.Add_Loaded({
    Update-InstallStatus; Update-Selection; Start-Reveal $PageContent
    if ([Windows.SystemParameters]::ClientAreaAnimation) {
        $spin=New-Object Windows.Media.Animation.DoubleAnimation
        $spin.From=0; $spin.To=360; $spin.Duration=[TimeSpan]::FromSeconds(26)
        $spin.RepeatBehavior=[Windows.Media.Animation.RepeatBehavior]::Forever
        $Orbit.RenderTransform.BeginAnimation([Windows.Media.RotateTransform]::AngleProperty,$spin)
    }
})
# Folder cleanup is a separate, reviewed operation. No recursive deletion is used.
$script:folderPlan=$null
$script:folderAppRoot=$null # This edition has no installed program directory.
$script:imageExtensions=@('.jpg','.jpeg','.jpe','.jfif','.png','.gif','.bmp','.webp','.tif','.tiff','.heic','.heif','.avif','.svg','.ico','.dng','.raw','.cr2','.cr3','.nef','.nrw','.arw','.srf','.sr2','.orf','.rw2','.raf','.pef','.psd','.psb','.jxl','.jp2','.j2k')

function Test-InFolder([string]$Path,[string]$Root,[switch]$IncludeRoot) {
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Root)) { return $false }
    $full=[IO.Path]::GetFullPath($Path).TrimEnd('\')
    $base=[IO.Path]::GetFullPath($Root).TrimEnd('\')
    return (($IncludeRoot -and $full.Equals($base,[StringComparison]::OrdinalIgnoreCase)) -or $full.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase))
}
function Get-CleanupFolderRoot([ValidateSet('Downloads','Pictures')][string]$Kind) {
    $path=[AppCleaner.MemoryKnownFolders]::Get($Kind)
    if ([string]::IsNullOrWhiteSpace($path) -or -not [IO.Path]::IsPathRooted($path)) { throw "Cannot resolve the $Kind folder." }
    return [IO.Path]::GetFullPath($path).TrimEnd('\')
}
function Assert-CleanupRoot([string]$Root) {
    if (-not [IO.Path]::IsPathRooted($Root) -or $Root.StartsWith('\\')) { throw 'Only local Windows folders are supported.' }
    $full=[IO.Path]::GetFullPath($Root).TrimEnd('\')
    if ($full -eq [IO.Path]::GetPathRoot($full).TrimEnd('\')) { throw 'A drive root cannot be cleaned.' }
    if ($env:USERPROFILE -and (Test-InFolder $env:USERPROFILE $full -IncludeRoot)) { throw 'A user profile root or its parent cannot be cleaned.' }
    foreach ($protected in @($env:WINDIR,$env:ProgramFiles,${env:ProgramFiles(x86)},$env:ProgramData,$env:APPDATA,$env:LOCALAPPDATA)) {
        if ($protected -and ((Test-InFolder $full $protected -IncludeRoot) -or (Test-InFolder $protected $full -IncludeRoot))) { throw 'This folder overlaps a system or application-data location.' }
    }
    if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw 'The folder does not exist.' }
    Assert-FolderPathNoLinks $full
}
function Assert-FolderPathNoLinks([string]$Path) {
    $node=Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    while ($null -ne $node) {
        if ($node.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw ('Linked or cloud-placeholder path skipped: ' + $node.FullName) }
        if ($node -is [IO.FileInfo]) { $node=$node.Directory } else { $node=$node.Parent }
    }
}
function Format-FileSize([long]$Bytes) {
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}
function New-FolderEntry($Item,[string]$Root) {
    $directory=[bool]$Item.PSIsContainer
    $length=$(if ($directory) {0L} else {[long]$Item.Length})
    [pscustomobject]@{
        Path=$Item.FullName; RelativePath=$Item.FullName.Substring($Root.Length).TrimStart('\')
        IsDirectory=$directory; Length=$length; LastWriteTicks=$Item.LastWriteTimeUtc.Ticks
        CreationTicks=$Item.CreationTimeUtc.Ticks
        Type=$(if ($directory) {'Folder'} else {'File'})
        Size=$(if ($directory) {'Empty after cleanup'} else {Format-FileSize $length})
        Action='Delete'
    }
}
function New-FolderCleanupPlan([ValidateSet('Downloads','Pictures')][string]$Kind) {
    $root=Get-CleanupFolderRoot $Kind
    Assert-CleanupRoot $root
    $files=New-Object 'System.Collections.Generic.List[object]'
    $directories=New-Object 'System.Collections.Generic.List[object]'
    $skipped=New-Object 'System.Collections.Generic.List[object]'
    $pending=New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push($root)
    $total=0L; $visited=0
    while ($pending.Count) {
        $folder=$pending.Pop()
        try {
            if (Test-InFolder $folder $script:folderAppRoot -IncludeRoot) { throw 'AppCleaner program folder is protected.' }
            Assert-FolderPathNoLinks $folder
            $children=@(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop)
        } catch {
            $skipped.Add([pscustomobject]@{Path=$folder;RelativePath=$folder;Type='Skipped';Size='';Action=$_.Exception.Message}); continue
        }
        foreach ($item in $children) {
            $visited++
            if ($visited % 100 -eq 0) { Invoke-UiPump }
            if (-not (Test-InFolder $item.FullName $root)) { throw 'An item is outside the selected folder.' }
            $reason=$null
            if (Test-InFolder $item.FullName $script:folderAppRoot -IncludeRoot) { $reason='AppCleaner program folder is protected.' }
            elseif ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { $reason='Linked or cloud-placeholder item; not followed.' }
            if ($reason) {
                $skipped.Add([pscustomobject]@{Path=$item.FullName;RelativePath=$item.FullName.Substring($root.Length).TrimStart('\');Type='Skipped';Size='';Action=$reason}); continue
            }
            if ($item.PSIsContainer) {
                $pending.Push($item.FullName)
                if ($Kind -eq 'Downloads' -and -not (Test-InFolder $script:folderAppRoot $item.FullName -IncludeRoot)) { $directories.Add((New-FolderEntry $item $root)) }
            } elseif ($Kind -eq 'Downloads' -or $item.Extension.ToLowerInvariant() -in $script:imageExtensions) {
                $entry=New-FolderEntry $item $root
                $files.Add($entry); $total+=$entry.Length
            }
        }
    }
    [pscustomobject]@{Kind=$Kind;Root=$root;Files=$files.ToArray();Directories=$directories.ToArray();Skipped=$skipped.ToArray();TotalBytes=$total;Preview=$false;ScannedAt=(Get-Date)}
}
function New-PreviewFolderPlan([ValidateSet('Downloads','Pictures')][string]$Kind) {
    $root=Get-CleanupFolderRoot $Kind
    $names=$(if ($Kind -eq 'Downloads') {@('installer.exe','archive.zip','Notes\readme.txt')} else {@('photo.jpg','Screenshots\capture.png','holiday.webp')})
    $files=@(for ($i=0;$i -lt $names.Count;$i++) {
        [pscustomobject]@{Path=(Join-Path $root $names[$i]);RelativePath=$names[$i];IsDirectory=$false;Length=(($i+1)*1MB);LastWriteTicks=0L;CreationTicks=0L;Type='Sample file';Size=(Format-FileSize (($i+1)*1MB));Action='Preview only'}
    })
    [pscustomobject]@{Kind=$Kind;Root=$root;Files=$files;Directories=@();Skipped=@();TotalBytes=6MB;Preview=$true;ScannedAt=(Get-Date)}
}
function Assert-FolderPlanEntry($Plan,$Entry) {
    if (-not (Test-InFolder $Entry.Path $Plan.Root)) { throw 'The target is outside the reviewed folder.' }
    if ((Test-InFolder $Entry.Path $script:folderAppRoot -IncludeRoot) -or ($Entry.IsDirectory -and (Test-InFolder $script:folderAppRoot $Entry.Path -IncludeRoot))) { throw 'AppCleaner program files are protected.' }
    if ($Plan.Kind -eq 'Pictures' -and ($Entry.IsDirectory -or [IO.Path]::GetExtension($Entry.Path).ToLowerInvariant() -notin $script:imageExtensions)) { throw 'Pictures cleanup only accepts image files.' }
    Assert-FolderPathNoLinks $Entry.Path
    $current=Get-Item -LiteralPath $Entry.Path -Force -ErrorAction Stop
    if ([bool]$current.PSIsContainer -ne $Entry.IsDirectory -or $current.CreationTimeUtc.Ticks -ne $Entry.CreationTicks) { throw 'The item was replaced after the scan; scan again.' }
    if (-not $Entry.IsDirectory -and ($current.Length -ne $Entry.Length -or $current.LastWriteTimeUtc.Ticks -ne $Entry.LastWriteTicks)) { throw 'The file changed after the scan; scan again.' }
}
function Invoke-FolderCleanupPlan($Plan) {
    if ($null -eq $Plan -or $Plan.Kind -notin @('Downloads','Pictures')) { throw 'Scan a supported folder first.' }
    $deleted=0; $removedFolders=0; $failed=0; $bytes=0L
    if ($script:isPreview) {
        foreach ($entry in $Plan.Files) { Write-Log ('[PREVIEW] Would permanently delete: ' + $entry.Path); Wait-PreviewStep }
        return [pscustomobject]@{Deleted=$Plan.Files.Count;Folders=0;Failed=0;Bytes=$Plan.TotalBytes;Preview=$true}
    }
    if ($Plan.Preview) { throw 'Preview plans cannot be used for real deletion. Scan again.' }
    $currentRoot=Get-CleanupFolderRoot $Plan.Kind
    if (-not $currentRoot.Equals($Plan.Root,[StringComparison]::OrdinalIgnoreCase)) { throw 'The Windows folder location changed. Scan again.' }
    Assert-CleanupRoot $Plan.Root
    foreach ($entry in @($Plan.Files)+@($Plan.Directories | Sort-Object { $_.Path.Length } -Descending)) {
        try {
            Assert-FolderPlanEntry $Plan $entry
            if ($entry.IsDirectory) {
                # Never recurse: new, skipped, protected or locked contents keep the directory intact.
                [IO.Directory]::Delete($entry.Path,$false)
                $removedFolders++
            } else {
                $attributes=[IO.File]::GetAttributes($entry.Path)
                if ($attributes -band [IO.FileAttributes]::ReadOnly) { [IO.File]::SetAttributes($entry.Path,($attributes -band (-bnot [IO.FileAttributes]::ReadOnly))) }
                # File.Delete cannot recurse or remove a directory if the file was replaced.
                [IO.File]::Delete($entry.Path)
                $deleted++; $bytes+=$entry.Length
            }
            Write-Log ('Deleted permanently: ' + $entry.Path)
        } catch { $failed++; Write-Log ('KEPT: ' + $entry.Path + ' - ' + $_.Exception.Message) }
        Invoke-UiPump
    }
    [pscustomobject]@{Deleted=$deleted;Folders=$removedFolders;Failed=$failed;Bytes=$bytes;Preview=$false}
}

function Update-FolderActions {
    if ($script:busy) { return }
    $FolderDeleteButton.IsEnabled=($null -ne $script:folderPlan -and ($script:folderPlan.Files.Count+$script:folderPlan.Directories.Count) -gt 0)
}
function Invoke-FolderScan([ValidateSet('Downloads','Pictures')][string]$Kind) {
    if ($script:busy) { return }
    $script:folderPlan=$null
    $FolderItems.ItemsSource=$null
    $FolderDeleteButton.IsEnabled=$false
    $FolderSummary.Text='Scanning ' + $Kind + '...'
    Invoke-Work {
        try {
            $plan=$(if ($script:isPreview) { New-PreviewFolderPlan $Kind } else { New-FolderCleanupPlan $Kind })
            $script:folderPlan=$plan
            $FolderPath.Text=$plan.Root
            $FolderItems.ItemsSource=@($plan.Files)+@($plan.Directories)+@($plan.Skipped)
            $FolderSummary.Text=('{0}: {1} files / {2} folders / {3} / {4} skipped' -f $Kind,$plan.Files.Count,$plan.Directories.Count,(Format-FileSize $plan.TotalBytes),$plan.Skipped.Count)
            $FolderDetail.Text=$(if ($script:isPreview) {'Sample results only. Preview never reads or deletes your folder contents.'} else {'Review the full list below. Only this scan is deleted; new or changed files are kept. Skipped items are listed at the end.'})
            $ReadyTitle.Text='Folder scan finished.'; $ReadyDetail.Text='Review the Folder cleanup page.'
            Write-Log $FolderSummary.Text
        } catch { $FolderSummary.Text='Scan could not finish.'; $FolderDetail.Text=$_.Exception.Message; throw }
    }
    Update-FolderActions
}
function Show-FolderDeleteReview {
    $plan=$script:folderPlan
    if ($script:busy -or $null -eq $plan -or ($plan.Files.Count+$plan.Directories.Count) -eq 0) { return }
    $scope=$(if ($plan.Kind -eq 'Downloads') {'All reviewed files and subfolders inside Downloads.'} else {'Image files only, including images in subfolders. Other file types and folders stay.'})
    $body="Folder: $($plan.Root)`r`n`r`n$scope`r`n$($plan.Files.Count) files / $($plan.Directories.Count) folders / $(Format-FileSize $plan.TotalBytes)`r`n$($plan.Skipped.Count) skipped items will stay.`r`n`r`nPERMANENT DELETE. No Recycle Bin. Files cannot be restored through this app. Synced folder deletions may also sync to the cloud and other devices.`r`n`r`nLinked paths are excluded. This memory session has no program folder to exclude. The Downloads/Pictures root folder itself is kept."
    Show-Review ('Delete from ' + $plan.Kind + '?') $body 'Delete permanently' {
        $approvedPlan=$script:folderPlan
        $script:folderPlan=$null
        Invoke-Work {
            try {
                $result=Invoke-FolderCleanupPlan $approvedPlan
                $FolderSummary.Text=$(if ($result.Preview) {'Preview complete. No files were changed.'} else {'Folder cleanup finished.'})
                $FolderDetail.Text=('{0} files deleted / {1} folders removed / {2} failed or changed / {3} skipped. Scan again to refresh.' -f $result.Deleted,$result.Folders,$result.Failed,$approvedPlan.Skipped.Count)
                $ReadyTitle.Text=$FolderSummary.Text; $ReadyDetail.Text=$FolderDetail.Text
                $FolderItems.ItemsSource=$null
                Write-Log $FolderDetail.Text
                $script:cleanupSucceeded=($result.Failed -eq 0 -and ($result.Deleted+$result.Folders) -gt 0)
            } catch { $FolderSummary.Text='Cleanup stopped.'; $FolderDetail.Text=$_.Exception.Message; throw }
        }
        Update-FolderActions
    }
}
$DownloadsScanButton.Add_Click({ Invoke-FolderScan 'Downloads' })
$PicturesScanButton.Add_Click({ Invoke-FolderScan 'Pictures' })
$FolderDeleteButton.Add_Click({ Show-FolderDeleteReview })
$FoldersNav.Add_Click({ Set-Page 'Folders' })

# One review covers checked apps plus both user folders. Scanning never deletes.
$script:allReviewPlan=$null
$script:lastAllOutcome=$null

function Get-AllInstallSignature($Entries) {
    return (@($Entries | Sort-Object Name,Command,QuietCommand | Select-Object Name,Command,QuietCommand,Machine) | ConvertTo-Json -Depth 4 -Compress)
}
function New-AllCleanupPlan {
    $selectedApps=@(Get-SelectedApps)
    $appPlans=New-Object 'System.Collections.Generic.List[object]'
    $review=New-Object 'System.Collections.Generic.List[object]'
    foreach ($app in $selectedApps) {
        $installs=@(Get-Installs $app)
        $targets=@(foreach ($kind in @('Paths','Keys')) {
            foreach ($path in ($app[$kind] | Select-Object -Unique)) {
                if ($script:isPreview -or $installs.Count -gt 0 -or (Test-Path -LiteralPath $path)) {
                    [pscustomobject]@{Kind=$kind;Path=$path}
                }
            }
        })
        $appPlans.Add([pscustomobject]@{App=$app;Installs=$installs;Signature=(Get-AllInstallSignature $installs);Targets=$targets})
        $action=$(if ($installs.Count) {'Close + uninstall + clean known locations'} elseif ($targets.Count) {'Close + delete reviewed leftover locations'} else {'No supported installation or leftovers; skip'})
        $review.Add([pscustomobject]@{Scope=$app.Name;Path=$app.Name;Action=$action})
        foreach ($target in $targets) { $review.Add([pscustomobject]@{Scope=$app.Name;Path=$target.Path;Action='Permanently delete known location and its contents'}) }
        Invoke-UiPump
    }
    $folderPlans=@(foreach ($kind in @('Downloads','Pictures')) {
        Write-Log ('All cleanup: scanning '+$kind)
        if ($script:isPreview) { New-PreviewFolderPlan $kind } else { New-FolderCleanupPlan $kind }
    })
    # Known folders can overlap after redirection; delete each reviewed file once.
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($plan in $folderPlans) {
        $unique=New-Object 'System.Collections.Generic.List[object]'
        $total=0L
        foreach ($file in $plan.Files) {
            if ($seen.Add($file.Path)) { $unique.Add($file); $total+=$file.Length }
        }
        $plan.Files=$unique.ToArray(); $plan.TotalBytes=$total
        foreach ($entry in @($plan.Files)+@($plan.Directories)+@($plan.Skipped)) {
            $review.Add([pscustomobject]@{Scope=$plan.Kind;Path=$entry.Path;Action=$entry.Action})
        }
    }
    return [pscustomobject]@{Apps=$appPlans.ToArray();Folders=$folderPlans;Review=$review.ToArray();Preview=$script:isPreview;Used=$false}
}
function Invoke-AllAppCleanup($Entry) {
    $app=$Entry.App
    $current=@(Get-Installs $app)
    if ((Get-AllInstallSignature $current) -cne $Entry.Signature) { throw 'Installation changed after review. Scan again.' }
    if ($Entry.Installs.Count) { Invoke-FastCleanup $app; return }
    Stop-AppProcesses $app
    if (@(Get-Installs $app).Count) { throw 'App became installed after review. Leftover deletion stopped.' }
    foreach ($target in $Entry.Targets) {
        Assert-Stopped $app
        Remove-AppLeftover $app $target.Kind $target.Path
    }
}
function Invoke-AllCleanupPlan($Plan) {
    if ($script:busy) { return }
    if ($null -eq $Plan -or $Plan.Used -or $Plan.Preview -ne $script:isPreview) { throw 'This review is no longer valid. Scan again.' }
    $Plan.Used=$true
    $script:lastAllOutcome=$null
    Invoke-Work {
        Clear-Results
        $script:folderPlan=$null; $FolderItems.ItemsSource=$null
        $FolderSummary.Text='Combined cleanup in progress.'
        $done=0; $skipped=0; $failedApps=0; $files=0; $folders=0; $failedFiles=0; $excluded=0
        $stage=0; $totalStages=$Plan.Apps.Count+$Plan.Folders.Count
        foreach ($entry in $Plan.Apps) {
            $stage++
            $ReadyTitle.Text="[$stage/$totalStages] " + $entry.App.Name
            Write-Log ('All cleanup: '+$entry.App.Name); Invoke-UiPump
            if (-not $entry.Installs.Count -and -not $entry.Targets.Count) { $skipped++; continue }
            try { Invoke-AllAppCleanup $entry; $done++ }
            catch { $failedApps++; Write-Log ('KEPT app data: '+$entry.App.Name+' - '+$_.Exception.Message) }
        }
        foreach ($folderPlan in $Plan.Folders) {
            $stage++
            $ReadyTitle.Text="[$stage/$totalStages] " + $folderPlan.Kind
            Write-Log ('All cleanup: '+$folderPlan.Kind); Invoke-UiPump
            $excluded+=$folderPlan.Skipped.Count
            if (-not ($folderPlan.Files.Count+$folderPlan.Directories.Count)) { continue }
            try {
                $result=Invoke-FolderCleanupPlan $folderPlan
                $files+=$result.Deleted; $folders+=$result.Folders; $failedFiles+=$result.Failed
            } catch { $failedFiles++; Write-Log ('Folder cleanup stopped: '+$folderPlan.Kind+' - '+$_.Exception.Message) }
        }
        $script:lastAllOutcome=[pscustomobject]@{Apps=$done;SkippedApps=$skipped;FailedApps=$failedApps;Files=$files;Folders=$folders;FailedFiles=$failedFiles;SkippedFiles=$excluded;Preview=$script:isPreview}
        $script:cleanupSucceeded=($failedApps+$failedFiles+$excluded+$skipped -eq 0 -and $done+$files+$folders -gt 0)
    }
    if ($null -ne $script:lastAllOutcome) {
        $outcome=$script:lastAllOutcome
        $ReadyTitle.Text=$(if ($outcome.Preview) {'Preview complete. Nothing was deleted.'} elseif ($outcome.FailedApps+$outcome.FailedFiles) {'Combined cleanup finished with issues.'} else {'Combined cleanup complete.'})
        $ReadyDetail.Text=('{0} apps / {1} files / {2} folders / {3} skipped / {4} failed. See Activity log.' -f $outcome.Apps,$outcome.Files,$outcome.Folders,($outcome.SkippedApps+$outcome.SkippedFiles),($outcome.FailedApps+$outcome.FailedFiles))
        $FolderSummary.Text=$ReadyTitle.Text; $FolderDetail.Text=$ReadyDetail.Text
        Write-Log $ReadyTitle.Text; Write-Log $ReadyDetail.Text
    }
}
function Show-AllCleanupReview {
    if ($script:busy) { return }
    $script:allReviewPlan=$null
    Invoke-Work {
        $ReadyTitle.Text='Scanning all cleanup targets...'
        $script:allReviewPlan=New-AllCleanupPlan
        $ReadyTitle.Text='All targets scanned. Review before deleting.'
    }
    $plan=$script:allReviewPlan
    if ($null -eq $plan) { return }
    $folderSummary=(@($plan.Folders | ForEach-Object { '{0}: {1} files / {2} folders / {3} skipped / {4}{5}{6}' -f $_.Kind,$_.Files.Count,$_.Directories.Count,$_.Skipped.Count,(Format-FileSize $_.TotalBytes),[Environment]::NewLine,$_.Root }) -join "`r`n`r`n")
    $appSummary=if ($plan.Apps.Count) {
        'CHECKED APPS ('+$plan.Apps.Count+'): '+(($plan.Apps | ForEach-Object {$_.App.Name}) -join ', ')+".`r`nClose these apps, run their uninstallers, and delete known local profiles, saved sign-ins, bookmarks and chats. Existing leftover locations are included even when no installer is found. Unsaved work may be lost."
    } else { 'NO APPS SELECTED. Only Downloads contents and Pictures image files are included.' }
    $body="$appSummary`r`n`r`n$folderSummary`r`n`r`nDownloads: reviewed contents. Pictures: image files only. PERMANENT DELETE. No Recycle Bin. Synced deletions may propagate to cloud storage. New or changed folder files are kept. Review the full target list below; skipped items stay."
    Show-Review 'Clean everything in one flow?' $body 'Confirm all cleanup' { param($ApprovedPlan) Invoke-AllCleanupPlan $ApprovedPlan }
    $DialogCard.Width=700; $DialogBodyScroll.MaxHeight=210
    $DialogReviewItems.ItemsSource=$plan.Review
    $DialogReviewPanel.Visibility='Visible'
}
$AllCleanupButton.Add_Click({ Show-AllCleanupReview })

$LogPathLabel.Text='MEMORY SESSION / Activity is shown here only; no application log file is written.'
Write-Log $(if ($Preview) {'Preview ready. All cleanup actions are simulated.'} else {'Memory session ready. Select apps and review before confirming cleanup.'})
Update-Selection
if (-not $NoShow) { [void]$window.ShowDialog() }




