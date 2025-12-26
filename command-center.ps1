# command-center.ps1
# encoding: utf-8-bom
# Toggle dashboard: Ctrl+Alt+D
# Run:
#   powershell -STA -NoProfile -ExecutionPolicy Bypass -File ".\command-center.ps1"
#
# Modules:
# - Calendar (custom mini calendar – reference look)
# - Quick Notes (auto-saved to notes.json)
# - Focus / Pomodoro
# - Unread mails (Outlook COM; may be blocked by policy)

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class Win32 {
  [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd,int id,uint fsModifiers,uint vk);
  [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd,int id);

  public const int WM_HOTKEY = 0x0312;
  public const uint MOD_ALT = 0x0001;
  public const uint MOD_CONTROL = 0x0002;
}
"@

# ---------------- Hotkey ----------------
$HOTKEY_ID = 1
$VK_D = 0x44   # D

# ---------------- Paths ----------------
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$NotesPath    = Join-Path $ScriptDir "notes.json"
$TodoPath     = Join-Path $ScriptDir "todos.json"
$ErrorLogPath = Join-Path $ScriptDir "cc-errors.log"

trap {
  try {
    ("[{0}] " -f (Get-Date)) + ($_ | Out-String) | Out-File -FilePath $ErrorLogPath -Append -Encoding utf8
  } catch {}
  break
}

# ---------------- State ----------------
$global:DashboardWin = $null
$global:HotkeyHostHwnd = [IntPtr]::Zero

$global:Pomodoro = [pscustomobject]@{
  Mode = "IDLE"          # IDLE | FOCUS | BREAK
  Total = 1500           # 25 min
  Remaining = 1500
  Running = $false
}

# Calendar state
$global:CalDisplay = (Get-Date -Day 1).Date
$global:CalSelected = (Get-Date).Date
$global:CalButtons = @()

function Format-Time([int]$sec){
  if($sec -lt 0){ $sec = 0 }
  $ts = [TimeSpan]::FromSeconds($sec)
  return "{0:00}:{1:00}" -f [int]$ts.TotalMinutes, $ts.Seconds
}

function Get-UnreadMailsTop5 {
  try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6) # 6=Inbox
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    $result = @()
    foreach($it in $items){
      if($result.Count -ge 5){ break }
      try {
        if($it.UnRead -eq $true){
          $result += [pscustomobject]@{
            Sender   = $it.SenderName
            Subject  = $it.Subject
            Received = $it.ReceivedTime
            EntryID  = $it.EntryID
            StoreID  = $it.Parent.StoreID
            Item     = $it
          }
        }
      } catch {}
    }
    return $result
  } catch {
    return $null
  }
}

function Open-MailItem($item){
  try {
    if($null -ne $item){
      $item.Display() | Out-Null
    } else {
      [System.Windows.MessageBox]::Show("Mail bulunamadı.", "Command Center") | Out-Null
    }
  } catch {
    [System.Windows.MessageBox]::Show("Mail açılamadı: " + $_.Exception.Message, "Command Center") | Out-Null
  }
}

function Ensure-NotesFile {
  if(!(Test-Path $NotesPath)){
    "[]" | Out-File -FilePath $NotesPath -Encoding utf8
  }
}
function Load-NotesData {
  Ensure-NotesFile
  try {
    $raw = Get-Content $NotesPath -Raw -ErrorAction Stop
    $obj = $raw | ConvertFrom-Json

    if($null -eq $obj){ return @() }

    # legacy single-note format: {"text":"..."}
    if($obj.PSObject.Properties.Name -contains "text" -and -not ($obj -is [System.Collections.IEnumerable])){
      $created = Get-Date
      $note = [pscustomobject]@{
        id = 1
        title = "Not"
        text = [string]$obj.text
        createdAt = $created
        updatedAt = $created
        date = $created.Date
      }
      return @($note)
    }

    if($obj -is [System.Collections.IEnumerable]){
      $arr = @($obj)
      foreach($n in $arr){
        if(-not $n.PSObject.Properties["date"]){
          $created = $null
          try { $created = [datetime]$n.createdAt } catch {}
          if($created -eq $null){ $created = Get-Date }
          $n | Add-Member -NotePropertyName date -NotePropertyValue $created.Date -Force
        }
      }
      return $arr
    }

    return @()
  } catch {
    return @()
  }
}

function Save-NotesData([array]$notes){
  try {
    $payload = $notes | ConvertTo-Json -Depth 5
    $payload | Out-File -FilePath $NotesPath -Encoding utf8
  } catch {}
}

# legacy helpers kept for compatibility (no longer used)
function Load-NotesText { return "" }
function Save-NotesText([string]$t){}

# ---------------- UI ----------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1120" Height="700"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="False">

  <Window.Resources>

    <!-- pill icon button -->
    <Style x:Key="IconBtn" TargetType="{x:Type Button}">
      <Setter Property="Width" Value="30"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="#8FA3BC"/>
      <Setter Property="FontSize" Value="18"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Border x:Name="Bg" CornerRadius="12" Background="Transparent">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#CFE0F7"/>
                <Setter TargetName="Bg" Property="Background" Value="#1E2A38"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- primary action button -->
    <Style x:Key="PrimaryActionBtn" TargetType="{x:Type Button}">
      <Setter Property="Padding" Value="22,9"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Background" Value="#2382E8"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Border x:Name="Bg" CornerRadius="8" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bg" Property="Background" Value="#3391F0"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bg" Property="Background" Value="#1B6FCC"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bg" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- secondary action button -->
    <Style x:Key="SecondaryActionBtn" TargetType="{x:Type Button}">
      <Setter Property="Padding" Value="20,9"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Background" Value="#1A2028"/>
      <Setter Property="Foreground" Value="#CFE0F7"/>
      <Setter Property="BorderBrush" Value="#2B3A4E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Normal"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Border x:Name="Bg" CornerRadius="8" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bg" Property="Background" Value="#232B36"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bg" Property="Background" Value="#151B22"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bg" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- calendar day button -->
    <Style x:Key="CalDayBtn" TargetType="{x:Type Button}">
      <Setter Property="Width" Value="36"/>
      <Setter Property="Height" Value="36"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#CFE0F7"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Grid>
              <Border x:Name="Bg" CornerRadius="18" Background="{TemplateBinding Background}"/>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bg" Property="Background" Value="#1E2A38"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Foreground" Value="#3E4D60"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Quick Notes TextBox -->
    <Style x:Key="NotesBoxStyle" TargetType="{x:Type TextBox}">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="#CFE0F7"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="CaretBrush" Value="#CFE0F7"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="AcceptsReturn" Value="True"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type TextBox}">
            <Grid>
              <ScrollViewer x:Name="PART_ContentHost" Background="Transparent"/>
              <TextBlock x:Name="Watermark"
                         Text="Type a quick note..."
                         Foreground="#5F738A"
                         IsHitTestVisible="False"
                         Visibility="Collapsed"
                         Margin="0,2,0,0"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="Text" Value="">
                <Setter TargetName="Watermark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="Text" Value="{x:Null}">
                <Setter TargetName="Watermark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="Watermark" Property="Visibility" Value="Collapsed"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Quick Notes list styles -->
    <Style x:Key="NotesListItemStyle" TargetType="{x:Type ListBoxItem}">
      <Setter Property="Padding" Value="8,4"/>
      <Setter Property="Margin" Value="4,2"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Foreground" Value="#CFE0F7"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type ListBoxItem}">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#2382E8"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1E2A38"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="NotesListStyle" TargetType="{x:Type ListBox}">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource NotesListItemStyle}"/>
      <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type ListBox}">
            <Border Background="Transparent">
              <ScrollViewer Focusable="False"
                           Padding="0"
                           VerticalScrollBarVisibility="Auto"
                           HorizontalScrollBarVisibility="Disabled">
                <ItemsPresenter />
              </ScrollViewer>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Tasks checkbox style -->
    <Style x:Key="TodoCheckBoxStyle" TargetType="{x:Type CheckBox}">
      <Setter Property="Foreground" Value="#CFE0F7"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type CheckBox}">
            <StackPanel Orientation="Horizontal">
              <Border x:Name="Box" Width="14" Height="14" CornerRadius="4"
                      Background="Transparent" BorderBrush="#4A5A70" BorderThickness="1" Margin="0,1,6,0">
                <Path x:Name="CheckMark" Data="M 2 7 L 6 11 L 12 3"
                      Stroke="#CFE0F7" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                      Visibility="Collapsed" />
              </Border>
              <ContentPresenter VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Box" Property="Background" Value="#2382E8"/>
                <Setter TargetName="Box" Property="BorderBrush" Value="#2382E8"/>
                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Box" Property="Background" Value="#1E2A38"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid Margin="14">
    <Border CornerRadius="22" Background="#151B22" BorderBrush="#223041" BorderThickness="1">
      <Grid Margin="22">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="18"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

          <StackPanel>
            <TextBlock x:Name="TimeText" Text="--:--" FontSize="46" FontWeight="SemiBold" Foreground="#EAF2FF"/>
            <TextBlock x:Name="GreetText" Text="Good Morning, İrem" FontSize="16" Foreground="#9FB0C6" Margin="2,4,0,0"/>
          </StackPanel>

          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top" HorizontalAlignment="Right">
            <Border CornerRadius="18" Background="#1E2A38" BorderBrush="#2B3A4E" BorderThickness="1" Padding="12,8" Margin="0,0,10,0">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="⌨" FontSize="14" Foreground="#CFE0F7" Margin="0,0,8,0"/>
                <TextBlock Text="Ctrl+Alt+D" FontSize="14" Foreground="#CFE0F7"/>
              </StackPanel>
            </Border>

            <Button x:Name="CloseBtn" Width="42" Height="36" Background="#1E2A38" Foreground="#CFE0F7" BorderBrush="#2B3A4E" BorderThickness="1"
                    FontSize="14" Content="✕" Padding="0" />
          </StackPanel>
        </Grid>

        <!-- Body -->
        <Grid Grid.Row="2">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="340"/>
            <ColumnDefinition Width="22"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="22"/>
            <ColumnDefinition Width="360"/>
          </Grid.ColumnDefinitions>

          <!-- LEFT: Calendar + Quick Notes -->
          <StackPanel Grid.Column="0">
            <!-- Calendar card (reference look) -->
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16" Margin="0,0,0,16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="14"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="10"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Header row -->
                <Grid Grid.Row="0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="40"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="40"/>
                  </Grid.ColumnDefinitions>
                  <Button x:Name="PrevMonthBtn" Style="{StaticResource IconBtn}" Content="‹"/>
                  <TextBlock x:Name="CalTitle" Grid.Column="1" Text="Month YYYY" Foreground="#EAF2FF" FontSize="16" FontWeight="SemiBold"
                             HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  <Button x:Name="NextMonthBtn" Grid.Column="2" Style="{StaticResource IconBtn}" Content="›"/>
                </Grid>

                <!-- Day names -->
                <UniformGrid Grid.Row="2" Columns="7">
                  <TextBlock Text="S" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="M" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="T" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="W" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="T" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="F" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                  <TextBlock Text="S" Foreground="#6F839A" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                </UniformGrid>

                <!-- Days grid -->
                <UniformGrid Grid.Row="4" x:Name="DaysGrid" Columns="7" Rows="6" />
              </Grid>
            </Border>

            <!-- Quick Notes card -->
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="12"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="8"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="🗒" FontSize="14" Foreground="#F2B84B" Margin="0,0,10,0"/>
                    <TextBlock Text="Quick Notes" Foreground="#EAF2FF" FontSize="16" FontWeight="SemiBold"/>
                  </StackPanel>
                  <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="NewNoteBtn" Style="{StaticResource IconBtn}" Content="＋" Width="28" Height="28" Margin="0,0,4,0" ToolTip="Yeni not (bugün)"/>
                    <Button x:Name="NewNoteForSelectedBtn" Style="{StaticResource IconBtn}" Content="📅" Width="28" Height="28" Margin="0,0,4,0" ToolTip="Yeni notu seçili güne kaydet" Visibility="Collapsed"/>
                    <Button x:Name="DeleteNoteBtn" Style="{StaticResource IconBtn}" Content="🗑" Width="28" Height="28"/>
                  </StackPanel>
                </Grid>

                <Grid Grid.Row="2">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>

                  <Border Grid.Column="0" Background="#151B22" BorderBrush="#223041" BorderThickness="1" CornerRadius="10">
                    <ListBox x:Name="NotesList"
                             Style="{StaticResource NotesListStyle}">
                    </ListBox>
                  </Border>

                  <TextBox x:Name="NotesBox" Grid.Column="2" Style="{StaticResource NotesBoxStyle}"/>
                </Grid>

                <TextBlock Grid.Row="4" x:Name="NoteMetaText" Text="AUTO-SAVED" Foreground="#6F839A" FontSize="11" HorizontalAlignment="Right"/>
              </Grid>
            </Border>
          </StackPanel>

          <!-- CENTER: Pomodoro + System Status + Tasks -->
          <StackPanel Grid.Column="2">
            <!-- Pomodoro -->
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16" Margin="0,0,0,16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="10"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="12"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="14"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="⏱" FontSize="14" Foreground="#F2B84B" Margin="0,0,8,0"/>
                    <TextBlock Text="Focus / Pomodoro" Foreground="#EAF2FF" FontSize="16" FontWeight="SemiBold"/>
                  </StackPanel>
                  <Border Grid.Column="1" Padding="8,2" Background="#131820" CornerRadius="10" VerticalAlignment="Center">
                    <TextBlock x:Name="PomodoroMode" Text="IDLE" Foreground="#BFE3FF" FontSize="12" FontWeight="SemiBold"/>
                  </Border>
                </Grid>

                <TextBlock Grid.Row="2" x:Name="PomodoroTime" Text="25:00" FontSize="36" FontWeight="SemiBold" Foreground="#EAF2FF"/>

                <TextBlock Grid.Row="4" Text="Stay focused in short sprints" Foreground="#8FA3BC" FontSize="12"/>

                <ProgressBar Grid.Row="5" x:Name="PomodoroBar" Height="6" Minimum="0" Maximum="1" Value="0"
                             Margin="0,10,0,0"
                             Background="#131820" Foreground="#4BA3FF" BorderThickness="0"/>

                <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                  <Button x:Name="FocusBtn" Content="Start Focus" Style="{StaticResource PrimaryActionBtn}"/>
                  <Button x:Name="BreakBtn" Content="Break" Style="{StaticResource SecondaryActionBtn}"/>
                  <Button x:Name="ResetBtn" Content="Reset" Style="{StaticResource SecondaryActionBtn}" Margin="0,0,0,0"/>
                </StackPanel>
              </Grid>
            </Border>

            <!-- System Status -->
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16" Margin="0,0,0,16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="10"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="8"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="8"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Text="SYSTEM STATUS" Foreground="#8FA3BC" FontSize="12" FontWeight="SemiBold"/>

                <!-- CPU -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Stretch" VerticalAlignment="Center">
                  <TextBlock Text="CPU" Foreground="#CFE0F7" FontSize="12"/>
                  <TextBlock x:Name="SysCpuText" Text="-" Foreground="#CFE0F7" FontSize="12" HorizontalAlignment="Right" Margin="8,0,0,0"/>
                </StackPanel>
                <ProgressBar Grid.Row="3" x:Name="SysCpuBar" Height="6" Minimum="0" Maximum="100" Value="0"
                             Margin="0,4,0,0" Background="#131820" Foreground="#4BA3FF" BorderThickness="0"/>

                <!-- RAM -->
                <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Stretch" VerticalAlignment="Center">
                  <TextBlock Text="RAM" Foreground="#CFE0F7" FontSize="12"/>
                  <TextBlock x:Name="SysMemText" Text="-" Foreground="#CFE0F7" FontSize="12" HorizontalAlignment="Right" Margin="8,0,0,0"/>
                </StackPanel>
                <ProgressBar Grid.Row="6" x:Name="SysMemBar" Height="6" Minimum="0" Maximum="100" Value="0"
                             Margin="0,4,0,0" Background="#131820" Foreground="#4BA3FF" BorderThickness="0"/>

                <!-- Meta: battery + network -->
                <TextBlock Grid.Row="8" x:Name="SysMetaText" Text="Battery: -   •   Network: -"
                           Foreground="#8FA3BC" FontSize="11" TextTrimming="CharacterEllipsis"/>
              </Grid>
            </Border>

            <!-- Mini To-Do / Tasks -->
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="10"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="8"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="☑" FontSize="14" Foreground="#F2B84B" Margin="0,0,8,0"/>
                    <TextBlock Text="TODAY'S TASKS" Foreground="#EAF2FF" FontSize="16" FontWeight="SemiBold"/>
                  </StackPanel>
                  <Button x:Name="TodoAddBtn" Grid.Column="1" Style="{StaticResource IconBtn}" Content="＋" Width="28" Height="28"/>
                </Grid>

                    <Border Grid.Row="2" x:Name="TodoInputContainer"
                    Background="#151B22" BorderBrush="#223041" BorderThickness="1"
                    CornerRadius="8" Padding="6,3" Visibility="Collapsed">
                  <TextBox x:Name="TodoInput"
                           Background="Transparent" BorderThickness="0"
                           Foreground="#CFE0F7" FontSize="13"
                           CaretBrush="#CFE0F7"
                           VerticalContentAlignment="Center"
                           ToolTip="Yeni görev yazıp Enter'a basın"/>
                    </Border>

                <ScrollViewer Grid.Row="4" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                  <StackPanel x:Name="TodoList" />
                </ScrollViewer>
              </Grid>
            </Border>
          </StackPanel>

          <!-- RIGHT: Unread Mails -->
          <StackPanel Grid.Column="4">
            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16" Margin="0,0,0,16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="12"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Text="UNREAD MAILS" Foreground="#EAF2FF" FontSize="16" FontWeight="SemiBold"/>
                  <Button x:Name="RefreshMailBtn" Grid.Column="1" Content="↻" Style="{StaticResource IconBtn}" Margin="8,0,0,0"/>
                </Grid>

                <StackPanel Grid.Row="2" x:Name="MailList" />
              </Grid>
            </Border>

            <Border CornerRadius="18" Background="#1A2028" BorderBrush="#223041" BorderThickness="1" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="18"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Text="QUICK LAUNCH" Foreground="#8FA3BC" FontSize="12" FontWeight="SemiBold"/>

                <!-- 6 küçük ikon tek satır, ortalanmış -->
                <UniformGrid Grid.Row="2" Columns="6" Margin="0,4,0,0" HorizontalAlignment="Center">
                  <!-- Outlook -->
                  <Button x:Name="QlOutlook" Margin="0,0,10,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#2F4B78" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="📧" FontSize="18" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="Outlook" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>

                  <!-- Teams -->
                  <Button x:Name="QlTeams" Margin="0,0,10,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#51338E" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="👥" FontSize="18" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="Teams" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>

                  <!-- File Explorer -->
                  <Button x:Name="QlExplorer" Margin="0,0,10,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#6C4720" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="📁" FontSize="18" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="Explorer" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>

                  <!-- Edge -->
                  <Button x:Name="QlEdge" Margin="0,0,10,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#14686F" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="🌐" FontSize="16" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="Edge" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>

                  <!-- VS Code -->
                  <Button x:Name="QlVSCode" Margin="0,0,10,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#0B6FA4" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="&lt;/&gt;" FontSize="16" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="VS Code" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>

                  <!-- Excel -->
                  <Button x:Name="QlExcel" Margin="0,0,0,0" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                    <StackPanel>
                      <Border CornerRadius="12" Background="#17653E" Width="40" Height="40" HorizontalAlignment="Center">
                        <TextBlock Text="X" FontSize="16" FontWeight="Bold" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                      </Border>
                      <TextBlock Text="Excel" Foreground="#CFE0F7" FontSize="10" Margin="0,4,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                  </Button>
                </UniformGrid>
              </Grid>
            </Border>
          </StackPanel>

        </Grid>
      </Grid>
    </Border>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$win = [Windows.Markup.XamlReader]::Load($reader)

# ------- UI refs -------
$TimeText = $win.FindName("TimeText")
$CloseBtn = $win.FindName("CloseBtn")

# Calendar refs
$PrevMonthBtn = $win.FindName("PrevMonthBtn")
$NextMonthBtn = $win.FindName("NextMonthBtn")
$CalTitle = $win.FindName("CalTitle")
$DaysGrid = $win.FindName("DaysGrid")

# Notes
$NotesBox = $win.FindName("NotesBox")
$NotesList = $win.FindName("NotesList")
$NewNoteBtn = $win.FindName("NewNoteBtn")
$NewNoteForSelectedBtn = $win.FindName("NewNoteForSelectedBtn")
$DeleteNoteBtn = $win.FindName("DeleteNoteBtn")
$NoteMetaText = $win.FindName("NoteMetaText")

# Pomodoro
$PomodoroMode = $win.FindName("PomodoroMode")
$PomodoroTime = $win.FindName("PomodoroTime")
$PomodoroBar  = $win.FindName("PomodoroBar")
$FocusBtn = $win.FindName("FocusBtn")
$BreakBtn = $win.FindName("BreakBtn")
$ResetBtn = $win.FindName("ResetBtn")

# System Status
$SysCpuText  = $win.FindName("SysCpuText")
$SysCpuBar   = $win.FindName("SysCpuBar")
$SysMemText  = $win.FindName("SysMemText")
$SysMemBar   = $win.FindName("SysMemBar")
$SysMetaText = $win.FindName("SysMetaText")

# Tasks
$TodoInputContainer = $win.FindName("TodoInputContainer")
$TodoInput  = $win.FindName("TodoInput")
$TodoList   = $win.FindName("TodoList")
$TodoAddBtn = $win.FindName("TodoAddBtn")

# Mail
$MailList = $win.FindName("MailList")
$RefreshMailBtn = $win.FindName("RefreshMailBtn")

# Quick Launch
$QlOutlook  = $win.FindName("QlOutlook")
$QlTeams    = $win.FindName("QlTeams")
$QlExplorer = $win.FindName("QlExplorer")
$QlEdge     = $win.FindName("QlEdge")
$QlVSCode   = $win.FindName("QlVSCode")
$QlExcel    = $win.FindName("QlExcel")

# ---------------- Notes model & helpers ----------------
$global:Notes = @(Load-NotesData)
$global:CurrentNote = $null
$global:NotesUpdatingFromCode = $false
$global:NoteRelocateNoteId = $null  # takvimden "ait olduğu tarihi" güncellemek için geçici id

# ---------------- Tasks model & helpers ----------------
$global:Todos = @()
$global:EditingTodoId = $null

function Ensure-TodoFile {
  if(!(Test-Path $TodoPath)){
    "[]" | Out-File -FilePath $TodoPath -Encoding utf8
  }
}

function Load-TodosData {
  Ensure-TodoFile
  try {
    $raw = Get-Content $TodoPath -Raw -ErrorAction Stop
    $obj = $raw | ConvertFrom-Json
    if($null -eq $obj){ return @() }
    if($obj -is [System.Collections.IEnumerable]){ return @($obj) }
    return @()
  } catch { return @() }
}

function Save-TodosData([array]$todos){
  try {
    ($todos | ConvertTo-Json -Depth 5) | Out-File -FilePath $TodoPath -Encoding utf8
  } catch {}
}

function Get-NextTodoId {
  if(-not $global:Todos -or $global:Todos.Count -eq 0){ return 1 }
  return ( ($global:Todos | Measure-Object -Property id -Maximum).Maximum + 1 )
}

function Begin-EditTodo($id){
  if(-not $TodoInput){ return }
  $todo = $global:Todos | Where-Object { $_.id -eq $id } | Select-Object -First 1
  if(-not $todo){ return }

  $global:EditingTodoId = $id
  $TodoInput.Text = [string]$todo.text
  if($TodoInputContainer){
    $TodoInputContainer.Visibility = [System.Windows.Visibility]::Visible
  }
  $TodoInput.Focus() | Out-Null
  $TodoInput.CaretIndex = $TodoInput.Text.Length
}

function Render-TodoList {
  if(-not $TodoList){ return }
  $TodoList.Children.Clear()

  foreach($t in ($global:Todos | Sort-Object createdAt)){
    $row = New-Object System.Windows.Controls.Border
    $row.CornerRadius = "8"
    $row.Background = "#151B22"
    $row.Padding = "6,4"
    $row.Margin = "0,0,0,4"

    $inner = New-Object System.Windows.Controls.StackPanel
    $inner.Orientation = "Horizontal"

    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.IsChecked = [bool]$t.done
    $cb.Margin = "0,0,6,0"
    $cb.Tag = $t.id
    try { $cb.Style = $win.FindResource("TodoCheckBoxStyle") } catch {}

    $txt = New-Object System.Windows.Controls.TextBlock
    $txt.Text = $t.text
    $txt.Foreground = if($t.done){ "#6F839A" } else { "#CFE0F7" }
    if($t.done){ $txt.TextDecorations = [System.Windows.TextDecorations]::Strikethrough }
    $txt.Tag = $t.id
    $txt.Cursor = [System.Windows.Input.Cursors]::IBeam

    $inner.Children.Add($cb) | Out-Null
    $inner.Children.Add($txt) | Out-Null

    $cb.add_Checked({
      param($sender,$args)
      $id = $sender.Tag
      $todo = $global:Todos | Where-Object { $_.id -eq $id } | Select-Object -First 1
      if($todo){ $todo.done = $true; Save-TodosData $global:Todos; Render-TodoList }
    })

    $txt.add_MouseLeftButtonUp({
      param($sender,$args)
      $id = $sender.Tag
      if($id){ Begin-EditTodo $id }
    })
    $cb.add_Unchecked({
      param($sender,$args)
      $id = $sender.Tag
      $todo = $global:Todos | Where-Object { $_.id -eq $id } | Select-Object -First 1
      if($todo){ $todo.done = $false; Save-TodosData $global:Todos; Render-TodoList }
    })

    $row.Child = $inner
    $TodoList.Children.Add($row) | Out-Null
  }
}

function Add-TodoFromInput {
  if(-not $TodoInput){ return }
  $text = ($TodoInput.Text).Trim()
  if([string]::IsNullOrWhiteSpace($text)){
    # Boş metin: yeni görev ekleme, düzenleme modunda ise sil
    if($global:EditingTodoId -ne $null){
      $global:Todos = @($global:Todos | Where-Object { $_.id -ne $global:EditingTodoId })
      Save-TodosData $global:Todos
      Render-TodoList
    }
  } else {
    if($global:EditingTodoId -ne $null){
      # Mevcut görevi güncelle
      $todo = $global:Todos | Where-Object { $_.id -eq $global:EditingTodoId } | Select-Object -First 1
      if($todo){ $todo.text = $text }
      Save-TodosData $global:Todos
      Render-TodoList
    } else {
      # Yeni görev ekle
      $todo = [pscustomobject]@{
        id = Get-NextTodoId
        text = $text
        done = $false
        createdAt = Get-Date
      }
      $global:Todos += $todo
      Save-TodosData $global:Todos
      Render-TodoList
    }
  }

  $TodoInput.Text = ""
  $global:EditingTodoId = $null
  if($TodoInputContainer){ $TodoInputContainer.Visibility = 'Collapsed' }
}

function Get-NextNoteId {
  if(-not $global:Notes -or $global:Notes.Count -eq 0){ return 1 }
  return ( ($global:Notes | Measure-Object -Property id -Maximum).Maximum + 1 )
}

function Convert-ToLocalDateOnly([object]$value){
  if($null -eq $value){ return $null }

  # Zaten DateTime ise doğrudan yerel tarihe indir
  if($value -is [datetime]){
    if($value.Kind -eq [System.DateTimeKind]::Utc){
      return ($value).ToLocalTime().Date
    }
    return ($value).Date
  }

  $s = [string]$value

  # JSON /Date(XXXXXXXXXXXX)/ formatını yakala (ms since Unix epoch)
  if($s -match "\\/Date\\((\d+)\\)\\/"){
    try {
      $ms = [int64]$matches[1]
      $dto = [System.DateTimeOffset]::FromUnixTimeMilliseconds($ms)
      # Değeri senin yerel zaman dilimine (GMT+03:00) çevirip
      # sadece tarih kısmını alıyoruz; böylece JSON'da saklanan
      # tarih hangi makinede okunursa okunsun, takvimde yerel gün
      # olarak doğru hücreye oturuyor.
      return $dto.ToLocalTime().Date
    } catch {
      # devam et, aşağıdaki genel parse'a düşsün
    }
  }

  # ISO veya diğer tanınan formatlar için genel dönüşüm
  try {
    return ([datetime]$value).Date
  } catch {
    return $null
  }
}

function Get-NoteDate($note){
  if($null -eq $note){ return $null }
  if($note.PSObject.Properties["date"]){
    return (Convert-ToLocalDateOnly $note.date)
  }
  if($note.PSObject.Properties["createdAt"]){
    return (Convert-ToLocalDateOnly $note.createdAt)
  }
  return $null
}

function Get-NotesForCurrentDay {
  $target = $global:CalSelected.Date
  return @($global:Notes | Where-Object {
    $d = Get-NoteDate $_
    $d -ne $null -and $d -eq $target
  })
}

function New-NoteForDate([datetime]$day){
  $now = Get-Date
  return [pscustomobject]@{
    id        = Get-NextNoteId
    title     = "Yeni Not"
    text      = ""
    createdAt = $now
    updatedAt = $now
    date      = $day.Date
  }
}

function Update-NotesAreaVisibility {
  if(-not $NotesBox -or -not $NotesList){ return }

  $dayNotes = Get-NotesForCurrentDay
  if($dayNotes.Count -eq 0){
    $NotesBox.Visibility = [System.Windows.Visibility]::Collapsed
    $NotesList.Visibility = [System.Windows.Visibility]::Collapsed
    if($NoteMetaText){ $NoteMetaText.Text = "" }
  } else {
    $NotesBox.Visibility = [System.Windows.Visibility]::Visible
    $NotesList.Visibility = [System.Windows.Visibility]::Visible
  }

  if($NewNoteForSelectedBtn){
    # Yeni davranış: mevcutta bir not seçiliyse takvim butonu görünsün,
    # ilk gün seçimi bu notun "ait olduğu tarih" alanını güncelleyecek.
    $NewNoteForSelectedBtn.Visibility = if($global:CurrentNote){ [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
  }
}

function Update-NotesListUI {
  $global:NotesUpdatingFromCode = $true
  $currentId = $null
  if($NotesList.SelectedItem){ $currentId = $NotesList.SelectedItem.Tag }

  $NotesList.Items.Clear()
  $dayNotes = Get-NotesForCurrentDay | Sort-Object createdAt -Descending
  foreach($n in $dayNotes){
    $item = New-Object System.Windows.Controls.ListBoxItem
    $item.Content = if([string]::IsNullOrWhiteSpace($n.title)){"Not $($n.id)"}else{$n.title}
    $item.Tag = $n.id
    $NotesList.Items.Add($item) | Out-Null
  }

  if($currentId -ne $null){
    $exists = $dayNotes | Where-Object { $_.id -eq $currentId } | Select-Object -First 1
    if($exists){
      for($i=0; $i -lt $NotesList.Items.Count; $i++){
        if($NotesList.Items[$i].Tag -eq $currentId){ $NotesList.SelectedIndex = $i; break }
      }
    } else {
      $NotesList.SelectedIndex = -1
    }
  }

  $global:NotesUpdatingFromCode = $false
  Update-NotesAreaVisibility
}

function Set-CurrentNote($note){
  $global:CurrentNote = $note
  if($null -eq $note){
    $NotesBox.Text = ""
    $NoteMetaText.Text = "No note selected"
    Update-NotesAreaVisibility
    return
  }

  $NotesBox.Text = $note.text
  $NoteMetaText.Text = "Created: {0}" -f ([datetime]$note.createdAt).ToString("dd.MM.yyyy HH:mm")

    $global:NotesUpdatingFromCode = $false
  for($i=0; $i -lt $NotesList.Items.Count; $i++){
    if($NotesList.Items[$i].Tag -eq $note.id){ $NotesList.SelectedIndex = $i; break }
  }
  $global:NotesUpdatingFromCode = $false
  Update-NotesAreaVisibility
}

function Save-CurrentNote {
  if($null -eq $global:CurrentNote){ return }
  $global:CurrentNote.text = $NotesBox.Text

  $firstLine = ($NotesBox.Text -split "(`r`n|`n)" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -First 1)
  if([string]::IsNullOrWhiteSpace($firstLine)){
    $global:CurrentNote.title = "Not $($global:CurrentNote.id)"
  } else {
    $global:CurrentNote.title = $firstLine
  }

  $global:CurrentNote.updatedAt = Get-Date
  Save-NotesData $global:Notes
  Update-NotesListUI
  Update-NotesAreaVisibility
}

function Ensure-AtLeastOneNote {
  if($global:Notes.Count -eq 0){
    $today = (Get-Date).Date
    $note = New-NoteForDate $today
    $global:Notes = @($global:Notes) + @($note)
    Save-NotesData $global:Notes
  }
}

Ensure-AtLeastOneNote
Update-NotesListUI
$firstForDay = Get-NotesForCurrentDay | Sort-Object createdAt -Descending | Select-Object -First 1
if($firstForDay){
  Set-CurrentNote $firstForDay
} else {
  Set-CurrentNote $null
}

# ---------------- Notes autosave (debounced) ----------------
$saveTimer = New-Object System.Windows.Threading.DispatcherTimer
$saveTimer.Interval = [TimeSpan]::FromMilliseconds(700)
$saveTimer.Add_Tick({
  $saveTimer.Stop()
  Save-CurrentNote
})
$NotesBox.Add_TextChanged({
  $saveTimer.Stop()
  $saveTimer.Start()
})

$NotesList.Add_SelectionChanged({
  if($global:NotesUpdatingFromCode){ return }

  $saveTimer.Stop()
  Save-CurrentNote

  $item = $NotesList.SelectedItem
  if($null -eq $item){
    Set-CurrentNote $null
    return
  }

  $id = $item.Tag
  $note = $global:Notes | Where-Object { $_.id -eq $id } | Select-Object -First 1
  Set-CurrentNote $note
})

$NewNoteBtn.Add_Click({
  $saveTimer.Stop()
  Save-CurrentNote

  $today = (Get-Date).Date
  $note = New-NoteForDate $today
  $global:Notes = @($global:Notes) + @($note)
  Save-NotesData $global:Notes

  $global:CalDisplay = Get-Date -Year $today.Year -Month $today.Month -Day 1
  $global:CalSelected = $today
  Render-Calendar
  Update-NotesListUI
  Set-CurrentNote $note
})

if($NewNoteForSelectedBtn){
  $NewNoteForSelectedBtn.Add_Click({
    # Mevcut seçili notun "ait olduğu tarih" alanını değiştirmek için
    # bir sonraki takvim günü tıklamasını yakala.
    if($null -eq $global:CurrentNote){ return }
    $saveTimer.Stop()
    Save-CurrentNote

    $global:NoteRelocateNoteId = $global:CurrentNote.id
    if($NoteMetaText){
      $NoteMetaText.Text = "Takvimden yeni bir gün seçin (ait olduğu tarih güncellenecek)."
    }
  })
}

$DeleteNoteBtn.Add_Click({
  $saveTimer.Stop()
  if($null -eq $global:CurrentNote){ return }

  $idToRemove = $global:CurrentNote.id
  $global:Notes = @($global:Notes | Where-Object { $_.id -ne $idToRemove })
  Save-NotesData $global:Notes
  Update-NotesListUI

  if($global:Notes.Count -gt 0){
    $next = Get-NotesForCurrentDay | Sort-Object createdAt -Descending | Select-Object -First 1
    if($next){
      Set-CurrentNote $next
    } else {
      Set-CurrentNote $null
    }
  } else {
    Set-CurrentNote $null
  }
})

# ---------------- Pomodoro ----------------
function Update-PomodoroUI {
  $PomodoroMode.Text = $global:Pomodoro.Mode
  $PomodoroTime.Text = Format-Time $global:Pomodoro.Remaining
  $progress = 1.0 - ($global:Pomodoro.Remaining / [double]$global:Pomodoro.Total)
  if($progress -lt 0){ $progress = 0 }
  if($progress -gt 1){ $progress = 1 }
  $PomodoroBar.Value = $progress

  if($global:Pomodoro.Mode -eq "FOCUS"){
    $FocusBtn.Content = $(if($global:Pomodoro.Running){"Pause"} else {"Resume"})
  } else {
    $FocusBtn.Content = "Start Focus"
  }
}

function Start-Focus {
  if($global:Pomodoro.Mode -ne "FOCUS"){
    $global:Pomodoro.Mode = "FOCUS"
    $global:Pomodoro.Total = 25*60
    $global:Pomodoro.Remaining = $global:Pomodoro.Total
    $global:Pomodoro.Running = $true
  } else {
    $global:Pomodoro.Running = -not $global:Pomodoro.Running
  }
  Update-PomodoroUI
}

function Start-Break {
  $global:Pomodoro.Mode = "BREAK"
  $global:Pomodoro.Total = 5*60
  $global:Pomodoro.Remaining = $global:Pomodoro.Total
  $global:Pomodoro.Running = $true
  Update-PomodoroUI
}

function Reset-Pomodoro {
  $global:Pomodoro.Mode = "IDLE"
  $global:Pomodoro.Total = 25*60
  $global:Pomodoro.Remaining = $global:Pomodoro.Total
  $global:Pomodoro.Running = $false
  Update-PomodoroUI
}

function Get-SystemStatus {
  $cpu = $null
  $mem = $null
  $battery = $null
  $netOnline = $false

  try {
    $c = Get-Counter '\Processor(_Total)\% Processor Time'
    $cpu = [math]::Round($c.CounterSamples.CookedValue, 1)
  } catch {}

  try {
    $os = Get-CimInstance Win32_OperatingSystem
    $total = [double]$os.TotalVisibleMemorySize * 1KB
    $free  = [double]$os.FreePhysicalMemory * 1KB
    $used  = $total - $free
    if($total -gt 0){
      $memPct = [math]::Round(($used / $total) * 100, 1)
    } else {
      $memPct = 0
    }
    $mem = [pscustomobject]@{
      UsedGB  = [math]::Round($used  / 1GB, 1)
      TotalGB = [math]::Round($total / 1GB, 1)
      Percent = $memPct
    }
  } catch {}

  try {
    $b = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    if($b){
      $battery = [pscustomobject]@{
        Percent = [int]$b.EstimatedChargeRemaining
        Status  = $b.BatteryStatus
      }
    }
  } catch {}

  try {
    $netOnline = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
  } catch {}

  return [pscustomobject]@{
    CpuPercent   = $cpu
    Memory       = $mem
    Battery      = $battery
    NetworkOnline = $netOnline
  }
}

function Update-SystemStatusUI {
  if(-not $SysCpuText){ return }

  $s = Get-SystemStatus

  # CPU
  if($null -ne $s.CpuPercent){
    $SysCpuText.Text = "{0:N1}%" -f $s.CpuPercent
    $SysCpuBar.Value = [double]$s.CpuPercent
  } else {
    $SysCpuText.Text = "-"
    $SysCpuBar.Value = 0
  }

  # Memory
  if($null -ne $s.Memory){
    $SysMemText.Text = "{0:N1} / {1:N1} GB  ({2:N1}%)" -f $s.Memory.UsedGB, $s.Memory.TotalGB, $s.Memory.Percent
    $SysMemBar.Value = [double]$s.Memory.Percent
  } else {
    $SysMemText.Text = "-"
    $SysMemBar.Value = 0
  }

  # Battery & network
  $batteryText = "No battery"
  if($null -ne $s.Battery){
    $batteryText = "{0}%" -f $s.Battery.Percent
  }

  $netText = if($s.NetworkOnline){ "Online" } else { "Offline" }

  $SysMetaText.Text = "Battery: {0}   •   Network: {1}" -f $batteryText, $netText
}

# ---------------- Mail UI ----------------
function Clear-MailList { $MailList.Children.Clear() }

function Add-MailRow($m){
  $row = New-Object System.Windows.Controls.Border
  $row.CornerRadius = "14"
  $row.Background = "#131820"
  $row.BorderBrush = "#223041"
  $row.BorderThickness = "1"
  $row.Padding = "12"
  $row.Margin = "0,0,0,10"
  $row.Cursor = [System.Windows.Input.Cursors]::Hand

  $grid = New-Object System.Windows.Controls.Grid
  $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
  $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" })) | Out-Null

  $stack = New-Object System.Windows.Controls.StackPanel

  $sender = New-Object System.Windows.Controls.TextBlock
  $sender.Text = $m.Sender
  $sender.Foreground = "#CFE0F7"
  $sender.FontWeight = "SemiBold"

  $subj = New-Object System.Windows.Controls.TextBlock
  $subj.Text = $m.Subject
  $subj.Foreground = "#8FA3BC"
  $subj.TextTrimming = "CharacterEllipsis"
  $subj.MaxWidth = 250
  $subj.Margin = "0,4,0,0"

  $time = New-Object System.Windows.Controls.TextBlock
  $time.Text = ([datetime]$m.Received).ToString("HH:mm")
  $time.Foreground = "#6F839A"
  $time.VerticalAlignment = "Top"
  [System.Windows.Controls.Grid]::SetColumn($time,1)

  $stack.Children.Add($sender) | Out-Null
  $stack.Children.Add($subj) | Out-Null

  $grid.Children.Add($stack) | Out-Null
  $grid.Children.Add($time) | Out-Null

  $row.Child = $grid
  $row.Tag = $m.Item
  $row.add_MouseLeftButtonUp({
    param($sender, $args)
    $item = $sender.Tag
    if($item){ Open-MailItem $item }
  })

  $row.add_MouseEnter({ param($sender, $args) $sender.Background = "#1E2A38" })
  $row.add_MouseLeave({ param($sender, $args) $sender.Background = "#131820" })

  $MailList.Children.Add($row) | Out-Null
}

function Update-MailUI {
  Clear-MailList
  $mails = Get-UnreadMailsTop5
  if($mails -eq $null){
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = "Outlook unread mail data unavailable. (Outlook not installed or COM restricted by policy.)"
    $tb.Foreground = "#8FA3BC"
    $tb.TextWrapping = "Wrap"
    $MailList.Children.Add($tb) | Out-Null
    return
  }
  if($mails.Count -eq 0){
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = "No unread mails 🎉"
    $tb.Foreground = "#8FA3BC"
    $MailList.Children.Add($tb) | Out-Null
    return
  }
  foreach($m in $mails){ Add-MailRow $m }
}

# --------- Calendar rendering (custom) ---------
function Ensure-CalendarButtons {
  if($global:CalButtons.Count -eq 42){ return }
  $DaysGrid.Children.Clear()
  $global:CalButtons = @()
  for($i=0; $i -lt 42; $i++){
    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $win.FindResource("CalDayBtn")
    $btn.Content = ""
    $btn.Tag = $null
    $btn.Add_Click({
      $d = $this.Tag
      if($d -eq $null){ return }

      $saveTimer.Stop()
      Save-CurrentNote

      # Eğer bir not için "ait olduğu tarih" değiştirme modu aktifse,
      # ilk gün seçimi o notun date alanını günceller.
      if($global:NoteRelocateNoteId -ne $null){
        $note = $global:Notes | Where-Object { $_.id -eq $global:NoteRelocateNoteId } | Select-Object -First 1
        $global:NoteRelocateNoteId = $null

        if($note){
          $newDate = ([datetime]$d).Date
          $note.date = $newDate
          $note.updatedAt = Get-Date
          Save-NotesData $global:Notes

          $global:CalSelected = $newDate
          $global:CalDisplay  = Get-Date -Year $newDate.Year -Month $newDate.Month -Day 1
          Render-Calendar
          Update-NotesListUI
          Set-CurrentNote $note
        } else {
          # Not bulunamazsa, normal davranışa geri dön
          $sel = ([datetime]$d).Date
          $global:CalSelected = $sel
          $global:CalDisplay  = Get-Date -Year $sel.Year -Month $sel.Month -Day 1
          Render-Calendar
          Update-NotesListUI

          $dayNotes = Get-NotesForCurrentDay | Sort-Object createdAt -Descending
          if($dayNotes.Count -gt 0){
            Set-CurrentNote ($dayNotes | Select-Object -First 1)
          } else {
            Set-CurrentNote $null
          }
        }
      } else {
        # Normal tıklama: seçili günü değiştir, o güne ait en son notu aç
        $sel = ([datetime]$d).Date
        $global:CalSelected = $sel
        $global:CalDisplay  = Get-Date -Year $sel.Year -Month $sel.Month -Day 1
        Render-Calendar
        Update-NotesListUI

        $dayNotes = Get-NotesForCurrentDay | Sort-Object createdAt -Descending
        if($dayNotes.Count -gt 0){
          Set-CurrentNote ($dayNotes | Select-Object -First 1)
        } else {
          Set-CurrentNote $null
        }
      }
    })
    $DaysGrid.Children.Add($btn) | Out-Null
    $global:CalButtons += $btn
  }
}

function Render-Calendar {
  Ensure-CalendarButtons

  $CalTitle.Text = $global:CalDisplay.ToString("MMMM yyyy")

  $first = $global:CalDisplay
  $offset = [int]$first.DayOfWeek  # Sunday=0
  $start = $first.AddDays(-$offset)

  $today = (Get-Date).Date

  for($i=0; $i -lt 42; $i++){
    $d = $start.AddDays($i)
    $btn = $global:CalButtons[$i]
    $btn.Tag = $d
    $btn.Content = $d.Day.ToString()

    $inMonth = ($d.Month -eq $first.Month) -and ($d.Year -eq $first.Year)

    $btn.IsEnabled = $true
    if($inMonth){
      $btn.Foreground = "#CFE0F7"
    } else {
      $btn.Foreground = "#3E4D60"
    }
    $btn.Background = "Transparent"

    if($d.Date -eq $today){
      $btn.Background = "#162536"
    }
    if($d.Date -eq $global:CalSelected.Date){
      $btn.Background = "#2382E8"
      $btn.Foreground = "White"
    }
  }
}

$PrevMonthBtn.Add_Click({
  $global:CalDisplay = $global:CalDisplay.AddMonths(-1)
  Render-Calendar
})
$NextMonthBtn.Add_Click({
  $global:CalDisplay = $global:CalDisplay.AddMonths(1)
  Render-Calendar
})

# ------- Buttons -------
$CloseBtn.Add_Click({ 
  # ensure notes saved before hide
  $saveTimer.Stop()
  Save-CurrentNote
  $win.Hide()
})
$FocusBtn.Add_Click({ Start-Focus })
$BreakBtn.Add_Click({ Start-Break })
$ResetBtn.Add_Click({ Reset-Pomodoro })
$RefreshMailBtn.Add_Click({ Update-MailUI })

if($TodoAddBtn){
  $TodoAddBtn.Add_Click({
    if(-not $TodoInputContainer){ return }
    if($TodoInputContainer.Visibility -eq [System.Windows.Visibility]::Visible){
      Add-TodoFromInput
    } else {
      $TodoInputContainer.Visibility = [System.Windows.Visibility]::Visible
      if($TodoInput){ $TodoInput.Focus() | Out-Null }
    }
  })
}
if($TodoInput){
  $TodoInput.Add_KeyDown({
    param($sender,$e)
    if($e.Key -eq 'Enter'){
      Add-TodoFromInput
      $e.Handled = $true
    } elseif($e.Key -eq 'Escape'){
      $TodoInput.Text = ""
      if($TodoInputContainer){ $TodoInputContainer.Visibility = 'Collapsed' }
      $e.Handled = $true
    }
  })
}

# ------- Quick Launch actions -------
function Start-ProcessSafe([string]$target, [string]$args=""){ 
  try {
    if([string]::IsNullOrWhiteSpace($args)){
      Start-Process $target | Out-Null
    } else {
      Start-Process -FilePath $target -ArgumentList $args | Out-Null
    }
  } catch {}
}

function Start-Teams {
  # 1) Yeni Teams protokolü
  Start-ProcessSafe "msteams:" ; if($?){ return }

  # 2) Eski Teams protokolü
  Start-ProcessSafe "ms-teams:" ; if($?){ return }

  # 3) Klasik teams komutu (PATH'te varsa)
  Start-ProcessSafe "teams" ; if($?){ return }

  # 4) Sık kullanılan kurulum yolları
  $candidates = @(
    Join-Path $env:LOCALAPPDATA "Microsoft\Teams\current\Teams.exe",
    "C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe",
    "C:\Program Files\Microsoft Teams\current\Teams.exe"
  )
  foreach($p in $candidates){
    if(Test-Path $p){
      Start-ProcessSafe $p
      return
    }
  }

  # 5) En son çare: web sürümü
  Start-ProcessSafe "https://teams.microsoft.com"
}

if($QlOutlook){  $QlOutlook.Add_Click({ Start-ProcessSafe "outlook" }) }
if($QlTeams){    $QlTeams.Add_Click({ Start-Teams }) }
if($QlExplorer){ $QlExplorer.Add_Click({ Start-ProcessSafe "explorer.exe" }) }
if($QlEdge){     $QlEdge.Add_Click({ Start-ProcessSafe "microsoft-edge:" }) }
if($QlVSCode){   $QlVSCode.Add_Click({ Start-ProcessSafe "code" }) }
if($QlExcel){    $QlExcel.Add_Click({ Start-ProcessSafe "excel" }) }

# ------- Timers -------
$clock = New-Object System.Windows.Threading.DispatcherTimer
$clock.Interval = [TimeSpan]::FromSeconds(1)
$clock.add_Tick({ $TimeText.Text = (Get-Date).ToString("hh:mm tt") })
$clock.Start()

$pom = New-Object System.Windows.Threading.DispatcherTimer
$pom.Interval = [TimeSpan]::FromSeconds(1)
$pom.add_Tick({
  if($global:Pomodoro.Running -and $global:Pomodoro.Mode -ne "IDLE"){
    $global:Pomodoro.Remaining -= 1
    if($global:Pomodoro.Remaining -le 0){
      $global:Pomodoro.Remaining = 0
      $global:Pomodoro.Running = $false
    }
    Update-PomodoroUI
  }
})
$pom.Start()

$sysT = New-Object System.Windows.Threading.DispatcherTimer
$sysT.Interval = [TimeSpan]::FromSeconds(5)
$sysT.add_Tick({ Update-SystemStatusUI })
$sysT.Start()

$mailT = New-Object System.Windows.Threading.DispatcherTimer
$mailT.Interval = [TimeSpan]::FromSeconds(30)
$mailT.add_Tick({ if($win.IsVisible){ Update-MailUI } })
$mailT.Start()

# initial UI
$global:Todos = @(Load-TodosData)
Render-TodoList
Reset-Pomodoro
Update-SystemStatusUI
Update-MailUI
Render-Calendar

# start hidden (centered)
$wa = [System.Windows.SystemParameters]::WorkArea
$win.Left = $wa.Left + (($wa.Width - $win.Width) / 2)
$win.Top  = $wa.Top  + (($wa.Height - $win.Height) / 2)
$win.Hide()
$global:DashboardWin = $win

function Toggle-Dashboard {
  if($global:DashboardWin.IsVisible){
    # save notes when hiding
    $saveTimer.Stop()
    Save-CurrentNote
    $global:DashboardWin.Hide()
  } else {
    $global:DashboardWin.Show()
    $global:DashboardWin.Activate() | Out-Null
  }
}

# Hotkey host (invisible)
$hotkeyHost = New-Object System.Windows.Window
$hotkeyHost.WindowStyle = 'None'
$hotkeyHost.AllowsTransparency = $true
$hotkeyHost.Background = 'Transparent'
$hotkeyHost.Width = 1; $hotkeyHost.Height = 1
$hotkeyHost.ShowInTaskbar = $false
$hotkeyHost.Topmost = $true
$hotkeyHost.Left = -10000; $hotkeyHost.Top = -10000
$hotkeyHost.Opacity = 0.0

$hotkeyHost.Add_SourceInitialized({
  param($sender,$e)
  $global:HotkeyHostHwnd = [System.Windows.Interop.WindowInteropHelper]::new($sender).Handle
  $src = [System.Windows.Interop.HwndSource]::FromHwnd($global:HotkeyHostHwnd)

  $null = $src.AddHook({
    param($hwnd, $msg, $wParam, $lParam, [ref]$handled)
    if($msg -eq [Win32]::WM_HOTKEY -and $wParam.ToInt32() -eq $HOTKEY_ID){
      Toggle-Dashboard
      $handled.Value = $true
    }
    return [IntPtr]::Zero
  })

  [Win32]::RegisterHotKey(
    $global:HotkeyHostHwnd, $HOTKEY_ID,
    ([Win32]::MOD_CONTROL -bor [Win32]::MOD_ALT),
    $VK_D
  ) | Out-Null
})

$hotkeyHost.Add_Closed({
  param($sender,$e)
  if($global:HotkeyHostHwnd -ne [IntPtr]::Zero){
    [Win32]::UnregisterHotKey($global:HotkeyHostHwnd, $HOTKEY_ID) | Out-Null
  }
  if($global:DashboardWin){ $global:DashboardWin.Close() }
})

$null = $hotkeyHost.Show()

$app = New-Object System.Windows.Application

# Swallow PipelineStoppedException coming from WPF dispatcher so the
# PowerShell host doesn’t show an unhandled exception on shutdown.
$app.add_DispatcherUnhandledException({
  param($sender, $e)
  if($e.Exception -is [System.Management.Automation.PipelineStoppedException]){
    $e.Handled = $true
  }
})

$null = $app.Run()

exit 0
