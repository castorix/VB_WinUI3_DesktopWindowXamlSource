Imports System.Collections.ObjectModel
Imports System.ComponentModel.DataAnnotations
Imports System.IO
Imports System.Media
Imports System.Numerics
Imports System.Runtime.InteropServices
Imports AnimatedVisuals
Imports CommunityToolkit.WinUI
Imports CommunityToolkit.WinUI.Lottie
Imports Microsoft.UI
Imports Microsoft.UI.Composition
Imports Microsoft.UI.Dispatching
Imports Microsoft.UI.Xaml.Hosting
Imports Microsoft.UI.Xaml.Markup
Imports Microsoft.UI.Xaml.Media
Imports Microsoft.UI.Xaml.Media.Imaging
Imports Windows.Devices.Enumeration
Imports Windows.Graphics
Imports Windows.Media.Core
Imports Windows.Media.Playback
Imports Windows.Storage
Imports Application = Microsoft.UI.Xaml.Application
Imports Button = Microsoft.UI.Xaml.Controls.Button
Imports Image = Microsoft.UI.Xaml.Controls.Image
Imports Panel = Microsoft.UI.Xaml.Controls.Panel
Imports ToolTip = Microsoft.UI.Xaml.Controls.ToolTip

Public Class Form1

    Public Enum DWMWINDOWATTRIBUTE
        DWMWA_NCRENDERING_ENABLED = 1 ' [get] Is non - client rendering enabled/disabled
        DWMWA_NCRENDERING_POLICY                   ' [set] DWMNCRENDERINGPOLICY - Non-client rendering policy
        DWMWA_TRANSITIONS_FORCEDISABLED            ' [set] Potentially enable/forcibly disable transitions
        DWMWA_ALLOW_NCPAINT                        ' [set] Allow contents rendered In the non-client area To be visible On the DWM-drawn frame.
        DWMWA_CAPTION_BUTTON_BOUNDS                ' [get] Bounds Of the caption button area In window-relative space.
        DWMWA_NONCLIENT_RTL_LAYOUT                 ' [set] Is non-client content RTL mirrored
        DWMWA_FORCE_ICONIC_REPRESENTATION          ' [set] Force this window To display iconic thumbnails.
        DWMWA_FLIP3D_POLICY                         ' [set] Designates how Flip3D will treat the window.
        DWMWA_EXTENDED_FRAME_BOUNDS                 ' [get] Gets the extended frame bounds rectangle In screen space
        DWMWA_HAS_ICONIC_BITMAP                     ' [set] Indicates an available bitmap When there Is no better thumbnail representation.
        DWMWA_DISALLOW_PEEK                         ' [set] Don't invoke Peek on the window.
        DWMWA_EXCLUDED_FROM_PEEK                    ' [set] LivePreview exclusion information
        DWMWA_CLOAK                                 ' [set] Cloak Or uncloak the window
        DWMWA_CLOAKED                               ' [get] Gets the cloaked state Of the window
        DWMWA_FREEZE_REPRESENTATION                 ' [set] BOOL, Force this window To freeze the thumbnail without live update
        DWMWA_PASSIVE_UPDATE_MODE                   ' [set] BOOL, Updates the window only When desktop composition runs For other reasons
        DWMWA_USE_HOSTBACKDROPBRUSH                 ' [set] BOOL, Allows the use Of host backdrop brushes For the window.
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20          ' [set] BOOL, Allows a window To either use the accent color, Or dark, according To the user Color Mode preferences.
        DWMWA_WINDOW_CORNER_PREFERENCE = 33         ' [set] WINDOW_CORNER_PREFERENCE, Controls the policy that rounds top-level window corners
        DWMWA_BORDER_COLOR                          ' [set] COLORREF, The color Of the thin border around a top-level window
        DWMWA_CAPTION_COLOR                         ' [set] COLORREF, The color Of the caption
        DWMWA_TEXT_COLOR                            ' [set] COLORREF, The color Of the caption text
        DWMWA_VISIBLE_FRAME_BORDER_THICKNESS        ' [get] UINT, width Of the visible border around a thick frame window
        DWMWA_SYSTEMBACKDROP_TYPE                   ' [get, Set] SYSTEMBACKDROP_TYPE, Controls the system-drawn backdrop material Of a window, including behind the non-client area.
        DWMWA_LAST
    End Enum

    <DllImport("Dwmapi.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function DwmSetWindowAttribute(hwnd As IntPtr, dwAttribute As Integer, <[In]> pvAttribute() As Integer, cbAttribute As Integer) As Integer
    End Function

    Dim m_dwxs As DesktopWindowXamlSource = Nothing
    Dim m_BackgroundBrush As SolidColorBrush = Nothing
    Dim m_LinearGradientBrush As New LinearGradientBrush With {
        .StartPoint = New Windows.Foundation.Point(0, 0),
        .EndPoint = New Windows.Foundation.Point(0, 1),
        .GradientStops = New GradientStopCollection From {
        New GradientStop With {.Color = Colors.LightBlue, .Offset = 0},
        New GradientStop With {.Color = Colors.Blue, .Offset = 1}
    }
    }

    Private m_MainGrid As Grid = Nothing
    Private m_NavigationView As NavigationView
    Private m_Button As Button
    Private m_Tooltip As ToolTip
    Private m_CalendarView As CalendarView
    Private m_CalendarText As TextBlock
    Private m_MPE As MediaPlayerElement
    Private m_ViewBoxRatingControl As Viewbox
    Private m_RatingControl As RatingControl
    Private m_RatingControlText As TextBlock
    Private m_Ring As ProgressRing
    'Private m_AboutText As TextBlock
    Private m_ContentGrid As Grid
    Private m_TeachingTip As TeachingTip
    Private m_ItemsView As ItemsView

    Private m_MP As New MediaPlayer()
    Private m_WhatSource As MediaSource
    Private m_BookPageSource As MediaSource
    Private m_PotionSource As MediaSource
    Private m_SwordSource As MediaSource
    Private m_DiamondSource As MediaSource
    Private m_ChestOpenSource As MediaSource

    Private m_LottiePlayer As AnimatedVisualPlayer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WindowsXamlManager.InitializeForCurrentThread()
        m_dwxs = New DesktopWindowXamlSource()
        Dim myWndId As WindowId = Win32Interop.GetWindowIdFromWindow(Me.Handle)
        m_dwxs.Initialize(myWndId)

        Application.Current.Resources("ButtonBackgroundPointerOver") = New SolidColorBrush(Colors.SteelBlue)
        Application.Current.Resources("ButtonBackgroundPressed") = New SolidColorBrush(Colors.MidnightBlue)
        Application.Current.Resources("ButtonForegroundPressed") = New SolidColorBrush(Colors.White)
        SetCustomThemeBrushes() ' To test Dark/Light resources

        CreateXamlUI()

        Dim root As FrameworkElement = CType(m_dwxs.Content, FrameworkElement)
        AddHandler root.ActualThemeChanged, AddressOf Root_ActualThemeChanged

        SetBackgroundBrush()
        SetFormBackgroundColor()
        LoadMP3s()

        'Me.DoubleBuffered = True ' Bad when resizing

        ResizeXamlIsland()
        Me.Text = "VB.NET - WinUI3 DesktopWindowXamlSource"
        Me.Icon = New Icon("Assets\Logo-WinUI.ico")
        Dim nUseDarkMode As Integer = 1
        DwmSetWindowAttribute(Me.Handle, DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, New Integer() {nUseDarkMode}, Marshal.SizeOf(GetType(Integer)))
        CenterToScreen()
    End Sub

    Private Sub Root_ActualThemeChanged(sender As FrameworkElement, args As Object)
        SetBackgroundBrush()
        SetFormBackgroundColor()
        UpdateTooltipTheme(m_Tooltip)
    End Sub

    Private Sub SetBackgroundBrush()
        m_BackgroundBrush = TryCast(Application.Current.Resources("ApplicationPageBackgroundThemeBrush"), SolidColorBrush)
        If m_BackgroundBrush IsNot Nothing Then
            m_MainGrid.Background = m_BackgroundBrush
            Dim selectedItem = TryCast(m_NavigationView.SelectedItem, NavigationViewItem)
            If selectedItem IsNot Nothing AndAlso selectedItem.Content.ToString() = "Composition 1" Then
                m_ContentGrid.Background = m_LinearGradientBrush
            Else
                m_ContentGrid.Background = m_BackgroundBrush
            End If
        End If
    End Sub

    Private Sub SetFormBackgroundColor()
        If m_BackgroundBrush IsNot Nothing Then
            Dim formColor As Color = Color.FromArgb(m_BackgroundBrush.Color.A, m_BackgroundBrush.Color.R, m_BackgroundBrush.Color.G, m_BackgroundBrush.Color.B)
            Me.BackColor = formColor
        End If
    End Sub

    Private Sub CreateXamlUI()
        m_MainGrid = New Grid()

        m_NavigationView = New NavigationView With {
            .PaneDisplayMode = NavigationViewPaneDisplayMode.Left,
            .IsPaneOpen = True,
            .CompactPaneLength = 40,
            .IsSettingsVisible = True
        }
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "Button",
                              .Icon = New FontIcon With {
                              .Glyph = ChrW(&HEECA),
                              .FontFamily = New FontFamily("Segoe MDL2 Assets")}})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "CalendarView",
                              .Icon = New SymbolIcon(Symbol.Calendar)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "MediaPlayerElement",
                              .Icon = New SymbolIcon(Symbol.Play)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "RatingControl",
                              .Icon = New SymbolIcon(Symbol.Favorite)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "ProgressRing",
                              .Icon = New SymbolIcon(Symbol.Sync)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "Lottie",
                              .Icon = New FontIcon With {
                              .Glyph = ChrW(&HED54),
                              .FontFamily = New FontFamily("Segoe MDL2 Assets")}})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "ItemsView",
                              .Icon = New SymbolIcon(Symbol.List)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "Composition 1",
                              .Icon = New SymbolIcon(Symbol.SwitchApps)})
        m_NavigationView.MenuItems.Add(New NavigationViewItem With {
                              .Content = "Composition 2",
                              .Icon = New SymbolIcon(Symbol.Orientation)})

        AddHandler m_NavigationView.ItemInvoked, AddressOf NavView_ItemInvoked

        m_ContentGrid = New Grid()

        m_Button = New Button With {
            .Content = "Click me",
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
            .VerticalAlignment = Xaml.VerticalAlignment.Center
        }
        m_Tooltip = New ToolTip With {
            .Content = "Click to display a Teaching Tip",
            .Background = CType(Application.Current.Resources("MyTooltipBackgroundBrush"), Brush),
            .Foreground = CType(Application.Current.Resources("MyTooltipForegroundBrush"), Brush)
        }

        'XamlBindingHelper.SetPropertyFromString(m_Tooltip, ToolTip.BackgroundProperty, "{ThemeResource MyTooltipBackgroundBrush}")
        'XamlBindingHelper.SetPropertyFromString(m_Tooltip, ToolTip.ForegroundProperty, "{ThemeResource MyTooltipForegroundBrush}")

        ToolTipService.SetToolTip(m_Button, m_Tooltip)
        AddHandler m_Button.Click, AddressOf Button_Click

        m_CalendarView = New CalendarView()
        m_CalendarView.HorizontalAlignment = Xaml.HorizontalAlignment.Center
        m_CalendarView.VerticalAlignment = Xaml.VerticalAlignment.Top
        m_CalendarView.Margin = New Thickness(10, 10, 10, 0)
        m_CalendarView.Height = 300

        m_CalendarText = New TextBlock()
        m_CalendarText.HorizontalAlignment = Xaml.HorizontalAlignment.Center
        m_CalendarText.VerticalAlignment = Xaml.VerticalAlignment.Bottom
        m_CalendarText.Margin = New Thickness(10)
        m_CalendarText.FontSize = 24

        AddHandler m_CalendarView.SelectedDatesChanged, Sub(sender As Object, e As CalendarViewSelectedDatesChangedEventArgs)
                                                            ' Get the first selected date and display it in the TextBlock
                                                            If m_CalendarView.SelectedDates.Count > 0 Then
                                                                m_CalendarText.Text = "Selected Date: " & m_CalendarView.SelectedDates(0).ToString("MMMM dd, yyyy")
                                                            End If
                                                        End Sub
        m_MPE = New MediaPlayerElement With {
        .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
        .VerticalAlignment = Xaml.VerticalAlignment.Center, '.Background = New SolidColorBrush(Colors.Black),
        .AreTransportControlsEnabled = True,
        .Source = MediaSource.CreateFromUri(New Uri("https://storage.googleapis.com/gtv-videos-bucket/sample/ElephantsDream.mp4"))
        }

        'Dim xamlTemplate As String = "
        '<ControlTemplate xmlns='using:Microsoft.UI.Xaml.Controls' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' TargetType='RatingControl'>
        '    <Grid>
        '        <StackPanel Orientation='Horizontal' HorizontalAlignment='Center'>
        '            <!-- Heart icon 1 -->
        '            <TextBlock FontFamily='Segoe UI Symbol' Text='?' Foreground='Red' FontSize='30'/>
        '            <!-- Heart icon 2 -->
        '            <TextBlock FontFamily='Segoe UI Symbol' Text='?' Foreground='Red' FontSize='30'/>
        '            <!-- Heart icon 3 -->
        '            <TextBlock FontFamily='Segoe UI Symbol' Text='?' Foreground='Red' FontSize='30'/>
        '            <!-- Heart icon 4 -->
        '            <TextBlock FontFamily='Segoe UI Symbol' Text='?' Foreground='Red' FontSize='30'/>
        '            <!-- Heart icon 5 -->
        '            <TextBlock FontFamily='Segoe UI Symbol' Text='?' Foreground='Red' FontSize='30'/>
        '        </StackPanel>
        '    </Grid>
        '</ControlTemplate>"
        'Dim controlTemplate As ControlTemplate = CType(XamlReader.Load(xamlTemplate), ControlTemplate)

        m_ViewBoxRatingControl = New Viewbox() With {.Height = 100}
        m_RatingControl = New RatingControl() With {
            .MaxRating = 5,
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
            .VerticalAlignment = Xaml.VerticalAlignment.Center '.Template = controlTemplate
        }
        m_ViewBoxRatingControl.Child = m_RatingControl

        m_RatingControlText = New TextBlock()
        m_RatingControlText.HorizontalAlignment = Xaml.HorizontalAlignment.Center
        m_RatingControlText.VerticalAlignment = Xaml.VerticalAlignment.Bottom
        m_RatingControlText.Margin = New Thickness(10)
        m_RatingControlText.FontSize = 24

        AddHandler m_RatingControl.ValueChanged, Sub(sender As Object, e As Object)
                                                     m_RatingControlText.Text = "Value: " & sender.Value.ToString()
                                                 End Sub

        m_Ring = New ProgressRing With {
            .IsIndeterminate = True,
            .IsActive = True,
            .Width = 100,
            .Height = 100,
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
            .VerticalAlignment = Xaml.VerticalAlignment.Center
        }

        'm_AboutText = New TextBlock With {
        '    .Text = "About TextBlock",
        '    .FontSize = 24,
        '    .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
        '    .VerticalAlignment = Xaml.VerticalAlignment.Center
        '}

        ' Initial view : Button + TeachingTip
        m_ContentGrid.Children.Add(m_Button)

        Dim image As New Image With {
            .Source = New BitmapImage(New Uri("ms-appx:///Assets/waterfall.gif")),
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
            .VerticalAlignment = Xaml.VerticalAlignment.Center,
            .Margin = New Thickness(5, 5, 5, 0)
        }
        Dim spTipContent As New StackPanel()
        spTipContent.Children.Add(image)

        m_TeachingTip = New TeachingTip With {
            .Title = "❤️ VB.NET + WinUI 3 ❤️",
            .Subtitle = "Test TeachingTip",
            .IsOpen = False,
            .Target = m_Button,
            .PreferredPlacement = TeachingTipPlacementMode.BottomRight,
            .Background = New SolidColorBrush(Colors.RoyalBlue),
            .XamlRoot = m_MainGrid.XamlRoot
        }
        m_TeachingTip.Content = spTipContent
        m_ContentGrid.Children.Add(m_TeachingTip)

        m_LottiePlayer = New AnimatedVisualPlayer With {
            .Width = 400,
            .Height = 400,
            .Stretch = Stretch.Uniform,
            .AutoPlay = True,
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center,
            .VerticalAlignment = Xaml.VerticalAlignment.Center}
        Dim cowSource As New Cow()
        m_LottiePlayer.Source = cowSource

        'x:      DataType ='local:CustomDataObject'
        'The property 'DataType' was not found in type 'DataTemplate'. [Line: 4 Position: 16]
        Dim dtXaml As String = "<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' 
               xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' 
               xmlns:local='using:VB_WinUI3_DesktopWindowXamlSource'>
<ItemContainer AutomationProperties.Name='{Binding Title}'>
        <Grid>
            <Image Source='{Binding ImageLocation}' Stretch='UniformToFill' HorizontalAlignment='Center' VerticalAlignment='Center' MinWidth='70'/>
            <StackPanel Orientation='Vertical' Height='40' VerticalAlignment='Bottom' Padding='5,1,5,1' Background='{ThemeResource SystemControlBackgroundBaseMediumBrush}' Opacity='.75'>
                <TextBlock Text='{Binding Title}' Foreground='{ThemeResource SystemControlForegroundAltHighBrush}'/>
            </StackPanel>
        </Grid>
    </ItemContainer>
</DataTemplate>"

        Dim itemTemplate As DataTemplate = CType(XamlReader.Load(dtXaml), DataTemplate)
        Dim linedFlowLayout As New LinedFlowLayout() With {
            .ItemsStretch = LinedFlowLayoutItemsStretch.Fill,
            .LineHeight = 220,
            .LineSpacing = 5,
            .MinItemSpacing = 5
        }

        m_ItemsView = New ItemsView() With {
            .HorizontalAlignment = Xaml.HorizontalAlignment.Stretch,
            .VerticalAlignment = Xaml.VerticalAlignment.Stretch,
            .ItemTemplate = itemTemplate,
            .Layout = linedFlowLayout
        }
        Dim tempList As List(Of CustomDataObject) = CustomDataObject.GetDataObjects("Assets\Animals")
        Dim items As New ObservableCollection(Of CustomDataObject)(tempList)
        m_ItemsView.ItemsSource = items

        m_NavigationView.Content = m_ContentGrid
        m_MainGrid.Children.Add(m_NavigationView)

        m_dwxs.Content = m_MainGrid
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        ' If enable native code debugging :  0x80073D54 : 'Le processus n’a aucune identité de package.'.
        m_TeachingTip.IsOpen = True
        'SystemSounds.Asterisk.Play()
    End Sub

    ' In generic.xaml : ToolTipBackgroundThemeBrush, ToolTipBorderThemeBrush, ToolTipForegroundThemeBrush
    Private Sub SetCustomThemeBrushes()
        ' Create the theme-specific dictionaries
        Dim lightTheme = New ResourceDictionary()
        lightTheme("MyTooltipBackgroundBrush") = New SolidColorBrush(Colors.LightYellow)
        lightTheme("MyTooltipForegroundBrush") = New SolidColorBrush(Colors.Black)

        Dim darkTheme = New ResourceDictionary()
        darkTheme("MyTooltipBackgroundBrush") = New SolidColorBrush(Colors.Black)
        darkTheme("MyTooltipForegroundBrush") = New SolidColorBrush(Colors.Yellow)

        ' Add them into the existing ThemeDictionaries of a new dictionary
        Dim themedResources = New ResourceDictionary()
        themedResources.ThemeDictionaries("Light") = lightTheme
        themedResources.ThemeDictionaries("Dark") = darkTheme

        ' Merge it into the app's global resources
        Application.Current.Resources.MergedDictionaries.Add(themedResources)
    End Sub

    Private Sub UpdateTooltipTheme(tooltip As ToolTip)
        Dim currentTheme = Application.Current.RequestedTheme.ToString()
        Dim themedResources = Application.Current.Resources.MergedDictionaries.LastOrDefault()
        If themedResources IsNot Nothing AndAlso themedResources.ThemeDictionaries.ContainsKey(currentTheme) Then
            Dim themeDict = CType(themedResources.ThemeDictionaries(currentTheme), ResourceDictionary)
            Dim bg = TryCast(themeDict("MyTooltipBackgroundBrush"), SolidColorBrush)
            Dim fg = TryCast(themeDict("MyTooltipForegroundBrush"), SolidColorBrush)
            If bg IsNot Nothing Then tooltip.Background = bg
            If fg IsNot Nothing Then tooltip.Foreground = fg
        End If
    End Sub

    Private Sub ShowInfoBar()
        For Each child As UIElement In m_MainGrid.Children
            If TypeOf child Is InfoBar Then
                m_MP.Source = m_WhatSource
                m_MP.Play()
                Return
            End If
        Next

        Dim infoBar As New InfoBar With {
            .Severity = InfoBarSeverity.Error,
            .Title = "No Settings",
            .Message = "Settings are not implemented",
            .IsOpen = True,
            .IsClosable = True
        }
        m_MainGrid.Children.Add(infoBar)

        AddHandler infoBar.Closing, Sub(s, e)
                                        m_MainGrid.Children.Remove(infoBar)
                                    End Sub
    End Sub

    Private Sub NavView_ItemInvoked(sender As NavigationView, args As NavigationViewItemInvokedEventArgs)
        If args.IsSettingsInvoked Then
            ShowInfoBar()
        Else
            m_TeachingTip.IsOpen = False
            m_MPE.MediaPlayer.Pause()
            m_ContentGrid.Children.Clear()
            SetBackgroundBrush()

            StopCompositionAnimation()

            Select Case args.InvokedItem.ToString()
                Case "Button"
                    m_ContentGrid.Children.Add(m_Button)
                    m_ContentGrid.Children.Add(m_TeachingTip)
                Case "CalendarView"
                    m_ContentGrid.Children.Add(m_CalendarView)
                    m_ContentGrid.Children.Add(m_CalendarText)
                Case "MediaPlayerElement"
                    m_ContentGrid.Children.Add(m_MPE)
                Case "RatingControl"
                    m_ContentGrid.Children.Add(m_ViewBoxRatingControl)
                    m_ContentGrid.Children.Add(m_RatingControlText)
                Case "ProgressRing"
                    m_ContentGrid.Children.Add(m_Ring)
                Case "Lottie"
                    m_ContentGrid.Children.Add(m_LottiePlayer)
                Case "ItemsView"
                    m_ContentGrid.Children.Add(m_ItemsView)
                Case "Composition 1"
                    ShowImageButtons(m_ContentGrid)
                Case "Composition 2"
                    ShowCompositionAnimation(m_ContentGrid, imageUris(m_nCurrentIndex))
            End Select
        End If
    End Sub

    Private Sub ShowImageButtons(Host As FrameworkElement)

        Dim sp As New StackPanel With {
            .Orientation = Xaml.Controls.Orientation.Horizontal,
            .HorizontalAlignment = Xaml.HorizontalAlignment.Center
        }

        Dim btnBook = New ImageButton(New Uri("ms-appx:///Assets/Fantasy/Book.png"), Nothing, 100)
        btnBook.Margin = New Thickness(10, 0, 0, 0)
        btnBook.VerticalAlignment = VerticalAlignment.Center
        AddHandler btnBook.Clicked, Sub(s, e)
                                        m_MP.Source = m_BookPageSource
                                        m_MP.Play()
                                    End Sub
        sp.Children.Add(btnBook)
        'btnBook.ButtonLabel = "Test"
        'btnBook.IsEnabled = False

        Dim btnPotion = New ImageButton(New Uri("ms-appx:///Assets/Fantasy/Potion.png"), Nothing, 100)
        btnPotion.Margin = New Thickness(10, 0, 0, 0)
        btnPotion.VerticalAlignment = VerticalAlignment.Center
        AddHandler btnPotion.Clicked, Sub(s, e)
                                          m_MP.Source = m_PotionSource
                                          m_MP.Play()
                                      End Sub
        sp.Children.Add(btnPotion)

        Dim btnSword = New ImageButton(New Uri("ms-appx:///Assets/Fantasy/Sword.png"), Nothing, 100)
        btnSword.Margin = New Thickness(10, 0, 0, 0)
        btnSword.VerticalAlignment = VerticalAlignment.Center
        AddHandler btnSword.Clicked, Sub(s, e)
                                         m_MP.Source = m_SwordSource
                                         m_MP.Play()
                                     End Sub
        sp.Children.Add(btnSword)

        Dim btnDiamond = New ImageButton(New Uri("ms-appx:///Assets/Fantasy/Diamond.png"), Nothing, 100)
        btnDiamond.Margin = New Thickness(10, 0, 0, 0)
        btnDiamond.VerticalAlignment = VerticalAlignment.Center
        AddHandler btnDiamond.Clicked, Sub(s, e)
                                           m_MP.Source = m_DiamondSource
                                           m_MP.Play()
                                       End Sub
        sp.Children.Add(btnDiamond)

        Dim btnChestOpen = New ImageButton(New Uri("ms-appx:///Assets/Fantasy/Chest_open.png"), Nothing, 100)
        btnChestOpen.Margin = New Thickness(10, 0, 0, 0)
        btnChestOpen.VerticalAlignment = VerticalAlignment.Center
        AddHandler btnChestOpen.Clicked, Sub(s, e)
                                             m_MP.Source = m_ChestOpenSource
                                             m_MP.Play()
                                         End Sub
        sp.Children.Add(btnChestOpen)

        DirectCast(Host, Panel).Children.Add(sp)
        DirectCast(Host, Grid).Background = m_LinearGradientBrush
    End Sub

    Private imageUris As List(Of Uri) = LoadImageUrisFromAssets()

    Private Function LoadImageUrisFromAssets() As List(Of Uri)
        Dim imageUris As New List(Of Uri)()
        Dim sExeDir = AppContext.BaseDirectory
        Dim sImageFolder = Path.Combine(sExeDir, "Assets", "Landscapes")

        If Directory.Exists(sImageFolder) Then
            For Each file In Directory.EnumerateFiles(sImageFolder)
                Dim fileUri = New Uri("file:///" & file.Replace("\", "/"))
                imageUris.Add(fileUri)
            Next
        End If
        Return imageUris
    End Function

    'Private imageUris As List(Of Uri) = New List(Of Uri) From {
    '    New Uri("ms-appx:///Assets/Landscapes/Mountain_lake.jpeg"),
    '    New Uri("ms-appx:///Assets/Landscapes/Lake_snow_sunset.jpg")
    '}
    Private m_nCurrentIndex As Integer = 0

    Private m_ContainerVisual As ContainerVisual = Nothing
    Private m_SpriteVisual As SpriteVisual = Nothing
    Private m_ActiveScaleAnimation As Vector3KeyFrameAnimation = Nothing
    Private m_ActiveRotationAnimation As ScalarKeyFrameAnimation = Nothing

    ' Try to do same effect I did in Direct2D  (https://github.com/castorix/WinUI3_Direct2D_Effects/blob/master/MainWindow.xaml.cs#L2710)
    ' with the help of ChatGPT

    Private Sub ShowCompositionAnimation(AnimationHost As FrameworkElement, imageUri As Uri)
        Dim visual = ElementCompositionPreview.GetElementVisual(AnimationHost)
        Dim compositor = visual.Compositor

        If m_ContainerVisual Is Nothing Then
            m_ContainerVisual = compositor.CreateContainerVisual()
            ElementCompositionPreview.SetElementChildVisual(AnimationHost, m_ContainerVisual)
        End If

        'Dim surface = LoadedImageSurface.StartLoadFromUri(New Uri("ms-appx:///Assets/Landscapes/Mountain_lake.jpeg"))
        Dim surface = LoadedImageSurface.StartLoadFromUri(imageUri)
        Dim brush = compositor.CreateSurfaceBrush(surface)

        m_SpriteVisual = compositor.CreateSpriteVisual()
        m_SpriteVisual.Brush = brush

        ' Add the sprite visual to the container visual
        m_ContainerVisual.Children.InsertAtTop(m_SpriteVisual)

        ' When the image is loaded, apply proper size and center rotation
        AddHandler surface.LoadCompleted, Sub(s, e)
                                              If m_SpriteVisual IsNot Nothing Then
                                                  Dim scopedBatch As CompositionScopedBatch = compositor.CreateScopedBatch(CompositionBatchTypes.Animation)
                                                  ' Get image natural size
                                                  Dim naturalSize = surface.NaturalSize
                                                  Dim hostSize = AnimationHost.RenderSize

                                                  ' Scale factor from image to host size
                                                  Dim scaleX As Single = CSng(hostSize.Width) / CSng(naturalSize.Width)
                                                  Dim scaleY As Single = CSng(hostSize.Height) / CSng(naturalSize.Height)

                                                  ' Set final image size to match host
                                                  m_SpriteVisual.Size = New Vector2(CSng(naturalSize.Width), CSng(naturalSize.Height))

                                                  ' Set center point for rotation (center of image)
                                                  m_SpriteVisual.CenterPoint = New Vector3(m_SpriteVisual.Size.X / 2, m_SpriteVisual.Size.Y / 2, 0)

                                                  ' Center the visual in the host
                                                  m_SpriteVisual.Offset = New Vector3((CSng(hostSize.Width) - m_SpriteVisual.Size.X) / 2,
                                                                                (CSng(hostSize.Height) - m_SpriteVisual.Size.Y) / 2, 0)

                                                  ' Start the scale animation from 0 to 1
                                                  Dim scaleAnim = compositor.CreateVector3KeyFrameAnimation()
                                                  scaleAnim.InsertKeyFrame(0.0F, New Vector3(0, 0, 1))
                                                  scaleAnim.InsertKeyFrame(0.5F, New Vector3(scaleX, scaleY, 1))
                                                  scaleAnim.InsertKeyFrame(1.0F, New Vector3(scaleX, scaleY, 1))
                                                  scaleAnim.Duration = TimeSpan.FromSeconds(5)
                                                  'scaleAnim.IterationBehavior = Composition.AnimationIterationBehavior.Forever
                                                  scaleAnim.IterationBehavior = Composition.AnimationIterationBehavior.Count
                                                  scaleAnim.IterationCount = 1
                                                  m_SpriteVisual.StartAnimation("Scale", scaleAnim)

                                                  m_ActiveScaleAnimation = scaleAnim

                                                  ' Start the rotation animation
                                                  Dim rotationAnim = compositor.CreateScalarKeyFrameAnimation()
                                                  rotationAnim.InsertKeyFrame(0.0F, 0.0F)
                                                  rotationAnim.InsertKeyFrame(0.5F, 360.0F)
                                                  rotationAnim.InsertKeyFrame(1.0F, 360.0F)
                                                  rotationAnim.Duration = TimeSpan.FromSeconds(5)
                                                  'rotationAnim.IterationBehavior = Composition.AnimationIterationBehavior.Forever
                                                  rotationAnim.IterationBehavior = Composition.AnimationIterationBehavior.Count
                                                  rotationAnim.IterationCount = 1
                                                  m_SpriteVisual.StartAnimation("RotationAngleInDegrees", rotationAnim)

                                                  m_ActiveRotationAnimation = rotationAnim

                                                  scopedBatch.End()

                                                  AddHandler scopedBatch.Completed, Sub(sender As Object, arg As CompositionBatchCompletedEventArgs)
                                                                                        If m_SpriteVisual IsNot Nothing Then
                                                                                            StopCompositionAnimation()
                                                                                            m_nCurrentIndex = (m_nCurrentIndex + 1) Mod imageUris.Count
                                                                                            ShowCompositionAnimation(m_ContentGrid, imageUris(m_nCurrentIndex))
                                                                                        End If
                                                                                    End Sub
                                              End If
                                          End Sub

        ' Listen for size changes and update scale and position (random crashes)

        'AddHandler AnimationHost.SizeChanged, Sub(sender, e)
        '                                          If m_SpriteVisual IsNot Nothing Then

        '                                              'spriteVisual.StopAnimation("Scale")
        '                                              'spriteVisual.StopAnimation("RotationAngleInDegrees")

        '                                              ' Recalculate scale based on new size
        '                                              Dim newHostSize = AnimationHost.RenderSize
        '                                              Dim newScaleX As Single = CSng(newHostSize.Width) / CSng(surface.NaturalSize.Width)
        '                                              Dim newScaleY As Single = CSng(newHostSize.Height) / CSng(surface.NaturalSize.Height)

        '                                              ' Update the scale animation with new values
        '                                              Dim newScaleAnim = compositor.CreateVector3KeyFrameAnimation()
        '                                              newScaleAnim.InsertKeyFrame(0.0F, New Vector3(0, 0, 1))
        '                                              newScaleAnim.InsertKeyFrame(0.5F, New Vector3(newScaleX, newScaleY, 1))
        '                                              newScaleAnim.InsertKeyFrame(1.0F, New Vector3(newScaleX, newScaleY, 1))
        '                                              newScaleAnim.Duration = TimeSpan.FromSeconds(5)
        '                                              'newScaleAnim.IterationBehavior = Composition.AnimationIterationBehavior.Forever
        '                                              newScaleAnim.IterationBehavior = Composition.AnimationIterationBehavior.Count
        '                                              newScaleAnim.IterationCount = 1
        '                                              m_SpriteVisual.StartAnimation("Scale", newScaleAnim)

        '                                              m_ActiveScaleAnimation = newScaleAnim

        '                                              ' Update the position to keep the visual centered
        '                                              m_SpriteVisual.Offset = New Vector3((CSng(newHostSize.Width) - m_SpriteVisual.Size.X) / 2,
        '                                                                           (CSng(newHostSize.Height) - m_SpriteVisual.Size.Y) / 2, 0)

        '                                              Dim rotationAnim = compositor.CreateScalarKeyFrameAnimation()
        '                                              rotationAnim.InsertKeyFrame(0.0F, 0.0F)
        '                                              rotationAnim.InsertKeyFrame(0.5F, 360.0F)
        '                                              rotationAnim.InsertKeyFrame(1.0F, 360.0F)
        '                                              rotationAnim.Duration = TimeSpan.FromSeconds(5)
        '                                              'rotationAnim.IterationBehavior = Composition.AnimationIterationBehavior.Forever
        '                                              rotationAnim.IterationBehavior = Composition.AnimationIterationBehavior.Count
        '                                              rotationAnim.IterationCount = 1
        '                                              m_SpriteVisual.StartAnimation("RotationAngleInDegrees", rotationAnim)

        '                                              m_ActiveRotationAnimation = rotationAnim
        '                                          End If
        '                                      End Sub
    End Sub

    Private Sub StopCompositionAnimation()
        If m_ActiveScaleAnimation IsNot Nothing Then
            m_SpriteVisual.StopAnimation("Scale")
            m_ActiveScaleAnimation = Nothing
        End If

        If m_ActiveRotationAnimation IsNot Nothing Then
            m_SpriteVisual.StopAnimation("RotationAngleInDegrees")
            m_ActiveRotationAnimation = Nothing
        End If
        If m_SpriteVisual IsNot Nothing Then
            m_ContainerVisual.Children.Remove(m_SpriteVisual)
            m_SpriteVisual = Nothing
        End If
    End Sub

    Private Async Sub LoadMP3s()
        Await PreloadMP3s()
    End Sub

    Private Async Function PreloadMP3s() As Task
        Dim basePath As String = AppContext.BaseDirectory
        Dim whatFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Whatdoyouwant.mp3"))
        m_WhatSource = MediaSource.CreateFromStorageFile(whatFile)
        Dim bookPageFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Book_page.mp3"))
        m_BookPageSource = MediaSource.CreateFromStorageFile(bookPageFile)
        Dim potionFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Potion.mp3"))
        m_PotionSource = MediaSource.CreateFromStorageFile(potionFile)
        Dim swordFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Sword.mp3"))
        m_SwordSource = MediaSource.CreateFromStorageFile(swordFile)
        Dim diamondFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Diamond.mp3"))
        m_DiamondSource = MediaSource.CreateFromStorageFile(diamondFile)
        Dim chestOpenFile As StorageFile = Await StorageFile.GetFileFromPathAsync(System.IO.Path.Combine(basePath, "Assets\Sounds\Chest_open.mp3"))
        m_ChestOpenSource = MediaSource.CreateFromStorageFile(chestOpenFile)
    End Function

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        ResizeXamlIsland()
    End Sub

    'Protected Overrides Sub OnLayout(levent As LayoutEventArgs)
    '    MyBase.OnLayout(levent)
    '    ResizeXamlIsland()
    'End Sub

    'Protected Overrides Sub OnSizeChanged(e As EventArgs)
    '    MyBase.OnSizeChanged(e)
    '    ResizeXamlIsland()
    'End Sub

    'Protected Overrides Sub OnPaintBackground(pevent As PaintEventArgs)
    '    pevent.Graphics.Clear(Color.Red) '
    'End Sub

    Private Sub ResizeXamlIsland()
        If m_dwxs?.SiteBridge Is Nothing Then Return

        Dim sb = m_dwxs.SiteBridge
        Dim siteView = sb.SiteView
        Dim nScale As Double = 1.0

        If siteView IsNot Nothing Then
            nScale = siteView.RasterizationScale
        End If

        Dim rect As New RectInt32 With {
            .X = 0,
            .Y = 0,
            .Width = CInt(Me.ClientSize.Width * nScale),
            .Height = CInt(Me.ClientSize.Height * nScale)
        }
        sb.MoveAndResize(rect)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        StopCompositionAnimation()
    End Sub
End Class



