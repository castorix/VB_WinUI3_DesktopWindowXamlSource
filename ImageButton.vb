Imports System.Numerics
Imports System.Reflection.Emit
Imports Microsoft.UI
Imports Microsoft.UI.Composition
Imports Microsoft.UI.Xaml.Hosting
Imports Microsoft.UI.Xaml.Media
Imports Microsoft.UI.Xaml.Media.Imaging
Imports Image = Microsoft.UI.Xaml.Controls.Image

' Done with help of ChatGPT

Public Class ImageButton
    Inherits StackPanel

    Private ReadOnly _image As Image
    Private ReadOnly _imageBorder As Border
    Private ReadOnly _textBlock As TextBlock
    Private ReadOnly _container As Grid
    Private ReadOnly _visual As Visual
    Private ReadOnly _compositor As Compositor
    Private ReadOnly _ripple As Border

    Private ReadOnly _shadowSprite As SpriteVisual
    Private ReadOnly _shadowHost As Grid
    Private ReadOnly _shadowVisualHost As Grid

    Private ReadOnly _containerBackground As SolidColorBrush = New SolidColorBrush(Colors.LightGray)
    Private ReadOnly _containerBackgroundOver As SolidColorBrush = New SolidColorBrush(Colors.White)
    Private ReadOnly _dropShadowColor As Windows.UI.Color = Colors.Black

    Private _isEnabled As Boolean = True
    Private _isPointerDown As Boolean = False
    Private _pressedInside As Boolean = False

    Private _imageUri As Uri
    Private _label As String

    Public Event Clicked As EventHandler

    Public Sub New(imageUri As Uri, label As String, Optional size As Double = 100)
        Me.Orientation = Controls.Orientation.Vertical
        Me.HorizontalAlignment = Xaml.HorizontalAlignment.Center
        Me.Spacing = 4

        _imageUri = imageUri
        _label = label

        _shadowHost = New Grid() With {
            .Width = size,
            .Height = size
        }

        ' Shadow host that will contain only the composition shadow visual
        _shadowVisualHost = New Grid() With {
            .Width = size,
            .Height = size,
            .IsHitTestVisible = False ' Make sure it doesn't block input
        }

        _container = New Grid() With {
            .Width = size,
            .Height = size,
            .CornerRadius = New CornerRadius(size / 5),
            .Background = _containerBackground
        }
        ' Order: shadow layer below, then content
        _shadowHost.Children.Add(_shadowVisualHost)
        _shadowHost.Children.Add(_container)

        _compositor = ElementCompositionPreview.GetElementVisual(_shadowVisualHost).Compositor

        ' Create a hidden masking Border with the same size and corner radius
        Dim maskBorder = New Border() With {
            .Width = size,
            .Height = size,
            .CornerRadius = New CornerRadius(size / 5),
            .Background = New SolidColorBrush(Colors.White),
            .Opacity = 0.0,
            .IsHitTestVisible = False
        }
        _shadowHost.Children.Add(maskBorder)

        ' Create a VisualSurface for the mask
        Dim surface = _compositor.CreateVisualSurface()
        surface.SourceVisual = ElementCompositionPreview.GetElementVisual(maskBorder)
        surface.SourceSize = New Vector2(CSng(size), CSng(size))

        ' Create the mask brush
        Dim maskBrush = _compositor.CreateSurfaceBrush(surface)

        ' Create the shadow
        Dim dropShadow = _compositor.CreateDropShadow()
        dropShadow.BlurRadius = 15
        dropShadow.Offset = New Vector3(6, 6, 0)
        dropShadow.Color = _dropShadowColor
        dropShadow.Mask = maskBrush ' Use the shape of the border

        _shadowSprite = _compositor.CreateSpriteVisual()
        _shadowSprite.Size = New Vector2(CSng(size), CSng(size))
        _shadowSprite.Shadow = dropShadow

        ElementCompositionPreview.SetElementChildVisual(_shadowVisualHost, _shadowSprite)

        ' Ripple overlay
        '_ripple = New Border() With {
        '    .Background = New SolidColorBrush(Colors.White),
        '    .CornerRadius = New CornerRadius(size / 5),
        '    .Opacity = 0
        '}
        '_container.Children.Add(_ripple)

        ' Image
        _image = New Image() With {
            .Source = New BitmapImage(imageUri),
            .Stretch = Stretch.Uniform,
            .HorizontalAlignment = HorizontalAlignment.Center,
            .VerticalAlignment = VerticalAlignment.Center
        }
        '_container.Children.Add(_image)

        ' Create Border that wraps the Image
        _imageBorder = New Border() With {
            .BorderBrush = New SolidColorBrush(Colors.Transparent),
            .BorderThickness = New Thickness(0),
            .CornerRadius = New CornerRadius(size / 5),
            .HorizontalAlignment = HorizontalAlignment.Stretch,
            .VerticalAlignment = VerticalAlignment.Stretch,
            .Child = _image
        }
        _container.Children.Add(_imageBorder)

        Me.Children.Add(_shadowHost)

        _textBlock = New TextBlock() With {
            .Text = label,
            .HorizontalAlignment = HorizontalAlignment.Center,
            .TextAlignment = TextAlignment.Center,
            .FontSize = 20
        }
        Me.Children.Add(_textBlock)

        _visual = ElementCompositionPreview.GetElementVisual(_container)
        AddHandler _container.SizeChanged, Sub(s, e)
                                               _visual.CenterPoint = New Vector3(CSng(_container.ActualWidth / 2), CSng(_container.ActualHeight / 2), 0)
                                           End Sub

        AddHandler _container.PointerEntered, Sub(s, e)
                                                  If _isEnabled AndAlso _isPointerDown Then
                                                      _pressedInside = True
                                                      AnimateScale(0.9F)
                                                      _container.Background = _containerBackgroundOver
                                                  ElseIf _isEnabled Then
                                                      _container.Background = _containerBackgroundOver
                                                  End If
                                              End Sub

        AddHandler _container.PointerExited, Sub(s, e)
                                                 If _isEnabled AndAlso _isPointerDown Then
                                                     _pressedInside = False
                                                     AnimateScale(1.0F)
                                                     _container.Background = _containerBackground
                                                 ElseIf _isEnabled Then
                                                     _container.Background = _containerBackground
                                                 End If
                                             End Sub

        AddHandler _container.PointerPressed, Sub(s, e)
                                                  If _isEnabled Then
                                                      _isPointerDown = True
                                                      _pressedInside = True
                                                      _container.CapturePointer(e.Pointer)
                                                      AnimateScale(0.9F)
                                                      'SetShadowOffset(2) ' Smaller shadow to simulate "pressed"
                                                  End If
                                              End Sub

        AddHandler _container.PointerReleased, Sub(s, e)
                                                   If _isEnabled Then
                                                       AnimateScale(1.0F)
                                                       'SetShadowOffset(6) ' Restore shadow
                                                       If _pressedInside Then
                                                           RaiseEvent Clicked(Me, EventArgs.Empty)
                                                       End If
                                                       _container.ReleasePointerCapture(e.Pointer)
                                                       _isPointerDown = False
                                                       '_pressedInside = False
                                                   End If
                                               End Sub
        AddHandler _container.PointerCaptureLost, Sub(s, e)
                                                      _isPointerDown = False
                                                      AnimateScale(1.0F)
                                                      If _isEnabled Then
                                                          If _pressedInside Then
                                                              _container.Background = _containerBackgroundOver
                                                          Else
                                                              _container.Background = _containerBackground
                                                          End If
                                                      End If
                                                      '_pressedInside = False
                                                      'SetShadowOffset(6)
                                                  End Sub
    End Sub

    Private Sub SetShadowOffset(offsetY As Single)
        Dim shadow = TryCast(_shadowSprite.Shadow, DropShadow)
        If shadow IsNot Nothing Then
            shadow.Offset = New Vector3(0, offsetY, 0)
        End If
    End Sub

    Private Sub AnimateScale(targetScale As Single)
        Dim scaleAnim = _compositor.CreateVector3KeyFrameAnimation()
        scaleAnim.InsertKeyFrame(1.0F, New Vector3(targetScale, targetScale, 1))
        scaleAnim.Duration = TimeSpan.FromMilliseconds(100)
        _visual.StartAnimation("Scale", scaleAnim)
        ' Also animate the shadow
        Dim shadowVisual = ElementCompositionPreview.GetElementVisual(_shadowVisualHost)
        shadowVisual.CenterPoint = New Vector3(CSng(_shadowVisualHost.ActualWidth / 2), CSng(_shadowVisualHost.ActualHeight / 2), 0)
        shadowVisual.StartAnimation("Scale", scaleAnim)
    End Sub

    Private Sub ShowRipple()
        Dim visual = ElementCompositionPreview.GetElementVisual(_ripple)
        Dim compositor = visual.Compositor

        ' Set the starting opacity
        visual.Opacity = 0.3F

        ' Create the animation
        Dim fadeOutAnim = compositor.CreateScalarKeyFrameAnimation()
        fadeOutAnim.InsertKeyFrame(1.0F, 0.0F)
        fadeOutAnim.Duration = TimeSpan.FromMilliseconds(300)

        ' Start the animation
        visual.StartAnimation("Opacity", fadeOutAnim)
    End Sub

    Public Property IsEnabled As Boolean
        Get
            Return _isEnabled
        End Get
        Set(value As Boolean)
            _isEnabled = value
            _textBlock.Opacity = If(value, 1.0, 0.5)
            _image.Opacity = If(value, 1.0, 0.5)
        End Set
    End Property

    Public Property ButtonLabel As String
        Get
            Return _label
        End Get
        Set(value As String)
            _label = value
            _textBlock.Text = value
        End Set
    End Property

    Public Property ButtonImageSource As Uri
        Get
            Return _imageUri
        End Get
        Set(value As Uri)
            _imageUri = value
            _image.Source = New BitmapImage(value)
        End Set
    End Property

    Public Property ButtonBorderBrush As Brush
        Get
            Return _imageBorder.BorderBrush
        End Get
        Set(value As Brush)
            _imageBorder.BorderBrush = value
        End Set
    End Property

    Public Property ButtonBorderThickness As Thickness
        Get
            Return _imageBorder.BorderThickness
        End Get
        Set(value As Thickness)
            _imageBorder.BorderThickness = value
        End Set
    End Property

End Class
