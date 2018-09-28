<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.HeadlettersGenChartInstaller = New System.ServiceProcess.ServiceProcessInstaller()
        Me.HeadlettersGenChartService = New System.ServiceProcess.ServiceInstaller()
        '
        'HeadlettersGenChartInstaller
        '
        Me.HeadlettersGenChartInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.HeadlettersGenChartInstaller.Password = Nothing
        Me.HeadlettersGenChartInstaller.Username = Nothing
        '
        'HeadlettersGenChartService
        '
        Me.HeadlettersGenChartService.Description = "Headletters Generate Chart triggers in every 1 minute"
        Me.HeadlettersGenChartService.DisplayName = "Headletters_GenChart"
        Me.HeadlettersGenChartService.ServiceName = "HeadlettersGenChartService"
        Me.HeadlettersGenChartService.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.HeadlettersGenChartInstaller, Me.HeadlettersGenChartService})

    End Sub

    Friend WithEvents HeadlettersGenChartInstaller As ServiceProcess.ServiceProcessInstaller
    Friend WithEvents HeadlettersGenChartService As ServiceProcess.ServiceInstaller
End Class
