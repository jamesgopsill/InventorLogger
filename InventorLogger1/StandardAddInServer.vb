Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace InventorLogger1
    <ProgIdAttribute("InventorLogger1.StandardAddInServer"), _
    GuidAttribute("43940efe-7931-4e2a-8e3c-707b4ef86b1c")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private m_inventorApplication As Inventor.Application

        ' My Extras
        Private WithEvents m_ApplicationEvents As Inventor.ApplicationEvents
        Private WithEvents m_AssemblyEvents As Inventor.AssemblyEvents
        Private WithEvents m_FileAccessEvents As Inventor.FileAccessEvents
        Private WithEvents m_FileManagerEvents As Inventor.FileManagerEvents
        Private WithEvents m_FileUIEvents As Inventor.FileUIEvents
        Private WithEvents m_ModelingEvents As Inventor.ModelingEvents
        Private WithEvents m_PanelBar As Inventor.PanelBar
        Private WithEvents m_UserInputEvents As Inventor.UserInputEvents
        Private WithEvents m_UserInterfaceEvents As Inventor.UserInterfaceEvents
        Private WithEvents m_RepresentationEvents As Inventor.RepresentationEvents
        Private WithEvents m_SketchEvents As Inventor.SketchEvents
        Private WithEvents m_StyleEvents As Inventor.StyleEvents
        Private WithEvents m_TransactionEvents As Inventor.TransactionEvents

        Private Shared logFileName As String = Nothing
        Private Shared logFileFullName As String = Nothing


#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application


            ' Log File Name
            logFileName = String.Concat(System.Environment.UserName, "_(", Date.Now().ToString, ")_inventor.log")
            logFileName = logFileName.Replace(" ", "_")
            logFileName = logFileName.Replace(":", "_")
            logFileName = logFileName.Replace("/", "_")

            logFileFullName = String.Concat("C:\ProgramData\InventorLogs\", logFileName)

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.
            Log(Date.Now)
            Log("- Application Loaded -")
            Log(String.Concat("User: ", System.Environment.UserName))
            Log(String.Concat("Machine Name: ", System.Environment.MachineName))

            m_ApplicationEvents = m_inventorApplication.ApplicationEvents
            m_AssemblyEvents = CType(m_inventorApplication.AssemblyEvents, Inventor.AssemblyEvents)
            m_UserInputEvents = m_inventorApplication.CommandManager.UserInputEvents
            m_UserInterfaceEvents = m_inventorApplication.UserInterfaceManager.UserInterfaceEvents
            m_FileAccessEvents = m_inventorApplication.FileAccessEvents
            m_FileManagerEvents = m_inventorApplication.FileManager.FileManagerEvents
            m_FileUIEvents = m_inventorApplication.FileUIEvents
            m_ModelingEvents = m_inventorApplication.ModelingEvents
            m_PanelBar = m_inventorApplication.UserInterfaceManager.ActiveEnvironment.PanelBar
            m_RepresentationEvents = m_inventorApplication.RepresentationEvents
            m_SketchEvents = m_inventorApplication.SketchEvents
            m_StyleEvents = m_inventorApplication.StyleEvents
            m_TransactionEvents = m_inventorApplication.TransactionManager.TransactionEvents


            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.

        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Opening the access rights to the files for all users (Does not work)
            'Dim fSecurity As System.Security.AccessControl.FileSecurity = System.IO.File.GetAccessControl(logFileFullName)
            'Dim accessRule As System.Security.AccessControl.FileSystemAccessRule = New System.Security.AccessControl.FileSystemAccessRule("IT051731\Users", System.Security.AccessControl.FileSystemRights.Modify, System.Security.AccessControl.AccessControlType.Allow)
            'fSecurity.AddAccessRule(accessRule)
            'System.IO.File.SetAccessControl(logFileFullName, fSecurity)

            'Dim moveFileTo As String = String.Concat("C:\ProgramData\InventorLogs\archived\", logFileName)
            'Try
            'Dim moveFileTo As String = String.Concat("\\ads.bris.ac.uk\filestore\Engineering\TeachingLabs\data\MENG\MENG26000\Logs\", logFileName)
            'System.IO.File.Move(logFileFullName, moveFileTo)
            'Catch ex As Exception
            ' if it cannot be moved then just delete the log so it doesn't fill up all the computers
            'Try
            'System.IO.File.Delete(logFileFullName)
            'Catch exc As Exception

            'End Try
            'End Try

            ' Release objects.
            Marshal.ReleaseComObject(m_inventorApplication)
            m_inventorApplication = Nothing

            Marshal.ReleaseComObject(m_ApplicationEvents)
            m_ApplicationEvents = Nothing

            Marshal.ReleaseComObject(m_AssemblyEvents)
            m_AssemblyEvents = Nothing

            Marshal.ReleaseComObject(m_UserInputEvents)
            m_UserInputEvents = Nothing

            Marshal.ReleaseComObject(m_UserInterfaceEvents)
            m_UserInterfaceEvents = Nothing

            Marshal.ReleaseComObject(m_FileAccessEvents)
            m_FileAccessEvents = Nothing

            Marshal.ReleaseComObject(m_FileManagerEvents)
            m_FileManagerEvents = Nothing

            Marshal.ReleaseComObject(m_FileUIEvents)
            m_FileUIEvents = Nothing

            Marshal.ReleaseComObject(m_ModelingEvents)
            m_ModelingEvents = Nothing

            Marshal.ReleaseComObject(m_PanelBar)
            m_PanelBar = Nothing

            Marshal.ReleaseComObject(m_RepresentationEvents)
            m_RepresentationEvents = Nothing

            Marshal.ReleaseComObject(m_SketchEvents)
            m_SketchEvents = Nothing

            Marshal.ReleaseComObject(m_StyleEvents)
            m_StyleEvents = Nothing

            Marshal.ReleaseComObject(m_TransactionEvents)
            m_TransactionEvents = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()

        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

        ' Log File Creation and Appending
        Public Shared Sub Log(logMessage As String)
            Using w As System.IO.StreamWriter = System.IO.File.AppendText(logFileFullName)
                w.WriteLine(logMessage)
            End Using
        End Sub

        ' Capture Application Events

        Private Sub m_ApplicationEvents_OnActivateDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnActivateDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnActivateDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnActivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnActivateView
            Log(Date.Now)
            Log("ApplicationEvents.OnActivateView")
        End Sub

        Private Sub m_ApplicationEvents_OnActiveProjectChanged(ByVal ProjectObject As Inventor.DesignProject, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnActiveProjectChanged
            Log(Date.Now)
            Log("ApplicationEvents.OnActiveProjectChanged")
        End Sub

        Private Sub m_ApplicationEvents_OnApplicationOptionChange(ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnApplicationOptionChange
            Log(Date.Now)
            Log("ApplicationEvents.OnApplicationOptionChange")
        End Sub

        Private Sub m_ApplicationEvents_OnCloseDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnCloseDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnCloseDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnCloseView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnCloseView
            Log(Date.Now)
            Log("ApplicationEvents.OnCloseView")
        End Sub

        Private Sub m_ApplicationEvents_OnDeactivateDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnDeactivateDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnDeactivateDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnDeactivateView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnDeactivateView
            Log(Date.Now)
            Log("ApplicationEvents.OnDeactivateView")
        End Sub

        Private Sub m_ApplicationEvents_OnDisplayModeChange(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnDisplayModeChange
            Log(Date.Now)
            Log("ApplicationEvents.OnDisplayModeChange")
        End Sub

        Private Sub m_ApplicationEvents_OnDocumentChange(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal ReasonsForChange As Inventor.CommandTypesEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnDocumentChange
            Log(Date.Now)
            Log("ApplicationEvents.OnDocumentChange")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnInitializeDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnInitializeDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnInitializeDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnMigrateDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnMigrateDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnMigrateDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnNewDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnNewDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnNewDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnNewEditObject(ByVal EditObject As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnNewEditObject
            Log(Date.Now)
            Log("ApplicationEvents.OnNewEditObject")
        End Sub

        Private Sub m_ApplicationEvents_OnNewView(ByVal ViewObject As Inventor.View, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnNewView
            Log(Date.Now)
            Log("ApplicationEvents.OnNewView")
        End Sub

        Private Sub m_ApplicationEvents_OnOpenDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnOpenDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnOpenDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnQuit(ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnQuit
            Log(Date.Now)
            Log("ApplicationEvents.OnQuit")
        End Sub

        Private Sub m_ApplicationEvents_OnReady(ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnReady
            Log(Date.Now)
            Log("ApplicationEvents.OnReady")
        End Sub

        Private Sub m_ApplicationEvents_OnRestart32BitHost(ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnRestart32BitHost
            Log(Date.Now)
            Log("ApplicationEvents.OnRestart32BitHostt")
        End Sub

        Private Sub m_ApplicationEvents_OnSaveDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnSaveDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnSaveDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnTerminateDocument(ByVal DocumentObject As Inventor._Document, ByVal FullDocumentName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnTerminateDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnTerminateDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        Private Sub m_ApplicationEvents_OnTranslateDocument(ByVal TranslatingIn As Boolean, ByVal DocumentObject As Inventor._Document, ByVal FullFileName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ApplicationEvents.OnTranslateDocument
            Log(Date.Now)
            Log("ApplicationEvents.OnTranslateDocument")
            Log(String.Concat("Document Display Name: ", DocumentObject.DisplayName))
        End Sub

        ' Assembly Events
        Private Sub m_AssemblyEvents_OnAssemblyChange(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnAssemblyChange
            Log(Date.Now)
            Log("AssemblyEvents.OnAssemblyChange")
        End Sub

        Private Sub m_AssemblyEvents_OnAssemblyChanged(ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnAssemblyChanged
            Log(Date.Now)
            Log("AssemblyEvents.OnAssemblyChanged")
        End Sub

        Private Sub m_AssemblyEvents_OnAssemblySolve(ByVal Solver As Inventor._AssemblySolver, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap) Handles m_AssemblyEvents.OnAssemblySolve
            Log(Date.Now)
            Log("AssemblyEvents.OnAssemblySolve")
        End Sub

        Private Sub m_AssemblyEvents_OnDelete(ByVal DocumentObject As Inventor._Document, ByVal Entity As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnDelete
            Log(Date.Now)
            Log("AssemblyEvents.OnDelete")
        End Sub

        Private Sub m_AssemblyEvents_OnNewConstraint(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Constraint As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnNewConstraint
            Log(Date.Now)
            Log("AssemblyEvents.OnNewConstraint")
        End Sub

        Private Sub m_AssemblyEvents_OnNewOccurrence(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Occurrence As Inventor.ComponentOccurrence, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnNewOccurrence
            Log(Date.Now)
            Log("AssemblyEvents.OnNewOccurrence")
        End Sub

        Private Sub m_AssemblyEvents_OnOccurrenceChange(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Occurrence As Inventor.ComponentOccurrence, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_AssemblyEvents.OnOccurrenceChange
            Log(Date.Now)
            Log("AssemblyEvents.OnOccurrenceChange")
        End Sub

        ' File Access Events
        Private Sub m_FileAccessEvents_OnFileDirty(ByVal RelativeFileName As String, ByVal LibraryName As String, ByRef CustomLogicalName() As Byte, ByVal FullFileName As String, ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileAccessEvents.OnFileDirty
            Log(Date.Now)
            Log("FileAccessEvents.OnFileDirty")
            Log(String.Concat("Relative File Name: ", RelativeFileName))
            Log(String.Concat("Library Name: ", LibraryName))
            Log(String.Concat("Full File Name: ", FullFileName))
        End Sub

        Private Sub m_FileAccessEvents_OnFileResolution(ByVal RelativeFileName As String, ByVal LibraryName As String, ByRef CustomLogicalName() As Byte, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef FullFileName As String, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileAccessEvents.OnFileResolution
            Log(Date.Now)
            Log("FileAccessEvents.OnFileResolution")
            Log(String.Concat("Relative File Name: ", RelativeFileName))
            Log(String.Concat("Library Name: ", LibraryName))
            Log(String.Concat("Full File Name: ", FullFileName))
        End Sub

        ' File Manager Events
        Private Sub m_FileManagerEvents_OnFileDelete(ByVal FullFileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileManagerEvents.OnFileDelete
            Log(Date.Now)
            Log("FileManagerEvents.OnFileDelete")
            Log(String.Concat("Full File Name: ", FullFileName))
        End Sub

        Private Sub m_FileManagerEvents_OnFileCopy(ByVal SourceFullFileName As String, ByVal DestinationFullFileName As String, ByVal Copy As Boolean, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileManagerEvents.OnFileCopy
            Log(Date.Now)
            Log("FileManagerEvents.OnFileCopy")
            Log(String.Concat("Source Full File Name: ", SourceFullFileName))
            Log(String.Concat("Destination Full File Name: ", DestinationFullFileName))
        End Sub

        ' File UI Events
        Private Sub m_FileUIEvents_OnFileInsertDialog(ByRef FileTypes() As String, ByVal DocumentObject As Inventor._Document, ByVal ParentHWND As Integer, ByRef FileName As String, ByRef RelativeFileName As String, ByRef LibraryName As String, ByRef CustomLogicalName() As Byte, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileInsertDialog
            Log(Date.Now)
            Log("FileUIEvents.OnFileInsertDialog")
            Log(String.Concat("File Name: ", FileName))
            Log(String.Concat("Relative File Name: ", RelativeFileName))
            Log(String.Concat("Library Name: ", LibraryName))
        End Sub

        Private Sub m_FileUIEvents_OnFileInsertNewDialog(ByVal TemplateDir As String, ByRef FileTypes() As String, ByVal DocumentObject As Inventor._Document, ByVal ParentHWND As Integer, ByRef TemplateFileName As String, ByRef FileName As String, ByRef RelativeFileName As String, ByRef LibraryName As String, ByRef CustomLogicalName() As Byte, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileInsertNewDialog
            Log(Date.Now)
            Log("FileUIEvents.OnFileInsertNewDialog")
            Log(String.Concat("Template Directory Name: ", TemplateDir))
            Log(String.Concat("Template File Name: ", TemplateFileName))
            Log(String.Concat("File Name: ", FileName))
            Log(String.Concat("Relative File Name: ", RelativeFileName))
            Log(String.Concat("Library Name: ", LibraryName))
        End Sub

        Private Sub m_FileUIEvents_OnFileNew(ByVal DocumentType As Inventor.DocumentTypeEnum, ByRef TemplateFileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileNew
            Log(Date.Now)
            Log("FileUIEvents.OnFileNew")
            Log(String.Concat("Template File Name: ", TemplateFileName))
        End Sub

        Private Sub m_FileUIEvents_OnFileNewDialog(ByVal TemplateDir As String, ByVal ParentHWND As Integer, ByRef TemplateFileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileNewDialog
            Log(Date.Now)
            Log("FileUIEvents.OnFileNewDialog")
            Log(String.Concat("Template Directory Name: ", TemplateDir))
            Log(String.Concat("Template File Name: ", TemplateFileName))
        End Sub

        Private Sub m_FileUIEvents_OnFileOpenDialog(ByRef FileTypes() As String, ByVal ParentHWND As Integer, ByRef FileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileOpenDialog
            Log(Date.Now)
            Log("FileUIEvents.OnFileOpenDialog")
            ' Log(String.Concat("File Types: ", FileTypes))
            Log(String.Concat("File Name: ", FileName))
        End Sub

        Private Sub m_FileUIEvents_OnFileOpenFromMRU(ByRef FullFileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileOpenFromMRU
            Log(Date.Now)
            Log("FileUIEvents.OnFileOpenFromMRU")
            Log(String.Concat("Full File Name: ", FullFileName))
        End Sub

        Private Sub m_FileUIEvents_OnFileSaveAsDialog(ByRef FileTypes() As String, ByVal SaveCopyAs As Boolean, ByVal ParentHWND As Integer, ByRef FileName As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnFileSaveAsDialog
            Log(Date.Now)
            Log("FileUIEvents.OnFileSaveAsDialog")
            ' Log(String.Concat("File Types: ", FileTypes))
            Log(String.Concat("File Name: ", FileName))
        End Sub

        Private Sub m_FileUIEvents_OnPopulateFileMetadata(ByVal FileMetadataObjects As Inventor.ObjectsEnumerator, ByVal Formulae As String, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_FileUIEvents.OnPopulateFileMetadata
            Log(Date.Now)
            Log("FileUIEvents.OnPopulateFileMetadata")
            Log(String.Concat("Formulae: ", Formulae))
        End Sub

        ' Modeling Events
        Private Sub m_ModelingEvents_OnClientFeatureDoubleClick(ByVal DocumentObject As Inventor._Document, ByVal Feature As Inventor.ClientFeature, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnClientFeatureDoubleClick
            Log(Date.Now)
            Log("ModelingEvents.OnClientFeatureDoubleClick")
        End Sub

        Private Sub m_ModelingEvents_OnDelete(ByVal DocumentObject As Inventor._Document, ByVal Entity As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnDelete
            Log(Date.Now)
            Log("ModelingEvents.OnDelete")
        End Sub

        Private Sub m_ModelingEvents_OnFeatureChange(ByVal DocumentObject As Inventor._Document, ByVal Feature As Inventor.PartFeature, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnFeatureChange
            Log(Date.Now)
            Log("ModelingEvents.OnFeatureChange")
        End Sub

        Private Sub m_ModelingEvents_OnGenerateMember(ByVal FactoryDocumentObject As Inventor._Document, ByVal MemberName As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnGenerateMember
            Log(Date.Now)
            Log("ModelingEvents.OnGenerateMember")
            Log(String.Concat("Member Name: ", MemberName))
        End Sub

        Private Sub m_ModelingEvents_OnNewFeature(ByVal DocumentObject As Inventor._Document, ByVal Feature As Inventor.PartFeature, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnNewFeature
            Log(Date.Now)
            Log("ModelingEvents.OnNewFeature")
        End Sub

        Private Sub m_ModelingEvents_OnNewParameter(ByVal DocumentObject As Inventor._Document, ByVal Parameter As Inventor.Parameter, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnNewParameter
            Log(Date.Now)
            Log("ModelingEvents.OnNewParameter")
        End Sub

        Private Sub m_ModelingEvents_OnParameterChange(ByVal DocumentObject As Inventor._Document, ByVal Parameter As Inventor.Parameter, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_ModelingEvents.OnParameterChange
            Log(Date.Now)
            Log("ModelingEvents.OnParameterChange")
        End Sub

        ' Panel Bar Events
        Private Sub m_PanelBar_OnCommandBarSelection(ByVal CommandBarObject As Inventor.CommandBar, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_PanelBar.OnCommandBarSelection
            Log(Date.Now)
            Log("PanelBar.OnCommandBarSelection")
            Log(String.Concat("Command Bar Object Display Name: ", CommandBarObject.DisplayName))
        End Sub

        ' Representation Events
        Private Sub m_RepresentationEvents_OnActivateDesignViewRepresentation(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Representation As Inventor.DesignViewRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnActivateDesignViewRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnActivateDesignViewRepresentation")
        End Sub

        Private Sub m_RepresentationEvents_OnActivateLevelOfDetailRepresentation(ByVal DocumentObject As Inventor._Document, ByVal Representation As Inventor.LevelOfDetailRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnActivateLevelOfDetailRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnActivateLevelOfDetailRepresentation")
        End Sub

        Private Sub m_RepresentationEvents_OnActivatePositionalRepresentation(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Representation As Inventor.PositionalRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnActivatePositionalRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnActivatePositionalRepresentation")
        End Sub

        Private Sub m_RepresentationEvents_OnDelete(ByVal DocumentObject As Inventor._Document, ByVal Entity As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnDelete
            Log(Date.Now)
            Log("RepresentationEvents.OnDelete")
        End Sub

        Private Sub m_RepresentationEvents_OnNewDesignViewRepresentation(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Representation As Inventor.DesignViewRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnNewDesignViewRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnNewDesignViewRepresentation")
        End Sub

        Private Sub m_RepresentationEvents_OnNewLevelOfDetailRepresentation(ByVal DocumentObject As Inventor._Document, ByVal Representation As Inventor.LevelOfDetailRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnNewLevelOfDetailRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnNewLevelOfDetailRepresentation")
        End Sub

        Private Sub m_RepresentationEvents_OnNewPositionalRepresentation(ByVal DocumentObject As Inventor._AssemblyDocument, ByVal Representation As Inventor.PositionalRepresentation, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_RepresentationEvents.OnNewPositionalRepresentation
            Log(Date.Now)
            Log("RepresentationEvents.OnNewPositionalRepresentation")
        End Sub

        ' Sketch Events
        Private Sub m_SketchEvents_OnDelete(ByVal DocumentObject As Inventor._Document, ByVal Entity As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnDelete
            Log(Date.Now)
            Log("SketchEvents.OnDelete")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_SketchEvents_OnNewSketch(ByVal DocumentObject As Inventor._Document, ByVal Sketch As Inventor.Sketch, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnNewSketch
            Log(Date.Now)
            Log("SketchEvents.OnNewSketch")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_SketchEvents_OnNewSketch3D(ByVal DocumentObject As Inventor._Document, ByVal Sketch3D As Inventor.Sketch3D, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnNewSketch3D
            Log(Date.Now)
            Log("SketchEvents.OnNewSketch3D")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_SketchEvents_OnSketch3DChange(ByVal DocumentObject As Inventor._Document, ByVal Sketch3D As Inventor.Sketch3D, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnSketch3DChange
            Log(Date.Now)
            Log("SketchEvents.OnSketch3DChange")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_SketchEvents_OnSketch3DSolve(ByVal DocumentObject As Inventor._Document, ByVal Sketch As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnSketch3DSolve
            Log(Date.Now)
            Log("SketchEvents.OnSketch3DSolve")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_SketchEvents_OnSketchChange(ByVal DocumentObject As Inventor._Document, ByVal Sketch As Inventor.Sketch, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_SketchEvents.OnSketchChange
            Log(Date.Now)
            Log("SketchEvents.OnSketchChange")
            Log(String.Concat("Display Name:", DocumentObject.DisplayName))
            Log(String.Concat("File Name:", DocumentObject.FullFileName))
        End Sub

        Private Sub m_StyleEvents_OnActivateStyle(ByVal DocumentObject As Inventor._Document, ByVal Style As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_StyleEvents.OnActivateStyle
            Log(Date.Now)
            Log("StyleEvents.OnActivateStyle")
        End Sub

        Private Sub m_StyleEvents_OnDelete(ByVal DocumentObject As Inventor._Document, ByVal Style As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_StyleEvents.OnDelete
            Log(Date.Now)
            Log("StyleEvents.OnDelete")
        End Sub

        Private Sub m_StyleEvents_OnMigrateStyleLibrary(ByVal LibraryPath As String, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_StyleEvents.OnMigrateStyleLibrary
            Log(Date.Now)
            Log("StyleEvents.OnMigrateStyleLibrary")
            Log(String.Concat("Library Path:", LibraryPath))
        End Sub

        Private Sub m_StyleEvents_OnNewStyle(ByVal DocumentObject As Inventor._Document, ByVal Style As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_StyleEvents.OnNewStyle
            Log(Date.Now)
            Log("StyleEvents.OnNewStyle")
        End Sub

        Private Sub m_StyleEvents_OnStyleChange(ByVal DocumentObject As Inventor._Document, ByVal Style As Object, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_StyleEvents.OnStyleChange
            Log(Date.Now)
            Log("StyleEvents.OnStyleChange")
        End Sub

        ' Transaction Events
        Private Sub m_TransactionEvents_OnAbort(ByVal TransactionObject As Inventor.Transaction, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum) Handles m_TransactionEvents.OnAbort
            Log(Date.Now)
            Log("TransactionEvents.OnAbort")
        End Sub

        Private Sub m_TransactionEvents_OnCommit(ByVal TransactionObject As Inventor.Transaction, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_TransactionEvents.OnCommit
            Log(Date.Now)
            Log("TransactionEvents.OnCommit")
        End Sub

        Private Sub m_TransactionEvents_OnDelete(ByVal TransactionObject As Inventor.Transaction, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum) Handles m_TransactionEvents.OnDelete
            Log(Date.Now)
            Log("TransactionEvents.OnDelete")
        End Sub

        Private Sub m_TransactionEvents_OnRedo(ByVal TransactionObject As Inventor.Transaction, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_TransactionEvents.OnRedo
            Log(Date.Now)
            Log("TransactionEvents.OnRedo")
        End Sub

        Private Sub m_TransactionEvents_OnUndo(ByVal TransactionObject As Inventor.Transaction, ByVal Context As Inventor.NameValueMap, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_TransactionEvents.OnUndo
            Log(Date.Now)
            Log("TransactionEvents.OnUndo")
        End Sub

        ' User Input Events Capture
        Private Sub m_UserInputEvents_OnActivateCommand(ByVal CommandName As String, ByVal Context As Inventor.NameValueMap) Handles m_UserInputEvents.OnActivateCommand
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnActivateCommand")
            'Console.WriteLine(CommandName)
            Log(Date.Now)
            Log("UserInputEvents.OnActivateCommand")
            Log(CommandName)
        End Sub

        Private Sub m_UserInputEvents_OnContextMenu(ByVal SelectionDevice As Inventor.SelectionDeviceEnum, ByVal AdditionalInfo As Inventor.NameValueMap, ByVal CommandBar As Inventor.CommandBar) Handles m_UserInputEvents.OnContextMenu
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnContextMenu")
            Log(Date.Now)
            Log("UserInputEvents.OnContextMenu")
        End Sub

        Private Sub m_UserInputEvents_OnContextMenuOld(ByVal SelectionDevice As Inventor.SelectionDeviceEnum, ByVal AdditionalInfo As Inventor.NameValueMap, ByVal CommandBar As Inventor.CommandBar) Handles m_UserInputEvents.OnContextMenuOld
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnContextMenuOld")
            Log(Date.Now)
            Log("UserInputEvents.OnContextMenuOld")
        End Sub

        Private Sub m_UserInputEvents_OnDoubleClick(ByVal SelectedEntities As Inventor.ObjectsEnumerator, ByVal SelectionDevice As Inventor.SelectionDeviceEnum, ByVal Button As Inventor.MouseButtonEnum, ByVal ShiftKeys As Inventor.ShiftStateEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Inventor.Point2d, ByVal View As Inventor.View, ByVal AdditionalInfo As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_UserInputEvents.OnDoubleClick
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnDoubleClick")
            Log(Date.Now)
            Log("UserInputEvents.OnDoubleClick")
        End Sub

        Private Sub m_UserInputEvents_OnDrag(ByVal DragState As Inventor.DragStateEnum, ByVal ShiftKeys As Inventor.ShiftStateEnum, ByVal ModelPosition As Inventor.Point, ByVal ViewPosition As Inventor.Point2d, ByVal View As Inventor.View, ByVal AdditionalInfo As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_UserInputEvents.OnDrag
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnDrag")
            Log(Date.Now)
            Log("UserInputEvents.OnDrag")
        End Sub

        Private Sub m_UserInputEvents_OnTerminateCommand(ByVal CommandName As String, ByVal Context As Inventor.NameValueMap) Handles m_UserInputEvents.OnTerminateCommand
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInputEvents.OnTerminateCommand")
            'Console.WriteLine(CommandName)
            Log(Date.Now)
            Log("UserInputEvents.OnTerminateCommand")
            Log(CommandName)
        End Sub

        ' User Interface Events Capture
        Private Sub m_UserInterfaceEvents_OnEnvironmentChange(ByVal Environment As Inventor.Environment, ByVal EnvironmentState As Inventor.EnvironmentStateEnum, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_UserInterfaceEvents.OnEnvironmentChange
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInterfaceEvents.OnEnvironmentChange")
            Log(Date.Now)
            Log("UserInterfaceEvents.OnEnvironmentChange")
            Log(String.Concat("Display Name:", Environment.DisplayName))
        End Sub

        Private Sub m_UserInterfaceEvents_OnResetCommandBars(ByVal CommandBars As Inventor.ObjectsEnumerator, ByVal Context As Inventor.NameValueMap) Handles m_UserInterfaceEvents.OnResetCommandBars
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInterfaceEvents.OnResetCommandBars")
            Log(Date.Now)
            Log("UserInterfaceEvents.OnResetCommandBars")
        End Sub

        Private Sub m_UserInterfaceEvents_OnResetEnvironments(ByVal Environments As Inventor.ObjectsEnumerator, ByVal Context As Inventor.NameValueMap) Handles m_UserInterfaceEvents.OnResetEnvironments
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInterfaceEvents.OnResetEnvironments")
            Log(Date.Now)
            Log("UserInterfaceEvents.OnResetEnvironments")
        End Sub

        Private Sub m_UserInterfaceEvents_OnResetShortcuts(ByVal Context As Inventor.NameValueMap) Handles m_UserInterfaceEvents.OnResetShortcuts
            'Console.WriteLine(Date.Now)
            'Console.WriteLine("UserInterfaceEvents.OnResetShortcuts")
            Log(Date.Now)
            Log("UserInterfaceEvents.OnResetShortcuts")
        End Sub


#End Region

    End Class

End Namespace

