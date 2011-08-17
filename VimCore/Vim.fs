﻿#light

namespace Vim
open Microsoft.VisualStudio.Text
open Microsoft.VisualStudio.Text.Editor
open Microsoft.VisualStudio.Text.Operations
open Microsoft.VisualStudio.Text.Outlining
open Microsoft.VisualStudio.Text.Classification
open System.ComponentModel.Composition
open Vim.Modes

type internal VimData() =

    let mutable _commandHistory = HistoryList()
    let mutable _searchHistory = HistoryList()
    let mutable _lastSubstituteData : SubstituteData option = None
    let mutable _lastPatternData = { Pattern = StringUtil.empty; Path = Path.Forward }
    let mutable _lastCharSearch : (CharSearchKind * Path * char) option = None
    let mutable _lastMacroRun : char option = None
    let mutable _lastCommand : StoredCommand option = None
    let _lastPatternDataChanged = Event<PatternData>()
    let _highlightSearchOneTimeDisabled = Event<unit>()

    interface IVimData with 
        member x.CommandHistory
            with get () = _commandHistory
            and set value = _commandHistory <- value
        member x.SearchHistory 
            with get () = _searchHistory
            and set value = _searchHistory <- value
        member x.LastSubstituteData 
            with get () = _lastSubstituteData
            and set value = _lastSubstituteData <- value
        member x.LastCommand 
            with get () = _lastCommand
            and set value = _lastCommand <- value
        member x.LastPatternData 
            with get () = _lastPatternData
            and set value = 
                _lastPatternData <- value
                _lastPatternDataChanged.Trigger value
        member x.LastCharSearch 
            with get () = _lastCharSearch
            and set value = _lastCharSearch <- value
        member x.LastMacroRun 
            with get () = _lastMacroRun
            and set value = _lastMacroRun <- value
        member x.RaiseHighlightSearchOneTimeDisable () = _highlightSearchOneTimeDisabled.Trigger ()
        [<CLIEvent>]
        member x.LastPatternDataChanged = _lastPatternDataChanged.Publish
        [<CLIEvent>]
        member x.HighlightSearchOneTimeDisabled = _highlightSearchOneTimeDisabled.Publish

[<Export(typeof<IVimBufferFactory>)>]
type internal VimBufferFactory

    [<ImportingConstructor>]
    (
        _host : IVimHost,
        _editorOperationsFactoryService : IEditorOperationsFactoryService,
        _editorOptionsFactoryService : IEditorOptionsFactoryService,
        _outliningManagerService : IOutliningManagerService,
        _completionWindowBrokerFactoryService : IDisplayWindowBrokerFactoryService,
        _commonOperationsFactory : ICommonOperationsFactory,
        _wordUtilFactory : IWordUtilFactory,
        _textChangeTrackerFactory : ITextChangeTrackerFactory,
        _textSearchService : ITextSearchService,
        _smartIndentationService : ISmartIndentationService,
        _tlcService : ITrackingLineColumnService,
        _undoManagerProvider : ITextBufferUndoManagerProvider,
        _statusUtilFactory : IStatusUtilFactory,
        _foldManagerFactory : IFoldManagerFactory,
        _keyboardDevice : IKeyboardDevice,
        _mouseDevice : IMouseDevice,
        _wordCompletionSessionFactoryService : IWordCompletionSessionFactoryService
    ) = 

    /// Create an IVimTextBuffer instance for the provided ITextBuffer
    member x.CreateVimTextBuffer textBuffer (vim : IVim) = 
        let localSettings = LocalSettings(vim.GlobalSettings) :> IVimLocalSettings
        let jumpList = JumpList(_tlcService) :> IJumpList
        let wordUtil = _wordUtilFactory.GetWordUtil textBuffer
        let wordNavigator = wordUtil.CreateTextStructureNavigator WordKind.NormalWord
        VimTextBuffer(textBuffer, localSettings, jumpList, wordNavigator, vim)

    /// Create an IVimBuffer instance for the provided ITextView and IVimTextBuffer which is associated
    /// with the ITextBuffer of the ITextView
    member x.CreateVimBuffer (textView : ITextView) (vimTextBuffer : IVimTextBuffer) =
        Contract.Requires (vimTextBuffer.TextBuffer = textView.TextBuffer)

        let vim = vimTextBuffer.Vim
        let textBuffer = textView.TextBuffer
        let editOperations = _editorOperationsFactoryService.GetEditorOperations(textView)
        let editOptions = _editorOptionsFactoryService.GetOptions(textView)
        let statusUtil = _statusUtilFactory.GetStatusUtil textView
        let localSettings = vimTextBuffer.LocalSettings
        let jumpList = vimTextBuffer.JumpList

        let undoRedoOperations = 
            let history = 
                let manager = _undoManagerProvider.GetTextBufferUndoManager textBuffer
                if manager = null then None
                else manager.TextBufferUndoHistory |> Some
            UndoRedoOperations(statusUtil, history, editOperations) :> IUndoRedoOperations
        let bufferData : VimBufferData = {
            TextView = textView
            StatusUtil = statusUtil
            UndoRedoOperations = undoRedoOperations
            VimTextBuffer = vimTextBuffer }
        let commonOperations = _commonOperationsFactory.GetCommonOperations bufferData
        let wordUtil = _wordUtilFactory.GetWordUtil textBuffer

        let wordNav = wordUtil.CreateTextStructureNavigator WordKind.NormalWord
        let incrementalSearch = 
            IncrementalSearch(
                commonOperations,
                localSettings, 
                wordNav, 
                statusUtil,
                vim.VimData) :> IIncrementalSearch
        let capture = MotionCapture(vim.VimHost, textView, incrementalSearch, localSettings) :> IMotionCapture

        let textChangeTracker = _textChangeTrackerFactory.GetTextChangeTracker bufferData
        let motionUtil = MotionUtil(textView, vim.MarkMap, localSettings, vim.SearchService, wordNav, jumpList, statusUtil, wordUtil, vim.VimData) :> IMotionUtil
        let foldManager = _foldManagerFactory.GetFoldManager textView
        let insertUtil = InsertUtil(bufferData, commonOperations) :> IInsertUtil
        let commandUtil = CommandUtil(bufferData, motionUtil, commonOperations, _smartIndentationService, foldManager, wordNav, insertUtil) :> ICommandUtil
        let windowSettings = WindowSettings(vim.GlobalSettings, textView)

        let bufferRaw = VimBuffer(bufferData, incrementalSearch, motionUtil, wordNav, windowSettings)
        let buffer = bufferRaw :> IVimBuffer

        /// Create the selection change tracker so that it will begin to monitor
        /// selection events.  
        ///
        /// TODO: This feels wrong.  Either the result should be stored somewhere
        /// or it should be exposed as a MEF service that listens to buffer 
        /// creation events.
        let selectionChangeTracker = SelectionChangeTracker(buffer)

        let createCommandRunner kind = CommandRunner (textView, vim.RegisterMap, capture, commandUtil, statusUtil, kind) :>ICommandRunner
        let broker = _completionWindowBrokerFactoryService.CreateDisplayWindowBroker textView
        let bufferOptions = _editorOptionsFactoryService.GetOptions(textView.TextBuffer)
        let commandOpts = Modes.Command.DefaultOperations(commonOperations, textView, editOperations, jumpList, localSettings, undoRedoOperations, vim.KeyMap, vim.VimData, vim.VimHost, statusUtil) :> Modes.Command.IOperations
        let commandProcessor = Modes.Command.CommandProcessor(buffer, commonOperations, commandOpts, statusUtil, FileSystem() :> IFileSystem, foldManager) :> Modes.Command.ICommandProcessor
        let visualOptsFactory kind = 
            let kind = VisualKind.OfModeKind kind |> Option.get
            let tracker = Modes.Visual.SelectionTracker(textView, vim.GlobalSettings, incrementalSearch, kind) :> Modes.Visual.ISelectionTracker
            (tracker, commonOperations)

        let visualModeList =
            [ ModeKind.VisualBlock; ModeKind.VisualCharacter; ModeKind.VisualLine ]
            |> Seq.ofList
            |> Seq.map (fun kind -> 
                let tracker, opts = visualOptsFactory kind
                let visualKind = VisualKind.OfModeKind kind |> Option.get
                ((Modes.Visual.VisualMode(buffer, opts, kind, createCommandRunner visualKind,capture, tracker)) :> IMode) )
            |> List.ofSeq
    
        // Normal mode values
        let modeList = 
            [
                ((Modes.Normal.NormalMode(buffer, commonOperations, statusUtil,broker, createCommandRunner VisualKind.Character, capture)) :> IMode)
                ((Modes.Command.CommandMode(buffer, commandProcessor, commonOperations)) :> IMode)
                ((Modes.Insert.InsertMode(buffer, commonOperations, broker, editOptions, undoRedoOperations, textChangeTracker, insertUtil, false, _keyboardDevice, _mouseDevice, wordUtil, _wordCompletionSessionFactoryService)) :> IMode)
                ((Modes.Insert.InsertMode(buffer, commonOperations, broker, editOptions, undoRedoOperations, textChangeTracker, insertUtil, true, _keyboardDevice, _mouseDevice, wordUtil, _wordCompletionSessionFactoryService)) :> IMode)
                ((Modes.SubstituteConfirm.SubstituteConfirmMode(buffer, commonOperations) :> IMode))
                (DisabledMode(buffer) :> IMode)
                (ExternalEditMode(buffer) :> IMode)
            ] @ visualModeList
        modeList |> List.iter (fun m -> bufferRaw.AddMode m)
        buffer.SwitchMode vimTextBuffer.ModeKind ModeArgument.None |> ignore
        bufferRaw

    interface IVimBufferFactory with
        member x.CreateVimTextBuffer textBuffer vim = x.CreateVimTextBuffer textBuffer vim :> IVimTextBuffer
        member x.CreateVimBuffer textView vimTextBuffer = x.CreateVimBuffer textView vimTextBuffer :> IVimBuffer

/// Default implementation of IVim 
[<Export(typeof<IVim>)>]
type internal Vim
    (
        _host : IVimHost,
        _bufferFactoryService : IVimBufferFactory,
        _bufferCreationListeners : Lazy<IVimBufferCreationListener> list,
        _globalSettings : IVimGlobalSettings,
        _markMap : IMarkMap,
        _keyMap : IKeyMap,
        _clipboardDevice : IClipboardDevice,
        _search : ISearchService,
        _fileSystem : IFileSystem,
        _vimData : IVimData ) =

    /// Key for IVimTextBuffer instances inside of the ITextBuffer property bag
    let _vimTextBufferKey = System.Object()

    /// Holds an IVimBuffer and the DisposableBag for event handlers on the IVimBuffer.  This
    /// needs to be removed when we're done with the IVimBuffer to avoid leaks
    let _bufferMap = new System.Collections.Generic.Dictionary<ITextView, IVimBuffer * DisposableBag>()

    /// Holds the active stack of IVimBuffer instances
    let mutable _activeBufferStack : IVimBuffer list = List.empty

    /// Holds the setting information which was stored when loading the VimRc file.  This 
    /// is applied to IVimBuffer instances which are created when there is no active IVimBuffer
    let mutable _vimRcLocalSettings = LocalSettings(_globalSettings) :> IVimLocalSettings
    let mutable _vimRcWindowSettings : IVimWindowSettings option = None

    let _registerMap =
        let currentFileNameFunc() = 
            match _activeBufferStack with
            | [] -> None
            | h::_ -> 
                let name = _host.GetName h.TextBuffer 
                let name = System.IO.Path.GetFileName(name)
                Some name
        RegisterMap(_clipboardDevice, currentFileNameFunc) :> IRegisterMap

    let _recorder = MacroRecorder(_registerMap)

    /// Add the IMacroRecorder to the list of IVimBufferCreationListeners.  
    let _bufferCreationListeners =
        let item = Lazy<IVimBufferCreationListener>(fun () -> _recorder :> IVimBufferCreationListener)
        item :: _bufferCreationListeners

    do
        // When the 'history' setting is changed it impacts our history limits.  Keep track of 
        // them here
        //
        // Up cast here to work around the F# bug which prevents accessing a CLIEvent from
        // a derived type

        (_globalSettings :> IVimSettings).SettingChanged 
        |> Event.filter (fun args -> StringUtil.isEqual args.Name GlobalSettingNames.HistoryName)
        |> Event.add (fun _ -> 
            _vimData.SearchHistory.Limit <- _globalSettings.History
            _vimData.CommandHistory.Limit <- _globalSettings.History)

    [<ImportingConstructor>]
    new(
        host : IVimHost,
        bufferFactoryService : IVimBufferFactory,
        tlcService : ITrackingLineColumnService,
        [<ImportMany>] bufferCreationListeners : Lazy<IVimBufferCreationListener> seq,
        search : ITextSearchService,
        fileSystem : IFileSystem,
        clipboard : IClipboardDevice ) =
        let markMap = MarkMap(tlcService)
        let vimData = VimData() :> IVimData
        let globalSettings = GlobalSettings() :> IVimGlobalSettings
        let listeners = 
            [markMap :> IVimBufferCreationListener]
            |> Seq.map (fun t -> new Lazy<IVimBufferCreationListener>(fun () -> t))
            |> Seq.append bufferCreationListeners 
            |> List.ofSeq
        Vim(
            host,
            bufferFactoryService,
            listeners,
            globalSettings,
            markMap :> IMarkMap,
            KeyMap() :> IKeyMap,
            clipboard,
            SearchService(search, globalSettings) :> ISearchService,
            fileSystem,
            vimData)

    member x.ActiveBuffer = ListUtil.tryHeadOnly _activeBufferStack

    member x.Buffers = _bufferMap.Values |> Seq.map fst |> List.ofSeq

    member x.FocusedBuffer = 
        match _host.GetFocusedTextView() with
        | None -> 
            None
        | Some textView -> 
            let found, (buffer, _) = _bufferMap.TryGetValue(textView)
            if found then Some buffer
            else None

    /// Get the IVimLocalSettings which should be the basis for a newly created IVimTextBuffer
    member x.GetLocalSettingsForNewTextBuffer () =
        match x.ActiveBuffer with
        | Some buffer -> Some buffer.LocalSettings
        | None -> Some _vimRcLocalSettings

    /// Get the IVimWindowSettings which should be the basis for a newly created IVimBuffer
    member x.GetWindowSettingsForNewBuffer () =
        match x.ActiveBuffer with
        | Some buffer -> Some buffer.WindowSettings
        | None -> _vimRcWindowSettings

    /// Create an IVimTextBuffer for the given ITextBuffer.  If an IVimLocalSettings instance is 
    /// provided then attempt to copy them into the created IVimTextBuffer copy of the 
    /// IVimLocalSettings
    member x.CreateVimTextBuffer (textBuffer : ITextBuffer) (localSettings : IVimLocalSettings option) = 
        if textBuffer.Properties.ContainsProperty _vimTextBufferKey then
            invalidArg "textBuffer" Resources.Vim_TextViewAlreadyHasVimBuffer

        let vimTextBuffer = _bufferFactoryService.CreateVimTextBuffer textBuffer x

        // Apply the specified local buffer settings
        match localSettings with
        | None -> 
            ()
        | Some localSettings ->
            localSettings.AllSettings
            |> Seq.filter (fun s -> not s.IsGlobal && not s.IsValueCalculated)
            |> Seq.iter (fun s -> vimTextBuffer.LocalSettings.TrySetValue s.Name s.Value |> ignore)

        // Put the IVimTextBuffer into the ITextBuffer property bag so we can query for it in the future
        textBuffer.Properties.[_vimTextBufferKey] <- vimTextBuffer

        vimTextBuffer

    /// Create an IVimBuffer for the given ITextView and associated IVimTextBuffer
    member x.CreateVimBuffer textView (windowSettings : IVimWindowSettings option) =
        if _bufferMap.ContainsKey(textView) then 
            invalidArg "textView" Resources.Vim_TextViewAlreadyHasVimBuffer

        let vimTextBuffer = x.GetOrCreateVimTextBuffer textView.TextBuffer
        let buffer = _bufferFactoryService.CreateVimBuffer textView vimTextBuffer

        // Apply the specified window settings
        match windowSettings with
        | None -> 
            ()
        | Some windowSettings ->
            windowSettings.AllSettings
            |> Seq.filter (fun s -> not s.IsGlobal && not s.IsValueCalculated)
            |> Seq.iter (fun s -> buffer.WindowSettings.TrySetValue s.Name s.Value |> ignore)

        // Setup the handlers for KeyInputStart and KeyInputEnd to accurately track the active
        // IVimBuffer instance
        let eventBag = DisposableBag()
        buffer.KeyInputStart
        |> Observable.subscribe (fun _ -> _activeBufferStack <- buffer :: _activeBufferStack )
        |> eventBag.Add

        buffer.KeyInputEnd 
        |> Observable.subscribe (fun _ -> 
            _activeBufferStack <- 
                match _activeBufferStack with
                | h::t -> t
                | [] -> [] )
        |> eventBag.Add

        _bufferMap.Add(textView, (buffer,eventBag))
        _bufferCreationListeners |> Seq.iter (fun x -> x.Value.VimBufferCreated buffer)
        buffer

    member x.RemoveBuffer view = 
        let found, tuple = _bufferMap.TryGetValue(view)
        if found then 
            let _,bag = tuple
            bag.DisposeAll()
        _bufferMap.Remove(view)

    member x.GetVimTextBuffer (textBuffer : ITextBuffer) =
        let found, vimTextBuffer = textBuffer.Properties.TryGetProperty<IVimTextBuffer>(_vimTextBufferKey)
        if found then
            Some vimTextBuffer
        else
            None

    member x.GetVimBuffer textView =
        let tuple = _bufferMap.TryGetValue textView
        match tuple with 
        | (true,(buffer,_)) -> Some buffer
        | (false,_) -> None

    member x.GetOrCreateVimTextBuffer textBuffer =
        match x.GetVimTextBuffer textBuffer with
        | Some vimTextBuffer ->
            vimTextBuffer
        | None ->
            let settings = x.GetLocalSettingsForNewTextBuffer()
            x.CreateVimTextBuffer textBuffer settings

    member x.GetOrCreateVimBuffer textView =
        match x.GetVimBuffer textView with
        | Some buffer -> 
            buffer
        | None -> 
            let settings = x.GetWindowSettingsForNewBuffer()
            x.CreateVimBuffer textView settings

    member x.LoadVimRc (createViewFunc : (unit -> ITextView)) =
        _globalSettings.VimRc <- System.String.Empty
        _globalSettings.VimRcPaths <- _fileSystem.GetVimRcDirectories() |> String.concat ";"

        match _fileSystem.LoadVimRc() with
        | None -> false
        | Some(path,lines) ->
            _globalSettings.VimRc <- path
            let view = createViewFunc()
            let buffer = x.GetOrCreateVimBuffer view
            let mode = buffer.CommandMode
            lines |> Seq.iter (fun input -> mode.RunCommand input |> ignore)
            _vimRcLocalSettings <- LocalSettings.Copy buffer.LocalSettings
            _vimRcWindowSettings <- Some (WindowSettings.Copy buffer.WindowSettings)
            view.Close()
            true

    interface IVim with
        member x.ActiveBuffer = x.ActiveBuffer
        member x.Buffers = x.Buffers
        member x.FocusedBuffer = x.FocusedBuffer
        member x.VimData = _vimData
        member x.VimHost = _host
        member x.VimRcLocalSettings 
            with get() = _vimRcLocalSettings
            and set value = _vimRcLocalSettings <- LocalSettings.Copy value
        member x.MacroRecorder = _recorder :> IMacroRecorder
        member x.MarkMap = _markMap
        member x.KeyMap = _keyMap
        member x.SearchService = _search
        member x.IsVimRcLoaded = not (System.String.IsNullOrEmpty(_globalSettings.VimRc))
        member x.RegisterMap = _registerMap 
        member x.GlobalSettings = _globalSettings
        member x.CreateVimBuffer textView = x.CreateVimBuffer textView (x.GetWindowSettingsForNewBuffer())
        member x.CreateVimTextBuffer textBuffer = x.CreateVimTextBuffer textBuffer (x.GetLocalSettingsForNewTextBuffer())
        member x.GetOrCreateVimBuffer textView = x.GetOrCreateVimBuffer textView
        member x.GetOrCreateVimTextBuffer textBuffer = x.GetOrCreateVimTextBuffer textBuffer
        member x.RemoveBuffer textView = x.RemoveBuffer textView
        member x.GetVimBuffer textView = x.GetVimBuffer textView
        member x.GetVimTextBuffer textBuffer = x.GetVimTextBuffer textBuffer
        member x.LoadVimRc createViewFunc = x.LoadVimRc createViewFunc

