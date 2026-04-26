# ============================================================
#  DetecteurDoublons.ps1  v2.0
#  Détecteur de doublons VIDÉO — Interface WPF asynchrone
#
#  AMÉLIORATIONS vs v1 :
#  ─ Bug Levenshtein corrigé (rolling-array, plus de tableau 2D)
#  ─ Filtrage strict aux extensions vidéo connues
#  ─ Normalisation avancée : résolution, codec, source, année,
#    épisodes, groupes de release, HDR, audio…
#  ─ Bucketing par préfixe normalisé → O(n·k) au lieu de O(n²)
#  ─ Passe 1 : doublons exacts par taille de fichier
#  ─ Passe 2 : fuzzy par nom dans chaque bucket
#  ─ Scan en arrière-plan (Runspace) — UI réactive en permanence
#  ─ Barre de progression + % + ETA mis à jour en temps réel
#  ─ Journal d'erreurs live (accès refusés, exceptions PS…)
#  ─ Bouton Annuler fonctionnel
# ============================================================
#Requires -Version 5.1

# ── Masquer la fenêtre console PowerShell ────────────────
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
'@
[Native.Win32]::ShowWindow([Native.Win32]::GetConsoleWindow(), 0) | Out-Null

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ── Dossier racine du scan ────────────────────────────────
$script:scriptRoot = $PSScriptRoot
if (-not $script:scriptRoot) { $script:scriptRoot = (Get-Location).Path }

# ══════════════════════════════════════════════════════════
#  INTERNATIONALISATION  (détection automatique langue OS)
# ══════════════════════════════════════════════════════════
$script:UILang = try {
    $ci = [System.Globalization.CultureInfo]::CurrentUICulture
    $ci.TwoLetterISOLanguageName.ToLower()
} catch { 'en' }

$script:T = switch ($script:UILang) {

    'fr' { @{
        WindowTitle          = 'Détecteur de Doublons'
        AppTitle             = 'DÉTECTEUR DE DOUBLONS'
        Folder               = 'Dossier'
        BtnScan              = '⟳  SCANNER'
        BtnScanning          = '⟳  EN COURS...'
        BtnCancel            = '✕  ANNULER'
        BtnDelete            = '🗑  SUPPRIMER LA SÉLECTION'
        BtnSelectAuto        = 'Sélection auto (copies)'
        BtnUncheckAll        = 'Tout décocher'
        BtnGetMediaInfo      = '⬇  Télécharger MediaInfo CLI'
        ModeLabelPrefix      = 'MODE :'
        ModeVideo            = 'Tout type de vidéos'
        ModeFilms            = 'Films'
        ModeSeries           = 'Séries'
        ModeFichiers         = 'Tout type de fichiers'
        ModeMusique          = 'Musiques'
        StatusReady          = 'Prêt — cliquez sur SCANNER.'
        StatusInit           = 'Initialisation...'
        StatusCancelling     = 'Annulation en cours...'
        StatusScanning       = 'Scan en cours...'
        StatusFilesIndexed   = 'fichiers indexés'
        StatusDone           = 'Scan terminé —'
        StatusFilesAnalyzed  = 'fichiers analysés —'
        StatusGroupsFound    = 'groupe(s) trouvé(s). Cochez les fichiers à supprimer.'
        StatusNoDupe         = '✓ Aucun doublon détecté parmi'
        StatusNoDupeSuffix   = 'fichiers.'
        StatusCancelled      = '⊘ Scan annulé après'
        StatusCancelledSuffix= 'fichiers indexés.'
        StatsGroups          = 'groupe(s) de doublons'
        StatsFiles           = 'fichiers impliqués'
        StatsWaste           = 'Gaspillage potentiel :'
        BadgeSimilar         = 'Similaire'
        BadgeWaste           = 'Gaspillage :'
        BadgeFiles           = 'fichiers'
        ReasonExact          = 'Copie exacte'
        ReasonFilm           = 'Même film (qualité/codec différent ?)'
        ReasonSerie          = 'Même épisode (qualité/codec différent ?)'
        ReasonFichier        = 'Nom similaire (version différente ?)'
        ReasonMusique        = 'Même morceau (metadata)'
        ReasonVideoTitle     = 'Même titre (qualité/codec différent ?)'
        LogTitle             = '▸ JOURNAL / ERREURS'
        AccessDenied         = 'Accès refusé :'
        CriticalError        = 'ERREUR CRITIQUE :'
        FileInUse            = "Fichier en cours d'utilisation"
        AccessDeniedMsg      = 'Accès refusé'
        FileNotFound         = 'Introuvable'
        ConfirmTitle         = '⚠  Supprimer'
        ConfirmFileSuffix    = 'fichier(s) ?'
        ConfirmSub           = 'Total :'
        ConfirmIrreversible  = '— Action irréversible.'
        ConfirmCancel        = 'Annuler'
        ConfirmDelete        = 'SUPPRIMER DÉFINITIVEMENT'
        ReportTitle          = 'RAPPORT DE SUPPRESSION'
        ReportSummary        = 'supprimé(s)  •'
        ReportIgnored        = 'ignoré(s)'
        ReportStats          = 'fichier(s) supprimé(s)  •'
        ReportFailed         = 'échec(s)'
        SelectedFiles        = 'fichier(s) sélectionné(s) —'
        Go                   = 'Go'
        Mo                   = 'Mo'
        Ko                   = 'Ko'
        OpenFolder           = 'Ouvrir le dossier'
        OpenFile             = 'Ouvrir :'
        EtaLabel             = 'ETA :'
        EnumFiles            = 'Énumération des fichiers...'
        FilesFound           = 'fichiers trouvés...'
        Normalizing          = 'Normalisation'
        ExactCopies          = 'Recherche des copies exactes...'
        CopiesOf             = 'Copies exactes :'
        CopiesX              = 'fichiers ×'
        AnalyzeTitles        = 'Analyse titres'
        ScanDone             = 'Scan terminé.'
        FolderBrowserDesc    = 'Choisissez le dossier à scanner'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) introuvable — requis pour la détection par métadonnées (Vidéos, Musiques).'
        MediaInfoCLINote     = '⚑  Télécharger le CLI (pas le GUI) sur mediaarea.net'
        MediaInfoPaths       = 'Chemins recherchés :'
        ChooseFolderTip      = 'Choisir un autre dossier de scan'
    }}

    'de' { @{
        WindowTitle          = 'Duplikat-Erkennung'
        AppTitle             = 'DUPLIKAT-ERKENNUNG'
        Folder               = 'Ordner'
        BtnScan              = '⟳  SCANNEN'
        BtnScanning          = '⟳  LÄUFT...'
        BtnCancel            = '✕  ABBRECHEN'
        BtnDelete            = '🗑  AUSWAHL LÖSCHEN'
        BtnSelectAuto        = 'Auto-Auswahl (Kopien)'
        BtnUncheckAll        = 'Alles abwählen'
        BtnGetMediaInfo      = '⬇  MediaInfo CLI herunterladen'
        ModeLabelPrefix      = 'MODUS:'
        ModeVideo            = 'Alle Videodateien'
        ModeFilms            = 'Filme'
        ModeSeries           = 'Serien'
        ModeFichiers         = 'Alle Dateitypen'
        ModeMusique          = 'Musik'
        StatusReady          = 'Bereit — SCANNEN klicken.'
        StatusInit           = 'Initialisierung...'
        StatusCancelling     = 'Wird abgebrochen...'
        StatusScanning       = 'Scan läuft...'
        StatusFilesIndexed   = 'Dateien indiziert'
        StatusDone           = 'Scan abgeschlossen —'
        StatusFilesAnalyzed  = 'Dateien analysiert —'
        StatusGroupsFound    = 'Gruppe(n) gefunden. Wählen Sie zu löschende Dateien.'
        StatusNoDupe         = '✓ Keine Duplikate gefunden unter'
        StatusNoDupeSuffix   = 'Dateien.'
        StatusCancelled      = '⊘ Scan abgebrochen nach'
        StatusCancelledSuffix= 'indizierten Dateien.'
        StatsGroups          = 'Duplikatgruppe(n)'
        StatsFiles           = 'betroffene Dateien'
        StatsWaste           = 'Potenzieller Speicherverlust:'
        BadgeSimilar         = 'Ähnlich'
        BadgeWaste           = 'Verlust:'
        BadgeFiles           = 'Dateien'
        ReasonExact          = 'Exakte Kopie'
        ReasonFilm           = 'Gleicher Film (anderer Codec?)'
        ReasonSerie          = 'Gleiche Episode (anderer Codec?)'
        ReasonFichier        = 'Ähnlicher Name (andere Version?)'
        ReasonMusique        = 'Gleicher Track (Metadaten)'
        ReasonVideoTitle     = 'Gleicher Titel (anderer Codec?)'
        LogTitle             = '▸ PROTOKOLL / FEHLER'
        AccessDenied         = 'Zugriff verweigert:'
        CriticalError        = 'KRITISCHER FEHLER:'
        FileInUse            = 'Datei wird verwendet'
        AccessDeniedMsg      = 'Zugriff verweigert'
        FileNotFound         = 'Nicht gefunden'
        ConfirmTitle         = '⚠  Löschen von'
        ConfirmFileSuffix    = 'Datei(en)?'
        ConfirmSub           = 'Gesamt:'
        ConfirmIrreversible  = '— Nicht rückgängig zu machen.'
        ConfirmCancel        = 'Abbrechen'
        ConfirmDelete        = 'ENDGÜLTIG LÖSCHEN'
        ReportTitle          = 'LÖSCHBERICHT'
        ReportSummary        = 'gelöscht  •'
        ReportIgnored        = 'übersprungen'
        ReportStats          = 'Datei(en) gelöscht  •'
        ReportFailed         = 'Fehler'
        SelectedFiles        = 'Datei(en) ausgewählt —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Ordner öffnen'
        OpenFile             = 'Öffnen:'
        EtaLabel             = 'Rest:'
        EnumFiles            = 'Dateien aufzählen...'
        FilesFound           = 'Dateien gefunden...'
        Normalizing          = 'Normalisierung'
        ExactCopies          = 'Suche exakter Kopien...'
        CopiesOf             = 'Exakte Kopien:'
        CopiesX              = 'Dateien ×'
        AnalyzeTitles        = 'Titel analysieren'
        ScanDone             = 'Scan abgeschlossen.'
        FolderBrowserDesc    = 'Zu scannenden Ordner wählen'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) nicht gefunden — erforderlich für Metadaten-Erkennung (Video, Musik).'
        MediaInfoCLINote     = '⚑  CLI herunterladen (nicht GUI) auf mediaarea.net'
        MediaInfoPaths       = 'Gesuchte Pfade:'
        ChooseFolderTip      = 'Anderen Scan-Ordner wählen'
    }}

    'es' { @{
        WindowTitle          = 'Detector de Duplicados'
        AppTitle             = 'DETECTOR DE DUPLICADOS'
        Folder               = 'Carpeta'
        BtnScan              = '⟳  ESCANEAR'
        BtnScanning          = '⟳  ESCANEANDO...'
        BtnCancel            = '✕  CANCELAR'
        BtnDelete            = '🗑  ELIMINAR SELECCIÓN'
        BtnSelectAuto        = 'Selección auto (copias)'
        BtnUncheckAll        = 'Desmarcar todo'
        BtnGetMediaInfo      = '⬇  Descargar MediaInfo CLI'
        ModeLabelPrefix      = 'MODO:'
        ModeVideo            = 'Todo tipo de vídeos'
        ModeFilms            = 'Películas'
        ModeSeries           = 'Series'
        ModeFichiers         = 'Todo tipo de archivos'
        ModeMusique          = 'Música'
        StatusReady          = 'Listo — haga clic en ESCANEAR.'
        StatusInit           = 'Inicializando...'
        StatusCancelling     = 'Cancelando...'
        StatusScanning       = 'Escaneando...'
        StatusFilesIndexed   = 'archivos indexados'
        StatusDone           = 'Escaneo completado —'
        StatusFilesAnalyzed  = 'archivos analizados —'
        StatusGroupsFound    = 'grupo(s) encontrado(s). Marque los archivos a eliminar.'
        StatusNoDupe         = '✓ Sin duplicados detectados entre'
        StatusNoDupeSuffix   = 'archivos.'
        StatusCancelled      = '⊘ Escaneo cancelado tras'
        StatusCancelledSuffix= 'archivos indexados.'
        StatsGroups          = 'grupo(s) de duplicados'
        StatsFiles           = 'archivos implicados'
        StatsWaste           = 'Desperdicio potencial:'
        BadgeSimilar         = 'Similar'
        BadgeWaste           = 'Desperdicio:'
        BadgeFiles           = 'archivos'
        ReasonExact          = 'Copia exacta'
        ReasonFilm           = 'Misma película (¿codec diferente?)'
        ReasonSerie          = 'Mismo episodio (¿codec diferente?)'
        ReasonFichier        = 'Nombre similar (¿versión diferente?)'
        ReasonMusique        = 'Misma pista (metadatos)'
        ReasonVideoTitle     = 'Mismo título (¿codec diferente?)'
        LogTitle             = '▸ REGISTRO / ERRORES'
        AccessDenied         = 'Acceso denegado:'
        CriticalError        = 'ERROR CRÍTICO:'
        FileInUse            = 'Archivo en uso'
        AccessDeniedMsg      = 'Acceso denegado'
        FileNotFound         = 'No encontrado'
        ConfirmTitle         = '⚠  Eliminar'
        ConfirmFileSuffix    = 'archivo(s)?'
        ConfirmSub           = 'Total:'
        ConfirmIrreversible  = '— Acción irreversible.'
        ConfirmCancel        = 'Cancelar'
        ConfirmDelete        = 'ELIMINAR DEFINITIVAMENTE'
        ReportTitle          = 'INFORME DE ELIMINACIÓN'
        ReportSummary        = 'eliminado(s)  •'
        ReportIgnored        = 'omitido(s)'
        ReportStats          = 'archivo(s) eliminado(s)  •'
        ReportFailed         = 'error(es)'
        SelectedFiles        = 'archivo(s) seleccionado(s) —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Abrir carpeta'
        OpenFile             = 'Abrir:'
        EtaLabel             = 'ETA:'
        EnumFiles            = 'Enumerando archivos...'
        FilesFound           = 'archivos encontrados...'
        Normalizing          = 'Normalizando'
        ExactCopies          = 'Buscando copias exactas...'
        CopiesOf             = 'Copias exactas:'
        CopiesX              = 'archivos ×'
        AnalyzeTitles        = 'Analizando títulos'
        ScanDone             = 'Escaneo completado.'
        FolderBrowserDesc    = 'Elija la carpeta a escanear'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) no encontrado — necesario para detección por metadatos (Vídeo, Música).'
        MediaInfoCLINote     = '⚑  Descargar el CLI (no el GUI) en mediaarea.net'
        MediaInfoPaths       = 'Rutas buscadas:'
        ChooseFolderTip      = 'Elegir otra carpeta de escaneo'
    }}

    'it' { @{
        WindowTitle          = 'Rilevatore di Duplicati'
        AppTitle             = 'RILEVATORE DI DUPLICATI'
        Folder               = 'Cartella'
        BtnScan              = '⟳  SCANSIONA'
        BtnScanning          = '⟳  IN CORSO...'
        BtnCancel            = '✕  ANNULLA'
        BtnDelete            = '🗑  ELIMINA SELEZIONE'
        BtnSelectAuto        = 'Selezione auto (copie)'
        BtnUncheckAll        = 'Deseleziona tutto'
        BtnGetMediaInfo      = '⬇  Scarica MediaInfo CLI'
        ModeLabelPrefix      = 'MODALITÀ:'
        ModeVideo            = 'Tutti i video'
        ModeFilms            = 'Film'
        ModeSeries           = 'Serie'
        ModeFichiers         = 'Tutti i file'
        ModeMusique          = 'Musica'
        StatusReady          = 'Pronto — fare clic su SCANSIONA.'
        StatusInit           = 'Inizializzazione...'
        StatusCancelling     = 'Annullamento...'
        StatusScanning       = 'Scansione in corso...'
        StatusFilesIndexed   = 'file indicizzati'
        StatusDone           = 'Scansione completata —'
        StatusFilesAnalyzed  = 'file analizzati —'
        StatusGroupsFound    = 'gruppo/i trovato/i. Selezionare i file da eliminare.'
        StatusNoDupe         = '✓ Nessun duplicato tra'
        StatusNoDupeSuffix   = 'file.'
        StatusCancelled      = '⊘ Scansione annullata dopo'
        StatusCancelledSuffix= 'file indicizzati.'
        StatsGroups          = 'gruppo/i di duplicati'
        StatsFiles           = 'file coinvolti'
        StatsWaste           = 'Spreco potenziale:'
        BadgeSimilar         = 'Simile'
        BadgeWaste           = 'Spreco:'
        BadgeFiles           = 'file'
        ReasonExact          = 'Copia esatta'
        ReasonFilm           = 'Stesso film (codec diverso?)'
        ReasonSerie          = 'Stesso episodio (codec diverso?)'
        ReasonFichier        = 'Nome simile (versione diversa?)'
        ReasonMusique        = 'Stesso brano (metadati)'
        ReasonVideoTitle     = 'Stesso titolo (codec diverso?)'
        LogTitle             = '▸ LOG / ERRORI'
        AccessDenied         = 'Accesso negato:'
        CriticalError        = 'ERRORE CRITICO:'
        FileInUse            = 'File in uso'
        AccessDeniedMsg      = 'Accesso negato'
        FileNotFound         = 'Non trovato'
        ConfirmTitle         = '⚠  Eliminare'
        ConfirmFileSuffix    = 'file?'
        ConfirmSub           = 'Totale:'
        ConfirmIrreversible  = '— Azione irreversibile.'
        ConfirmCancel        = 'Annulla'
        ConfirmDelete        = 'ELIMINA DEFINITIVAMENTE'
        ReportTitle          = 'REPORT ELIMINAZIONE'
        ReportSummary        = 'eliminato/i  •'
        ReportIgnored        = 'ignorato/i'
        ReportStats          = 'file eliminato/i  •'
        ReportFailed         = 'errore/i'
        SelectedFiles        = 'file selezionato/i —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Apri cartella'
        OpenFile             = 'Apri:'
        EtaLabel             = 'ETA:'
        EnumFiles            = 'Ricerca file...'
        FilesFound           = 'file trovati...'
        Normalizing          = 'Normalizzazione'
        ExactCopies          = 'Ricerca copie esatte...'
        CopiesOf             = 'Copie esatte:'
        CopiesX              = 'file ×'
        AnalyzeTitles        = 'Analisi titoli'
        ScanDone             = 'Scansione completata.'
        FolderBrowserDesc    = 'Scegliere la cartella da scansionare'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) non trovato — necessario per il rilevamento tramite metadati (Video, Musica).'
        MediaInfoCLINote     = '⚑  Scaricare il CLI (non la GUI) su mediaarea.net'
        MediaInfoPaths       = 'Percorsi cercati:'
        ChooseFolderTip      = 'Scegliere un''altra cartella'
    }}

    'pt' { @{
        WindowTitle          = 'Detetor de Duplicados'
        AppTitle             = 'DETETOR DE DUPLICADOS'
        Folder               = 'Pasta'
        BtnScan              = '⟳  ANALISAR'
        BtnScanning          = '⟳  A ANALISAR...'
        BtnCancel            = '✕  CANCELAR'
        BtnDelete            = '🗑  ELIMINAR SELEÇÃO'
        BtnSelectAuto        = 'Seleção auto (cópias)'
        BtnUncheckAll        = 'Desmarcar tudo'
        BtnGetMediaInfo      = '⬇  Descarregar MediaInfo CLI'
        ModeLabelPrefix      = 'MODO:'
        ModeVideo            = 'Todo o tipo de vídeos'
        ModeFilms            = 'Filmes'
        ModeSeries           = 'Séries'
        ModeFichiers         = 'Todo o tipo de ficheiros'
        ModeMusique          = 'Música'
        StatusReady          = 'Pronto — clique em ANALISAR.'
        StatusInit           = 'A inicializar...'
        StatusCancelling     = 'A cancelar...'
        StatusScanning       = 'Análise em curso...'
        StatusFilesIndexed   = 'ficheiros indexados'
        StatusDone           = 'Análise concluída —'
        StatusFilesAnalyzed  = 'ficheiros analisados —'
        StatusGroupsFound    = 'grupo(s) encontrado(s). Marque os ficheiros a eliminar.'
        StatusNoDupe         = '✓ Nenhum duplicado detetado entre'
        StatusNoDupeSuffix   = 'ficheiros.'
        StatusCancelled      = '⊘ Análise cancelada após'
        StatusCancelledSuffix= 'ficheiros indexados.'
        StatsGroups          = 'grupo(s) de duplicados'
        StatsFiles           = 'ficheiros envolvidos'
        StatsWaste           = 'Desperdício potencial:'
        BadgeSimilar         = 'Similar'
        BadgeWaste           = 'Desperdício:'
        BadgeFiles           = 'ficheiros'
        ReasonExact          = 'Cópia exata'
        ReasonFilm           = 'Mesmo filme (codec diferente?)'
        ReasonSerie          = 'Mesmo episódio (codec diferente?)'
        ReasonFichier        = 'Nome similar (versão diferente?)'
        ReasonMusique        = 'Mesma faixa (metadados)'
        ReasonVideoTitle     = 'Mesmo título (codec diferente?)'
        LogTitle             = '▸ REGISTO / ERROS'
        AccessDenied         = 'Acesso negado:'
        CriticalError        = 'ERRO CRÍTICO:'
        FileInUse            = 'Ficheiro em utilização'
        AccessDeniedMsg      = 'Acesso negado'
        FileNotFound         = 'Não encontrado'
        ConfirmTitle         = '⚠  Eliminar'
        ConfirmFileSuffix    = 'ficheiro(s)?'
        ConfirmSub           = 'Total:'
        ConfirmIrreversible  = '— Ação irreversível.'
        ConfirmCancel        = 'Cancelar'
        ConfirmDelete        = 'ELIMINAR DEFINITIVAMENTE'
        ReportTitle          = 'RELATÓRIO DE ELIMINAÇÃO'
        ReportSummary        = 'eliminado(s)  •'
        ReportIgnored        = 'ignorado(s)'
        ReportStats          = 'ficheiro(s) eliminado(s)  •'
        ReportFailed         = 'erro(s)'
        SelectedFiles        = 'ficheiro(s) selecionado(s) —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Abrir pasta'
        OpenFile             = 'Abrir:'
        EtaLabel             = 'ETA:'
        EnumFiles            = 'A enumerar ficheiros...'
        FilesFound           = 'ficheiros encontrados...'
        Normalizing          = 'Normalização'
        ExactCopies          = 'Procura de cópias exatas...'
        CopiesOf             = 'Cópias exatas:'
        CopiesX              = 'ficheiros ×'
        AnalyzeTitles        = 'Análise de títulos'
        ScanDone             = 'Análise concluída.'
        FolderBrowserDesc    = 'Escolha a pasta a analisar'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) não encontrado — necessário para deteção por metadados (Vídeo, Música).'
        MediaInfoCLINote     = '⚑  Descarregar o CLI (não a GUI) em mediaarea.net'
        MediaInfoPaths       = 'Caminhos pesquisados:'
        ChooseFolderTip      = 'Escolher outra pasta'
    }}

    'nl' { @{
        WindowTitle          = 'Duplicaten Detector'
        AppTitle             = 'DUPLICATEN DETECTOR'
        Folder               = 'Map'
        BtnScan              = '⟳  SCANNEN'
        BtnScanning          = '⟳  BEZIG...'
        BtnCancel            = '✕  ANNULEREN'
        BtnDelete            = '🗑  SELECTIE VERWIJDEREN'
        BtnSelectAuto        = 'Auto-selectie (kopieën)'
        BtnUncheckAll        = 'Alles deselecteren'
        BtnGetMediaInfo      = '⬇  MediaInfo CLI downloaden'
        ModeLabelPrefix      = 'MODUS:'
        ModeVideo            = 'Alle videobestanden'
        ModeFilms            = 'Films'
        ModeSeries           = 'Series'
        ModeFichiers         = 'Alle bestandstypen'
        ModeMusique          = 'Muziek'
        StatusReady          = 'Klaar — klik op SCANNEN.'
        StatusInit           = 'Initialiseren...'
        StatusCancelling     = 'Annuleren...'
        StatusScanning       = 'Scannen...'
        StatusFilesIndexed   = 'bestanden geïndexeerd'
        StatusDone           = 'Scan voltooid —'
        StatusFilesAnalyzed  = 'bestanden geanalyseerd —'
        StatusGroupsFound    = 'groep(en) gevonden. Selecteer bestanden om te verwijderen.'
        StatusNoDupe         = '✓ Geen duplicaten gevonden onder'
        StatusNoDupeSuffix   = 'bestanden.'
        StatusCancelled      = '⊘ Scan afgebroken na'
        StatusCancelledSuffix= 'geïndexeerde bestanden.'
        StatsGroups          = 'duplicaatgroep(en)'
        StatsFiles           = 'betrokken bestanden'
        StatsWaste           = 'Potentieel verlies:'
        BadgeSimilar         = 'Vergelijkbaar'
        BadgeWaste           = 'Verlies:'
        BadgeFiles           = 'bestanden'
        ReasonExact          = 'Exacte kopie'
        ReasonFilm           = 'Zelfde film (andere codec?)'
        ReasonSerie          = 'Zelfde aflevering (andere codec?)'
        ReasonFichier        = 'Vergelijkbare naam (andere versie?)'
        ReasonMusique        = 'Zelfde track (metadata)'
        ReasonVideoTitle     = 'Zelfde titel (andere codec?)'
        LogTitle             = '▸ LOG / FOUTEN'
        AccessDenied         = 'Toegang geweigerd:'
        CriticalError        = 'KRITIEKE FOUT:'
        FileInUse            = 'Bestand in gebruik'
        AccessDeniedMsg      = 'Toegang geweigerd'
        FileNotFound         = 'Niet gevonden'
        ConfirmTitle         = '⚠  Verwijderen van'
        ConfirmFileSuffix    = 'bestand(en)?'
        ConfirmSub           = 'Totaal:'
        ConfirmIrreversible  = '— Onomkeerbare actie.'
        ConfirmCancel        = 'Annuleren'
        ConfirmDelete        = 'DEFINITIEF VERWIJDEREN'
        ReportTitle          = 'VERWIJDERRAPPORT'
        ReportSummary        = 'verwijderd  •'
        ReportIgnored        = 'overgeslagen'
        ReportStats          = 'bestand(en) verwijderd  •'
        ReportFailed         = 'fout(en)'
        SelectedFiles        = 'bestand(en) geselecteerd —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Map openen'
        OpenFile             = 'Openen:'
        EtaLabel             = 'ETA:'
        EnumFiles            = 'Bestanden opsommen...'
        FilesFound           = 'bestanden gevonden...'
        Normalizing          = 'Normalisering'
        ExactCopies          = 'Exacte kopieën zoeken...'
        CopiesOf             = 'Exacte kopieën:'
        CopiesX              = 'bestanden ×'
        AnalyzeTitles        = 'Titels analyseren'
        ScanDone             = 'Scan voltooid.'
        FolderBrowserDesc    = 'Kies de te scannen map'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) niet gevonden — vereist voor metadata-detectie (Video, Muziek).'
        MediaInfoCLINote     = '⚑  Download de CLI (niet de GUI) op mediaarea.net'
        MediaInfoPaths       = 'Gezochte paden:'
        ChooseFolderTip      = 'Andere scanmap kiezen'
    }}

    default { @{   # English fallback
        WindowTitle          = 'Duplicate Detector'
        AppTitle             = 'DUPLICATE DETECTOR'
        Folder               = 'Folder'
        BtnScan              = '⟳  SCAN'
        BtnScanning          = '⟳  SCANNING...'
        BtnCancel            = '✕  CANCEL'
        BtnDelete            = '🗑  DELETE SELECTION'
        BtnSelectAuto        = 'Auto-select (copies)'
        BtnUncheckAll        = 'Uncheck all'
        BtnGetMediaInfo      = '⬇  Download MediaInfo CLI'
        ModeLabelPrefix      = 'MODE:'
        ModeVideo            = 'All video types'
        ModeFilms            = 'Movies'
        ModeSeries           = 'TV Series'
        ModeFichiers         = 'All file types'
        ModeMusique          = 'Music'
        StatusReady          = 'Ready — click SCAN.'
        StatusInit           = 'Initializing...'
        StatusCancelling     = 'Cancelling...'
        StatusScanning       = 'Scanning...'
        StatusFilesIndexed   = 'files indexed'
        StatusDone           = 'Scan complete —'
        StatusFilesAnalyzed  = 'files analyzed —'
        StatusGroupsFound    = 'group(s) found. Check files to delete.'
        StatusNoDupe         = '✓ No duplicates found among'
        StatusNoDupeSuffix   = 'files.'
        StatusCancelled      = '⊘ Scan cancelled after'
        StatusCancelledSuffix= 'files indexed.'
        StatsGroups          = 'duplicate group(s)'
        StatsFiles           = 'files involved'
        StatsWaste           = 'Potential waste:'
        BadgeSimilar         = 'Similar'
        BadgeWaste           = 'Waste:'
        BadgeFiles           = 'files'
        ReasonExact          = 'Exact copy'
        ReasonFilm           = 'Same movie (different codec?)'
        ReasonSerie          = 'Same episode (different codec?)'
        ReasonFichier        = 'Similar name (different version?)'
        ReasonMusique        = 'Same track (metadata)'
        ReasonVideoTitle     = 'Same title (different codec?)'
        LogTitle             = '▸ LOG / ERRORS'
        AccessDenied         = 'Access denied:'
        CriticalError        = 'CRITICAL ERROR:'
        FileInUse            = 'File in use by another process'
        AccessDeniedMsg      = 'Access denied'
        FileNotFound         = 'Not found'
        ConfirmTitle         = '⚠  Delete'
        ConfirmFileSuffix    = 'file(s)?'
        ConfirmSub           = 'Total:'
        ConfirmIrreversible  = '— This action is irreversible.'
        ConfirmCancel        = 'Cancel'
        ConfirmDelete        = 'DELETE PERMANENTLY'
        ReportTitle          = 'DELETION REPORT'
        ReportSummary        = 'deleted  •'
        ReportIgnored        = 'skipped'
        ReportStats          = 'file(s) deleted  •'
        ReportFailed         = 'failure(s)'
        SelectedFiles        = 'file(s) selected —'
        Go                   = 'GB'
        Mo                   = 'MB'
        Ko                   = 'KB'
        OpenFolder           = 'Open folder'
        OpenFile             = 'Open:'
        EtaLabel             = 'ETA:'
        EnumFiles            = 'Enumerating files...'
        FilesFound           = 'files found...'
        Normalizing          = 'Normalizing'
        ExactCopies          = 'Searching exact copies...'
        CopiesOf             = 'Exact copies:'
        CopiesX              = 'files ×'
        AnalyzeTitles        = 'Analyzing titles'
        ScanDone             = 'Scan complete.'
        FolderBrowserDesc    = 'Choose the folder to scan'
        MediaInfoMissing     = 'MediaInfo CLI (MediaInfo.exe) not found — required for metadata detection (Video, Music).'
        MediaInfoCLINote     = '⚑  Download the CLI (not the GUI) at mediaarea.net'
        MediaInfoPaths       = 'Searched paths:'
        ChooseFolderTip      = 'Choose another scan folder'
    }}
}

# ── Extensions vidéo supportées ──────────────────────────
$script:VIDEO_EXT = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        '.mp4','.mkv','.avi','.wmv','.mov','.flv','.m4v','.ts',
        '.mpeg','.mpg','.m2ts','.webm','.3gp','.divx','.ogv',
        '.vob','.rm', '.rmvb','.asf', '.f4v','.mts', '.m2t',
        '.tp', '.trp','.bdmv','.xvid','.264', '.265', '.hevc',
        '.dav','.rec','.wtv', '.dvr-ms'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ── Extensions audio supportées ─────────────────────────
$script:MUSIC_EXT = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        '.mp3','.flac','.aac','.ogg','.opus','.wma','.m4a','.wav',
        '.aiff','.aif','.ape','.wv','.tta','.mka','.dsf','.dff',
        '.alac','.m4b','.mpc','.ofr','.spx','.ra','.mid','.midi',
		'.dts'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ── Chemin MediaInfo CLI ──────────────────────────────────
# MediaInfo.exe = version ligne de commande (silencieuse)
$script:MediaInfoPath = $null
$_miCmd     = Get-Command 'MediaInfo.exe' -ErrorAction SilentlyContinue
$_miCmdPath = if ($_miCmd) { $_miCmd.Source } else { $null }
$_miPaths   = @(
    "$env:ProgramFiles\MediaInfo\MediaInfo.exe",
    "${env:ProgramFiles(x86)}\MediaInfo\MediaInfo.exe",
    "$env:LOCALAPPDATA\Programs\MediaInfo\MediaInfo.exe",
    $_miCmdPath
)
# Mise à jour du 4e chemin affiché avec la valeur réelle résolue via PATH
$script:MediaInfoSearchPaths = @(
    "$env:ProgramFiles\MediaInfo\MediaInfo.exe",
    "${env:ProgramFiles(x86)}\MediaInfo\MediaInfo.exe",
    "$env:LOCALAPPDATA\Programs\MediaInfo\MediaInfo.exe",
    $(if ($_miCmdPath) { $_miCmdPath } else { '$env:PATH\MediaInfo.exe  (non trouvé)' })
)
foreach ($p in $_miPaths) {
    if ($p -and (Test-Path $p -ErrorAction SilentlyContinue)) {
        $script:MediaInfoPath = $p; break
    }
}
Remove-Variable _miCmd, _miCmdPath, _miPaths -ErrorAction SilentlyContinue

# ── État partagé thread-safe ──────────────────────────────
$script:State = [hashtable]::Synchronized(@{
    Phase      = 'idle'
    Progress   = 0
    StepMsg    = ''
    Errors     = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    Groups     = $null
    StartTime  = [datetime]::Now
    FilesTotal = 0
    Cancelled  = $false
    Mode       = 'video'   # video | films | series | fichiers | musique
})

# ══════════════════════════════════════════════════════════
#  XAML
# ══════════════════════════════════════════════════════════
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$($script:T.WindowTitle)"
    Height="860" Width="1140" MinHeight="540" MinWidth="780"
    WindowStartupLocation="CenterScreen"
    Background="#0D0D0F">

  <Window.Resources>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="6"/>
      <Setter Property="Background" Value="#1A1A1F"/>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E0E0E8"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <!-- ProgressBar épurée -->
    <Style TargetType="ProgressBar">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border Background="#1A1A26" CornerRadius="4"/>
              <Border x:Name="PART_Track"     Background="Transparent" CornerRadius="4"/>
              <Border x:Name="PART_Indicator" Background="#E63946"     CornerRadius="4"
                      HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"      Value="#E63946"/>
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="FontFamily"      Value="Consolas"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="Padding"         Value="22,10"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#FF4D5A"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Background" Value="#3A3A42"/>
                <Setter Property="Foreground" Value="#666"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnSecondary" TargetType="Button">
      <Setter Property="Background"      Value="#1E1E26"/>
      <Setter Property="Foreground"      Value="#A0A0B0"/>
      <Setter Property="FontFamily"      Value="Consolas"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Padding"         Value="18,10"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush"     Value="#2E2E3A"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#2A2A34"/>
                <Setter Property="Foreground" Value="#D0D0E0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnCancel" TargetType="Button">
      <Setter Property="Background"      Value="#1E0808"/>
      <Setter Property="Foreground"      Value="#FF7070"/>
      <Setter Property="FontFamily"      Value="Consolas"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="Padding"         Value="16,10"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush"     Value="#4A1818"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#2A1010"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Foreground" Value="#555"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>   <!-- 0 Header      -->
      <RowDefinition Height="Auto"/>   <!-- 1 Mode        -->
      <RowDefinition Height="Auto"/>   <!-- 2 MediaInfo   -->
      <RowDefinition Height="Auto"/>   <!-- 3 Statut      -->
      <RowDefinition Height="Auto"/>   <!-- 4 Progress    -->
      <RowDefinition Height="Auto"/>   <!-- 5 Log         -->
      <RowDefinition Height="*"    />  <!-- 6 Résultats   -->
      <RowDefinition Height="Auto"/>   <!-- 7 Stats       -->
      <RowDefinition Height="Auto"/>   <!-- 8 Delete      -->
    </Grid.RowDefinitions>

    <!-- ══ HEADER ══ -->
    <Border Grid.Row="0" Background="#0D0D0F" Padding="28,20,28,10">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="10"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="$($script:T.AppTitle)" FontFamily="Consolas"
                     FontSize="19" FontWeight="Bold" Foreground="#E63946"/>
          <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
            <TextBlock x:Name="txtRootPath" FontFamily="Consolas" FontSize="11"
                       Foreground="#445566" VerticalAlignment="Center"/>
            <Button x:Name="btnBrowse" Content="📁" ToolTip="$($script:T.ChooseFolderTip)"
                    Background="#1A1A26" Foreground="#7788AA" BorderThickness="1"
                    BorderBrush="#2A2A38" Cursor="Hand" FontSize="13"
                    Width="28" Height="22" Margin="8,0,0,0" Padding="0">
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border Background="{TemplateBinding Background}"
                          BorderBrush="{TemplateBinding BorderBrush}"
                          BorderThickness="{TemplateBinding BorderThickness}"
                          CornerRadius="3" Padding="2">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter Property="Background" Value="#252535"/>
                      <Setter Property="Foreground" Value="#AABBDD"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
            </Button>
          </StackPanel>
        </StackPanel>
        <Button x:Name="btnCancel" Grid.Column="1" Style="{StaticResource BtnCancel}"
                Content="$($script:T.BtnCancel)" VerticalAlignment="Center" Visibility="Collapsed"/>
        <Button x:Name="btnScan"   Grid.Column="3" Style="{StaticResource BtnPrimary}"
                Content="$($script:T.BtnScan)" VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <!-- ══ MODE DE DÉTECTION ══ -->
    <Border Grid.Row="1" Background="#0A0A12" Padding="28,10"
            BorderBrush="#1C1C2E" BorderThickness="0,0,0,1">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="$($script:T.ModeLabelPrefix)" FontFamily="Consolas" FontSize="11"
                   Foreground="#445566" VerticalAlignment="Center" Margin="0,0,14,0"/>
        <RadioButton x:Name="rbModeFilms"   Content="$($script:T.ModeFilms)"                GroupName="ScanMode"
                     Foreground="#AABBCC" FontFamily="Consolas" FontSize="12"
                     VerticalContentAlignment="Center" Margin="0,0,20,0"/>
        <RadioButton x:Name="rbModeSeries"  Content="$($script:T.ModeSeries)"               GroupName="ScanMode"
                     Foreground="#AABBCC" FontFamily="Consolas" FontSize="12"
                     VerticalContentAlignment="Center" Margin="0,0,20,0"/>
		<RadioButton x:Name="rbModeMusique" Content="$($script:T.ModeMusique)" GroupName="ScanMode"
                     Foreground="#AABBCC" FontFamily="Consolas" FontSize="12"
                     VerticalContentAlignment="Center" Margin="0,0,20,0"/>
		<RadioButton x:Name="rbModeVideo"    Content="$($script:T.ModeVideo)"  GroupName="ScanMode"
                     Foreground="#AABBCC" FontFamily="Consolas" FontSize="12"
                     VerticalContentAlignment="Center" Margin="0,0,20,0"/>
        <RadioButton x:Name="rbModeFichiers" Content="$($script:T.ModeFichiers)" GroupName="ScanMode"
                     Foreground="#AABBCC" FontFamily="Consolas" FontSize="12"
                     VerticalContentAlignment="Center" Margin="0,0,20,0"/>
              </StackPanel>
    </Border>

    <!-- ══ BANNIÈRE MEDIAINFO MANQUANT ══ -->
<Border x:Name="panelMediaInfo" Grid.Row="2" Visibility="Collapsed"
        Background="#1A0D00" Padding="28,10"
        BorderBrush="#4A2800" BorderThickness="0,1,0,1">

  <Grid>
    <!-- Colonnes -->
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <!-- Lignes -->
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Ligne 1 : avertissement principal -->
    <StackPanel Orientation="Horizontal" Grid.Row="0" VerticalAlignment="Center">
      <TextBlock Text="⚠  " FontFamily="Consolas" FontSize="12"
                 Foreground="#FF9900" VerticalAlignment="Center"/>
      <TextBlock Text="$($script:T.MediaInfoMissing)"
                 FontFamily="Consolas" FontSize="11" Foreground="#CC8833"
                 VerticalAlignment="Center"/>
    </StackPanel>

    <!-- Ligne 2 : rappel CLI vs GUI mis en évidence -->
    <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="18,2,0,2">
      <TextBlock Text="$($script:T.MediaInfoCLINote)"
                 FontFamily="Consolas" FontSize="11" FontWeight="Bold"
                 Foreground="#FF9900" VerticalAlignment="Center"/>
    </StackPanel>

    <!-- Ligne 3 : chemins recherchés (bruts) -->
    <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="18,2,0,0">
      <TextBlock Text="$($script:T.MediaInfoPaths) "
                 FontFamily="Consolas" FontSize="10"
                 Foreground="#7A5A30" VerticalAlignment="Center"/>
      <TextBlock Text="$($script:MediaInfoSearchPaths[0])  |  $($script:MediaInfoSearchPaths[1])  |  $($script:MediaInfoSearchPaths[2])  |  $($script:MediaInfoSearchPaths[3])"
                 FontFamily="Consolas" FontSize="10"
                 Foreground="#7A5A30"
                 VerticalAlignment="Center"
                 TextWrapping="NoWrap"/>
    </StackPanel>

    <!-- Bouton (sur 2 lignes uniquement) -->
    <Button x:Name="btnGetMediaInfo"
            Grid.Column="1"
            Grid.Row="0"
            Grid.RowSpan="2"
            VerticalAlignment="Center"
            Content="$($script:T.BtnGetMediaInfo)"
            Background="#2A1800"
            Foreground="#FF9900"
            FontFamily="Consolas"
            FontSize="11"
            FontWeight="Bold"
            BorderThickness="1"
            BorderBrush="#4A2800"
            Cursor="Hand"
            Padding="14,6">

      <Button.Template>
        <ControlTemplate TargetType="Button">
          <Border Background="{TemplateBinding Background}"
                  BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}"
                  CornerRadius="4"
                  Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center"
                              VerticalAlignment="Center"/>
          </Border>

          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter Property="Background" Value="#3A2200"/>
              <Setter Property="Foreground" Value="#FFBB44"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Button.Template>

    </Button>

  </Grid>
</Border>
    <!-- ══ BARRE D'ÉTAT ══ -->
    <Border Grid.Row="3" Background="#13131A" Padding="28,8">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="txtStatus" FontFamily="Consolas" FontSize="12"
                   Foreground="#5566AA" VerticalAlignment="Center"
                   Text="$($script:T.StatusReady)"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button x:Name="btnSelectAll"  Style="{StaticResource BtnSecondary}"
                  Content="$($script:T.BtnSelectAuto)" IsEnabled="False" Margin="0,0,10,0"/>
          <Button x:Name="btnSelectNone" Style="{StaticResource BtnSecondary}"
                  Content="$($script:T.BtnUncheckAll)" IsEnabled="False"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ══ PROGRESSION (masqué hors scan) ══ -->
    <Border Grid.Row="4" x:Name="panelProgress" Visibility="Collapsed"
            Background="#0F0F18" Padding="28,14,28,14"
            BorderBrush="#1C1C2E" BorderThickness="0,0,0,1">
      <StackPanel>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <ProgressBar x:Name="prgScan" Height="10" Value="0" Maximum="100"
                       VerticalAlignment="Center"/>
          <TextBlock x:Name="txtPct" Grid.Column="1" Margin="14,0,0,0"
                     FontFamily="Consolas" FontSize="13" FontWeight="Bold"
                     Foreground="#E63946" VerticalAlignment="Center" MinWidth="50"/>
          <TextBlock x:Name="txtETA" Grid.Column="2" Margin="20,0,0,0"
                     FontFamily="Consolas" FontSize="11" Foreground="#445577"
                     VerticalAlignment="Center"/>
        </Grid>
        <TextBlock x:Name="txtCurrentFile" Margin="0,9,0,0"
                   FontFamily="Consolas" FontSize="10" Foreground="#3A4A5A"
                   TextTrimming="CharacterEllipsis"/>
      </StackPanel>
    </Border>

    <!-- ══ LOG / ERREURS (masqué si vide) ══ -->
    <Border Grid.Row="5" x:Name="panelLog" Visibility="Collapsed"
            Background="#080810" MaxHeight="145"
            BorderBrush="#2A1A1A" BorderThickness="0,0,0,1">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Background="#100808" Padding="28,5,28,5">
          <TextBlock x:Name="txtLogTitle" Text="$($script:T.LogTitle)"
                     FontFamily="Consolas" FontSize="10"
                     Foreground="#AA4444" FontWeight="Bold"/>
        </Border>
        <ScrollViewer Grid.Row="1" x:Name="svLog"
                      VerticalScrollBarVisibility="Auto" Background="#080810">
          <TextBlock x:Name="tbLog" FontFamily="Consolas" FontSize="10"
                     Foreground="#CC5555" Padding="28,7,28,7" TextWrapping="Wrap"/>
        </ScrollViewer>
      </Grid>
    </Border>

    <!-- ══ RÉSULTATS ══ -->
    <ScrollViewer Grid.Row="6" x:Name="svResults"
                  VerticalScrollBarVisibility="Auto"
                  Background="#0D0D0F" Padding="14,8,14,8">
      <StackPanel x:Name="panelGroups" Margin="14,0"/>
    </ScrollViewer>

    <!-- ══ STATS ══ -->
    <Border Grid.Row="7" Background="#13131A" Padding="28,10">
      <TextBlock x:Name="txtStats" FontFamily="Consolas" FontSize="11"
                 Foreground="#445566"/>
    </Border>

    <!-- ══ DELETE FOOTER ══ -->
    <Border Grid.Row="8" Background="#0D0D0F" Padding="28,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="txtSelected" FontFamily="Consolas" FontSize="12"
                   Foreground="#E63946" VerticalAlignment="Center"/>
        <Button x:Name="btnDelete" Grid.Column="1" Style="{StaticResource BtnPrimary}"
                Content="$($script:T.BtnDelete)" IsEnabled="False"/>
      </Grid>
    </Border>
  </Grid>
</Window>
"@

# ── Chargement XAML ──────────────────────────────────────
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$btnScan        = $window.FindName('btnScan')
$btnCancel      = $window.FindName('btnCancel')
$btnDelete      = $window.FindName('btnDelete')
$btnSelectAll   = $window.FindName('btnSelectAll')
$btnSelectNone  = $window.FindName('btnSelectNone')
$panelGroups    = $window.FindName('panelGroups')
$txtStatus      = $window.FindName('txtStatus')
$txtStats       = $window.FindName('txtStats')
$txtSelected    = $window.FindName('txtSelected')
$txtRootPath    = $window.FindName('txtRootPath')
$panelProgress  = $window.FindName('panelProgress')
$prgScan        = $window.FindName('prgScan')
$txtPct         = $window.FindName('txtPct')
$txtCurrentFile = $window.FindName('txtCurrentFile')
$txtETA         = $window.FindName('txtETA')
$panelLog       = $window.FindName('panelLog')
$tbLog          = $window.FindName('tbLog')
$svLog          = $window.FindName('svLog')
$txtLogTitle    = $window.FindName('txtLogTitle')

$btnBrowse = $window.FindName('btnBrowse')

$txtRootPath.Text = "$($script:T.Folder) : $script:scriptRoot"

$rbModeVideo    = $window.FindName('rbModeVideo')
$rbModeFilms    = $window.FindName('rbModeFilms')
$rbModeSeries   = $window.FindName('rbModeSeries')
$rbModeFichiers = $window.FindName('rbModeFichiers')
$rbModeMusique  = $window.FindName('rbModeMusique')
$panelMediaInfo = $window.FindName('panelMediaInfo')
$btnGetMediaInfo= $window.FindName('btnGetMediaInfo')

$script:allCheckboxes = [System.Collections.Generic.List[object]]::new()
$script:fileMap       = @{}
$script:psInstance    = $null
$script:runspace      = $null
$script:asyncHandle   = $null
$script:timer         = $null
$script:logBuf        = [System.Text.StringBuilder]::new()
$script:errCount      = 0

# ── Bouton choisir dossier ────────────────────────────────
$btnBrowse.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description         = $script:T.FolderBrowserDesc
    $dlg.SelectedPath        = $script:scriptRoot
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:scriptRoot = $dlg.SelectedPath
        $txtRootPath.Text = "$($script:T.Folder) : $script:scriptRoot"
    }
})

# ── Bannière MediaInfo si absent ────────────────────────
if (-not $script:MediaInfoPath) {
    $panelMediaInfo.Visibility = 'Visible'
}
$btnGetMediaInfo.Add_Click({
    Start-Process "https://mediaarea.net/en/MediaInfo/Download/Windows"
})

# ── Afficher/masquer la bannière selon le mode sélectionné ──
$showMediaInfoBanner = {
    $needsMediaInfo = ($rbModeVideo.IsChecked -or $rbModeFilms.IsChecked -or
                       $rbModeSeries.IsChecked -or $rbModeMusique.IsChecked)
    if (-not $script:MediaInfoPath -and $needsMediaInfo) {
        $panelMediaInfo.Visibility = 'Visible'
    } else {
        $panelMediaInfo.Visibility = 'Collapsed'
    }
}
$rbModeVideo.Add_Checked($showMediaInfoBanner)
$rbModeFilms.Add_Checked($showMediaInfoBanner)
$rbModeSeries.Add_Checked($showMediaInfoBanner)
$rbModeMusique.Add_Checked($showMediaInfoBanner)
$rbModeFichiers.Add_Checked($showMediaInfoBanner)

# ══════════════════════════════════════════════════════════
#  HELPERS UI
# ══════════════════════════════════════════════════════════

function New-Badge {
    param([string]$text, [string]$bg = "#1E2A3A", [string]$fg = "#5588CC")
    $conv = [System.Windows.Media.BrushConverter]::new()
    $b  = [System.Windows.Controls.Border]::new()
    $b.Background        = $conv.ConvertFrom($bg)
    $b.CornerRadius      = [System.Windows.CornerRadius]::new(3)
    $b.Padding           = [System.Windows.Thickness]::new(7,2,7,2)
    $b.Margin            = [System.Windows.Thickness]::new(0,0,6,0)
    $b.VerticalAlignment = "Center"
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.Text       = $text
    $tb.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
    $tb.FontSize   = 10
    $tb.Foreground = $conv.ConvertFrom($fg)
    $b.Child = $tb
    return $b
}

function Update-SelectionCount {
    $checked = @($script:allCheckboxes | Where-Object { $_.IsChecked })
    $count   = $checked.Count
    $totalSz = ($checked | ForEach-Object {
        $f = $script:fileMap[$_.Name]; if ($f) { $f.Length } else { 0 }
    } | Measure-Object -Sum).Sum
    if ($count -gt 0) {
        $go = [Math]::Round($totalSz / 1GB, 3)
        $txtSelected.Text    = "$count $($script:T.SelectedFiles) $go $($script:T.Go)"
        $btnDelete.IsEnabled = $true
    } else {
        $txtSelected.Text    = ""
        $btnDelete.IsEnabled = $false
    }
}

function Build-UI {
    param($groups)
    $panelGroups.Children.Clear()
    $script:allCheckboxes.Clear()
    $script:fileMap.Clear()
    $conv = [System.Windows.Media.BrushConverter]::new()

    $totalFiles  = 0
    $totalGroups = 0
    $totalWaste  = [long]0

    foreach ($group in $groups) {
        $totalGroups++
        $files = @($group)

        $card = [System.Windows.Controls.Border]::new()
        $card.Background      = $conv.ConvertFrom("#13131A")
        $card.BorderBrush     = $conv.ConvertFrom("#1E1E2A")
        $card.BorderThickness = [System.Windows.Thickness]::new(1)
        $card.CornerRadius    = [System.Windows.CornerRadius]::new(6)
        $card.Margin          = [System.Windows.Thickness]::new(0,6,0,6)
        $card.Padding         = [System.Windows.Thickness]::new(16,12,16,12)

        $sp = [System.Windows.Controls.StackPanel]::new()

        # En-tête du groupe
        $header = [System.Windows.Controls.WrapPanel]::new()
        $header.Margin = [System.Windows.Thickness]::new(0,0,0,8)
        $reason = if ($files[0].PSObject.Properties["Reason"]) { $files[0].Reason } else { $script:T.BadgeSimilar }
        $header.Children.Add((New-Badge $reason "#1A2A1A" "#44CC88")) | Out-Null
        $header.Children.Add((New-Badge "$($files.Count) $($script:T.BadgeFiles)" "#1A1A2A" "#8888FF")) | Out-Null

        $minSize = ($files | Measure-Object -Property Length -Minimum).Minimum
        $waste   = ($files | Measure-Object -Property Length -Sum).Sum - $minSize
        $totalWaste += $waste
        $wStr = if ($waste -ge 1GB) { "$([Math]::Round($waste/1GB,2)) $($script:T.Go)" } else { "$([Math]::Round($waste/1MB,0)) $($script:T.Mo)" }
        $header.Children.Add((New-Badge "$($script:T.BadgeWaste) $wStr" "#2A1A1A" "#FF6655")) | Out-Null
        $sp.Children.Add($header) | Out-Null

        $sep = [System.Windows.Controls.Separator]::new()
        $sep.Background = $conv.ConvertFrom("#1E1E2A")
        $sep.Margin = [System.Windows.Thickness]::new(0,0,0,8)
        $sp.Children.Add($sep) | Out-Null

        foreach ($f in $files) {
            $totalFiles++
            $row = [System.Windows.Controls.Grid]::new()
            $c0 = [System.Windows.Controls.ColumnDefinition]::new()
            $c0.Width = [System.Windows.GridLength]::new(28)
            $c1 = [System.Windows.Controls.ColumnDefinition]::new()
            $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c2 = [System.Windows.Controls.ColumnDefinition]::new()
            $c2.Width = [System.Windows.GridLength]::new(110)
            $row.ColumnDefinitions.Add($c0)
            $row.ColumnDefinitions.Add($c1)
            $row.ColumnDefinitions.Add($c2)
            $row.Margin = [System.Windows.Thickness]::new(0,2,0,2)

            # Checkbox
            $cb  = [System.Windows.Controls.CheckBox]::new()
            $cbN = "cb_" + [System.Guid]::NewGuid().ToString("N")
            $cb.Name = $cbN; $cb.VerticalAlignment = "Center"
            $script:allCheckboxes.Add($cb)
            $script:fileMap[$cbN] = $f
            $cb.Add_Checked({   Update-SelectionCount })
            $cb.Add_Unchecked({ Update-SelectionCount })
            [System.Windows.Controls.Grid]::SetColumn($cb, 0)
            $row.Children.Add($cb) | Out-Null

            # Nom → exécute le fichier / Chemin → ouvre le dossier
            $vsp = [System.Windows.Controls.StackPanel]::new()

            # Lien nom → exécute le fichier
            $lnk = [System.Windows.Documents.Hyperlink]::new()
            $lnk.Foreground = $conv.ConvertFrom("#88AAEE")
            $lnk.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
            $lnk.FontSize   = 12
            $lnk.Inlines.Add($f.Name) | Out-Null
            $lnk.ToolTip    = "$($script:T.OpenFile) $($f.FullName)"
            $fp = $f.FullName
            $lnk.Add_Click({ Invoke-Item -LiteralPath $fp }.GetNewClosure())
            $tb1 = [System.Windows.Controls.TextBlock]::new()
            $tb1.Inlines.Add($lnk)

            # Lien chemin → ouvre le dossier dans l'explorateur
            $lnk2 = [System.Windows.Documents.Hyperlink]::new()
            $lnk2.Foreground = $conv.ConvertFrom("#7A8A9A")
            $lnk2.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
            $lnk2.FontSize   = 10
            $lnk2.Inlines.Add($f.DirectoryName) | Out-Null
            $lnk2.ToolTip    = $script:T.OpenFolder
            $dp = $f.DirectoryName
            $lnk2.Add_Click({ Invoke-Item -LiteralPath $dp }.GetNewClosure())
            $tb2 = [System.Windows.Controls.TextBlock]::new()
            $tb2.Inlines.Add($lnk2)
            $tb2.TextTrimming = "CharacterEllipsis"
            $vsp.Children.Add($tb1) | Out-Null
            $vsp.Children.Add($tb2) | Out-Null
            [System.Windows.Controls.Grid]::SetColumn($vsp, 1)
            $row.Children.Add($vsp) | Out-Null

            # Taille
            $sz = if ($f.Length -ge 1GB) { "$([Math]::Round($f.Length/1GB,2)) $($script:T.Go)" } `
                  elseif ($f.Length -ge 1MB) { "$([Math]::Round($f.Length/1MB,1)) $($script:T.Mo)" } `
                  else { "$([Math]::Round($f.Length/1KB,0)) $($script:T.Ko)" }
            $tbSz = [System.Windows.Controls.TextBlock]::new()
            $tbSz.Text                = $sz
            $tbSz.FontFamily          = [System.Windows.Media.FontFamily]::new("Consolas")
            $tbSz.FontSize            = 12
            $tbSz.Foreground          = $conv.ConvertFrom("#778899")
            $tbSz.HorizontalAlignment = "Right"
            $tbSz.VerticalAlignment   = "Center"
            [System.Windows.Controls.Grid]::SetColumn($tbSz, 2)
            $row.Children.Add($tbSz) | Out-Null

            $sp.Children.Add($row) | Out-Null
        }

        $card.Child = $sp
        $panelGroups.Children.Add($card) | Out-Null
    }

    $wGoTotal = if ($totalWaste -ge 1GB) { "$([Math]::Round($totalWaste/1GB,2)) $($script:T.Go)" } else { "$([Math]::Round($totalWaste/1MB,0)) $($script:T.Mo)" }
    $txtStats.Text = "$totalGroups $($script:T.StatsGroups)  •  $totalFiles $($script:T.StatsFiles)  •  $($script:T.StatsWaste) $wGoTotal"
    $btnSelectAll.IsEnabled  = ($totalGroups -gt 0)
    $btnSelectNone.IsEnabled = ($totalGroups -gt 0)
}

# ══════════════════════════════════════════════════════════
#  SCRIPTBLOCK DU RUNSPACE  (tout le code de scan ici)
# ══════════════════════════════════════════════════════════
$script:scanBlock = {
    param(
        [hashtable]$state,
        [string]$rootPath,
        [System.Collections.Generic.HashSet[string]]$videoExt,
        [System.Collections.Generic.HashSet[string]]$musicExt,
        [string]$mediaInfoPath,
        [hashtable]$T
    )

    $mode = $state.Mode   # video | films | series | fichiers | musique

    # ── Lecture metadata MediaInfo (retourne hashtable ou $null) ──────
    function Get-MediaInfoMeta {
        param([string]$fullPath)
        if (-not $mediaInfoPath) { return $null }
        try {
            $raw = & $mediaInfoPath --Output=CSV --Inform="General;%Title%|%Performer%|%Album%|%Track%|%Duration%|%BitRate%|%OverallBitRate%" `
                   $fullPath 2>$null
            if (-not $raw) { return $null }
            $p = $raw -split '\|'
            return @{
                Title    = if ($p[0]) { $p[0].Trim() } else { '' }
                Artist   = if ($p[1]) { $p[1].Trim() } else { '' }
                Album    = if ($p[2]) { $p[2].Trim() } else { '' }
                Track    = if ($p[3]) { $p[3].Trim() } else { '' }
                Duration = if ($p[4]) { [long]$p[4] }  else { 0 }
                BitRate  = if ($p[5]) { $p[5].Trim() } else { '' }
            }
        } catch { return $null }
    }

    # ── Clé de déduplication musique (metadata > nom) ─────────────────
    function Get-MusicKey {
        param([System.IO.FileInfo]$f)
        $meta = Get-MediaInfoMeta $f.FullName
        if ($meta -and ($meta.Title -or $meta.Artist)) {
            $title  = ($meta.Title  -replace '[^\w\s]', '' -replace '\s+', ' ').Trim().ToLower()
            $artist = ($meta.Artist -replace '[^\w\s]', '' -replace '\s+', ' ').Trim().ToLower()
            # Durée arrondie à la seconde (tolérance ±2 s gérée plus bas)
            $durSec = [int]($meta.Duration / 1000)
            if ($title)  { return "$artist|$title|$durSec" }
        }
        # Fallback : nom normalisé si pas de metadata
        $n = [System.IO.Path]::GetFileNameWithoutExtension($f.Name).ToLower()
        $n = $n -replace '\b(320|256|192|128|64)k?bps?\b', ''
        $n = $n -replace '\b(v0|v2|q[0-9])\b', ''
        $n = $n -replace '\b(mp3|flac|aac|ogg|opus|wma)\b', ''
        $n = $n -replace '[._\-\s]+', ' '
        return $n.Trim()
    }

    # ── Metadata vidéo via MediaInfo (durée + résolution) ─────────────
    function Get-VideoMeta {
        param([string]$fullPath)
        if (-not $mediaInfoPath) { return $null }
        try {
            $raw = & $mediaInfoPath --Output=CSV --Inform="General;%Duration%|%Width%|%Height%" `
                   $fullPath 2>$null
            if (-not $raw) { return $null }
            $p = $raw -split '\|'
            return @{
                Duration = if ($p[0]) { [long]$p[0] } else { 0 }
                Width    = if ($p[1]) { [int]$p[1]  } else { 0 }
                Height   = if ($p[2]) { [int]$p[2]  } else { 0 }
            }
        } catch { return $null }
    }

    # ── Normalisation commune (supprime codec/résolution/source) ────
    function Normalize-Base {
        param([string]$name)
        $n = [System.IO.Path]::GetFileNameWithoutExtension($name).ToLower()
        $n = $n -replace '\b(4k|2160p?|1080[pi]?|720p?|480p?|360p?|240p?|2k|8k|uhd|fhd|hd|sd)\b', ''
        $n = $n -replace '\b(x\.?264|x\.?265|h\.?264|h\.?265|xvid|divx|avc|hevc|av1|vp[89]|mpeg[-.]?[24]?|wmv[39]?|rv[34][05])\b', ''
        $n = $n -replace '\b(aac|mp3|ac3|e-?ac3|dts[-.]?(hd|ma|x)?|truehd|atmos|ddp?[257]?\.[01]|flac|opus|vorbis|pcm)\b', ''
        $n = $n -replace '\b(blu[-.]?ray|bdrip|bdremux|dvdrip|dvdscr|dvd|webrip|web[-.]?dl|webdl|hdtv|pdtv|dsr|hdrip|hdr10?[+]?|sdr|remux|uhdrip|amzn|nf|hulu|dsnp|atvp|pcok)\b', ''
        $n = $n -replace '\b(10[-.]?bit|8[-.]?bit|hdr|dovi|dolby[-.]?vision|hlg)\b', ''
        $n = $n -replace '\b(proper|repack|extended|theatrical|directors?[-.]?cut|unrated|dc|sample|trailer|bonus|featurette|extras?|complete|dubbed|subbed|multi|french|vf|vo|vostfr)\b', ''
        $n = $n -replace '\[.*?\]', ''
        $n = $n -replace '\(.*?\)', ''
        $n = $n -replace '\{.*?\}', ''
        $n = $n -replace '[._\-\s]+', ' '
        return $n.Trim()
    }

    # ── Normalisation vidéo générale (supprime aussi l'année et épisodes) ──
    function Normalize-VideoName {
        param([string]$name)
        $n = Normalize-Base $name
        $n = $n -replace '\b(19|20)\d{2}\b', ''
        $n = $n -replace '\bs\d{1,2}(e\d{1,2}){1,3}\b', ''
        $n = $n -replace '\bepisode?\s*\d+\b', ''
        $n = $n -replace '\bep\.?\s*\d+\b', ''
        $n = $n -replace '\bpart\.?\s*\d+\b', ''
        $n = $n -replace '[._\-\s]+', ' '
        return $n.Trim()
    }

    # ── Normalisation films (conserve l'année comme discriminant) ──
    function Normalize-Film {
        param([string]$name)
        $n = Normalize-Base $name
        # Extraire l'année si présente, la normaliser en suffixe fixe
        $year = ''
        if ($n -match '\b((19|20)\d{2})\b') { $year = $Matches[1] }
        $n = $n -replace '\b(19|20)\d{2}\b', ''
        $n = $n -replace '[._\-\s]+', ' '
        $n = $n.Trim()
        if ($year) { $n = "$n $year" }
        return $n
    }

    # ── Extraction du code episode canonique ────────────────
    # Reconnait : S01E02, 1x02, Episode 5, ep5, et numero absolu (Naruto 040)
    function Get-EpisodeKey {
        param([string]$name)
        $raw = [System.IO.Path]::GetFileNameWithoutExtension($name).ToLower()
        # Format SxxExx (prioritaire)
        if ($raw -match '(?<![a-z])s(\d{1,3})e(\d{1,3})(?:e\d{1,3})?(?![a-z\d])') {
            return "s$($Matches[1].PadLeft(2,'0'))e$($Matches[2].PadLeft(2,'0'))"
        }
        # Format NxNN (ex: 7x03, 1x02)
        if ($raw -match '(?<![a-z\d])(\d{1,2})x(\d{2,3})(?!\d)') {
            return "s$($Matches[1].PadLeft(2,'0'))e$($Matches[2].PadLeft(2,'0'))"
        }
        # Format "episode N" ou "ep N"
        if ($raw -match 'episode?\s*(\d+)') { return "ep$($Matches[1].PadLeft(3,'0'))" }
        if ($raw -match 'ep\.?\s*(\d+)')    { return "ep$($Matches[1].PadLeft(3,'0'))" }
        # Format numero absolu : un nombre de 2-4 chiffres isole dans le nom
        # ex: "Naruto 040 - titre", "Dragon Ball 123", "One Piece - 105 - titre"
        # On prend le PREMIER nombre trouve apres un separateur (espace, tiret, underscore)
        # en excluant les annees (1900-2099) et les tailles/resolutions connues
        if ($raw -match '(?:^|[\s._\-])(\d{2,4})(?:[\s._\-]|$)') {
            $num = $Matches[1]
            # Exclure annees et resolutions
            if ($num -notmatch '^(19|20)\d{2}$' -and
                $num -notmatch '^(480|576|720|1080|2160|4320)$') {
                return "ep$($num.PadLeft(3,'0'))"
            }
        }
        return ''
    }

    # ── Normalisation series (conserve saison+episode, supprime le reste) ──
    function Normalize-Serie {
        param([string]$name)
        # Extraire le code episode canonique sur le nom brut (avant Normalize-Base)
        $ep = Get-EpisodeKey $name

        $n = Normalize-Base $name
        # Supprimer tous les marqueurs d'episode (SxxExx, NxNN, episode N, ep N)
        $n = $n -replace 's\d{1,3}e\d{1,3}(?:e\d{1,3})?', ''
        $n = $n -replace '(?<![a-z\d])\d{1,2}x\d{2,3}(?!\d)', ''
        $n = $n -replace 'episode?\s*\d+', ''
        $n = $n -replace 'ep\.?\s*\d+', ''
        $n = $n -replace '(19|20)\d{2}', ''
        # Supprimer aussi les numeros absolus (2-4 chiffres isoles)
        $n = $n -replace '(?:^|[\s])(\d{2,4})(?:[\s]|$)', ' '
        $n = $n -replace '\s+', ' '
        $n = $n.Trim()
        # Reattacher le code episode canonique -> titre+episode = cle unique par episode
        if ($ep) { $n = "$n $ep" }
        return $n
    }

    # ──────────────────────────────────────────────────────
    try {
        $state.Phase    = 'scanning'
        $state.StepMsg  = $T.EnumFiles
        $state.Progress = 2

        # ════ PHASE 1 : Énumération (0→15%) ════
        $allFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        $enumErr  = @()

        Get-ChildItem -Path $rootPath -Recurse -File `
            -ErrorAction SilentlyContinue -ErrorVariable enumErr |
        ForEach-Object {
            if ($state.Cancelled) { return }
            # Ignorer les fichiers vides (0 octet) dans tous les modes
            if ($_.Length -eq 0) { return }
            $ext = $_.Extension.ToLower()
            if (
                ($mode -eq 'fichiers') -or
                ($mode -eq 'musique'  -and $musicExt.Contains($ext)) -or
                ($mode -ne 'fichiers' -and $mode -ne 'musique' -and $videoExt.Contains($ext))
            ) {
                $allFiles.Add($_)
            }
        }

        foreach ($e in $enumErr) { $state.Errors.Enqueue("$($T.AccessDenied) $($e.TargetObject)") }

        if ($state.Cancelled) { $state.Groups = [System.Collections.Generic.List[object]]::new(); $state.Phase = 'done'; return }

        $fileList = @($allFiles)
        $n        = $fileList.Count
        $state.FilesTotal = $n
        $state.Progress   = 15
        $state.StepMsg    = "$n $($T.FilesFound)"

        if ($n -eq 0) { $state.Groups = [System.Collections.Generic.List[object]]::new(); $state.Phase = 'done'; return }

        # ════ PHASE 2 : Normalisation + bucketing (15→30%) ════
        $norms   = [System.Collections.Generic.Dictionary[string,string]]::new($n)
        $buckets = [System.Collections.Generic.Dictionary[string,
                    System.Collections.Generic.List[object]]]::new()

        for ($i = 0; $i -lt $n; $i++) {
            if ($state.Cancelled) { break }
            $f = $fileList[$i]

            $norm = switch ($mode) {
                'films'    { Normalize-Film  $f.Name }
                'series'   { Normalize-Serie $f.Name }
                'fichiers' { [System.IO.Path]::GetFileNameWithoutExtension($f.Name).ToLower() -replace '[._\-\s]+', ' ' }
                'musique'  { Get-MusicKey $f }
                default    { Normalize-VideoName $f.Name }
            }
            $norms[$f.FullName] = $norm

            $words = ($norm -split ' ') | Where-Object { $_.Length -ge 3 }
            $key   = ($words | Select-Object -First 2) -join ' '
            if (-not $key -or $key.Length -lt 2) { $key = '~divers~' }

            if (-not $buckets.ContainsKey($key)) {
                $buckets[$key] = [System.Collections.Generic.List[object]]::new()
            }
            $buckets[$key].Add($f)

            $state.Progress = 15 + [int](($i / $n) * 15)
            if ($i % 50 -eq 0) { $state.StepMsg = "$($T.Normalizing) [$i/$n] : $($f.Name)" }
        }

        # ════ PHASE 3 : Copies exactes par taille (30→45%) ════
        $state.StepMsg  = $T.ExactCopies
        $state.Progress = 30

        $bySize  = [System.Collections.Generic.Dictionary[long, System.Collections.Generic.List[object]]]::new()
        foreach ($f in $fileList) {
            if ($f.Length -gt 0) {
                if (-not $bySize.ContainsKey($f.Length)) {
                    $bySize[$f.Length] = [System.Collections.Generic.List[object]]::new()
                }
                $bySize[$f.Length].Add($f)
            }
        }

        $groups  = [System.Collections.Generic.List[object]]::new()
        $used    = [System.Collections.Generic.HashSet[string]]::new()
        $sGroups = @($bySize.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
        $sg = $sGroups.Count; $sgD = 0

        foreach ($entry in $sGroups) {
            if ($state.Cancelled) { break }
            $sgD++
            $state.Progress = 30 + [int](($sgD / [Math]::Max($sg, 1)) * 15)
            $sameSize = @($entry.Value)
            $szStr = if ($entry.Key -ge 1GB) { "$([Math]::Round($entry.Key/1GB,2)) $($T.Go)" } `
                     else { "$([Math]::Round($entry.Key/1MB,1)) $($T.Mo)" }
            $state.StepMsg = "$($T.CopiesOf) $($sameSize.Count) $($T.CopiesX) $szStr"

            for ($i = 0; $i -lt $sameSize.Count; $i++) {
                $fi = $sameSize[$i]
                if ($used.Contains($fi.FullName)) { continue }
                $cluster = [System.Collections.Generic.List[object]]::new()
                $cluster.Add($fi)
                # Mode series : code episode de fi
                $epI = if ($mode -eq 'series') { Get-EpisodeKey $fi.Name } else { $null }
                # Mode video/films/series : durée MediaInfo de fi (si disponible)
                $isVideoMode = $mode -in @('video','films','series')
                $metaI = if ($isVideoMode -and $mediaInfoPath) { Get-VideoMeta $fi.FullName } else { $null }
                for ($j = $i+1; $j -lt $sameSize.Count; $j++) {
                    $fj = $sameSize[$j]
                    if ($used.Contains($fj.FullName)) { continue }
                    # Mode series : ignorer si les codes episode sont differents
                    if ($mode -eq 'series' -and $epI -ne '') {
                        $epJ = Get-EpisodeKey $fj.Name
                        if ($epJ -ne '' -and $epJ -ne $epI) { continue }
                    }
                    # Mode video : si MediaInfo dispo, vérifier durée (tolérance 1s)
                    # Deux fichiers identiques en taille mais de durées différentes = faux positif
                    if ($metaI -and $metaI.Duration -gt 0) {
                        $metaJ = Get-VideoMeta $fj.FullName
                        if ($metaJ -and $metaJ.Duration -gt 0) {
                            $durDiffMs = [Math]::Abs($metaI.Duration - $metaJ.Duration)
                            if ($durDiffMs -gt 1000) { continue }  # >1s d'écart = pas une copie
                        }
                    }
                    $r = "$($T.ReasonExact) ($szStr)"
                    $fj | Add-Member -NotePropertyName Reason -NotePropertyValue $r -Force
                    if ($cluster.Count -eq 1) { $fi | Add-Member -NotePropertyName Reason -NotePropertyValue $r -Force }
                    $cluster.Add($fj)
                    $used.Add($fj.FullName) | Out-Null
                }
                if ($cluster.Count -gt 1) { $used.Add($fi.FullName) | Out-Null; $groups.Add($cluster) }
            }
        }

        # ════ PHASE 4 : Détection par titre normalisé (45→100%) ════
            $bArr = @($buckets.GetEnumerator())
            $bT = $bArr.Count; $bD = 0

            foreach ($entry in $bArr) {
                if ($state.Cancelled) { break }
                $bD++
                $state.Progress = 45 + [int](($bD / [Math]::Max($bT, 1)) * 55)
                $bucket = @($entry.Value)
                if ($bucket.Count -lt 2) { continue }
                $state.StepMsg = "$($T.AnalyzeTitles) [$bD/$bT] '$($entry.Key)'"

                for ($i = 0; $i -lt $bucket.Count; $i++) {
                    $fi = $bucket[$i]
                    if ($used.Contains($fi.FullName)) { continue }
                    $normI = $norms[$fi.FullName]
                    if ($normI.Length -lt 3) { continue }

                    $cluster = [System.Collections.Generic.List[object]]::new()
                    $cluster.Add($fi)

                    for ($j = $i+1; $j -lt $bucket.Count; $j++) {
                        $fj = $bucket[$j]
                        if ($used.Contains($fj.FullName)) { continue }
                        $normJ = $norms[$fj.FullName]
                        if ($normJ.Length -lt 3) { continue }

                        # Logique de correspondance selon le mode
                        $match = $false
                        $matchDetail = ''

                        if ($mode -eq 'fichiers') {
                            # Préfixe commun (≥50% de mots consécutifs identiques)
                            $wordsI = $normI -split ' '
                            $wordsJ = $normJ -split ' '
                            if ($wordsI[0] -eq $wordsJ[0]) {
                                $common = 0
                                $minLen = [Math]::Min($wordsI.Count, $wordsJ.Count)
                                for ($k = 0; $k -lt $minLen; $k++) {
                                    if ($wordsI[$k] -eq $wordsJ[$k]) { $common++ } else { break }
                                }
                                $match = ($common / [Math]::Max($wordsI.Count, $wordsJ.Count)) -ge 0.5
                            }
                        } elseif ($mode -eq 'musique') {
                            # Clé musique : "artist|title|durSec"
                            # Correspondance exacte artiste+titre, tolérance ±3s sur durée
                            $partsI = $normI -split '\|'
                            $partsJ = $normJ -split '\|'
                            if ($partsI.Count -ge 3 -and $partsJ.Count -ge 3) {
                                $sameArtistTitle = ($partsI[0] -eq $partsJ[0]) -and ($partsI[1] -eq $partsJ[1])
                                $durDiff = [Math]::Abs([int]$partsI[2] - [int]$partsJ[2])
                                if ($sameArtistTitle -and $durDiff -le 3) {
                                    $match = $true
                                    $matchDetail = if ($durDiff -gt 0) { " (durée diff. ${durDiff}s)" } else { '' }
                                }
                            } elseif ($partsI.Count -lt 3 -and $partsJ.Count -lt 3) {
                                # Fallback nom normalisé (pas de metadata)
                                $match = ($normI -eq $normJ)
                            }
                        } else {
                            $match = ($normI -eq $normJ)
                        }

                        if ($match) {
                            $reason = switch ($mode) {
                                'films'    { $T.ReasonFilm }
                                'series'   { $T.ReasonSerie }
                                'fichiers' { $T.ReasonFichier }
                                'musique'  { "$($T.ReasonMusique)$matchDetail" }
                                default    { $T.ReasonVideoTitle }
                            }
                            $fj | Add-Member -NotePropertyName Reason -NotePropertyValue $reason -Force
                            if ($cluster.Count -eq 1) { $fi | Add-Member -NotePropertyName Reason -NotePropertyValue $reason -Force }
                            $cluster.Add($fj)
                            $used.Add($fj.FullName) | Out-Null
                        }
                    }

                    if ($cluster.Count -gt 1) { $used.Add($fi.FullName) | Out-Null; $groups.Add($cluster) }
                }
            }

        $state.Progress = 100
        $state.StepMsg  = $T.ScanDone
        $state.Groups   = $groups
        $state.Phase    = 'done'

    } catch {
        $state.Errors.Enqueue("$($T.CriticalError) $_")
        $state.Groups = [System.Collections.Generic.List[object]]::new()
        $state.Phase  = 'done'
    }
}

# ══════════════════════════════════════════════════════════
#  GESTION DU RUNSPACE
# ══════════════════════════════════════════════════════════

function Start-ScanAsync {
    Stop-Scan   # Nettoie un éventuel scan précédent (sans marquer comme annulé)

    # Reset état
    $script:State.Phase      = 'idle'
    $script:State.Progress   = 0
    $script:State.StepMsg    = $script:T.StatusInit
    $script:State.Groups     = $null
    $script:State.Errors     = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $script:State.FilesTotal = 0
    $script:State.Cancelled  = $false
    $script:State.StartTime  = [datetime]::Now
    $script:logBuf.Clear()
    $script:errCount = 0

    # UI → mode scan
    $btnScan.IsEnabled        = $false
    $btnScan.Content          = $script:T.BtnScanning
    $btnCancel.IsEnabled      = $true
    $btnCancel.Visibility     = 'Visible'
    $btnSelectAll.IsEnabled   = $false
    $btnSelectNone.IsEnabled  = $false
    $btnDelete.IsEnabled      = $false
    $panelProgress.Visibility = 'Visible'
    $panelLog.Visibility      = 'Collapsed'
    $tbLog.Text               = ''
    $prgScan.Value            = 0
    $txtPct.Text              = '0%'
    $txtETA.Text              = ''
    $txtCurrentFile.Text      = ''
    $txtStatus.Text           = $script:T.StatusInit
    $txtStats.Text            = ''
    $txtSelected.Text         = ''
    $panelGroups.Children.Clear()

    # Lire le mode sélectionné
    $script:State.Mode = if ($rbModeFilms.IsChecked)     { 'films' }
                         elseif ($rbModeSeries.IsChecked)   { 'series' }
                         elseif ($rbModeFichiers.IsChecked) { 'fichiers' }
                         elseif ($rbModeMusique.IsChecked)  { 'musique' }
                         else                               { 'video' }

    # Création du runspace
    $script:runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:runspace.Open()
    $script:psInstance = [powershell]::Create()
    $script:psInstance.Runspace = $script:runspace
    $script:psInstance.AddScript($script:scanBlock)        | Out-Null
    $script:psInstance.AddArgument($script:State)          | Out-Null
    $script:psInstance.AddArgument($script:scriptRoot)     | Out-Null
    $script:psInstance.AddArgument($script:VIDEO_EXT)      | Out-Null
    $script:psInstance.AddArgument($script:MUSIC_EXT)      | Out-Null
    $script:psInstance.AddArgument($script:MediaInfoPath)  | Out-Null
    $script:psInstance.AddArgument($script:T)              | Out-Null
    $script:asyncHandle = $script:psInstance.BeginInvoke()

    # Timer de polling (200 ms, thread UI)
    $script:timer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:timer.Interval = [TimeSpan]::FromMilliseconds(200)
    $script:timer.Add_Tick({ Poll-ScanProgress })
    $script:timer.Start()
}

function Stop-Scan {
    param([switch]$UserCancelled)
    if ($script:timer)      { $script:timer.Stop(); $script:timer = $null }
    if ($script:psInstance) {
        if ($UserCancelled) { $script:State.Cancelled = $true }
        try { $script:psInstance.Stop()    } catch {}
        try { $script:psInstance.Dispose() } catch {}
        $script:psInstance = $null
    }
    if ($script:runspace) {
        try { $script:runspace.Close()   } catch {}
        try { $script:runspace.Dispose() } catch {}
        $script:runspace = $null
    }
    $script:asyncHandle = $null
}

# ── Polling (exécuté sur le thread UI via DispatcherTimer) ─
function Poll-ScanProgress {
    $pct   = $script:State.Progress
    $msg   = $script:State.StepMsg
    $phase = $script:State.Phase

    # Barre + pourcentage
    $prgScan.Value       = $pct
    $txtPct.Text         = "$pct%"
    $txtCurrentFile.Text = $msg

    # ETA
    $elapsed = ([datetime]::Now - $script:State.StartTime).TotalSeconds
    if ($pct -gt 3 -and $elapsed -gt 0.5) {
        $estTotal  = $elapsed * 100.0 / $pct
        $remaining = [int]($estTotal - $elapsed)
        $txtETA.Text = "$($script:T.EtaLabel) $([TimeSpan]::FromSeconds([Math]::Max($remaining, 0)).ToString('mm\:ss'))"
    }

    # Statut texte
    $nf = $script:State.FilesTotal
    if ($nf -gt 0) { $txtStatus.Text = "$($script:T.StatusScanning) $nf $($script:T.StatusFilesIndexed)" }

    # Vidage de la queue d'erreurs vers le log
    [string]$errLine = $null
    while ($script:State.Errors.TryDequeue([ref]$errLine)) {
        $script:errCount++
        $script:logBuf.AppendLine($errLine) | Out-Null
    }
    if ($script:logBuf.Length -gt 0) {
        $tbLog.Text        = $script:logBuf.ToString()
        $txtLogTitle.Text  = "$($script:T.LogTitle) ($($script:errCount))"
        $panelLog.Visibility = 'Visible'
        $svLog.ScrollToBottom()
    }

    # ── Fin de scan ────────────────────────────────────
    if ($phase -eq 'done') {
        $script:timer.Stop(); $script:timer = $null

        # Récupérer les erreurs du stream PS (ex. exceptions non catchées)
        if ($script:psInstance) {
            try { $script:psInstance.EndInvoke($script:asyncHandle) } catch {}
            foreach ($e in $script:psInstance.Streams.Error) {
                $script:errCount++
                $script:logBuf.AppendLine("[PS] $($e.Exception.Message)") | Out-Null
            }
        }
        if ($script:logBuf.Length -gt 0) {
            $tbLog.Text        = $script:logBuf.ToString()
            $txtLogTitle.Text  = "$($script:T.LogTitle) ($($script:errCount))"
            $panelLog.Visibility = 'Visible'
        }

        $wasCancelled = $script:State.Cancelled
        Stop-Scan

        # UI → mode résultats
        $panelProgress.Visibility = 'Collapsed'
        $btnScan.IsEnabled        = $true
        $btnScan.Content          = $script:T.BtnScan
        $btnCancel.Visibility     = 'Collapsed'
        $btnCancel.IsEnabled      = $true

        $groups = $script:State.Groups
        $nf     = $script:State.FilesTotal

        if ($wasCancelled) {
            $txtStatus.Text = "$($script:T.StatusCancelled) $nf $($script:T.StatusCancelledSuffix)"
        } elseif ($groups -and $groups.Count -gt 0) {
            Build-UI -groups $groups
            $txtStatus.Text = "$($script:T.StatusDone) $nf $($script:T.StatusFilesAnalyzed) $($groups.Count) $($script:T.StatusGroupsFound)"
        } else {
            $txtStats.Text  = ""
            $txtStatus.Text = "$($script:T.StatusNoDupe) $nf $($script:T.StatusNoDupeSuffix)"
        }
    }
}

# ══════════════════════════════════════════════════════════
#  ÉVÉNEMENTS
# ══════════════════════════════════════════════════════════

$btnScan.Add_Click({ Start-ScanAsync })

$btnCancel.Add_Click({
    Stop-Scan -UserCancelled
    $btnCancel.IsEnabled = $false
    $txtStatus.Text      = $script:T.StatusCancelling
})

$btnSelectAll.Add_Click({
    # Pour chaque groupe de COPIES EXACTES uniquement :
    # → conserver le plus ancien, cocher tous les plus récents
    foreach ($group in $script:State.Groups) {
        $members = @($group)
        if ($members.Count -lt 2) { continue }

        # Vérifier que c'est bien un groupe de copies exactes
        $reason = if ($members[0].PSObject.Properties["Reason"]) { $members[0].Reason } else { '' }
        if ($reason -notmatch "^$([regex]::Escape($script:T.ReasonExact))") { continue }

        # Trier par LastWriteTime croissant → le plus ancien en [0]
        $sorted  = $members | Sort-Object -Property LastWriteTime
        $oldestPath = $sorted[0].FullName

        foreach ($cb in $script:allCheckboxes) {
            $f = $script:fileMap[$cb.Name]
            if (-not $f) { continue }
            $inGroup = $members | Where-Object { $_.FullName -eq $f.FullName }
            if ($inGroup) {
                $cb.IsChecked = ($f.FullName -ne $oldestPath)
            }
        }
    }
    Update-SelectionCount
})

$btnSelectNone.Add_Click({
    foreach ($cb in $script:allCheckboxes) { $cb.IsChecked = $false }
    Update-SelectionCount
})

$btnDelete.Add_Click({
    $toDelete = @(
        $script:allCheckboxes |
        Where-Object { $_.IsChecked } |
        ForEach-Object { $script:fileMap[$_.Name] } |
        Where-Object { $_ }
    )
    if ($toDelete.Count -eq 0) { return }

    $totalSz = ($toDelete | Measure-Object -Property Length -Sum).Sum

    # ── Fenêtre WPF de confirmation scrollable ──
    [xml]$xamlC = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:T.ConfirmTitle) $($toDelete.Count) $($script:T.ConfirmFileSuffix)"
        Height="600" Width="780" MinHeight="300" MinWidth="500"
        WindowStartupLocation="CenterOwner"
        Background="#0D0D0F" ResizeMode="CanResize">
  <Grid Margin="24">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Margin="0,0,0,12">
      <TextBlock x:Name="txtConfirmTitle" FontFamily="Consolas" FontSize="15"
                 FontWeight="Bold" Foreground="#E63946"/>
      <TextBlock x:Name="txtConfirmSub" FontFamily="Consolas" FontSize="11"
                 Foreground="#667788" Margin="0,4,0,0"/>
    </StackPanel>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Background="#13131A">
      <TextBlock x:Name="txtFileList" FontFamily="Consolas" FontSize="11"
                 Foreground="#AABBCC" Padding="12" TextWrapping="Wrap"/>
    </ScrollViewer>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
      <Button x:Name="btnAnnuler" Content="$($script:T.ConfirmCancel)" Width="110" Height="36" Margin="0,0,12,0"
              Background="#1E1E26" Foreground="#888" FontFamily="Consolas"
              BorderThickness="1" BorderBrush="#2E2E3A" Cursor="Hand"/>
      <Button x:Name="btnConfirmer" Content="$($script:T.ConfirmDelete)" Height="36"
              Background="#E63946" Foreground="White" FontFamily="Consolas"
              FontWeight="Bold" BorderThickness="0" Cursor="Hand" Padding="16,0"/>
    </StackPanel>
  </Grid>
</Window>
"@
    $rC   = [System.Xml.XmlNodeReader]::new($xamlC)
    $winC = [Windows.Markup.XamlReader]::Load($rC)
    $winC.Owner = $window

    $winC.FindName("txtConfirmTitle").Text = "$($script:T.ConfirmTitle) $($toDelete.Count) $($script:T.ConfirmFileSuffix)"
    $winC.FindName("txtConfirmSub").Text   = "$($script:T.ConfirmSub) $([Math]::Round($totalSz/1GB,3)) $($script:T.Go)  $($script:T.ConfirmIrreversible)"
    $winC.FindName("txtFileList").Text = ($toDelete | ForEach-Object {
        $sz = if ($_.Length -ge 1GB) { "$([Math]::Round($_.Length/1GB,2)) $($script:T.Go)" } `
              elseif ($_.Length -ge 1MB) { "$([Math]::Round($_.Length/1MB,1)) $($script:T.Mo)" } `
              else { "$([Math]::Round($_.Length/1KB,0)) $($script:T.Ko)" }
        "[$sz]  $($_.Name)`n  └ $($_.DirectoryName)"
    }) -join "`n`n"

    $script:_confirmed = $false
    $winC.FindName("btnAnnuler").Add_Click({  $winC.DialogResult = $false; $winC.Close() })
    $winC.FindName("btnConfirmer").Add_Click({ $script:_confirmed = $true; $winC.Close() })
    $winC.ShowDialog() | Out-Null

    if (-not $script:_confirmed) { return }

    if ($true) {
        $ok      = 0
        $skipped = 0
        $entries = [System.Collections.Generic.List[object]]::new()

        foreach ($f in $toDelete) {
            $entry = [PSCustomObject]@{ File = $f; Success = $false; Reason = '' }
            try {
                if (-not (Test-Path -LiteralPath $f.FullName)) {
                    $entry.Reason = $script:T.FileNotFound
                    $skipped++
                } else {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                    $entry.Success = $true
                    $ok++
                }
            }
            catch {
                $r = $_.Exception.Message
                $entry.Reason = if ($r -match 'utilisé par un autre processus|being used by another process') {
                    $script:T.FileInUse
                } elseif ($r -match 'access.*denied|accès.*refusé') {
                    $script:T.AccessDeniedMsg
                } else { $r }
                $skipped++
            }
            $entries.Add($entry) | Out-Null
        }

        # ── Rapport dans l'UI (remplace la liste des doublons) ──
        $conv = [System.Windows.Media.BrushConverter]::new()
        $panelGroups.Children.Clear()
        $script:allCheckboxes.Clear()
        $script:fileMap.Clear()

        $card = [System.Windows.Controls.Border]::new()
        $card.Background      = $conv.ConvertFrom("#13131A")
        $card.BorderBrush     = $conv.ConvertFrom("#1E1E2A")
        $card.BorderThickness = [System.Windows.Thickness]::new(1)
        $card.CornerRadius    = [System.Windows.CornerRadius]::new(6)
        $card.Margin          = [System.Windows.Thickness]::new(0,6,0,6)
        $card.Padding         = [System.Windows.Thickness]::new(16,12,16,12)

        $sp = [System.Windows.Controls.StackPanel]::new()

        $titre = [System.Windows.Controls.TextBlock]::new()
        $titre.Text       = $script:T.ReportTitle
        $titre.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
        $titre.FontSize   = 13
        $titre.FontWeight = "Bold"
        $titre.Foreground = $conv.ConvertFrom("#E63946")
        $titre.Margin     = [System.Windows.Thickness]::new(0,0,0,12)
        $sp.Children.Add($titre) | Out-Null

        $sep0 = [System.Windows.Controls.Separator]::new()
        $sep0.Background = $conv.ConvertFrom("#1E1E2A")
        $sep0.Margin     = [System.Windows.Thickness]::new(0,0,0,10)
        $sp.Children.Add($sep0) | Out-Null

        foreach ($entry in $entries) {
            $row = [System.Windows.Controls.Grid]::new()
            $c0 = [System.Windows.Controls.ColumnDefinition]::new(); $c0.Width = [System.Windows.GridLength]::new(28)
            $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::new(110)
            $row.ColumnDefinitions.Add($c0); $row.ColumnDefinitions.Add($c1); $row.ColumnDefinitions.Add($c2)
            $row.Margin = [System.Windows.Thickness]::new(0,3,0,3)

            # Icône ✔ verte / ✘ rouge
            $icon = [System.Windows.Controls.TextBlock]::new()
            $icon.FontFamily        = [System.Windows.Media.FontFamily]::new("Consolas")
            $icon.FontSize          = 15
            $icon.FontWeight        = "Bold"
            $icon.VerticalAlignment = "Center"
            if ($entry.Success) {
                $icon.Text       = "✔"
                $icon.Foreground = $conv.ConvertFrom("#44CC88")
            } else {
                $icon.Text       = "✘"
                $icon.Foreground = $conv.ConvertFrom("#E63946")
            }
            [System.Windows.Controls.Grid]::SetColumn($icon, 0)
            $row.Children.Add($icon) | Out-Null

            # Nom + sous-texte
            $vsp = [System.Windows.Controls.StackPanel]::new()
            $tbName = [System.Windows.Controls.TextBlock]::new()
            $tbName.Text         = $entry.File.Name
            $tbName.FontFamily   = [System.Windows.Media.FontFamily]::new("Consolas")
            $tbName.FontSize     = 12
            $tbName.Foreground   = if ($entry.Success) { $conv.ConvertFrom("#88AAEE") } else { $conv.ConvertFrom("#AA6666") }
            $tbName.TextTrimming = "CharacterEllipsis"
            $vsp.Children.Add($tbName) | Out-Null

            $tbSub = [System.Windows.Controls.TextBlock]::new()
            $tbSub.FontFamily   = [System.Windows.Media.FontFamily]::new("Consolas")
            $tbSub.FontSize     = 10
            $tbSub.TextTrimming = "CharacterEllipsis"
            if ($entry.Success) {
                $tbSub.Text       = $entry.File.DirectoryName
                $tbSub.Foreground = $conv.ConvertFrom("#6A7A8A")
            } else {
                $tbSub.Text       = $entry.Reason
                $tbSub.Foreground = $conv.ConvertFrom("#CC5544")
            }
            $vsp.Children.Add($tbSub) | Out-Null
            [System.Windows.Controls.Grid]::SetColumn($vsp, 1)
            $row.Children.Add($vsp) | Out-Null

            # Taille
            $f  = $entry.File
            $sz = if ($f.Length -ge 1GB) { "$([Math]::Round($f.Length/1GB,2)) $($script:T.Go)" } `
                  elseif ($f.Length -ge 1MB) { "$([Math]::Round($f.Length/1MB,1)) $($script:T.Mo)" } `
                  else { "$([Math]::Round($f.Length/1KB,0)) $($script:T.Ko)" }
            $tbSz = [System.Windows.Controls.TextBlock]::new()
            $tbSz.Text                = $sz
            $tbSz.FontFamily          = [System.Windows.Media.FontFamily]::new("Consolas")
            $tbSz.FontSize            = 12
            $tbSz.Foreground          = $conv.ConvertFrom("#778899")
            $tbSz.HorizontalAlignment = "Right"
            $tbSz.VerticalAlignment   = "Center"
            [System.Windows.Controls.Grid]::SetColumn($tbSz, 2)
            $row.Children.Add($tbSz) | Out-Null

            $sp.Children.Add($row) | Out-Null
        }

        $card.Child = $sp
        $panelGroups.Children.Add($card) | Out-Null

        $txtStatus.Text          = "$($script:T.ReportTitle) — $ok $($script:T.ReportSummary)  $skipped $($script:T.ReportIgnored)"
        $txtStats.Text           = "$ok $($script:T.ReportStats)  $skipped $($script:T.ReportFailed)"
        $txtSelected.Text        = ""
        $btnDelete.IsEnabled     = $false
        $btnSelectAll.IsEnabled  = $false
        $btnSelectNone.IsEnabled = $false
    }
})

$window.Add_Closing({ Stop-Scan })

$window.ShowDialog() | Out-Null
