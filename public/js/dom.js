// Central DOM element registry. Populated once by initEls() from main.js
// before any other module's setup runs; other modules import `els` (a
// live object reference) and read its properties freely afterwards.
export const els = {};

export function initEls() {
    const $ = (id) => document.getElementById(id);

    Object.assign(els, {
        // Topbar / stepper
        stepper: $('stepper'),
        stepPanels: document.querySelectorAll('.step-panel'),
        activityBtn: $('activityBtn'),
        accountChip: $('accountChip'),
        userNameDisplay: $('userNameDisplay'),
        userAvatar: $('userAvatar'),

        // Banners
        resumeBanner: $('resumeBanner'),
        resumeMeta: $('resumeMeta'),
        resumeBtn: $('resumeBtn'),
        discardResumeBtn: $('discardResumeBtn'),
        degradedBanner: $('degradedBanner'),
        degradedBannerText: $('degradedBannerText'),

        // Recipients step
        recipientSourceSeg: $('recipientSourceSeg'),
        singleFields: $('singleFields'),
        bulkCsv: $('bulkCsv'),
        bulkManual: $('bulkManual'),
        singleEmail: $('singleEmail'),
        singleName: $('singleName'),
        singleCustomFields: $('singleCustomFields'),
        addSingleFieldBtn: $('addSingleFieldBtn'),
        ccEmail: $('ccEmail'),
        bccEmail: $('bccEmail'),

        csvDrop: $('csvDrop'),
        csvFile: $('csvFile'),
        csvStatus: $('csvStatus'),
        downloadTmplBtn: $('downloadTemplateBtn'),
        mappingCard: $('mappingCard'),
        mappingRows: $('mappingRows'),
        validationCard: $('validationCard'),
        validationStats: $('validationStats'),
        validationDetails: $('validationDetails'),

        manualGrid: $('manualGrid'),
        addRowBtn: $('addRowBtn'),
        addColBtn: $('addColBtn'),
        clearGridBtn: $('clearGridBtn'),
        gridStatus: $('gridStatus'),
        deleteSelectedBtn: $('deleteSelectedBtn'),
        selectedCount: $('selectedCount'),
        pasteImportBtn: $('pasteImportBtn'),
        pasteImportArea: $('pasteImportArea'),
        pasteImportText: $('pasteImportText'),
        pasteImportApply: $('pasteImportApply'),
        pasteImportCancel: $('pasteImportCancel'),
        pasteImportAppend: $('pasteImportAppend'),
        exportGridBtn: $('exportGridBtn'),
        mappingCardManual: $('mappingCardManual'),
        mappingRowsManual: $('mappingRowsManual'),
        validationCardManual: $('validationCardManual'),
        validationStatsManual: $('validationStatsManual'),
        validationDetailsManual: $('validationDetailsManual'),

        // Compose step
        subject: $('subject'),
        subjectCounter: $('subjectCounter'),
        subjectWarn: $('subjectWarn'),
        preheader: $('preheader'),
        preheaderCounter: $('preheaderCounter'),
        placeholderToolbar: $('placeholderToolbar'),

        editorSeg: $('editorSeg'),
        htmlEditor: $('htmlEditor'),
        textEditor: $('textEditor'),
        richEditorContainer: $('richEditorParams'),
        htmlEditorContainer: $('htmlEditorParams'),
        textEditorContainer: $('textEditorParams'),

        templateSelect: $('templateSelect'),
        loadTemplateBtn: $('loadTemplateBtn'),
        saveTemplateBtn: $('saveTemplateBtn'),
        deleteTemplateBtn: $('deleteTemplateBtn'),
        exportTemplatesBtn: $('exportTemplatesBtn'),
        importTemplatesInput: $('importTemplatesInput'),
        saveTemplateModal: $('saveTemplateModal'),
        templateNameInput: $('templateNameInput'),
        cancelSaveTemplate: $('cancelSaveTemplate'),
        confirmSaveTemplate: $('confirmSaveTemplate'),

        attachDrop: $('attachDrop'),
        attachInput: $('attachments'),
        attachList: $('attachmentList'),
        attachMeter: $('attachMeter'),
        attachMeterFill: $('attachMeterFill'),
        attachMeterText: $('attachMeterText'),
        attachWarnLine: $('attachWarnLine'),

        // Review step
        summaryFrom: $('summaryFrom'),
        summaryRecipients: $('summaryRecipients'),
        summaryCopies: $('summaryCopies'),
        summaryAttachments: $('summaryAttachments'),
        preflightBody: $('preflightBody'),
        refreshPreflightBtn: $('refreshPreflightBtn'),
        quotaFill: $('quotaFill'),
        quotaText: $('quotaText'),
        delayMs: $('delayMs'),
        dedupToggle: $('dedupToggle'),
        sendBtn: $('sendBtn'),
        sendTestBtn: $('sendTestBtn'),

        // Preview pane
        previewPrevBtn: $('previewPrevBtn'),
        previewNextBtn: $('previewNextBtn'),
        previewRecipientLabel: $('previewRecipientLabel'),
        previewThemeSeg: $('previewThemeSeg'),
        previewViewportSeg: $('previewViewportSeg'),
        previewWarnLine: $('previewWarnLine'),
        previewFrame: $('previewFrame'),
        gmailFrame: $('gmailFrame'),
        gmSubject: $('gmSubject'),
        gmFromName: $('gmFromName'),
        gmFromEmail: $('gmFromEmail'),
        gmTo: $('gmTo'),
        gmAvatar: $('gmAvatar'),
        gmDate: $('gmDate'),

        // Drawers
        settingsDrawer: $('settingsDrawer'),
        settingsDrawerBackdrop: $('settingsDrawerBackdrop'),
        closeSettingsBtn: $('closeSettingsBtn'),
        settingsEmail: $('settingsEmail'),
        settingsPass: $('settingsPass'),
        settingsDisplayName: $('settingsDisplayName'),
        settingsReplyTo: $('settingsReplyTo'),
        saveSettingsBtn: $('saveSettingsBtn'),

        activityDrawer: $('activityDrawer'),
        activityDrawerBackdrop: $('activityDrawerBackdrop'),
        closeActivityBtn: $('closeActivityBtn'),
        logTerminal: $('logTerminal'),
        clearLogsBtn: $('clearLogs'),

        // Modals
        modeWarningModal: $('modeWarningModal'),
        confirmModeSwitch: $('confirmModeSwitch'),
        cancelModeSwitch: $('cancelModeSwitch'),

        // Send console
        overlay: $('progressOverlay'),
        consoleTitle: $('consoleTitle'),
        pauseBtn: $('pauseBtn'),
        resumeBtnConsole: $('resumeBtnConsole'),
        cancelBtn: $('cancelBtn'),
        progressBar: $('progressBar'),
        progressText: $('progressText'),
        progressPercent: $('progressPercent'),
        successCount: $('successCount'),
        failureCount: $('failureCount'),
        pendingCount: $('pendingCount'),
        resultsFilterSeg: $('resultsFilterSeg'),
        resultsTable: $('resultsTable'),
        retryFailedBtn: $('retryFailedBtn'),
        exportResultsBtn: $('exportResultsBtn'),
        closeConsoleBtn: $('closeConsoleBtn'),

        toastContainer: $('toastContainer'),
    });
}
