// Templates feature: public surface used by Compose and the library page.
//   TemplatePicker      searchable switcher (Compose toolbar)
//   SaveTemplateDialog  save / update from Compose
export { TemplatePicker, type TemplatePickerProps } from './TemplatePicker';
export { SaveTemplateDialog, type SaveTemplateDialogProps } from './SaveTemplateDialog';
export { TemplateThumbnail, templatePreviewHtml } from './TemplateThumbnail';
export { TemplateCard } from './TemplateCard';
export { RenameDialog, ShareDialog, DeleteDialog, ImportDialog, type PendingImport } from './LibraryDialogs';
export {
  buildExport, parseImport, parseImportText, legacyToRequest, legacyToRequests, templateToRequest, copyNameFor,
  LEGACY_STORAGE_KEY, type ExportFileV2,
} from './importExport';
export {
  categoriesOf, detailKey, fetchTemplate, primeTemplate, themeLabel, useAllTemplates, useTemplateDetail, useTemplates,
  visibilityLabel,
} from './hooks';
