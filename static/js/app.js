// Correspon Frontend Logic
document.addEventListener('DOMContentLoaded', () => {
    // State management
    const state = {
        activeProjectId: null,
        excelFilePath: null,
        excelFileName: null,
        excelColumns: [],
        excelRecords: [],
        docxFilePath: null,
        docxFileName: null,
        docxTemplates: [], // Lista de { filepath, filename, html }
        activeDocxIndex: 0,
        outputDirPath: null,
        zoomLevel: 1.0,
        currentJobId: null,
        currentOutputDir: null,
        pollInterval: null,
        activeView: 'word', // 'word' or 'pdf'
        excelMtime: 0,
        isAutoSyncing: false,
        settings: {
            output_format: 'docx',
            output_filename_pattern: '{{ index }}_documento.docx',
            outlook_enabled: false,
            email_column: 'Email',
            check_column: 'Enviado',
            doc_check_column: 'Doc_Generado',
            doc_check_value_mode: 'timestamp',
            doc_check_custom_text: 'Generado',
            all_fields_mandatory: false,
            mandatory_columns: [],
            template_rules_enabled: false,
            template_rules: [],
            send_mode: 'display',
            attachment_type: 'pdf',
            email_subject: 'Notificación Oficial - {{ Nombre }}',
            email_body_source: 'text',
            email_body: 'Estimado/a {{ Nombre }},\n\nAdjuntamos su documento oficial.\n\nSaludos cordiales.',
            single_email_override: '',
            custom_paragraph_rules: []
        }
    };

    // DOM Elements
    const excelDropzone = document.getElementById('excel-dropzone');
    const excelInput = document.getElementById('excel-input');
    const excelInfo = document.getElementById('excel-info');
    const excelFilename = document.getElementById('excel-filename');
    const excelRowsCount = document.getElementById('excel-rows-count');
    const btnRemoveExcel = document.getElementById('btn-remove-excel');
    const btnViewExcel = document.getElementById('btn-view-excel');
    const excelStatusPill = document.getElementById('excel-status-pill');

    const projectSelectDropdown = document.getElementById('project-select-dropdown');
    const btnOpenCreateProject = document.getElementById('btn-open-create-project');
    const createProjectModal = document.getElementById('create-project-modal');
    const btnCloseCreateProjModal = document.getElementById('btn-close-create-proj-modal');
    const btnCancelCreateProj = document.getElementById('btn-cancel-create-proj');
    const btnConfirmCreateProj = document.getElementById('btn-confirm-create-proj');
    const projNameInput = document.getElementById('proj-name-input');
    const projPathInput = document.getElementById('proj-path-input');

    const excelModal = document.getElementById('excel-modal');
    const btnCloseExcelModal = document.getElementById('btn-close-excel-modal');
    const btnCancelExcelEdit = document.getElementById('btn-cancel-excel-edit');
    const btnSaveExcelEdit = document.getElementById('btn-save-excel-edit');
    const btnAddRow = document.getElementById('btn-add-row');
    const btnAddCol = document.getElementById('btn-add-col');
    const excelGridThead = document.getElementById('excel-grid-thead');
    const excelGridTbody = document.getElementById('excel-grid-tbody');

    const docxDropzone = document.getElementById('docx-dropzone');
    const docxInput = document.getElementById('docx-input');
    const docxTemplatesList = document.getElementById('docx-templates-list');
    const btnAddDocxTrigger = document.getElementById('btn-add-docx-trigger');
    const docxStatusPill = document.getElementById('docx-status-pill');

    const variablesContainer = document.getElementById('variables-container');
    const varsCount = document.getElementById('vars-count');
    const btnOpenParagraphRules = document.getElementById('btn-open-paragraph-rules');
    const paperPage = document.getElementById('paper-page');
    const pdfViewerFrame = document.getElementById('pdf-viewer-frame');
    const pdfViewerWrapper = document.getElementById('pdf-viewer-wrapper');
    const pdfLoadingOverlay = document.getElementById('pdf-loading-overlay');
    const docTypeIcon = document.getElementById('doc-type-icon');
    const docTitleDisplay = document.getElementById('doc-title-display');
    const btnSaveDocx = document.getElementById('btn-save-docx');
    const saveDocxSplitGroup = document.getElementById('save-docx-split-group');
    const btnSaveAppToWord = document.getElementById('btn-save-app-to-word');
    const btnSyncWordToApp = document.getElementById('btn-sync-word-to-app');
    const btnViewWord = document.getElementById('btn-view-word');
    const btnViewPdf = document.getElementById('btn-view-pdf');
    const btnToggleBadges = document.getElementById('btn-toggle-badges');
    const sendEmailsSplitGroup = document.getElementById('send-emails-split-group');
    const btnOpenTemplateRulesTrigger = document.getElementById('btn-open-template-rules-trigger');

    const btnOpenSettings = document.getElementById('btn-open-settings');
    const settingsModal = document.getElementById('settings-modal');
    const btnCloseSettingsModal = document.getElementById('btn-close-settings-modal');
    const btnCancelSettings = document.getElementById('btn-cancel-settings');
    const btnSaveSettings = document.getElementById('btn-save-settings');

    const settingOutputFormat = document.getElementById('setting-output-format');
    const settingDocCheckEnabled = document.getElementById('setting-doc-check-enabled');
    const docCheckOptionsContainer = document.getElementById('doc-check-options-container');
    const settingDocCheckCol = document.getElementById('setting-doc-check-col');
    const settingDocCheckMode = document.getElementById('setting-doc-check-mode');
    const docCheckCustomTextContainer = document.getElementById('doc-check-custom-text-container');
    const settingDocCheckCustomText = document.getElementById('setting-doc-check-custom-text');
    const settingTemplateRulesEnabled = document.getElementById('setting-template-rules-enabled');
    const templateRulesOptionsContainer = document.getElementById('template-rules-options-container');
    const btnOpenTemplateRulesModal = document.getElementById('btn-open-template-rules-modal');
    const toggleAllFieldsMandatory = document.getElementById('toggle-all-fields-mandatory');

    const addColumnModal = document.getElementById('add-column-modal');
    const newColumnNameInput = document.getElementById('new-column-name-input');
    const btnCancelAddColumn = document.getElementById('btn-cancel-add-column');
    const btnConfirmAddColumn = document.getElementById('btn-confirm-add-column');

    const settingEmailBodySource = document.getElementById('setting-email-body-source');
    const outlookTextBodyContainer = document.getElementById('outlook-text-body-container');
    const settingEmailBodyTemplate = document.getElementById('setting-email-body-template');
    const outlookTemplateBodyContainer = document.getElementById('outlook-template-body-container');
    const settingOutlookEnabled = document.getElementById('setting-outlook-enabled');
    const outlookFieldsContainer = document.getElementById('outlook-fields-container');
    const settingEmailCol = document.getElementById('setting-email-col');
    const settingCheckCol = document.getElementById('setting-check-col');
    const settingSendMode = document.getElementById('setting-send-mode');
    const settingAttachmentType = document.getElementById('setting-attachment-type');
    const settingEmailSubject = document.getElementById('setting-email-subject');
    const settingEmailBody = document.getElementById('setting-email-body');
    const emailVarsList = document.getElementById('email-vars-list');

    const btnAddParagraphRule = document.getElementById('btn-add-paragraph-rule');
    const paragraphRulesContainer = document.getElementById('paragraph-rules-container');
    const paragraphRulesModal = document.getElementById('paragraph-rules-modal');
    const btnCloseParagraphModal = document.getElementById('btn-close-paragraph-modal');
    const btnCancelParagraphRules = document.getElementById('btn-cancel-paragraph-rules');
    const btnSaveParagraphRules = document.getElementById('btn-save-paragraph-rules');

    const templateRulesModal = document.getElementById('template-rules-modal');
    const btnCloseTemplateRulesModal = document.getElementById('btn-close-template-rules-modal');
    const btnCancelTemplateRules = document.getElementById('btn-cancel-template-rules');
    const btnSaveTemplateRules = document.getElementById('btn-save-template-rules');
    const btnAddTemplateRule = document.getElementById('btn-add-template-rule');
    const templateRulesListContainer = document.getElementById('template-rules-list-container');

    const docsSelectionModal = document.getElementById('docs-selection-modal');
    const btnCloseDocsSelectionModal = document.getElementById('btn-close-docs-selection-modal');
    const btnCancelDocsSelectionModal = document.getElementById('btn-cancel-docs-selection-modal');
    const docsSelectionIdCol = document.getElementById('docs-selection-id-col');
    const docsSelectionSearchFilter = document.getElementById('docs-selection-search-filter');
    const checkAllDocsSelection = document.getElementById('check-all-docs-selection');
    const docsSelectionCounterBadge = document.getElementById('docs-selection-counter-badge');
    const docsSelectionListContainer = document.getElementById('docs-selection-list-container');
    const btnGenerateMissingDocsAction = document.getElementById('btn-generate-missing-docs-action');
    const btnGenerateSelectedDocsAction = document.getElementById('btn-generate-selected-docs-action');

    const btnGenerateDocs = document.getElementById('btn-generate-docs');
    const btnGenerateDocsOptions = document.getElementById('btn-generate-docs-options');
    const btnSendEmails = document.getElementById('btn-send-emails');
    const btnSendEmailsOptions = document.getElementById('btn-send-emails-options');
    const patternInput = document.getElementById('pattern-input');

    const btnOpenNativeExcel = document.getElementById('btn-open-native-excel');
    const btnOpenNativeDocx = document.getElementById('btn-open-native-docx');
    const btnOpenProjectFolder = document.getElementById('btn-open-project-folder');
    const btnDeleteProject = document.getElementById('btn-delete-project');

    const recipientsSelectionModal = document.getElementById('recipients-selection-modal');
    const btnCloseRecipientsModal = document.getElementById('btn-close-recipients-modal');
    const btnCancelRecipientsModal = document.getElementById('btn-cancel-recipients-modal');
    const recipientIdCol = document.getElementById('recipient-id-col');
    const recipientSearchFilter = document.getElementById('recipient-search-filter');
    const checkAllRecipients = document.getElementById('check-all-recipients');
    const recipientsCounterBadge = document.getElementById('recipients-counter-badge');
    const recipientsListContainer = document.getElementById('recipients-list-container');
    const btnSendSelectedRecipients = document.getElementById('btn-send-selected-recipients');
    const btnCancelJob = document.getElementById('btn-cancel-job');

    const btnZoomIn = document.getElementById('btn-zoom-in');
    const btnZoomOut = document.getElementById('btn-zoom-out');
    const zoomLevelDisplay = document.getElementById('zoom-level');
    const pdfVariationSelect = document.getElementById('pdf-variation-select');

    const progressModal = document.getElementById('progress-modal');
    const progressBarFill = document.getElementById('progress-bar-fill');
    const progressText = document.getElementById('progress-text');
    const progressPercent = document.getElementById('progress-percent');
    const modalTitle = document.getElementById('modal-title');
    const modalSubtitle = document.getElementById('modal-subtitle');

    // MODALES IN-APP DE ELIMINACIÓN
    const deleteProjectModal = document.getElementById('delete-project-modal');
    const btnCloseDeleteProjModal = document.getElementById('btn-close-delete-proj-modal');
    const btnCancelDeleteProj = document.getElementById('btn-cancel-delete-proj');
    const btnConfirmDeleteProj = document.getElementById('btn-confirm-delete-proj');
    const deleteProjTargetName = document.getElementById('delete-proj-target-name');
    const deleteProjConfirmNameInput = document.getElementById('delete-proj-confirm-name-input');

    const confirmDeleteFileModal = document.getElementById('confirm-delete-file-modal');
    const btnCloseDeleteFileModal = document.getElementById('btn-close-delete-file-modal');
    const btnCancelDeleteFile = document.getElementById('btn-cancel-delete-file');
    const btnAcceptDeleteFile = document.getElementById('btn-accept-delete-file');
    const confirmDeleteFileMessage = document.getElementById('confirm-delete-file-message');

    let onFileDeleteAcceptCallback = null;

    function showInAppDeleteConfirm(messageHtml, onAccept) {
        if (confirmDeleteFileMessage) confirmDeleteFileMessage.innerHTML = messageHtml;
        onFileDeleteAcceptCallback = onAccept;
        showModal(confirmDeleteFileModal);
    }

    if (btnCloseDeleteFileModal) btnCloseDeleteFileModal.addEventListener('click', () => confirmDeleteFileModal.classList.add('hidden'));
    if (btnCancelDeleteFile) btnCancelDeleteFile.addEventListener('click', () => confirmDeleteFileModal.classList.add('hidden'));
    if (btnAcceptDeleteFile) {
        btnAcceptDeleteFile.addEventListener('click', async () => {
            if (confirmDeleteFileModal) confirmDeleteFileModal.classList.add('hidden');
            if (onFileDeleteAcceptCallback) {
                const cb = onFileDeleteAcceptCallback;
                onFileDeleteAcceptCallback = null;
                await cb();
            }
        });
    }
    const modalActions = document.getElementById('modal-actions');
    const btnOpenFolder = document.getElementById('btn-open-folder');
    const btnCloseModal = document.getElementById('btn-close-modal');

    // Show a modal instantly and reliably
    function showModal(modal) {
        if (!modal) return;
        document.querySelectorAll('.modal-backdrop').forEach(m => m.classList.add('hidden'));
        modal.classList.remove('hidden');
    }

    function renderHtmlInPaper(htmlContent) {
        if (!paperPage) return;
        paperPage.innerHTML = htmlContent || '';
    }

    // INITIALIZE PROJECTS LIST
    loadProjectsList();

    async function loadProjectsList() {
        try {
            const res = await fetch('/api/projects');
            const data = await res.json();
            
            projectSelectDropdown.innerHTML = '<option value="">-- Seleccionar Correspondencia / Ruta --</option>';
            const seenNames = new Set();
            const seenPaths = new Set();

            data.projects.forEach(p => {
                const normPath = (p.folder_path || '').toLowerCase();
                const normName = (p.name || '').toLowerCase();
                if ((normPath && seenPaths.has(normPath)) || (normName && seenNames.has(normName))) {
                    return;
                }
                if (normPath) seenPaths.add(normPath);
                if (normName) seenNames.add(normName);

                const opt = document.createElement('option');
                opt.value = p.id;
                opt.textContent = p.name;
                if (p.id === data.active_project_id) {
                    opt.selected = true;
                }
                projectSelectDropdown.appendChild(opt);
            });

            if (data.projects.length > 0) {
                const activeId = data.active_project_id || data.projects[0].id;
                const active = data.projects.find(p => p.id === activeId) || data.projects[0];
                if (active) {
                    projectSelectDropdown.value = active.id;
                    await loadSelectedProject(active.id);
                }
            }
        } catch (err) {
            console.error('Error cargando proyectos:', err);
        }
    }

    if (projectSelectDropdown) {
        projectSelectDropdown.addEventListener('change', (e) => {
            const val = e.target.value;
            if (val) loadSelectedProject(val);
        });
    }

    async function loadSelectedProject(projectId) {
        try {
            const res = await fetch('/api/select-project', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ project_id: projectId })
            });

            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error al cargar el proyecto');

            state.activeProjectId = projectId;
            state.projectFolderPath = data.project.folder_path;
            state.excelFilePath = data.project.excel_path;
            state.excelFileName = data.project.excel_path ? data.project.excel_path.split(/[\\/]/).pop() : null;
            state.excelColumns = data.columns || [];
            state.excelRecords = data.records || [];

            state.docxTemplates = data.docx_templates || [];
            state.activeDocxIndex = 0;

            if (state.docxTemplates.length > 0) {
                state.docxFilePath = state.docxTemplates[0].filepath;
                state.docxFileName = state.docxTemplates[0].filename;
            } else {
                state.docxFilePath = data.project.docx_path;
                state.docxFileName = null;
            }

            state.outputDirPath = data.project.output_path;

            if (data.project.settings) {
                state.settings = { ...state.settings, ...data.project.settings };
            }

            if (state.settings.output_filename_pattern && patternInput) {
                patternInput.value = state.settings.output_filename_pattern;
            }

            // UI updates for Excel
            if (state.excelFilePath && state.excelColumns.length > 0) {
                excelDropzone.classList.add('hidden');
                excelInfo.classList.remove('hidden');
                btnViewExcel.classList.remove('hidden');
                if (btnOpenNativeExcel) btnOpenNativeExcel.classList.remove('hidden');
                excelFilename.textContent = state.excelFileName || 'datos.xlsx';
                excelRowsCount.textContent = `${data.total_rows} registros cargados`;
                if (excelStatusPill) {
                    excelStatusPill.classList.remove('disabled');
                    excelStatusPill.classList.add('active');
                    excelStatusPill.querySelector('span').textContent = `${data.columns.length} Cols`;
                }
            } else {
                excelDropzone.classList.remove('hidden');
                excelInfo.classList.add('hidden');
                btnViewExcel.classList.add('hidden');
                if (btnOpenNativeExcel) btnOpenNativeExcel.classList.add('hidden');
                if (excelStatusPill) {
                    excelStatusPill.classList.add('disabled');
                    excelStatusPill.classList.remove('active');
                    excelStatusPill.querySelector('span').textContent = 'Sin Excel';
                }
            }

            if (btnOpenProjectFolder) btnOpenProjectFolder.classList.remove('hidden');
            if (btnDeleteProject) btnDeleteProject.classList.remove('hidden');

            renderDocxTemplatesList();
            
            if (state.docxTemplates.length > 0) {
                if (docTitleDisplay) docTitleDisplay.textContent = `${data.project.name} / ${state.docxTemplates[0].filename}`;
                renderHtmlInPaper(state.docxTemplates[0].html);
                if (saveDocxSplitGroup) saveDocxSplitGroup.classList.remove('hidden');
                if (btnSaveDocx) btnSaveDocx.classList.remove('hidden');
                if (btnOpenNativeDocx) btnOpenNativeDocx.classList.remove('hidden');
                if (docxStatusPill) {
                    docxStatusPill.classList.remove('disabled');
                    docxStatusPill.classList.add('active');
                    docxStatusPill.querySelector('span').textContent = `${state.docxTemplates.length} Plantilla(s)`;
                }
            } else {
                if (docTitleDisplay) docTitleDisplay.textContent = `${data.project.name} (Sin Plantilla Word)`;
                if (saveDocxSplitGroup) saveDocxSplitGroup.classList.add('hidden');
                if (btnSaveDocx) btnSaveDocx.classList.add('hidden');
                if (btnOpenNativeDocx) btnOpenNativeDocx.classList.add('hidden');
                if (docxStatusPill) {
                    docxStatusPill.classList.add('disabled');
                    docxStatusPill.classList.remove('active');
                    docxStatusPill.querySelector('span').textContent = 'Sin Plantilla';
                }
                paperPage.innerHTML = `
                    <div class="welcome-placeholder">
                        <div class="placeholder-content">
                            <i class="fa-solid fa-file-circle-plus"></i>
                            <h2>Carga tu documento .docx base</h2>
                            <p>El contenido de tu plantilla aparecerá aquí. Podrás arrastrar y soltar las fichas de variables de Excel directamente en cualquier párrafo o tabla.</p>
                        </div>
                    </div>`;
            }

            renderVariableChips(state.excelColumns);
            renderEmailMappableVars();
            updateGenerateButtonState();
            showToast(`Proyecto '${data.project.name}' cargado`, 'success');
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    function renderDocxTemplatesList() {
        docxTemplatesList.innerHTML = '';

        if (!state.docxTemplates || state.docxTemplates.length === 0) {
            docxTemplatesList.classList.add('hidden');
            btnAddDocxTrigger.classList.add('hidden');
            docxDropzone.classList.remove('hidden');
            if (btnSaveDocx) btnSaveDocx.classList.add('hidden');
            if (btnOpenNativeDocx) btnOpenNativeDocx.classList.add('hidden');
            if (docxStatusPill) {
                docxStatusPill.classList.add('disabled');
                docxStatusPill.classList.remove('active');
                docxStatusPill.querySelector('span').textContent = 'Sin Plantilla';
            }
            if (docTitleDisplay) docTitleDisplay.textContent = 'Documento Sin Título.docx';
            paperPage.innerHTML = `
                <div class="welcome-placeholder">
                    <div class="placeholder-content">
                        <i class="fa-solid fa-file-circle-plus"></i>
                        <h2>Carga tu documento .docx base</h2>
                        <p>El contenido de tu plantilla aparecerá aquí. Podrás arrastrar y soltar las fichas de variables de Excel directamente en cualquier párrafo o tabla.</p>
                    </div>
                </div>`;
            return;
        }

        docxTemplatesList.classList.remove('hidden');
        btnAddDocxTrigger.classList.remove('hidden');
        docxDropzone.classList.add('hidden');

        state.docxTemplates.forEach((tpl, idx) => {
            const div = document.createElement('div');
            div.className = `docx-template-chip ${idx === state.activeDocxIndex ? 'active' : ''}`;
            div.style.cursor = 'pointer';

            div.innerHTML = `
                <div class="tpl-info" style="flex: 1; display: flex; align-items: center; gap: 8px; overflow: hidden;">
                    <i class="fa-solid fa-file-word text-accent"></i>
                    <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(tpl.filename)}">${escapeHtml(tpl.filename)}</span>
                </div>
                <button class="btn-icon-danger btn-delete-tpl" title="Eliminar plantilla"><i class="fa-solid fa-xmark"></i></button>
            `;

            div.addEventListener('click', (e) => {
                if (e.target.closest('.btn-delete-tpl')) return;
                selectActiveDocxTemplate(idx);
            });

            const delBtn = div.querySelector('.btn-delete-tpl');
            if (delBtn) {
                delBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    deleteDocxTemplate(idx);
                });
            }

            docxTemplatesList.appendChild(div);
        });
    }

    function selectActiveDocxTemplate(index) {
        if (!state.docxTemplates || index < 0 || index >= state.docxTemplates.length) return;
        state.activeDocxIndex = index;
        const activeTpl = state.docxTemplates[index];
        state.docxFilePath = activeTpl.filepath;
        state.docxFileName = activeTpl.filename;

        const projName = state.activeProjectId ? (state.projectFolderPath ? state.projectFolderPath.split(/[\\/]/).pop() : 'Proyecto') : 'Proyecto';
        if (docTitleDisplay) docTitleDisplay.textContent = `${projName} / ${activeTpl.filename}`;

        renderHtmlInPaper(activeTpl.html);
        renderDocxTemplatesList();

        if (state.activeView === 'pdf') {
            loadPdfPreviewForActiveTemplate(true);
        }
    }

    async function deleteDocxTemplate(index) {
        const tplToDelete = state.docxTemplates[index];
        if (!tplToDelete) return;

        showInAppDeleteConfirm(`¿Estás seguro de que deseas desvincular la plantilla Word <strong>${escapeHtml(tplToDelete.filename)}</strong> del proyecto?`, async () => {
            try {
                if (state.activeProjectId) {
                    const res = await fetch('/api/unlink-docx', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            project_id: state.activeProjectId,
                            docx_path: tplToDelete.filepath
                        })
                    });

                    const data = await res.json();
                    if (!res.ok) throw new Error(data.detail || 'Error desvinculando plantilla');
                }

                state.docxTemplates.splice(index, 1);
                state.activeDocxIndex = Math.max(0, state.activeDocxIndex - 1);
                if (state.docxTemplates.length > 0) {
                    selectActiveDocxTemplate(state.activeDocxIndex);
                } else {
                    renderDocxTemplatesList();
                    if (saveDocxSplitGroup) saveDocxSplitGroup.classList.add('hidden');
                    if (btnOpenNativeDocx) btnOpenNativeDocx.classList.add('hidden');
                    if (docxStatusPill) {
                        docxStatusPill.classList.add('disabled');
                        docxStatusPill.classList.remove('active');
                        docxStatusPill.querySelector('span').textContent = 'Sin Plantilla';
                    }
                    paperPage.innerHTML = `
                        <div class="welcome-placeholder">
                            <div class="placeholder-content">
                                <i class="fa-solid fa-file-circle-plus"></i>
                                <h2>Carga tu documento .docx base</h2>
                                <p>El contenido de tu plantilla aparecerá aquí.</p>
                            </div>
                        </div>`;
                }
                updateGenerateButtonState();
                showToast(`Plantilla '${tplToDelete.filename}' desvinculada del proyecto (conservada en disco)`, 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    if (btnAddDocxTrigger) {
        btnAddDocxTrigger.addEventListener('click', () => {
            docxInput.click();
        });
    }

    // BOTONES DE APERTURA NATIVA DE ARCHIVOS
    if (btnOpenNativeExcel) {
        btnOpenNativeExcel.addEventListener('click', async (e) => {
            if (e) e.preventDefault();
            if (!state.excelFilePath) {
                showToast('No hay ningún archivo Excel cargado en este proyecto', 'error');
                return;
            }
            try {
                const res = await fetch('/api/open-file', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filepath: state.excelFilePath })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'No se pudo abrir el archivo Excel');
                showToast('Abriendo archivo Excel en aplicación nativa...', 'info');
            } catch (err) {
                showToast(err.message || 'Error de conexión al abrir Excel', 'error');
            }
        });
    }

    if (btnOpenNativeDocx) {
        btnOpenNativeDocx.addEventListener('click', async (e) => {
            if (e) e.preventDefault();
            if (!state.docxFilePath) {
                showToast('No hay ninguna plantilla Word cargada en este proyecto', 'error');
                return;
            }
            try {
                const res = await fetch('/api/open-file', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filepath: state.docxFilePath })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'No se pudo abrir la plantilla Word');
                showToast('Abriendo plantilla Word en aplicación nativa...', 'info');
            } catch (err) {
                showToast(err.message || 'Error de conexión al abrir Word', 'error');
            }
        });
    }

    if (btnOpenProjectFolder) {
        btnOpenProjectFolder.addEventListener('click', async (e) => {
            if (e) e.preventDefault();
            let folder = state.projectFolderPath;
            if (!folder && state.excelFilePath) {
                const idx = Math.max(state.excelFilePath.lastIndexOf('\\'), state.excelFilePath.lastIndexOf('/'));
                if (idx !== -1) folder = state.excelFilePath.substring(0, idx);
            }
            if (!folder) {
                showToast('No hay una carpeta de proyecto activa', 'error');
                return;
            }
            try {
                const res = await fetch('/api/open-output-folder', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ folder_path: folder })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'No se pudo abrir la carpeta');
                showToast('Abriendo carpeta raíz del proyecto...', 'info');
            } catch (err) {
                showToast(err.message || 'Error al abrir la carpeta del proyecto', 'error');
            }
        });
    }

    let currentProjectNameToDelete = '';

    if (btnDeleteProject) {
        btnDeleteProject.addEventListener('click', () => {
            if (!state.activeProjectId) {
                showToast('No hay ningún proyecto activo seleccionado', 'error');
                return;
            }

            const activeOpt = projectSelectDropdown.options[projectSelectDropdown.selectedIndex];
            currentProjectNameToDelete = activeOpt ? activeOpt.textContent.trim() : 'Proyecto';

            if (deleteProjTargetName) deleteProjTargetName.textContent = currentProjectNameToDelete;
            if (deleteProjConfirmNameInput) deleteProjConfirmNameInput.value = '';
            if (btnConfirmDeleteProj) btnConfirmDeleteProj.disabled = true;

            showModal(deleteProjectModal);
        });
    }

    if (deleteProjConfirmNameInput) {
        deleteProjConfirmNameInput.addEventListener('input', (e) => {
            const val = e.target.value.trim();
            if (btnConfirmDeleteProj) {
                btnConfirmDeleteProj.disabled = (val !== currentProjectNameToDelete);
            }
        });
    }

    if (btnCloseDeleteProjModal) btnCloseDeleteProjModal.addEventListener('click', () => deleteProjectModal.classList.add('hidden'));
    if (btnCancelDeleteProj) btnCancelDeleteProj.addEventListener('click', () => deleteProjectModal.classList.add('hidden'));

    if (btnConfirmDeleteProj) {
        btnConfirmDeleteProj.addEventListener('click', async () => {
            if (!state.activeProjectId) return;
            if (deleteProjectModal) deleteProjectModal.classList.add('hidden');

            showToast('Cerrando archivos y eliminando proyecto...', 'info');

            try {
                const res = await fetch('/api/delete-project', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ project_id: state.activeProjectId })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error eliminando el proyecto');

                state.activeProjectId = null;
                state.excelFilePath = null;
                state.docxFilePath = null;
                state.docxTemplates = [];

                if (btnDeleteProject) btnDeleteProject.classList.add('hidden');
                if (btnOpenProjectFolder) btnOpenProjectFolder.classList.add('hidden');

                await loadProjectsList();
                showToast('Proyecto y sus archivos eliminados exitosamente', 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    async function refreshExcelDataFromDisk(isAuto = false) {
        if (!state.excelFilePath || state.isAutoSyncing) return;
        state.isAutoSyncing = true;
        if (!isAuto) showToast('Refrescando datos desde el archivo Excel...', 'info');

        try {
            const res = await fetch('/api/reload-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: state.excelFilePath })
            });
            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error al recargar Excel');

            state.excelColumns = data.columns;
            state.excelRecords = data.all_records || data.records || [];
            excelRowsCount.textContent = `${data.total_rows} registros cargados`;
            excelStatusPill.querySelector('span').textContent = `${data.columns.length} Cols`;

            renderVariableChips(data.columns);
            renderEmailMappableVars();
            updateGenerateButtonState();

            if (excelModal && !excelModal.classList.contains('hidden')) {
                renderExcelGridTable();
            }
            if (recipientsSelectionModal && !recipientsSelectionModal.classList.contains('hidden')) {
                renderRecipientsList();
            }

            const stRes = await fetch('/api/check-excel-status', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: state.excelFilePath })
            });
            const stData = await stRes.json();
            if (stData.exists) state.excelMtime = stData.mtime;

            if (isAuto) {
                showToast('Datos de Excel sincronizados automáticamente', 'info');
            } else {
                showToast('Tabla e indicadores actualizados desde Excel', 'success');
            }
        } catch (e) {
            if (!isAuto) showToast(e.message, 'error');
        } finally {
            state.isAutoSyncing = false;
        }
    }

    // AUTOMATIC REAL-TIME EXCEL FILE WATCHER
    setInterval(async () => {
        if (!state.excelFilePath || state.isAutoSyncing) return;
        try {
            const res = await fetch('/api/check-excel-status', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: state.excelFilePath })
            });
            const data = await res.json();
            if (data.exists && data.mtime) {
                if (state.excelMtime === 0) {
                    state.excelMtime = data.mtime;
                } else if (data.mtime > state.excelMtime) {
                    state.excelMtime = data.mtime;
                    await refreshExcelDataFromDisk(true);
                }
            }
        } catch (e) {}
    }, 1500);

    window.addEventListener('focus', async () => {
        if (!state.excelFilePath || state.isAutoSyncing) return;
        try {
            const res = await fetch('/api/check-excel-status', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: state.excelFilePath })
            });
            const data = await res.json();
            if (data.exists && data.mtime && data.mtime > state.excelMtime) {
                state.excelMtime = data.mtime;
                await refreshExcelDataFromDisk(true);
            }
        } catch (e) {}
    });

    if (pdfVariationSelect) {
        pdfVariationSelect.addEventListener('change', (e) => {
            pdfViewerFrame.src = e.target.value;
        });
    }

    // OPEN PARAGRAPH RULES MODAL DIRECTLY FROM STEP 2
    if (btnOpenParagraphRules) {
        btnOpenParagraphRules.addEventListener('click', () => {
            renderParagraphRules();
            showModal(paragraphRulesModal);
        });
    }

    if (btnCloseParagraphModal) btnCloseParagraphModal.addEventListener('click', () => paragraphRulesModal.classList.add('hidden'));
    if (btnCancelParagraphRules) btnCancelParagraphRules.addEventListener('click', () => paragraphRulesModal.classList.add('hidden'));

    if (btnSaveParagraphRules) {
        btnSaveParagraphRules.addEventListener('click', async () => {
            if (paragraphRulesModal) paragraphRulesModal.classList.add('hidden');
            await saveCurrentProjectSettings();
            renderVariableChips(state.excelColumns);
        });
    }

    async function saveCurrentProjectSettings() {
        if (!state.activeProjectId) return;
        try {
            const res = await fetch('/api/save-project-settings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    project_id: state.activeProjectId,
                    settings: state.settings
                })
            });
            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error guardando ajustes');
            showToast('Ajustes guardados exitosamente en settings.json del proyecto', 'success');
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    // TOGGLE VIEWWORD VS VIEWPDF
    if (btnViewWord) {
        btnViewWord.addEventListener('click', () => {
            state.activeView = 'word';
            btnViewWord.classList.add('active');
            btnViewPdf.classList.remove('active');
            paperPage.classList.remove('hidden');
            pdfViewerWrapper.classList.add('hidden');
            if (btnToggleBadges) btnToggleBadges.classList.remove('hidden');
            if (docTypeIcon) {
                docTypeIcon.classList.remove('pdf-mode');
                docTypeIcon.innerHTML = '<i class="fa-solid fa-file-word"></i>';
            }
        });
    }

    if (btnViewPdf) {
        btnViewPdf.addEventListener('click', async () => {
            state.activeView = 'pdf';
            loadPdfPreviewForActiveTemplate(false);
        });
    }

    if (btnSaveAppToWord) {
        btnSaveAppToWord.addEventListener('click', async () => {
            const currentDocx = (state.docxTemplates && state.docxTemplates[state.activeDocxIndex]) ? state.docxTemplates[state.activeDocxIndex].filepath : state.docxFilePath;
            if (!currentDocx) {
                showToast('No hay ninguna plantilla Word cargada para guardar', 'error');
                return;
            }
            showToast('Guardando cambios en el archivo Word en disco...', 'info');
            try {
                const res = await fetch('/api/save-docx-content', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        docx_path: currentDocx,
                        html: paperPage.innerHTML
                    })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error al guardar cambios en Word');
                showToast('Cambios guardados exitosamente del Editor de la App al archivo Word (.docx)', 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    if (btnSyncWordToApp) {
        btnSyncWordToApp.addEventListener('click', async () => {
            const currentDocx = (state.docxTemplates && state.docxTemplates[state.activeDocxIndex]) ? state.docxTemplates[state.activeDocxIndex].filepath : state.docxFilePath;
            if (!currentDocx) {
                showToast('No hay ninguna plantilla Word cargada para recargar', 'error');
                return;
            }
            showToast('Recargando contenido Word desde disco...', 'info');
            try {
                const res = await fetch('/api/reload-docx', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ docx_path: currentDocx })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error al recargar el archivo Word');

                if (state.docxTemplates && state.docxTemplates[state.activeDocxIndex]) {
                    state.docxTemplates[state.activeDocxIndex].html = data.html;
                }
                renderHtmlInPaper(data.html);
                showToast('Contenido recargado exitosamente del archivo Word (.docx) a la App', 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    if (btnOpenTemplateRulesTrigger) {
        btnOpenTemplateRulesTrigger.addEventListener('click', () => {
            renderTemplateRulesList();
            showModal(templateRulesModal);
        });
    }

    async function loadPdfPreviewForActiveTemplate(forceRefresh = false) {
        if (!state.docxFilePath) {
            showToast('Primero debes cargar una plantilla Word', 'error');
            return;
        }

        btnViewPdf.classList.add('active');
        btnViewWord.classList.remove('active');

        // Mostrar overlay de carga y el wrapper del PDF
        paperPage.classList.add('hidden');
        pdfViewerWrapper.classList.remove('hidden');
        if (pdfLoadingOverlay) pdfLoadingOverlay.classList.remove('hidden');
        if (btnToggleBadges) btnToggleBadges.classList.add('hidden');
        if (docTypeIcon) {
            docTypeIcon.classList.add('pdf-mode');
            docTypeIcon.innerHTML = '<i class="fa-solid fa-file-pdf"></i>';
        }

        try {
            const res = await fetch('/api/get-pdf-preview', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    docx_path: state.docxFilePath,
                    excel_path: state.excelFilePath,
                    settings: state.settings,
                    force_refresh: forceRefresh
                })
            });

            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error generando PDF');

            pdfViewerFrame.src = data.pdf_url;
            // Ocultar overlay de carga cuando el PDF esté listo
            if (pdfLoadingOverlay) pdfLoadingOverlay.classList.add('hidden');

            showToast(data.cached ? 'Vista previa PDF cargada desde caché' : 'Vista previa PDF generada', 'success');
        } catch (err) {
            if (pdfLoadingOverlay) pdfLoadingOverlay.classList.add('hidden');
            showToast(err.message, 'error');
            btnViewWord.click();
        }
    }

    // SETTINGS MODAL & OUTLOOK LOGIC
    let initialSettingsSnapshot = null;

    function captureCurrentSettingsSnapshot() {
        return JSON.stringify({
            output_format: settingOutputFormat ? settingOutputFormat.value : '',
            outlook_enabled: settingOutlookEnabled ? settingOutlookEnabled.checked : false,
            email_body_source: settingEmailBodySource ? settingEmailBodySource.value : 'text',
            email_body_template: settingEmailBodyTemplate ? settingEmailBodyTemplate.value : '',
            doc_check_enabled: settingDocCheckEnabled ? settingDocCheckEnabled.checked : false,
            doc_check_column: settingDocCheckCol ? settingDocCheckCol.value : '',
            doc_check_value_mode: settingDocCheckMode ? settingDocCheckMode.value : '',
            doc_check_custom_text: settingDocCheckCustomText ? settingDocCheckCustomText.value : '',
            template_rules_enabled: settingTemplateRulesEnabled ? settingTemplateRulesEnabled.checked : false,
            email_column: settingEmailCol ? settingEmailCol.value : '',
            check_column: settingCheckCol ? settingCheckCol.value : '',
            send_mode: settingSendMode ? settingSendMode.value : '',
            attachment_type: settingAttachmentType ? settingAttachmentType.value : '',
            email_subject: settingEmailSubject ? settingEmailSubject.value : '',
            email_body: settingEmailBody ? settingEmailBody.value : ''
        });
    }

    function updateSaveSettingsButtonState() {
        if (!btnSaveSettings || !initialSettingsSnapshot) return;
        const currentSnap = captureCurrentSettingsSnapshot();
        btnSaveSettings.disabled = (currentSnap === initialSettingsSnapshot);
    }

    if (settingsModal) {
        settingsModal.addEventListener('input', updateSaveSettingsButtonState);
        settingsModal.addEventListener('change', updateSaveSettingsButtonState);
    }

    if (btnOpenSettings) {
        btnOpenSettings.addEventListener('click', () => {
            populateSettingsFields();
            initialSettingsSnapshot = captureCurrentSettingsSnapshot();
            if (btnSaveSettings) btnSaveSettings.disabled = true;
            showModal(settingsModal);
        });
    }

    if (btnCloseSettingsModal) {
        btnCloseSettingsModal.addEventListener('click', (e) => {
            if (e) e.preventDefault();
            settingsModal.classList.add('hidden');
        });
    }

    if (btnCancelSettings) {
        btnCancelSettings.addEventListener('click', (e) => {
            if (e) e.preventDefault();
            settingsModal.classList.add('hidden');
        });
    }

    if (settingOutlookEnabled) {
        settingOutlookEnabled.addEventListener('change', (e) => {
            state.settings.outlook_enabled = e.target.checked;
            if (e.target.checked) {
                if (outlookFieldsContainer) outlookFieldsContainer.classList.remove('hidden');
            } else {
                if (outlookFieldsContainer) outlookFieldsContainer.classList.add('hidden');
            }
            updateGenerateButtonState();
        });
    }

    if (settingEmailBodySource) {
        settingEmailBodySource.addEventListener('change', (e) => {
            const val = e.target.value;
            if (val === 'text') {
                if (outlookTextBodyContainer) outlookTextBodyContainer.classList.remove('hidden');
                if (outlookTemplateBodyContainer) outlookTemplateBodyContainer.classList.add('hidden');
            } else if (val === 'template') {
                if (outlookTextBodyContainer) outlookTextBodyContainer.classList.add('hidden');
                if (outlookTemplateBodyContainer) outlookTemplateBodyContainer.classList.remove('hidden');
            }
        });
    }

    if (settingDocCheckEnabled) {
        settingDocCheckEnabled.addEventListener('change', (e) => {
            if (e.target.checked) {
                if (docCheckOptionsContainer) docCheckOptionsContainer.classList.remove('hidden');
            } else {
                if (docCheckOptionsContainer) docCheckOptionsContainer.classList.add('hidden');
            }
        });
    }

    if (settingDocCheckMode) {
        settingDocCheckMode.addEventListener('change', (e) => {
            if (e.target.value === 'custom_text') {
                if (docCheckCustomTextContainer) docCheckCustomTextContainer.classList.remove('hidden');
            } else {
                if (docCheckCustomTextContainer) docCheckCustomTextContainer.classList.add('hidden');
            }
        });
    }

    if (settingTemplateRulesEnabled) {
        settingTemplateRulesEnabled.addEventListener('change', (e) => {
            if (e.target.checked) {
                if (templateRulesOptionsContainer) templateRulesOptionsContainer.classList.remove('hidden');
            } else {
                if (templateRulesOptionsContainer) templateRulesOptionsContainer.classList.add('hidden');
            }
        });
    }

    function populateSettingsFields() {
        settingOutputFormat.value = state.settings.output_format || 'docx';
        settingOutlookEnabled.checked = state.settings.outlook_enabled || false;
        
        if (settingOutlookEnabled.checked) {
            if (outlookFieldsContainer) outlookFieldsContainer.classList.remove('hidden');
        } else {
            if (outlookFieldsContainer) outlookFieldsContainer.classList.add('hidden');
        }

        if (settingEmailBodySource) {
            settingEmailBodySource.value = state.settings.email_body_source || 'text';
            if (settingEmailBodySource.value === 'text') {
                if (outlookTextBodyContainer) outlookTextBodyContainer.classList.remove('hidden');
                if (outlookTemplateBodyContainer) outlookTemplateBodyContainer.classList.add('hidden');
            } else if (settingEmailBodySource.value === 'template') {
                if (outlookTextBodyContainer) outlookTextBodyContainer.classList.add('hidden');
                if (outlookTemplateBodyContainer) outlookTemplateBodyContainer.classList.remove('hidden');
            }
        }

        if (settingEmailBodyTemplate) {
            settingEmailBodyTemplate.innerHTML = '';
            if (state.docxTemplates && state.docxTemplates.length > 0) {
                state.docxTemplates.forEach(tpl => {
                    const opt = document.createElement('option');
                    opt.value = tpl.filepath;
                    opt.textContent = tpl.filename;
                    if (tpl.filepath === state.settings.email_body_template) opt.selected = true;
                    settingEmailBodyTemplate.appendChild(opt);
                });
            } else {
                const opt = document.createElement('option');
                opt.value = '';
                opt.textContent = '-- Carga una plantilla Word en el proyecto --';
                settingEmailBodyTemplate.appendChild(opt);
            }
        }

        if (settingDocCheckEnabled) {
            settingDocCheckEnabled.checked = state.settings.doc_check_enabled !== false;
            if (settingDocCheckEnabled.checked) {
                if (docCheckOptionsContainer) docCheckOptionsContainer.classList.remove('hidden');
            } else {
                if (docCheckOptionsContainer) docCheckOptionsContainer.classList.add('hidden');
            }
        }

        if (settingDocCheckCol) {
            settingDocCheckCol.innerHTML = '';
            const defaultCheckCols = ['Doc_Generado', 'Estado_Generación', 'Estado'];
            const allCols = Array.from(new Set([...state.excelColumns, ...defaultCheckCols]));
            allCols.forEach(c => {
                const opt = document.createElement('option');
                opt.value = c;
                opt.textContent = c;
                if (c === (state.settings.doc_check_column || 'Doc_Generado')) opt.selected = true;
                settingDocCheckCol.appendChild(opt);
            });
        }

        if (settingDocCheckMode) {
            settingDocCheckMode.value = state.settings.doc_check_value_mode || 'timestamp';
            if (settingDocCheckMode.value === 'custom_text') {
                if (docCheckCustomTextContainer) docCheckCustomTextContainer.classList.remove('hidden');
            } else {
                if (docCheckCustomTextContainer) docCheckCustomTextContainer.classList.add('hidden');
            }
        }

        if (settingDocCheckCustomText) {
            settingDocCheckCustomText.value = state.settings.doc_check_custom_text || 'Generado';
        }

        if (settingTemplateRulesEnabled) {
            settingTemplateRulesEnabled.checked = state.settings.template_rules_enabled || false;
            if (settingTemplateRulesEnabled.checked) {
                if (templateRulesOptionsContainer) templateRulesOptionsContainer.classList.remove('hidden');
            } else {
                if (templateRulesOptionsContainer) templateRulesOptionsContainer.classList.add('hidden');
            }
        }

        settingEmailCol.innerHTML = '';
        settingCheckCol.innerHTML = '';

        state.excelColumns.forEach(c => {
            const opt1 = document.createElement('option');
            opt1.value = c;
            opt1.textContent = c;
            if (c === state.settings.email_column) opt1.selected = true;
            settingEmailCol.appendChild(opt1);

            const opt2 = document.createElement('option');
            opt2.value = c;
            opt2.textContent = c;
            if (c === state.settings.check_column) opt2.selected = true;
            settingCheckCol.appendChild(opt2);
        });

        settingSendMode.value = state.settings.send_mode || 'display';
        settingAttachmentType.value = state.settings.attachment_type || 'pdf';
        settingEmailSubject.value = state.settings.email_subject || 'Notificación Oficial - {{ Nombre }}';
        settingEmailBody.value = state.settings.email_body || '';

        renderEmailMappableVars();
    }

    function renderEmailMappableVars() {
        if (!emailVarsList) return;
        emailVarsList.innerHTML = '';

        if (!state.excelColumns || state.excelColumns.length === 0) {
            emailVarsList.innerHTML = '<span class="meta">Carga un Excel para ver las variables disponibles.</span>';
            return;
        }

        state.excelColumns.forEach(col => {
            const chip = document.createElement('span');
            chip.className = 'email-var-chip';
            chip.textContent = `{{ ${col} }}`;
            chip.title = "Haz clic para insertar en el campo de texto enfocado";

            chip.addEventListener('click', () => {
                const activeEl = document.activeElement;
                const tag = `{{ ${col} }}`;

                if (activeEl && (activeEl.id === 'setting-email-subject' || activeEl.id === 'setting-email-body')) {
                    const start = activeEl.selectionStart || 0;
                    const end = activeEl.selectionEnd || 0;
                    const val = activeEl.value;
                    activeEl.value = val.substring(0, start) + tag + val.substring(end);
                    activeEl.focus();
                    activeEl.selectionStart = activeEl.selectionEnd = start + tag.length;
                } else {
                    settingEmailBody.value += ` ${tag}`;
                    settingEmailBody.focus();
                }
            });

            emailVarsList.appendChild(chip);
        });
    }

    // CONDITIONAL PARAGRAPH RULES LOGIC
    if (btnAddParagraphRule) {
        btnAddParagraphRule.addEventListener('click', () => {
            const defaultCol = state.excelColumns[0] || 'Estado';
            if (!state.settings.custom_paragraph_rules) state.settings.custom_paragraph_rules = [];
            state.settings.custom_paragraph_rules.push({
                column: defaultCol,
                value: '',
                custom_text: ''
            });
            renderParagraphRules();
        });
    }

    function renderParagraphRules() {
        paragraphRulesContainer.innerHTML = '';

        if (!state.settings.custom_paragraph_rules || state.settings.custom_paragraph_rules.length === 0) {
            paragraphRulesContainer.innerHTML = '<div class="empty-state p-2"><p class="meta">Sin reglas de párrafos personalizadas configuradas.</p></div>';
            return;
        }

        state.settings.custom_paragraph_rules.forEach((rule, idx) => {
            const card = document.createElement('div');
            card.className = 'rule-card';

            let optionsHtml = '';
            state.excelColumns.forEach(c => {
                optionsHtml += `<option value="${escapeHtml(c)}" ${c === rule.column ? 'selected' : ''}>${escapeHtml(c)}</option>`;
            });

            card.innerHTML = `
                <div class="flex-justify-between">
                    <span class="filename">Regla #${idx + 1}</span>
                    <button class="btn-icon-danger btn-delete-rule" data-rule-idx="${idx}"><i class="fa-solid fa-trash-can"></i></button>
                </div>
                <div class="grid-2col">
                    <div class="form-group">
                        <label>Si Columna en Excel:</label>
                        <select class="form-control rule-col-select" data-rule-idx="${idx}">
                            ${optionsHtml}
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Es igual al valor:</label>
                        <input type="text" class="form-control rule-val-input" data-rule-idx="${idx}" value="${escapeHtml(rule.value)}" placeholder="ej. Aprobado">
                    </div>
                </div>
                <div class="form-group">
                    <label>Imprimir Párrafo Personalizado (Soporta {{ Variables }}):</label>
                    <textarea class="form-control text-area rule-text-input" data-rule-idx="${idx}" rows="2" placeholder="ej. Estimado/a {{ Nombre }}, le confirmamos que...">${escapeHtml(rule.custom_text)}</textarea>
                </div>
            `;

            paragraphRulesContainer.appendChild(card);
        });

        paragraphRulesContainer.querySelectorAll('.btn-delete-rule').forEach(btn => {
            btn.addEventListener('click', () => {
                const idx = parseInt(btn.dataset.ruleIdx, 10);
                state.settings.custom_paragraph_rules.splice(idx, 1);
                renderParagraphRules();
            });
        });

        paragraphRulesContainer.querySelectorAll('.rule-col-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.custom_paragraph_rules[idx].column = e.target.value;
            });
        });

        paragraphRulesContainer.querySelectorAll('.rule-val-input').forEach(inp => {
            inp.addEventListener('input', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.custom_paragraph_rules[idx].value = e.target.value;
            });
        });

        paragraphRulesContainer.querySelectorAll('.rule-text-input').forEach(txt => {
            txt.addEventListener('input', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.custom_paragraph_rules[idx].custom_text = e.target.value;
            });
        });
    }

    if (btnSaveSettings) {
        btnSaveSettings.addEventListener('click', async () => {
            state.settings.output_format = settingOutputFormat.value;
            if (patternInput) state.settings.output_filename_pattern = patternInput.value.trim() || '{{ index }}_documento.docx';
            if (settingDocCheckEnabled) state.settings.doc_check_enabled = settingDocCheckEnabled.checked;
            if (settingDocCheckCol) state.settings.doc_check_column = settingDocCheckCol.value;
            if (settingDocCheckMode) state.settings.doc_check_value_mode = settingDocCheckMode.value;
            if (settingDocCheckCustomText) state.settings.doc_check_custom_text = settingDocCheckCustomText.value.trim() || 'Generado';
            if (settingTemplateRulesEnabled) state.settings.template_rules_enabled = settingTemplateRulesEnabled.checked;

            if (settingEmailBodySource) state.settings.email_body_source = settingEmailBodySource.value;
            if (settingEmailBodyTemplate) state.settings.email_body_template = settingEmailBodyTemplate.value;
            state.settings.outlook_enabled = settingOutlookEnabled.checked;
            state.settings.email_column = settingEmailCol.value;
            state.settings.check_column = settingCheckCol.value;
            state.settings.send_mode = settingSendMode.value;
            state.settings.attachment_type = settingAttachmentType.value;
            state.settings.email_subject = settingEmailSubject.value;
            state.settings.email_body = settingEmailBody.value;

            settingsModal.classList.add('hidden');
            await saveCurrentProjectSettings();
            renderVariableChips(state.excelColumns);
        });
    }

    // CREATE PROJECT MODAL LOGIC
    if (btnOpenCreateProject) {
        btnOpenCreateProject.addEventListener('click', () => {
            projNameInput.value = '';
            projPathInput.value = '';
            showModal(createProjectModal);
        });
    }

    if (btnCloseCreateProjModal) btnCloseCreateProjModal.addEventListener('click', () => createProjectModal.classList.add('hidden'));
    if (btnCancelCreateProj) btnCancelCreateProj.addEventListener('click', () => createProjectModal.classList.add('hidden'));

    if (btnConfirmCreateProj) {
        btnConfirmCreateProj.addEventListener('click', async () => {
            const name = projNameInput.value.trim();
            if (!name) {
                showToast('Ingresa un nombre para el proyecto', 'error');
                return;
            }

            showToast('Creando carpeta de correspondencia...', 'info');

            try {
                const res = await fetch('/api/create-project', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        name: name,
                        parent_dir: projPathInput.value.trim() || null
                    })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error creando proyecto');

                createProjectModal.classList.add('hidden');
                await loadProjectsList();
                await loadSelectedProject(data.project.id);
                showToast(`Carpeta '${name}' creada exitosamente`, 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    // SAVE DOCX TEMPLATE CONTENT
    if (btnSaveDocx) {
        btnSaveDocx.addEventListener('click', async () => {
            if (!state.docxFilePath) return;

            showToast('Guardando plantilla Word en disco...', 'info');

            try {
                const res = await fetch('/api/save-docx-content', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        docx_path: state.docxFilePath,
                        content_html: paperPage.innerHTML
                    })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error guardando plantilla');

                if (state.docxTemplates[state.activeDocxIndex]) {
                    state.docxTemplates[state.activeDocxIndex].html = paperPage.innerHTML;
                }

                showToast('Plantilla Word guardada exitosamente', 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    // DROPZONE SETUP
    setupDropzone(excelDropzone, excelInput, handleExcelUpload);
    setupDropzone(docxDropzone, docxInput, handleDocxUpload);

    function setupDropzone(zone, input, handler) {
        ['dragenter', 'dragover'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.add('drag-over');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.remove('drag-over');
            });
        });

        zone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) handler(files[0]);
        });

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handler(e.target.files[0]);
        });
    }

    // EXCEL UPLOAD HANDLER
    async function handleExcelUpload(file) {
        const formData = new FormData();
        formData.append('file', file);
        if (state.activeProjectId) {
            formData.append('project_id', state.activeProjectId);
        }

        showToast('Leyendo archivo Excel...', 'info');

        try {
            const res = await fetch('/api/upload-excel', {
                method: 'POST',
                body: formData
            });

            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error al procesar Excel');

            state.excelFilePath = data.filepath;
            state.excelFileName = data.filename;
            state.excelColumns = data.columns;
            state.excelRecords = data.all_records || data.preview || [];

            excelDropzone.classList.add('hidden');
            excelInfo.classList.remove('hidden');
            btnViewExcel.classList.remove('hidden');
            if (btnOpenNativeExcel) btnOpenNativeExcel.classList.remove('hidden');
            excelFilename.textContent = data.filename;
            excelRowsCount.textContent = `${data.total_rows} registros encontrados`;
            excelStatusPill.classList.remove('disabled');
            excelStatusPill.classList.add('active');
            excelStatusPill.querySelector('span').textContent = `${data.columns.length} Cols`;

            renderVariableChips(data.columns);
            renderEmailMappableVars();
            updateGenerateButtonState();
            showToast(`Excel cargado: ${data.columns.length} columnas`, 'success');
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    // RENDER DRAGGABLE CHIPS
    function renderVariableChips(columns) {
        variablesContainer.innerHTML = '';

        if (!columns || columns.length === 0) {
            varsCount.textContent = '0';
            variablesContainer.innerHTML = `
                <div class="empty-state">
                    <i class="fa-solid fa-ghost"></i>
                    <p>No se encontraron columnas en el archivo.</p>
                </div>`;
            renderSidebarRules();
            return;
        }

        let totalVars = 0;
        const activeRules = state.settings.custom_paragraph_rules || [];

        // Renderizar variables directas. Si la columna tiene regla asociada, se tiñe de NARANJA (chip-orange)
        columns.forEach(col => {
            totalVars++;
            const colClean = String(col).trim();
            const hasRule = activeRules.some(r => String(r.column || '').trim().toLowerCase() === colClean.toLowerCase() && r.value && r.custom_text);

            const chip = document.createElement('div');
            chip.className = `variable-chip ${hasRule ? 'chip-orange' : ''}`;
            chip.draggable = true;
            chip.dataset.varName = colClean;

            const iconClass = hasRule ? 'fa-paragraph' : 'fa-grip-vertical';
            const badgeClass = hasRule ? 'docx-tag-badge docx-tag-badge-orange' : 'docx-tag-badge';

            chip.innerHTML = `
                <i class="fa-solid ${iconClass} drag-handle"></i>
                <span class="chip-name">${escapeHtml(colClean)}</span>
                <span class="chip-tag">{{ ${escapeHtml(colClean)} }}</span>
            `;

            chip.addEventListener('dragstart', (e) => {
                const tag = `{{ ${colClean} }}`;
                e.dataTransfer.setData('text/plain', tag);
                e.dataTransfer.setData('text/html', `&nbsp;<span class="${badgeClass}" contenteditable="false">${tag}</span>&nbsp;`);
                chip.style.opacity = '0.5';
            });

            chip.addEventListener('dragend', () => {
                chip.style.opacity = '1';
            });

            chip.addEventListener('click', () => {
                insertTagInPaper(`{{ ${colClean} }}`);
            });

            variablesContainer.appendChild(chip);
        });

        varsCount.textContent = totalVars;
        renderSidebarRules();
    }

    function renderSidebarRules() {
        const container = document.getElementById('sidebar-rules-container');
        if (!container) return;
        container.innerHTML = '';

        const rules = state.settings.custom_paragraph_rules || [];
        if (rules.length === 0) {
            container.innerHTML = '<span class="meta fs-11">Sin reglas de párrafos configuradas.</span>';
            return;
        }

        rules.forEach((rule, idx) => {
            const item = document.createElement('div');
            item.className = 'sidebar-rule-item';
            item.innerHTML = `
                <div class="sidebar-rule-header">
                    <span>{{ ${escapeHtml(rule.column)} }} = "${escapeHtml(rule.value)}"</span>
                    <button class="btn-icon-danger btn-delete-sidebar-rule" data-rule-idx="${idx}" title="Eliminar regla"><i class="fa-solid fa-xmark"></i></button>
                </div>
                <div class="sidebar-rule-preview" title="${escapeHtml(rule.custom_text)}">
                    <i class="fa-solid fa-file-pen"></i> ${escapeHtml(rule.custom_text)}
                </div>
            `;
            container.appendChild(item);
        });

        container.querySelectorAll('.btn-delete-sidebar-rule').forEach(btn => {
            btn.addEventListener('click', async () => {
                const idx = parseInt(btn.dataset.ruleIdx, 10);
                state.settings.custom_paragraph_rules.splice(idx, 1);
                await saveCurrentProjectSettings();
                renderVariableChips(state.excelColumns);
            });
        });
    }

    // DOCX UPLOAD HANDLER (ADDS TO N TEMPLATES IF PROJECT ACTIVE)
    async function handleDocxUpload(file) {
        const formData = new FormData();
        formData.append('file', file);

        showToast('Convirtiendo plantilla DOCX...', 'info');

        try {
            let res, data;
            if (state.activeProjectId) {
                formData.append('project_id', state.activeProjectId);
                res = await fetch('/api/add-docx-template', {
                    method: 'POST',
                    body: formData
                });
                data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error agregando plantilla');

                state.docxTemplates.push({
                    filepath: data.filepath,
                    filename: data.filename,
                    html: data.html
                });
                selectActiveDocxTemplate(state.docxTemplates.length - 1);
            } else {
                res = await fetch('/api/upload-docx', {
                    method: 'POST',
                    body: formData
                });
                data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error al convertir DOCX');

                state.docxFilePath = data.filepath;
                state.docxFileName = data.filename;
                state.docxTemplates = [{ filepath: data.filepath, filename: data.filename, html: data.html }];
                state.activeDocxIndex = 0;
                docTitleDisplay.textContent = data.filename;
                renderHtmlInPaper(data.html);
                renderDocxTemplatesList();
            }

            if (saveDocxSplitGroup) saveDocxSplitGroup.classList.remove('hidden');
            if (btnSaveDocx) btnSaveDocx.classList.remove('hidden');
            docxStatusPill.classList.remove('disabled');
            docxStatusPill.classList.add('active');
            docxStatusPill.querySelector('span').textContent = `${state.docxTemplates.length} Plantilla(s)`;

            updateGenerateButtonState();
            showToast('Plantilla Word agregada exitosamente', 'success');
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    // RENDER HTML IN PAPER WITH ATOMIC NON-EDITABLE BADGES
    function renderHtmlInPaper(htmlContent) {
        if (!htmlContent.trim()) {
            paperPage.innerHTML = '<p>El documento está vacío.</p>';
            return;
        }

        const cleanHtml = htmlContent.replace(/&nbsp;/g, ' ');

        const formattedHtml = cleanHtml.replace(/\{\{\s*([a-zA-Z0-9_\-áéíóúÁÉÍÓÚñÑ\s]+)\s*\}\}/g, (match, p1) => {
            const rawTag = p1.trim();
            const badgeClass = rawTag.startsWith('p_') ? 'docx-tag-badge docx-tag-badge-orange' : 'docx-tag-badge';
            return ` <span class="${badgeClass}" contenteditable="false">{{ ${rawTag} }}</span> `;
        });

        paperPage.innerHTML = formattedHtml;
    }

    let isPaperDragSetupDone = false;
    function setupPaperDragAndDrop() {
        if (isPaperDragSetupDone || !paperPage) return;
        isPaperDragSetupDone = true;
        paperPage.contentEditable = 'true';

        paperPage.addEventListener('dragover', (e) => {
            e.preventDefault();
            paperPage.classList.add('drag-active');
        });

        paperPage.addEventListener('dragleave', () => {
            paperPage.classList.remove('drag-active');
        });

        paperPage.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            paperPage.classList.remove('drag-active');

            const tagText = e.dataTransfer.getData('text/plain');
            if (tagText && tagText.startsWith('{{')) {
                insertTagAtDropPoint(e, tagText);
            }
        });

        paperPage.addEventListener('click', (e) => {
            if (e.target.classList.contains('docx-tag-badge')) {
                const sel = window.getSelection();
                const range = document.createRange();
                let next = e.target.nextSibling;
                if (!next || next.nodeType !== Node.TEXT_NODE) {
                    next = document.createTextNode('\u00A0');
                    e.target.parentNode.insertBefore(next, e.target.nextSibling);
                }
                range.setStart(next, 0);
                range.setCollapse(true);
                sel.removeAllRanges();
                sel.addRange(range);
            }
        });
    }

    function createTagSpan(tagText) {
        const span = document.createElement('span');
        const raw = tagText.replace(/[\{\}\s]/g, '');
        span.className = raw.startsWith('p_') ? 'docx-tag-badge docx-tag-badge-orange' : 'docx-tag-badge';
        span.contentEditable = 'false';
        span.textContent = tagText;
        return span;
    }

    function insertTagAtDropPoint(e, tagText) {
        let range;
        if (document.caretRangeFromPoint) {
            range = document.caretRangeFromPoint(e.clientX, e.clientY);
        } else if (e.rangeParent) {
            range = document.createRange();
            range.setStart(e.rangeParent, e.rangeOffset);
        }

        const span = createTagSpan(tagText);
        const spaceBefore = document.createTextNode('\u00A0');
        const spaceAfter = document.createTextNode('\u00A0');

        if (range) {
            range.insertNode(spaceAfter);
            range.insertNode(span);
            range.insertNode(spaceBefore);

            const sel = window.getSelection();
            const newRange = document.createRange();
            newRange.setStart(spaceAfter, 1);
            newRange.setCollapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
        } else {
            paperPage.appendChild(spaceBefore);
            paperPage.appendChild(span);
            paperPage.appendChild(spaceAfter);
        }
    }

    function insertTagInPaper(tagText) {
        paperPage.focus();
        const sel = window.getSelection();
        const span = createTagSpan(tagText);
        const spaceBefore = document.createTextNode('\u00A0');
        const spaceAfter = document.createTextNode('\u00A0');

        if (sel.rangeCount > 0 && paperPage.contains(sel.getRangeAt(0).commonAncestorContainer)) {
            const range = sel.getRangeAt(0);
            range.deleteContents();
            range.insertNode(spaceAfter);
            range.insertNode(span);
            range.insertNode(spaceBefore);

            const newRange = document.createRange();
            newRange.setStart(spaceAfter, 1);
            newRange.setCollapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
        } else {
            paperPage.appendChild(spaceBefore);
            paperPage.appendChild(span);
            paperPage.appendChild(spaceAfter);

            const newRange = document.createRange();
            newRange.setStart(spaceAfter, 1);
            newRange.setCollapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
        }
    }

    // REMOVE / UNLINK FILES
    if (btnRemoveExcel) {
        btnRemoveExcel.addEventListener('click', () => {
            const fn = state.excelFileName || 'datos.xlsx';
            showInAppDeleteConfirm(`¿Estás seguro de que deseas desvincular el archivo Excel <strong>${escapeHtml(fn)}</strong> del proyecto?`, async () => {
                if (state.activeProjectId && state.excelFilePath) {
                    try {
                        await fetch('/api/unlink-excel', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ project_id: state.activeProjectId })
                        });
                    } catch (e) {
                        console.error('Error desvinculando Excel:', e);
                    }
                }

                state.excelFilePath = null;
                state.excelFileName = null;
                state.excelColumns = [];
                state.excelRecords = [];
                excelDropzone.classList.remove('hidden');
                excelInfo.classList.add('hidden');
                btnViewExcel.classList.add('hidden');
                if (btnOpenNativeExcel) btnOpenNativeExcel.classList.add('hidden');
                if (excelStatusPill) {
                    excelStatusPill.classList.add('disabled');
                    excelStatusPill.classList.remove('active');
                    excelStatusPill.querySelector('span').textContent = 'Sin Excel';
                }
                renderVariableChips([]);
                renderEmailMappableVars();
                updateGenerateButtonState();
                showToast('Excel desvinculado del proyecto (archivo conservado en disco)', 'info');
            });
        });
    }

    // EXCEL GRID MODAL & EDITOR LOGIC
    if (btnViewExcel) {
        btnViewExcel.addEventListener('click', () => {
            if (!state.excelFilePath) return;
            renderExcelGridTable();
            showModal(excelModal);
        });
    }

    if (btnCloseExcelModal) btnCloseExcelModal.addEventListener('click', () => excelModal.classList.add('hidden'));
    if (btnCancelExcelEdit) btnCancelExcelEdit.addEventListener('click', () => excelModal.classList.add('hidden'));

    if (toggleAllFieldsMandatory) {
        toggleAllFieldsMandatory.addEventListener('change', async (e) => {
            const isChecked = e.target.checked;
            state.settings.all_fields_mandatory = isChecked;
            if (isChecked) {
                state.settings.mandatory_columns = [...state.excelColumns];
            } else {
                state.settings.mandatory_columns = [];
            }
            await saveCurrentProjectSettings();
            renderExcelGridTable();
            showToast(isChecked ? 'Todos los campos marcados como obligatorios (*)' : 'Desactivada obligatoriedad de campos', 'info');
        });
    }

    function renderExcelGridTable() {
        const isGlobalMandatory = !!state.settings.all_fields_mandatory;
        if (toggleAllFieldsMandatory) {
            toggleAllFieldsMandatory.checked = isGlobalMandatory;
        }

        const mandatoryPillStatusEl = document.getElementById('mandatory-pill-status');
        if (mandatoryPillStatusEl) {
            mandatoryPillStatusEl.textContent = isGlobalMandatory ? 'TODOS OBLIGATORIOS (*)' : 'DESACTIVADO';
        }

        const specialControlCols = [
            state.settings.doc_check_column,
            state.settings.check_column,
            'Doc_Generado',
            'Enviado',
            'Estado_Generación'
        ].filter(Boolean);

        const mandatoryCols = isGlobalMandatory ? [...state.excelColumns] : (state.settings.mandatory_columns || []);

        let trHead = '<tr><th class="row-index-cell">#</th>';
        state.excelColumns.forEach((col, idx) => {
            const isSpecial = specialControlCols.includes(col);
            const isMandatory = !isSpecial && (isGlobalMandatory || mandatoryCols.includes(col));
            const redStar = isMandatory ? ' <span class="mandatory-asterisk" title="Campo Obligatorio">*</span>' : '';
            const btnMandatoryHtml = isSpecial ? `<span class="fs-11 text-muted" title="Columna de control (No puede ser obligatoria)"><i class="fa-solid fa-lock"></i></span>` : `
                <button class="btn-toggle-mandatory-col ${isMandatory ? 'active' : ''}" data-col="${escapeHtml(col)}" title="${isMandatory ? 'Campo Obligatorio (Haz clic para desmarcar)' : 'Marcar Campo como Obligatorio'}">
                    <i class="fa-solid fa-asterisk"></i>
                </button>`;

            trHead += `
                <th>
                    <div class="col-header-wrapper">
                        ${btnMandatoryHtml}
                        <span class="col-title" contenteditable="true" data-col-idx="${idx}">${escapeHtml(col)}</span>${redStar}
                        <button class="btn-icon-danger btn-delete-col" data-col="${escapeHtml(col)}" title="Eliminar columna"><i class="fa-solid fa-trash-can"></i></button>
                    </div>
                </th>`;
        });
        trHead += '<th class="row-action-cell"><i class="fa-solid fa-gear"></i></th></tr>';
        excelGridThead.innerHTML = trHead;

        let trBody = '';
        state.excelRecords.forEach((row, rIdx) => {
            trBody += `<tr><td class="row-index-cell">${rIdx + 1}</td>`;
            state.excelColumns.forEach(col => {
                const val = row[col] !== undefined && row[col] !== null ? String(row[col]) : '';
                trBody += `<td contenteditable="true" data-row-idx="${rIdx}" data-col="${escapeHtml(col)}">${escapeHtml(val)}</td>`;
            });
            trBody += `
                <td class="row-action-cell">
                    <button class="btn-icon-danger btn-delete-row" data-row-idx="${rIdx}" title="Eliminar fila"><i class="fa-solid fa-minus"></i></button>
                </td></tr>`;
        });

        if (state.excelRecords.length === 0) {
            trBody = `<tr><td colspan="${state.excelColumns.length + 2}" class="text-center p-3">Sin registros. Haz clic en "Agregar Fila" para crear un registro.</td></tr>`;
        }

        excelGridTbody.innerHTML = trBody;

        excelGridThead.querySelectorAll('.btn-toggle-mandatory-col').forEach(btn => {
            btn.addEventListener('click', async () => {
                const col = btn.dataset.col;
                let currentMandatory = state.settings.all_fields_mandatory ? [...state.excelColumns] : (state.settings.mandatory_columns || []);
                const idx = currentMandatory.indexOf(col);

                if (idx > -1) {
                    currentMandatory.splice(idx, 1);
                    state.settings.all_fields_mandatory = false;
                    showToast(`Columna '${col}' ya NO es obligatoria`, 'info');
                } else {
                    currentMandatory.push(col);
                    if (state.excelColumns.every(c => currentMandatory.includes(c))) {
                        state.settings.all_fields_mandatory = true;
                    }
                    showToast(`Columna '${col}' marcada como OBLIGATORIA`, 'success');
                }

                state.settings.mandatory_columns = currentMandatory;
                await saveCurrentProjectSettings();
                renderExcelGridTable();
            });
        });

        excelGridThead.querySelectorAll('.btn-delete-col').forEach(btn => {
            btn.addEventListener('click', () => {
                const colToDelete = btn.dataset.col;
                state.excelColumns = state.excelColumns.filter(c => c !== colToDelete);
                renderExcelGridTable();
            });
        });

        excelGridTbody.querySelectorAll('.btn-delete-row').forEach(btn => {
            btn.addEventListener('click', () => {
                const rIdx = parseInt(btn.dataset.rowIdx, 10);
                state.excelRecords.splice(rIdx, 1);
                renderExcelGridTable();
            });
        });
    }

    if (btnAddRow) {
        btnAddRow.addEventListener('click', () => {
            const newRow = {};
            state.excelColumns.forEach(c => newRow[c] = '');
            state.excelRecords.push(newRow);
            renderExcelGridTable();
        });
    }

    // IN-APP MODAL: AGREGAR NUEVA COLUMNA A EXCEL (SIN BROWSER PROMPT)
    if (btnAddCol) {
        btnAddCol.addEventListener('click', () => {
            if (!state.excelFilePath) {
                showToast('Primero debes cargar un archivo Excel', 'error');
                return;
            }
            if (newColumnNameInput) {
                newColumnNameInput.value = `Columna_${state.excelColumns.length + 1}`;
            }
            if (addColumnModal) {
                showModal(addColumnModal);
                setTimeout(() => { if (newColumnNameInput) newColumnNameInput.focus(); }, 100);
            }
        });
    }

    if (btnCancelAddColumn) {
        btnCancelAddColumn.addEventListener('click', () => {
            if (addColumnModal) addColumnModal.classList.add('hidden');
        });
    }

    function confirmAddNewColumn() {
        if (!newColumnNameInput) return;
        const rawName = newColumnNameInput.value.trim();
        if (!rawName) {
            showToast('El nombre de la columna no puede estar vacío', 'error');
            return;
        }
        if (state.excelColumns.includes(rawName)) {
            showToast(`La columna '${rawName}' ya existe en la tabla`, 'error');
            return;
        }

        state.excelColumns.push(rawName);
        state.excelRecords.forEach(r => r[rawName] = '');
        if (addColumnModal) addColumnModal.classList.add('hidden');
        renderExcelGridTable();
        showToast(`Columna '${rawName}' agregada a la tabla`, 'success');
    }

    if (btnConfirmAddColumn) {
        btnConfirmAddColumn.addEventListener('click', confirmAddNewColumn);
    }

    if (newColumnNameInput) {
        newColumnNameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                confirmAddNewColumn();
            }
        });
    }

    if (btnSaveExcelEdit) {
        btnSaveExcelEdit.addEventListener('click', async () => {
            const newCols = [];
            excelGridThead.querySelectorAll('.col-title').forEach(span => {
                const txt = span.textContent.trim();
                if (txt && !newCols.includes(txt)) newCols.push(txt);
            });
            if (newCols.length > 0) state.excelColumns = newCols;

            const newRecords = [];
            const rows = excelGridTbody.querySelectorAll('tr');
            rows.forEach((tr) => {
                const rowData = {};
                state.excelColumns.forEach(col => {
                    const cell = tr.querySelector(`td[data-col="${col}"]`);
                    rowData[col] = cell ? cell.textContent.trim() : '';
                });
                newRecords.push(rowData);
            });

            state.excelRecords = newRecords;
            showToast('Guardando cambios en Excel...', 'info');

            try {
                const res = await fetch('/api/update-excel-data', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        filepath: state.excelFilePath,
                        columns: state.excelColumns,
                        records: state.excelRecords
                    })
                });

                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error guardando datos');

                state.excelColumns = data.columns;
                state.excelRecords = data.all_records || data.records || state.excelRecords;
                excelRowsCount.textContent = `${data.total_rows} registros cargados`;
                excelStatusPill.querySelector('span').textContent = `${data.columns.length} Cols`;
                renderVariableChips(data.columns);
                renderEmailMappableVars();
                updateGenerateButtonState();
                if (recipientsSelectionModal && !recipientsSelectionModal.classList.contains('hidden')) {
                    renderRecipientsList();
                }
                excelModal.classList.add('hidden');
                showToast('Tabla Excel actualizada y guardada con éxito en disco', 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    function updateGenerateButtonState() {
        const hasDocx = state.docxTemplates && state.docxTemplates.length > 0;
        const hasExcel = state.excelFilePath && state.excelColumns && state.excelColumns.length > 0;
        const enabled = hasExcel && hasDocx;

        btnGenerateDocs.disabled = !enabled;
        if (btnGenerateDocsOptions) btnGenerateDocsOptions.disabled = !enabled;
        btnSendEmails.disabled = !enabled;
        if (btnSendEmailsOptions) btnSendEmailsOptions.disabled = !enabled;

        if (sendEmailsSplitGroup) {
            if (state.settings.outlook_enabled && enabled) {
                sendEmailsSplitGroup.classList.remove('hidden');
            } else {
                sendEmailsSplitGroup.classList.add('hidden');
            }
        }
    }

    // RECIPIENTS SELECTION MODAL LOGIC & SESSION STORAGE
    function getStoredSelectedRecipients() {
        if (!state.activeProjectId) return null;
        try {
            const stored = sessionStorage.getItem('selected_recipients_' + state.activeProjectId);
            return stored ? JSON.parse(stored) : null;
        } catch (e) {
            return null;
        }
    }

    function saveStoredSelectedRecipients(indices) {
        if (!state.activeProjectId) return;
        try {
            sessionStorage.setItem('selected_recipients_' + state.activeProjectId, JSON.stringify(indices));
        } catch (e) {}
    }

    if (btnSendEmailsOptions) {
        btnSendEmailsOptions.addEventListener('click', () => {
            if (!state.excelRecords || state.excelRecords.length === 0) {
                showToast('Primero debes cargar un archivo Excel con datos', 'error');
                return;
            }
            openRecipientsModal();
        });
    }

    function openRecipientsModal() {
        if (!recipientsSelectionModal) return;

        recipientIdCol.innerHTML = '';
        state.excelColumns.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c;
            opt.textContent = c;
            const cLower = c.toLowerCase();
            if (cLower === 'nombre' || cLower === 'empresa' || cLower === 'cedula' || cLower === 'cédula' || cLower === 'cargo') {
                opt.selected = true;
            }
            recipientIdCol.appendChild(opt);
        });

        let selectedIndices = getStoredSelectedRecipients();
        if (!selectedIndices) {
            selectedIndices = state.excelRecords.map((_, i) => i);
            saveStoredSelectedRecipients(selectedIndices);
        }

        renderRecipientsList(selectedIndices);
        showModal(recipientsSelectionModal);
    }

    function renderRecipientsList(selectedIndices = null) {
        if (!recipientsListContainer) return;
        recipientsListContainer.innerHTML = '';

        if (!selectedIndices) {
            selectedIndices = getStoredSelectedRecipients() || state.excelRecords.map((_, i) => i);
        }

        const idCol = recipientIdCol.value || state.excelColumns[0] || '';
        const emailCol = state.settings.email_column || 'Email';
        const checkCol = state.settings.check_column || 'Enviado';
        const query = (recipientSearchFilter.value || '').trim().toLowerCase();

        let visibleCount = 0;
        let selectedCount = 0;

        state.excelRecords.forEach((rec, idx) => {
            const emailVal = String(rec[emailCol] || '').trim();
            const idVal = String(rec[idCol] || '').trim();
            const checkVal = String(rec[checkCol] || '').trim();

            if (query && !emailVal.toLowerCase().includes(query) && !idVal.toLowerCase().includes(query)) {
                return;
            }
            visibleCount++;

            const isChecked = selectedIndices.includes(idx);
            if (isChecked) selectedCount++;

            const item = document.createElement('div');
            item.className = 'recipient-row-item';

            const statusHtml = checkVal ? `<span class="recipient-status-badge"><i class="fa-solid fa-check-double"></i> Enviado (${escapeHtml(checkVal)})</span>` : `<span class="meta fs-11">Pendiente</span>`;

            item.innerHTML = `
                <div class="recipient-info">
                    <span class="recipient-identifier">${escapeHtml(idVal || `Fila #${idx + 1}`)}</span>
                    <span class="recipient-email"><i class="fa-solid fa-envelope"></i> ${escapeHtml(emailVal || 'Sin correo')}</span>
                </div>
                <div class="flex-align-center gap-3">
                    ${statusHtml}
                    <input type="checkbox" class="recipient-check custom-checkbox ml-3" data-idx="${idx}" ${isChecked ? 'checked' : ''}>
                </div>
            `;

            recipientsListContainer.appendChild(item);
        });

        recipientsCounterBadge.textContent = `${selectedCount} de ${state.excelRecords.length} seleccionados`;
        checkAllRecipients.checked = selectedCount === state.excelRecords.length && state.excelRecords.length > 0;

        recipientsListContainer.querySelectorAll('.recipient-check').forEach(chk => {
            chk.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.idx, 10);
                let current = getStoredSelectedRecipients() || [];
                if (e.target.checked) {
                    if (!current.includes(idx)) current.push(idx);
                } else {
                    current = current.filter(i => i !== idx);
                }
                saveStoredSelectedRecipients(current);
                renderRecipientsList(current);
            });
        });
    }

    if (recipientIdCol) recipientIdCol.addEventListener('change', () => renderRecipientsList());
    if (recipientSearchFilter) recipientSearchFilter.addEventListener('input', () => renderRecipientsList());

    if (checkAllRecipients) {
        checkAllRecipients.addEventListener('change', (e) => {
            let current = [];
            if (e.target.checked) {
                current = state.excelRecords.map((_, i) => i);
            }
            saveStoredSelectedRecipients(current);
            renderRecipientsList(current);
        });
    }

    if (btnCloseRecipientsModal) btnCloseRecipientsModal.addEventListener('click', () => recipientsSelectionModal.classList.add('hidden'));
    if (btnCancelRecipientsModal) btnCancelRecipientsModal.addEventListener('click', () => recipientsSelectionModal.classList.add('hidden'));

    if (btnSendSelectedRecipients) {
        btnSendSelectedRecipients.addEventListener('click', () => {
            const selectedIndices = getStoredSelectedRecipients() || [];
            if (selectedIndices.length === 0) {
                showToast('Debes seleccionar al menos un destinatario', 'error');
                return;
            }
            recipientsSelectionModal.classList.add('hidden');
            triggerGenerationProcess('send_emails', 'Enviando a Destinatarios Seleccionados...', `Procesando envío para ${selectedIndices.length} destinatario(s)...`, selectedIndices);
        });
    }

    // ZOOM CONTROLS
    if (btnZoomIn) {
        btnZoomIn.addEventListener('click', () => {
            if (state.zoomLevel < 1.6) {
                state.zoomLevel += 0.1;
                applyZoom();
            }
        });
    }

    if (btnZoomOut) {
        btnZoomOut.addEventListener('click', () => {
            if (state.zoomLevel > 0.6) {
                state.zoomLevel -= 0.1;
                applyZoom();
            }
        });
    }

    function applyZoom() {
        paperPage.style.transform = `scale(${state.zoomLevel})`;
        pdfViewerFrame.style.transform = `scale(${state.zoomLevel})`;
        zoomLevelDisplay.textContent = `${Math.round(state.zoomLevel * 100)}%`;
    }

    if (btnToggleBadges) {
        btnToggleBadges.addEventListener('click', () => {
            btnToggleBadges.classList.toggle('active');
            const badges = paperPage.querySelectorAll('.docx-tag-badge');
            badges.forEach(b => {
                if (btnToggleBadges.classList.contains('active')) {
                    b.style.backgroundColor = '#eff6ff';
                    b.style.color = '#1d4ed8';
                    b.style.border = '1px solid #bfdbfe';
                } else {
                    b.style.backgroundColor = 'transparent';
                    b.style.color = 'inherit';
                    b.style.border = 'none';
                }
            });
        });
    }

    if (patternInput) {
        patternInput.addEventListener('change', async () => {
            const val = patternInput.value.trim();
            if (val) {
                state.settings.output_filename_pattern = val;
                await saveCurrentProjectSettings();
            }
        });
    }

    // GENERATE DOCS ONLY (DEFAULT: MISSING ONLY)
    if (btnGenerateDocs) {
        btnGenerateDocs.addEventListener('click', () => {
            triggerGenerationProcess('generate_docs', 'Generando Documentos Faltantes...', 'Por favor espera mientras el motor reemplaza los datos y exporta los archivos pendientes.', null, 'missing_only');
        });
    }

    // SEND EMAILS ONLY
    if (btnSendEmails) {
        btnSendEmails.addEventListener('click', () => {
            triggerGenerationProcess('send_emails', 'Enviando Correos por Outlook...', 'Por favor espera mientras se procesan los destinatarios y se envían los mensajes.');
        });
    }

    // CANCEL PROCESS
    if (btnCancelJob) {
        btnCancelJob.addEventListener('click', async () => {
            if (state.currentJobId) {
                try {
                    await fetch('/api/cancel-job', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ job_id: state.currentJobId })
                    });
                } catch (e) {
                    console.error('Error enviando cancelación:', e);
                }
            }
            if (state.pollInterval) clearInterval(state.pollInterval);
            progressModal.classList.add('hidden');
            showToast('Proceso cancelado por el usuario', 'info');
        });
    }

    async function triggerGenerationProcess(actionType, title, subtitle, selectedRowIndices = null, generateMode = 'all') {
        if (!state.excelFilePath) return;

        const pattern = patternInput.value.trim() || state.settings.output_filename_pattern || '{{ index }}_documento.docx';
        state.settings.output_filename_pattern = pattern;
        saveCurrentProjectSettings();
        const docxPaths = state.docxTemplates.map(t => t.filepath);
        if (docxPaths.length === 0 && state.docxFilePath) {
            docxPaths.push(state.docxFilePath);
        }

        const modalIcon = progressModal ? progressModal.querySelector('.modal-icon') : null;
        if (modalIcon) {
            modalIcon.className = 'modal-icon spin';
            modalIcon.style.color = '';
            modalIcon.innerHTML = '<i class="fa-solid fa-gear"></i>';
        }

        modalTitle.textContent = title;
        modalSubtitle.textContent = subtitle;
        progressBarFill.style.width = '0%';
        progressText.textContent = 'Iniciando...';
        progressPercent.textContent = '0%';
        modalActions.classList.add('hidden');
        if (btnCancelJob) btnCancelJob.classList.remove('hidden');
        showModal(progressModal);

        try {
            const payload = {
                excel_path: state.excelFilePath,
                docx_paths: docxPaths,
                output_filename_pattern: pattern,
                output_dir_override: state.outputDirPath,
                settings: state.settings,
                action: actionType,
                generate_mode: generateMode
            };
            if (selectedRowIndices && Array.isArray(selectedRowIndices)) {
                payload.selected_row_indices = selectedRowIndices;
            }

            const res = await fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error iniciando proceso');

            state.currentJobId = data.job_id;
            startPollingProgress(data.job_id);
        } catch (err) {
            progressModal.classList.add('hidden');
            showToast(err.message, 'error');
        }
    }

    function startPollingProgress(jobId) {
        if (state.pollInterval) clearInterval(state.pollInterval);
        const modalIcon = progressModal ? progressModal.querySelector('.modal-icon') : null;

        state.pollInterval = setInterval(async () => {
            try {
                const res = await fetch(`/api/progress/${jobId}`);
                const data = await res.json();

                if (data.status === 'processing') {
                    progressBarFill.style.width = `${data.progress}%`;
                    progressText.textContent = `${data.current} de ${data.total} elementos`;
                    progressPercent.textContent = `${data.progress}%`;
                    state.currentOutputDir = data.output_dir;
                } else if (data.status === 'completed') {
                    clearInterval(state.pollInterval);
                    if (btnCancelJob) btnCancelJob.classList.add('hidden');
                    progressBarFill.style.width = '100%';
                    progressPercent.textContent = '100%';
                    progressText.textContent = `${data.total} de ${data.total} completados`;
                    state.currentOutputDir = data.output_dir;
                    modalActions.classList.remove('hidden');

                    const hasWarnings = data.has_warnings || (data.warnings && data.warnings.length > 0);
                    const isAllAlreadyGenerated = hasWarnings && data.warnings.every(w => w.includes('pendientes por generar') || w.includes('ya generados') || w.includes('0 registros'));

                    if (isAllAlreadyGenerated) {
                        progressModal.classList.add('hidden');
                        showToast('Todos los documentos ya han sido generados previamente', 'info');
                    } else {
                        if (modalIcon) {
                            modalIcon.classList.remove('spin');
                            if (hasWarnings) {
                                modalIcon.className = 'modal-icon text-warning';
                                modalIcon.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i>';
                            } else {
                                modalIcon.className = 'modal-icon text-success';
                                modalIcon.innerHTML = '<i class="fa-solid fa-circle-check"></i>';
                            }
                        }

                        if (hasWarnings) {
                            modalTitle.textContent = 'Proceso Completado con Advertencias';
                            modalSubtitle.textContent = `Se procesó el lote, pero ocurrieron ${data.warnings.length} advertencia(s).`;
                            showToast('Proceso finalizado con advertencias', 'error');
                        } else {
                            modalTitle.textContent = 'Proceso Completado!';
                            modalSubtitle.textContent = 'Se han generado los archivos o correos exitosamente.';
                            showToast('Proceso finalizado con éxito', 'success');
                        }
                    }

                    if (state.activeProjectId) {
                        loadSelectedProject(state.activeProjectId);
                    }
                } else if (data.status === 'canceled') {
                    clearInterval(state.pollInterval);
                    if (btnCancelJob) btnCancelJob.classList.add('hidden');
                    progressModal.classList.add('hidden');
                    showToast('Proceso cancelado por el usuario', 'info');
                } else if (data.status === 'error') {
                    clearInterval(state.pollInterval);
                    if (btnCancelJob) btnCancelJob.classList.add('hidden');
                    if (modalIcon) {
                        modalIcon.classList.remove('spin');
                        modalIcon.className = 'modal-icon text-danger';
                        modalIcon.innerHTML = '<i class="fa-solid fa-circle-xmark"></i>';
                    }
                    modalTitle.textContent = 'Error en el Proceso';
                    modalSubtitle.textContent = data.error || 'Ocurrió un error crítico durante la ejecución.';
                    modalActions.classList.remove('hidden');
                    btnOpenFolder.classList.add('hidden');
                }
            } catch (err) {
                console.error(err);
            }
        }, 300);
    }

    if (btnOpenFolder) {
        btnOpenFolder.addEventListener('click', async () => {
            if (!state.currentOutputDir) return;
            try {
                await fetch('/api/open-output-folder', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ folder_path: state.currentOutputDir })
                });
            } catch (err) {
                showToast('No se pudo abrir la carpeta', 'error');
            }
        });
    }

    if (btnCloseModal) {
        btnCloseModal.addEventListener('click', () => {
            progressModal.classList.add('hidden');
        });
    }

    // TEMPLATE RULES MODAL LOGIC ("Reglas de Plantilla")
    if (btnOpenTemplateRulesModal) {
        btnOpenTemplateRulesModal.addEventListener('click', () => {
            renderTemplateRulesList();
            showModal(templateRulesModal);
        });
    }

    if (btnCloseTemplateRulesModal) btnCloseTemplateRulesModal.addEventListener('click', () => templateRulesModal.classList.add('hidden'));
    if (btnCancelTemplateRules) btnCancelTemplateRules.addEventListener('click', () => templateRulesModal.classList.add('hidden'));

    if (btnAddTemplateRule) {
        btnAddTemplateRule.addEventListener('click', () => {
            const defCol = state.excelColumns[0] || 'Tipo';
            const defTpl = state.docxTemplates[0] ? state.docxTemplates[0].filename : '';
            if (!state.settings.template_rules) state.settings.template_rules = [];
            state.settings.template_rules.push({
                column: defCol,
                value: '',
                template_filename: defTpl
            });
            renderTemplateRulesList();
        });
    }

    function renderTemplateRulesList() {
        if (!templateRulesListContainer) return;
        templateRulesListContainer.innerHTML = '';

        const rules = state.settings.template_rules || [];
        if (rules.length === 0) {
            templateRulesListContainer.innerHTML = '<div class="empty-state p-3"><p class="meta">Sin reglas de selección de plantilla configuradas. Haz clic en "Agregar Regla de Plantilla".</p></div>';
            return;
        }

        rules.forEach((rule, idx) => {
            const card = document.createElement('div');
            card.className = 'rule-card p-3 mb-2 bg-darker rounded border-light';

            let colOptionsHtml = '';
            state.excelColumns.forEach(c => {
                colOptionsHtml += `<option value="${escapeHtml(c)}" ${c === rule.column ? 'selected' : ''}>${escapeHtml(c)}</option>`;
            });

            let tplOptionsHtml = '';
            if (state.docxTemplates && state.docxTemplates.length > 0) {
                state.docxTemplates.forEach(t => {
                    tplOptionsHtml += `<option value="${escapeHtml(t.filename)}" ${t.filename === rule.template_filename ? 'selected' : ''}>${escapeHtml(t.filename)}</option>`;
                });
            } else {
                tplOptionsHtml = `<option value="${escapeHtml(rule.template_filename || '')}">${escapeHtml(rule.template_filename || 'Sin plantilla')}</option>`;
            }

            card.innerHTML = `
                <div class="grid-3col gap-2 align-center">
                    <div class="form-group">
                        <label class="fs-12">Si la columna:</label>
                        <select class="form-control rule-tr-col-select" data-rule-idx="${idx}">
                            ${colOptionsHtml}
                        </select>
                    </div>
                    <div class="form-group">
                        <label class="fs-12">Es igual a este valor:</label>
                        <input type="text" class="form-control rule-tr-val-input" data-rule-idx="${idx}" value="${escapeHtml(rule.value)}" placeholder="ej. Examen">
                    </div>
                    <div class="form-group">
                        <label class="fs-12">Usar únicamente plantilla:</label>
                        <div class="flex-align-center gap-2">
                            <select class="form-control rule-tr-tpl-select" data-rule-idx="${idx}">
                                ${tplOptionsHtml}
                            </select>
                            <button class="btn-icon-danger btn-delete-tr-rule" data-rule-idx="${idx}" title="Eliminar regla"><i class="fa-solid fa-trash-can"></i></button>
                        </div>
                    </div>
                </div>
            `;
            templateRulesListContainer.appendChild(card);
        });

        templateRulesListContainer.querySelectorAll('.rule-tr-col-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.template_rules[idx].column = e.target.value;
            });
        });

        templateRulesListContainer.querySelectorAll('.rule-tr-val-input').forEach(inp => {
            inp.addEventListener('input', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.template_rules[idx].value = e.target.value;
            });
        });

        templateRulesListContainer.querySelectorAll('.rule-tr-tpl-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.ruleIdx, 10);
                state.settings.template_rules[idx].template_filename = e.target.value;
            });
        });

        templateRulesListContainer.querySelectorAll('.btn-delete-tr-rule').forEach(btn => {
            btn.addEventListener('click', () => {
                const idx = parseInt(btn.dataset.ruleIdx, 10);
                state.settings.template_rules.splice(idx, 1);
                renderTemplateRulesList();
            });
        });
    }

    if (btnSaveTemplateRules) {
        btnSaveTemplateRules.addEventListener('click', async () => {
            await saveCurrentProjectSettings();
            if (templateRulesModal) templateRulesModal.classList.add('hidden');
            showToast('Reglas de plantillas guardadas exitosamente', 'success');
        });
    }

    // DOCUMENT SELECTION & SMART REGENERATION MODAL ("Sub-Botón de Opciones")
    if (btnGenerateDocsOptions) {
        btnGenerateDocsOptions.addEventListener('click', () => {
            if (!state.excelRecords || state.excelRecords.length === 0) {
                showToast('Carga primero un archivo Excel con registros', 'error');
                return;
            }
            openDocsSelectionModal();
        });
    }

    function openDocsSelectionModal() {
        if (!docsSelectionModal) return;

        docsSelectionIdCol.innerHTML = '';
        state.excelColumns.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c;
            opt.textContent = c;
            const cLower = c.toLowerCase();
            if (cLower === 'nombre' || cLower === 'empresa' || cLower === 'cedula' || cLower === 'cédula' || cLower === 'cargo') {
                opt.selected = true;
            }
            docsSelectionIdCol.appendChild(opt);
        });

        let selectedIndices = getStoredSelectedDocs();
        if (!selectedIndices) {
            selectedIndices = state.excelRecords.map((_, i) => i);
            saveStoredSelectedDocs(selectedIndices);
        }

        renderDocsSelectionList(selectedIndices);
        showModal(docsSelectionModal);
    }

    function getStoredSelectedDocs() {
        if (!state.activeProjectId) return null;
        try {
            const stored = sessionStorage.getItem('selected_docs_' + state.activeProjectId);
            return stored ? JSON.parse(stored) : null;
        } catch (e) { return null; }
    }

    function saveStoredSelectedDocs(indices) {
        if (!state.activeProjectId) return;
        try {
            sessionStorage.setItem('selected_docs_' + state.activeProjectId, JSON.stringify(indices));
        } catch (e) {}
    }

    function renderDocsSelectionList(selectedIndices = null) {
        if (!docsSelectionListContainer) return;
        docsSelectionListContainer.innerHTML = '';

        if (!selectedIndices) {
            selectedIndices = getStoredSelectedDocs() || state.excelRecords.map((_, i) => i);
        }

        const idCol = docsSelectionIdCol.value || state.excelColumns[0] || '';
        const docCheckCol = state.settings.doc_check_column || 'Doc_Generado';
        const query = (docsSelectionSearchFilter.value || '').trim().toLowerCase();

        let visibleCount = 0;
        let selectedCount = 0;

        state.excelRecords.forEach((row, idx) => {
            const idVal = row[idCol] ? String(row[idCol]) : `Fila ${idx + 1}`;
            const docStatusVal = String(row[docCheckCol] || '').trim();
            const isGenerated = !!docStatusVal && !['none', 'nan', 'false', '0', 'no'].includes(docStatusVal.toLowerCase());

            const textSearch = (Object.values(row).join(' ') + (isGenerated ? ' generado ' + docStatusVal : ' sin generar faltante')).toLowerCase();
            if (query && !textSearch.includes(query)) return;

            visibleCount++;
            const isChecked = selectedIndices.includes(idx);
            if (isChecked) selectedCount++;

            const card = document.createElement('div');
            card.className = `recipient-item ${isChecked ? 'selected' : ''}`;

            card.innerHTML = `
                <div class="recipient-checkbox-wrapper">
                    <input type="checkbox" class="doc-row-check custom-checkbox" data-row-idx="${idx}" ${isChecked ? 'checked' : ''}>
                </div>
                <div class="recipient-info">
                    <span class="recipient-name">${escapeHtml(idVal)}</span>
                    <span class="recipient-meta font-mono fs-11">Fila #${idx + 1}</span>
                </div>
                <div class="recipient-status">
                    ${isGenerated ? `<span class="badge badge-success"><i class="fa-solid fa-file-circle-check"></i> ${escapeHtml(docStatusVal)}</span>` : `<span class="badge badge-warning"><i class="fa-solid fa-file-circle-exclamation"></i> Sin Generar (Faltante)</span>`}
                </div>
            `;

            docsSelectionListContainer.appendChild(card);
        });

        if (docsSelectionCounterBadge) {
            docsSelectionCounterBadge.textContent = `${selectedCount} de ${state.excelRecords.length} seleccionados`;
        }

        if (checkAllDocsSelection) {
            checkAllDocsSelection.checked = visibleCount > 0 && selectedCount === visibleCount;
        }

        docsSelectionListContainer.querySelectorAll('.doc-row-check').forEach(chk => {
            chk.addEventListener('change', (e) => {
                const rIdx = parseInt(e.target.dataset.rowIdx, 10);
                let current = getStoredSelectedDocs() || state.excelRecords.map((_, i) => i);
                if (e.target.checked) {
                    if (!current.includes(rIdx)) current.push(rIdx);
                } else {
                    current = current.filter(i => i !== rIdx);
                }
                saveStoredSelectedDocs(current);
                renderDocsSelectionList(current);
            });
        });
    }

    if (docsSelectionSearchFilter) docsSelectionSearchFilter.addEventListener('input', () => renderDocsSelectionList());
    if (docsSelectionIdCol) docsSelectionIdCol.addEventListener('change', () => renderDocsSelectionList());

    if (checkAllDocsSelection) {
        checkAllDocsSelection.addEventListener('change', (e) => {
            const isChecked = e.target.checked;
            const query = (docsSelectionSearchFilter.value || '').trim().toLowerCase();
            const idCol = docsSelectionIdCol.value || state.excelColumns[0] || '';
            const docCheckCol = state.settings.doc_check_column || 'Doc_Generado';

            let current = getStoredSelectedDocs() || state.excelRecords.map((_, i) => i);

            state.excelRecords.forEach((row, idx) => {
                const idVal = row[idCol] ? String(row[idCol]) : `Fila ${idx + 1}`;
                const docStatusVal = String(row[docCheckCol] || '').trim();
                const isGen = !!docStatusVal;
                const textSearch = (Object.values(row).join(' ') + (isGen ? ' generado' : ' sin generar')).toLowerCase();

                if (!query || textSearch.includes(query)) {
                    if (isChecked) {
                        if (!current.includes(idx)) current.push(idx);
                    } else {
                        current = current.filter(i => i !== idx);
                    }
                }
            });

            saveStoredSelectedDocs(current);
            renderDocsSelectionList(current);
        });
    }

    if (btnCloseDocsSelectionModal) btnCloseDocsSelectionModal.addEventListener('click', () => docsSelectionModal.classList.add('hidden'));
    if (btnCancelDocsSelectionModal) btnCancelDocsSelectionModal.addEventListener('click', () => docsSelectionModal.classList.add('hidden'));

    if (btnGenerateMissingDocsAction) {
        btnGenerateMissingDocsAction.addEventListener('click', async () => {
            if (docsSelectionModal) docsSelectionModal.classList.add('hidden');
            await triggerGenerationProcess('generate_docs', 'Generando Documentos Faltantes...', 'Procesando únicamente las filas pendientes...', null, 'missing_only');
        });
    }

    if (btnGenerateSelectedDocsAction) {
        btnGenerateSelectedDocsAction.addEventListener('click', async () => {
            const selectedIndices = getStoredSelectedDocs();
            if (!selectedIndices || selectedIndices.length === 0) {
                showToast('Selecciona al menos un registro para generar/regenerar', 'error');
                return;
            }
            if (docsSelectionModal) docsSelectionModal.classList.add('hidden');
            await triggerGenerationProcess('generate_docs', 'Generando / Regenerando Documentos Seleccionados...', `Procesando ${selectedIndices.length} documento(s)...`, selectedIndices, 'all');
        });
    }

    function showToast(message, type = 'info') {
        const container = document.getElementById('toast-container');
        if (!container) return;

        // Eliminar notificaciones anteriores inmediatamente para evitar spam/acumulación
        container.innerHTML = '';

        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;

        let icon = 'fa-circle-info';
        if (type === 'success') icon = 'fa-circle-check';
        if (type === 'error') icon = 'fa-triangle-exclamation';

        toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${escapeHtml(message)}</span>`;
        container.appendChild(toast);

        const duration = type === 'info' ? 1800 : 2500;

        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }

    function escapeHtml(str) {
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // MODAL BACKDROP AND ESCAPE KEY CLOSING LOGIC
    document.querySelectorAll('.modal-backdrop').forEach(modal => {
        modal.addEventListener('click', (e) => {
            if (e.target === modal && modal.id !== 'progress-modal') {
                modal.classList.add('hidden');
            }
        });
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal-backdrop:not(.hidden)').forEach(modal => {
                if (modal.id !== 'progress-modal') {
                    modal.classList.add('hidden');
                }
            });
        }
    });

    // IMPORT PROJECT LOGIC
    const btnOpenImportProject = document.getElementById('btn-open-import-project');
    const importProjectModal = document.getElementById('import-project-modal');
    const btnCloseImportProjModal = document.getElementById('btn-close-import-proj-modal');
    const btnCancelImportProj = document.getElementById('btn-cancel-import-proj');
    const btnConfirmImportProj = document.getElementById('btn-confirm-import-proj');
    const importProjPathInput = document.getElementById('import-proj-path-input');

    if (btnOpenImportProject) {
        btnOpenImportProject.addEventListener('click', () => {
            if (importProjPathInput) importProjPathInput.value = '';
            showModal(importProjectModal);
        });
    }
    if (btnCloseImportProjModal) btnCloseImportProjModal.addEventListener('click', () => importProjectModal.classList.add('hidden'));
    if (btnCancelImportProj) btnCancelImportProj.addEventListener('click', () => importProjectModal.classList.add('hidden'));

    if (btnConfirmImportProj) {
        btnConfirmImportProj.addEventListener('click', async () => {
            const pathVal = importProjPathInput.value.trim();
            if (!pathVal) {
                showToast('Ingresa la ruta de la carpeta del proyecto', 'error');
                return;
            }
            try {
                const res = await fetch('/api/import-project', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ folder_path: pathVal })
                });
                const data = await res.json();
                if (!res.ok) throw new Error(data.detail || 'Error al importar proyecto');

                importProjectModal.classList.add('hidden');
                await loadProjectsList();
                await loadSelectedProject(data.project.id);
                showToast(`Proyecto '${data.project.name}' importado exitosamente`, 'success');
            } catch (err) {
                showToast(err.message, 'error');
            }
        });
    }

    // GENERATED FILES LOGIC
    const btnOpenGeneratedFiles = document.getElementById('btn-open-generated-files');
    const generatedFilesModal = document.getElementById('generated-files-modal');
    const btnCloseGeneratedFilesModal = document.getElementById('btn-close-generated-files-modal');
    const btnCloseGeneratedFilesModalBtn = document.getElementById('btn-close-generated-files-modal-btn');
    const generatedFilesListContainer = document.getElementById('generated-files-list-container');
    const generatedFilesSubtitle = document.getElementById('generated-files-subtitle');
    const btnOpenOutputFolderDirect = document.getElementById('btn-open-output-folder-direct');

    if (btnOpenGeneratedFiles) {
        btnOpenGeneratedFiles.addEventListener('click', async () => {
            if (!state.activeProjectId) {
                showToast('Selecciona un proyecto para ver sus archivos generados', 'error');
                return;
            }
            showModal(generatedFilesModal);
            await renderGeneratedFilesList();
        });
    }

    if (btnCloseGeneratedFilesModal) btnCloseGeneratedFilesModal.addEventListener('click', () => generatedFilesModal.classList.add('hidden'));
    if (btnCloseGeneratedFilesModalBtn) btnCloseGeneratedFilesModalBtn.addEventListener('click', () => generatedFilesModal.classList.add('hidden'));

    if (btnOpenOutputFolderDirect) {
        btnOpenOutputFolderDirect.addEventListener('click', async () => {
            if (!state.outputDirPath) return;
            try {
                await fetch('/api/open-file', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filepath: state.outputDirPath })
                });
            } catch (e) {}
        });
    }

    async function renderGeneratedFilesList() {
        if (!generatedFilesListContainer || !state.activeProjectId) return;
        generatedFilesListContainer.innerHTML = '<span class="meta p-3">Cargando archivos generados...</span>';

        try {
            const res = await fetch('/api/list-generated-files', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ project_id: state.activeProjectId })
            });
            const data = await res.json();
            if (!res.ok) throw new Error(data.detail || 'Error leyendo carpeta salida');

            if (generatedFilesSubtitle) {
                generatedFilesSubtitle.textContent = `Mostrando ${data.files.length} archivo(s) en la carpeta salida`;
            }

            if (data.files.length === 0) {
                generatedFilesListContainer.innerHTML = '<div class="p-3 text-center meta">Aún no se han generado archivos en la carpeta de salida.</div>';
                return;
            }

            generatedFilesListContainer.innerHTML = '';
            data.files.forEach(f => {
                const isPdf = f.type === 'pdf';
                const iconClass = isPdf ? 'fa-file-pdf generated-file-icon-pdf' : 'fa-file-word generated-file-icon-docx';
                const card = document.createElement('div');
                card.className = 'generated-file-card';
                card.innerHTML = `
                    <div class="flex-align-center gap-3">
                        <i class="fa-solid ${iconClass}"></i>
                        <div>
                            <div class="generated-file-name">${escapeHtml(f.filename)}</div>
                            <div class="generated-file-meta">${f.size_kb} KB • ${f.mtime}</div>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary btn-sm btn-open-gen-file" data-path="${escapeHtml(f.filepath)}">
                        <i class="fa-solid fa-arrow-up-right-from-square"></i> Abrir
                    </button>
                `;
                generatedFilesListContainer.appendChild(card);
            });

            generatedFilesListContainer.querySelectorAll('.btn-open-gen-file').forEach(btn => {
                btn.addEventListener('click', async () => {
                    const fp = btn.dataset.path;
                    try {
                        await fetch('/api/open-file', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ filepath: fp })
                        });
                    } catch (e) {}
                });
            });
        } catch (err) {
            generatedFilesListContainer.innerHTML = `<span class="text-danger p-3">${escapeHtml(err.message)}</span>`;
        }
    }

    // STEP 1 & 2 THREE-DOTS MENU HANDLERS
    const btnExcelStepMenu = document.getElementById('btn-excel-step-menu');
    const excelStepDropdown = document.getElementById('excel-step-dropdown');
    const menuBtnViewExcel = document.getElementById('menu-btn-view-excel');
    const menuBtnOpenNativeExcel = document.getElementById('menu-btn-open-native-excel');
    const menuBtnUnlinkExcel = document.getElementById('menu-btn-unlink-excel');

    const btnDocxStepMenu = document.getElementById('btn-docx-step-menu');
    const docxStepDropdown = document.getElementById('docx-step-dropdown');
    const menuBtnAddDocx = document.getElementById('menu-btn-add-docx');
    const menuBtnOpenNativeDocx = document.getElementById('menu-btn-open-native-docx');
    const menuBtnTemplateRules = document.getElementById('menu-btn-template-rules');
    const menuBtnUnlinkDocx = document.getElementById('menu-btn-unlink-docx');

    if (btnExcelStepMenu && excelStepDropdown) {
        btnExcelStepMenu.addEventListener('click', (e) => {
            e.stopPropagation();
            if (docxStepDropdown) docxStepDropdown.classList.add('hidden');
            excelStepDropdown.classList.toggle('hidden');
        });
    }

    if (btnDocxStepMenu && docxStepDropdown) {
        btnDocxStepMenu.addEventListener('click', (e) => {
            e.stopPropagation();
            if (excelStepDropdown) excelStepDropdown.classList.add('hidden');
            docxStepDropdown.classList.toggle('hidden');
        });
    }

    document.addEventListener('click', (e) => {
        if (excelStepDropdown && !excelStepDropdown.contains(e.target) && e.target !== btnExcelStepMenu) {
            excelStepDropdown.classList.add('hidden');
        }
        if (docxStepDropdown && !docxStepDropdown.contains(e.target) && e.target !== btnDocxStepMenu) {
            docxStepDropdown.classList.add('hidden');
        }
    });

    if (menuBtnViewExcel) {
        menuBtnViewExcel.addEventListener('click', () => {
            if (excelStepDropdown) excelStepDropdown.classList.add('hidden');
            if (btnViewExcel) btnViewExcel.click();
        });
    }

    if (menuBtnOpenNativeExcel) {
        menuBtnOpenNativeExcel.addEventListener('click', () => {
            if (excelStepDropdown) excelStepDropdown.classList.add('hidden');
            if (btnOpenNativeExcel) btnOpenNativeExcel.click();
        });
    }

    if (menuBtnUnlinkExcel) {
        menuBtnUnlinkExcel.addEventListener('click', () => {
            if (excelStepDropdown) excelStepDropdown.classList.add('hidden');
            if (btnRemoveExcel) btnRemoveExcel.click();
        });
    }

    if (menuBtnAddDocx) {
        menuBtnAddDocx.addEventListener('click', () => {
            if (docxStepDropdown) docxStepDropdown.classList.add('hidden');
            if (docxInput) docxInput.click();
        });
    }

    if (menuBtnOpenNativeDocx) {
        menuBtnOpenNativeDocx.addEventListener('click', () => {
            if (docxStepDropdown) docxStepDropdown.classList.add('hidden');
            if (btnOpenNativeDocx) btnOpenNativeDocx.click();
        });
    }

    if (menuBtnTemplateRules) {
        menuBtnTemplateRules.addEventListener('click', () => {
            if (docxStepDropdown) docxStepDropdown.classList.add('hidden');
            if (btnOpenTemplateRulesTrigger) btnOpenTemplateRulesTrigger.click();
        });
    }

    if (menuBtnUnlinkDocx) {
        menuBtnUnlinkDocx.addEventListener('click', () => {
            if (docxStepDropdown) docxStepDropdown.classList.add('hidden');
            if (state.docxTemplates && state.docxTemplates.length > 0) {
                deleteDocxTemplate(state.activeDocxIndex);
            } else {
                showToast('No hay plantilla cargada para desvincular', 'info');
            }
        });
    }

    setupPaperDragAndDrop();
});

