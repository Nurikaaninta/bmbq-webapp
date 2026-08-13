function appData() {
    return {
        isLoggedIn: false,
        loginError: false,
        loginErrorMsg: '',
        loginForm: {
            user: '',
            pass: '',
            role: ''
        },
        
        currentDashboardTab: 'home',
        activeProject: 'PRJ-3000',
        allProjectsData: {},

        sheets: [
            'SP Items', 'Valve', 'Support', 'LineList', 'Tee', 'Single Branch Fitting', 'Pipe', 'Nozzle', 
            'Instrument', 'Flange', 'Elbow', 'Coupling', 'Pipe Run Component', 
            'Tap Weld', 'Socketweld', 'Gasket', 'Buttweld', 'Bolt Set', 
            'Fasteners', 'Vessel', 'Tank', 'Pump', 'Misc Equipment', 
            'Equipment', 'Piping and Equipment'
        ],
        activeSheet: 'Valve',
        addFormType: 'MTO',
        globalSearch: '',
        currentPage: 1,
        itemsPerPage: 15,
        tableRenderKey: 0,
        _columnsCacheKey: null,
        _columnsCacheValue: null,

        // Filter per kolom seperti Excel.
        // Key = nama kolom, value = daftar nilai yang dipilih.
        columnFilters: {},
        filterPopupOpen: false,
        activeFilterColumn: '',
        filterOptionSearch: '',
        filterOptions: [],
        filterDraftValues: [],
        filterPopupPosition: { left: 0, top: 0 },

        // Mode input data langsung di tabel.
        addingNewData: false,
        newInlineRow: {},

        showAddModal: false,
        showBoqModal: false,
        showCostBoqModal: false,
        showApproveModal: false,
        showFinalReportModal: false,
        taskView: 'active',
        approvalHistory: [],
        approvalNote: '',
        costRows: [],
        costNotes: '',
        // Snapshot Quantity Take-Off yang dikirim Engineer ke Estimator.
        qtoRows: [],
        qtoSourceSheet: '',
        qtoSubmittedAt: '',
        workflowStatus: 'DRAFT',
        showModuleModal: false,

        // State Create Project
        showCreateProjectModal: false,
        isCreatingProject: false,
        createProjectMessage: '',
        createProjectMessageType: 'error',
        createProjectForm: {
            code: '',
            name: '',
            description: '',
            initialization: 'empty',
            file: null,
            fileName: ''
        },

        activeModuleName: '',
        moduleDescription: '',

        // Calculate BOM: Quantity Take-Off (tanpa harga)
        boqCurrency: 'USD',
        bomMaterialType: 'Pipe',
        bomMaterialFilter: '',
        bomSizeFilter: '',
        bomSpecFilter: '',
        bomQuantityInput: 0,
        bomLength: 6,
        bomOD: 0,
        bomThickness: 0,
        
        newRowForm: {
            longDesc: '',
            material: 'CS',
            spec: 'CS150',
            size: '2"',
            pressure: '150',
            wasteFactor: '5',
            tag: ''
        },

        // State & Fungsi Pendukung Kalender
        calendarMonth: new Date().getMonth(), // 0 - 11
        calendarYear: new Date().getFullYear(),
        monthNames: ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'],

        get calendarDaysInMonth() {
            return new Date(this.calendarYear, this.calendarMonth + 1, 0).getDate();
        },

        get calendarFirstDayIndex() {
            return new Date(this.calendarYear, this.calendarMonth, 1).getDay();
        },

        prevMonth() {
            if (this.calendarMonth === 0) {
                this.calendarMonth = 11;
                this.calendarYear--;
            } else {
                this.calendarMonth--;
            }
        },

        nextMonth() {
            if (this.calendarMonth === 11) {
                this.calendarMonth = 0;
                this.calendarYear++;
            } else {
                this.calendarMonth++;
            }
        },

        // Tambahan State untuk Modul Interaktif Baru
        travelForm: {
            destination: 'Yogyakarta',
            date: '2026-08-15'
        },
        selectedCalendarDate: 4,
        newAgendaText: '',
        calendarAgendas: {
            '4-7-2026': ['Meeting Koordinasi Proyek BMBQ', 'Review MTO Valve']
        },
        policiesList: [
            { id: 1, title: 'SOP Keselamatan Kerja (HSE-01)', category: 'SOP', status: 'Aktif' },
            { id: 2, title: 'Panduan Pengadaan Material Pipa (PR-04)', category: 'Panduan', status: 'Aktif' }
        ],

        // Fungsi Pendukung Modul Baru
        submitTravelRequest() {
            alert(`Pengajuan perjalanan dinas ke ${this.travelForm.destination} tanggal ${this.travelForm.date} berhasil disimpan!`);
        },

        selectCalendarDate(day) {
            this.selectedCalendarDate = day;
        },

        addAgenda() {
            if (!this.newAgendaText.trim()) return;
            const key = `${this.selectedCalendarDate}-${this.calendarMonth}-${this.calendarYear}`;
            if (!this.calendarAgendas[key]) {
                this.calendarAgendas[key] = [];
            }
            this.calendarAgendas[key].push(this.newAgendaText.trim());
            this.newAgendaText = '';
            alert('Agenda berhasil ditambahkan!');
        },

        removeAgenda(key, index) {
            if (this.calendarAgendas[key]) {
                this.calendarAgendas[key].splice(index, 1);
            }
        },

        init() {
            this.$nextTick(() => {
                this.updateTableScrollbar();
                if (!this.tableScrollbarResizeObserver) {
                    const card = document.querySelector('.workspace-table-card');
                    if (card && window.ResizeObserver) {
                        this.tableScrollbarResizeObserver = new ResizeObserver(() => this.updateTableScrollbar());
                        this.tableScrollbarResizeObserver.observe(card);
                    }
                }
            });

            window.addEventListener('resize', () => this.updateTableScrollbar());

            const tableObserver = new MutationObserver(() => {
                this.$nextTick(() => this.updateTableScrollbar());
            });
            this.$nextTick(() => {
                const tableArea = document.querySelector('.excel-table-container');
                if (tableArea) {
                    tableObserver.observe(tableArea, { childList: true, subtree: true });
                }
            });

            // ==========================================================
            // DATABASE V4 - SEMUA TABEL DIMULAI KOSONG
            // Setiap tabel memiliki storage sendiri berdasarkan PROJECT + SHEET.
            // ==========================================================
            const STORAGE_VERSION = 'BMBQ_TABLE_STORAGE_V7_EXACT_EXCEL';
            const savedVersion = localStorage.getItem('tripatra_storage_version');

            if (savedVersion !== STORAGE_VERSION) {
                Object.keys(localStorage)
                    .filter(key => key.startsWith('tripatra_table_v4_') || key === 'tripatra_multiproject_db')
                    .forEach(key => localStorage.removeItem(key));
                localStorage.setItem('tripatra_storage_version', STORAGE_VERSION);
            }

            const defaultPRJ3000 = this.createBlankProject('PRJ-3000', 'Piping & Equipment', 'Project utama BMBQ WebApp');
            const defaultPRJ4000 = this.createBlankProject('PRJ-4000', 'Piping & Equipment', '');
            const defaultPRJ5000 = this.createBlankProject('PRJ-5000', 'Piping & Equipment', '');

            this.allProjectsData = {
                'PRJ-3000': defaultPRJ3000,
                'PRJ-4000': defaultPRJ4000,
                'PRJ-5000': defaultPRJ5000
            };

            this.loadProjectMetaStorage();
            this.loadTableStorage();

            this.normalizeAllProjectsData();
            this.refreshSheetList();
            this.loadApprovalHistory();
            this.loadWorkflowState();
            this.saveStorage();
        },

        // ==========================================================
        // EXCEL DATA MODEL
        // Data hasil import disimpan 1:1 dengan header Excel.
        // Tidak ada alias/rename/delete kolom pada data import.
        // ==========================================================
        normalizeHeaderExact(value) {
            return String(value ?? '')
                .replace(/\u00A0/g, ' ')
                .trim();
        },

        getExactSchema(sheetName) {
            const schemas = {
                'SP Items': [
                    'Long Description (Family)',
                    'Short Description',
                    'Spec',
                    'Available Size',
                    'Part Subtype',
                    'Item Count'
                ],
                'Support': [
                    'Long Description (Family)',
                    'Compatible Standard',
                    'Manufacturer',
                    'Material',
                    'Material Code',
                    'Long Description (Size)',
                    'Short Description',
                    'Spec',
                    'Size',
                    'Line Number Tag',
                    'Design Std',
                    'Content Iso Symbol Definition',
                    'Design Pressure Factor',
                    'End Type',
                    'Engagement Length',
                    'Facing',
                    'Flange Std',
                    'Gasket Std',
                    'Port Unit',
                    'Matching Pipe OD',
                    'Nominal Diameter',
                    'X Coordinate (Port 1)',
                    'Y Coordinate (Port 1)',
                    'COP Elevation (Port 1)',
                    'Pressure Class',
                    'Schedule',
                    'Status',
                    'Wall Thickness',
                    'Weight',
                    'Weight Unit',
                    'Required Spec',
                    'Insulation Thickness',
                    'Insulation Type',
                    'Service',
                    'Tag',
                    'Tie In Number',
                    'Shop/Field',
                    'DWG Number',
                    'PnPID',
                    'PnPGuid',
                    'Item Code',
                    'Flange Thickness',
                    'Center of Gravity X',
                    'Center of Gravity Y',
                    'Center of Gravity Z',
                    'Tracing Type',
                    'Tracing Spec',
                    'Insulation Spec',
                    'Spool Number',
                    'Unit',
                    'Top of Pipe',
                    'Bottom of Pipe',
                    'Part Subtype',
                    'Support Type',
                    'Support Detail',
                    'Fixed',
                    'Length',
                    'Min Length',
                    'Max Length',
                    'Comment',
                    'Reference'
                ],
                'LineList': [
                    'No',
                    'Line Size (Inch)',
                    'Process Fluid Identifier',
                    'Pipe.Spec',
                    'Seq. No',
                    'Type',
                    'Thickness [mm]',
                    'Complete Line No.',
                    'P&ID No',
                    'From',
                    'To',
                    'Fluid Service',
                    'Phase',
                    'Mass Flow\n[kg/h]',
                    'Volume Flow\n[m1.5/h]',
                    'Operating',
                    'Design',
                    'Operating Temperature',
                    'Design Temperature',
                    'Density\n[kg/m3]',
                    'Viscosity\n[cP]',
                    'RT',
                    'PT/MT',
                    'Medium',
                    'Pressure [Barg]',
                    'Painting Code',
                    'Remarks'
                ],
                'Valve': [
                    'Number',
                    'Long Description (Family)',
                    'Compatible Standard',
                    'Manufacturer',
                    'Material',
                    'Material Code',
                    'Long Description (Size)',
                    'Short Description',
                    'Spec',
                    'Size',
                    'Line Number Tag',
                    'Design Std',
                    'Content Iso Symbol Definition',
                    'Design Pressure Factor',
                    'End Type',
                    'Engagement Length',
                    'Facing',
                    'Flange Std',
                    'Gasket Std',
                    'Port Unit',
                    'Nominal Diameter',
                    'X Coordinate (Port 1)',
                    'Y Coordinate (Port 1)',
                    'COP Elevation (Port 1)',
                    'Pressure Class',
                    'Schedule',
                    'Status',
                    'Wall Thickness',
                    'Weight',
                    'Weight Unit',
                    'Required Spec',
                    'Insulation Thickness',
                    'Insulation Type',
                    'Service',
                    'Tag',
                    'Tie In Number',
                    'Shop/Field',
                    'DWG Number',
                    'PnPID',
                    'PnPGuid',
                    'Item Code',
                    'Matching Pipe OD',
                    'Flange Thickness',
                    'Center of Gravity X',
                    'Center of Gravity Y',
                    'Center of Gravity Z',
                    'Tracing Type',
                    'Tracing Spec',
                    'Insulation Spec',
                    'Spool Number',
                    'Unit',
                    'Top of Pipe',
                    'Bottom of Pipe',
                    'Length',
                    'Valve Alignment',
                    'Valve Detail',
                    'Valve Body Type',
                    'Flow Dependent',
                    'Offset',
                    'Operator Type',
                    'Actuator Type',
                    'Actuator Height',
                    'Actuator Width',
                    'Control Valve',
                    'Valve Code',
                    'Normally',
                    'Failure',
                    'End Connections',
                    'Code'
                ]
            };
            return schemas[sheetName] || null;
        },

        getImportTargetSheet() {
            const name = String(this.activeSheet || '').trim();
            if (name.toLowerCase() === 'support' || name.toLowerCase() === 'pipe support') return 'Support';
            if (name.toLowerCase() === 'linelist' || name.toLowerCase() === 'line list') return 'LineList';
            return name;
        },

        findWorkbookSheet(workbook, targetSheet) {
            const target = String(targetSheet || '').trim().toLowerCase();
            const aliases = {
                'support': ['support', 'pipe support'],
                'linelist': ['linelist', 'line list'],
                'sp items': ['sp items', 'sp_items', 'spitems'],
                'valve': ['valve']
            };
            const accepted = aliases[target] || [target];

            return workbook.SheetNames.find(name =>
                accepted.includes(String(name || '').replace(/\u00A0/g, ' ').trim().toLowerCase())
            ) || null;
        },

        worksheetToExactRows(worksheet) {
            if (!worksheet) throw new Error('Worksheet Excel tidak ditemukan.');

            const matrix = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: true,
                blankrows: false
            });

            // Cari baris header pertama yang mempunyai minimal 2 cell berisi.
            const headerIndex = matrix.findIndex(row =>
                Array.isArray(row) &&
                row.filter(v => String(v ?? '').trim() !== '').length >= 2
            );

            if (headerIndex < 0) {
                throw new Error('Header Excel tidak ditemukan.');
            }

            const rawHeaders = matrix[headerIndex] || [];
            const headers = [];
            const used = new Set();

            rawHeaders.forEach((value, colIndex) => {
                let header = this.normalizeHeaderExact(value);
                if (!header) header = `Column ${colIndex + 1}`;

                // Header duplikat tidak boleh membuat nilai tertimpa.
                let unique = header;
                let n = 2;
                while (used.has(unique.toLowerCase())) {
                    unique = `${header} (${n++})`;
                }
                used.add(unique.toLowerCase());
                headers.push(unique);
            });

            const rows = [];

            for (let r = headerIndex + 1; r < matrix.length; r++) {
                const sourceRow = Array.isArray(matrix[r]) ? matrix[r] : [];
                const hasData = headers.some((_, c) => {
                    const value = sourceRow[c];
                    return value !== null && value !== undefined && String(value).trim() !== '';
                });
                if (!hasData) continue;

                const row = {};
                headers.forEach((header, c) => {
                    row[header] = sourceRow[c] ?? '';
                });
                rows.push(row);
            }

            return { headers, rows, headerIndex };
        },

        normalizeAllProjectsData() {
            Object.keys(this.allProjectsData || {}).forEach(projectKey => {
                const project = this.allProjectsData[projectKey];
                if (!project || typeof project !== 'object') return;

                if (!project.meta || typeof project.meta !== 'object') {
                    project.meta = {
                        projectCode: projectKey,
                        projectName: 'Piping & Equipment',
                        description: '',
                        createdAt: '',
                        createdBy: '',
                        isApproved: false,
                        approvedAt: '',
                        version: 0,
                        progress: '0%'
                    };
                }

                // Hanya pastikan tipe data benar.
                // JANGAN mengubah nama header atau isi cell Excel.
                Object.keys(project).forEach(sheetName => {
                    if (sheetName === 'meta') return;
                    if (!Array.isArray(project[sheetName])) project[sheetName] = [];
                    project[sheetName] = project[sheetName]
                        .filter(row => row && typeof row === 'object');
                });

                this.sheets.forEach(sheetName => {
                    if (!Array.isArray(project[sheetName])) project[sheetName] = [];
                });
            });
        },

        refreshSheetList() {
            const importedSheets = new Set();
            Object.values(this.allProjectsData || {}).forEach(project => {
                if (!project || typeof project !== 'object') return;
                Object.keys(project).forEach(key => {
                    if (key !== 'meta') importedSheets.add(key);
                });
            });
            this.sheets = [...new Set([...this.sheets, ...importedSheets])];
        },

        getTableStorageKey(projectKey, sheetName) {
            return `tripatra_table_v4::${encodeURIComponent(String(projectKey || 'default'))}::${encodeURIComponent(String(sheetName || 'default'))}`;
        },

        getProjectMetaStorageKey() {
            return 'tripatra_project_meta_v4';
        },

        loadProjectMetaStorage() {
            try {
                const raw = localStorage.getItem(this.getProjectMetaStorageKey());
                if (!raw) return;
                const metas = JSON.parse(raw);
                if (!metas || typeof metas !== 'object') return;
                Object.entries(metas).forEach(([projectKey, meta]) => {
                    if (!this.allProjectsData[projectKey]) {
                        this.allProjectsData[projectKey] = this.createBlankProject(
                            projectKey, meta?.projectName || 'Piping & Equipment', meta?.description || ''
                        );
                    }
                    this.allProjectsData[projectKey].meta = {
                        ...this.allProjectsData[projectKey].meta, ...(meta || {})
                    };
                });
            } catch (error) {
                console.warn('Metadata project tidak dapat dibaca:', error);
            }
        },

        loadTableStorage() {
            const prefix = 'tripatra_table_v4::';

            // Baca semua tabel yang pernah disimpan, termasuk project/sheet baru.
            Object.keys(localStorage)
                .filter(key => key.startsWith(prefix))
                .forEach(key => {
                    const rest = key.slice(prefix.length);
                    const parts = rest.split('::');
                    if (parts.length < 2) return;

                    let projectKey = '';
                    let sheetName = '';
                    try {
                        projectKey = decodeURIComponent(parts[0]);
                        sheetName = decodeURIComponent(parts.slice(1).join('::'));
                    } catch (error) {
                        return;
                    }
                    if (!projectKey || !sheetName) return;

                    if (!this.allProjectsData[projectKey]) {
                        this.allProjectsData[projectKey] = this.createBlankProject(projectKey, 'Piping & Equipment', '');
                    }
                    if (!this.sheets.includes(sheetName)) this.sheets.push(sheetName);

                    try {
                        const raw = localStorage.getItem(key);
                        const rows = raw ? JSON.parse(raw) : [];
                        this.allProjectsData[projectKey][sheetName] = Array.isArray(rows) ? rows : [];
                    } catch (error) {
                        this.allProjectsData[projectKey][sheetName] = [];
                    }
                });

            // Pastikan setiap project mempunyai array terpisah untuk setiap sheet.
            Object.values(this.allProjectsData || {}).forEach(project => {
                this.sheets.forEach(sheetName => {
                    if (!Array.isArray(project[sheetName])) project[sheetName] = [];
                });
            });
        },

        saveTableStorage(projectKey, sheetName) {
            if (!projectKey || !sheetName) return;
            const project = this.allProjectsData?.[projectKey];
            if (!project) return;
            const rows = Array.isArray(project[sheetName]) ? project[sheetName] : [];
            localStorage.setItem(this.getTableStorageKey(projectKey, sheetName), JSON.stringify(rows));
        },

        saveStorage() {
            const metas = {};
            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                metas[projectKey] = project?.meta || {};
            });
            localStorage.setItem(this.getProjectMetaStorageKey(), JSON.stringify(metas));

            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                Object.keys(project || {}).forEach(sheetName => {
                    if (sheetName !== 'meta') this.saveTableStorage(projectKey, sheetName);
                });
            });
        },

        switchProject() {
            this.loadTableStorage();
            this.currentPage = 1;
            this.globalSearch = '';
            this.columnFilters = {};
            this.closeColumnFilter();
            this.loadWorkflowState();
        },

        // ==========================================================
        // CREATE PROJECT
        // Project baru selalu dibuat dari struktur kosong.
        // Tidak pernah menyalin data dari project lain.
        // ==========================================================
        openCreateProjectModal() {
            this.createProjectForm = {
                code: '',
                name: '',
                description: '',
                initialization: 'empty',
                file: null,
                fileName: ''
            };
            this.createProjectMessage = '';
            this.createProjectMessageType = 'error';
            this.isCreatingProject = false;
            this.showCreateProjectModal = true;

            this.$nextTick(() => {
                const input = document.querySelector('.create-project-modal input');
                if (input) input.focus();
            });
        },

        closeCreateProjectModal() {
            if (this.isCreatingProject) return;
            this.showCreateProjectModal = false;
            this.createProjectMessage = '';
        },

        handleCreateProjectFile(event) {
            const file = event.target.files?.[0];
            if (!file) {
                this.createProjectForm.file = null;
                this.createProjectForm.fileName = '';
                return;
            }

            const lowerName = file.name.toLowerCase();
            if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
                this.createProjectForm.file = null;
                this.createProjectForm.fileName = '';
                this.createProjectMessageType = 'error';
                this.createProjectMessage = 'File harus berformat .xlsx atau .xls.';
                event.target.value = '';
                return;
            }

            this.createProjectForm.file = file;
            this.createProjectForm.fileName = file.name;
            this.createProjectMessage = '';
        },

        createBlankProject(projectCode, projectName, description = '') {
            const project = {
                meta: {
                    projectCode,
                    projectName,
                    description,
                    createdAt: new Date().toLocaleString('id-ID'),
                    createdBy: this.loginForm.user || 'engineer@tripatra.com',
                    isApproved: false,
                    approvedAt: '',
                    version: 0,
                    progress: '0%'
                }
            };

            this.sheets.forEach(sheetName => {
                project[sheetName] = [];
            });

            return project;
        },

        async createProject() {
            if (this.isCreatingProject) return;

            const code = String(this.createProjectForm.code || '').trim();
            const name = String(this.createProjectForm.name || '').trim();
            const description = String(this.createProjectForm.description || '').trim();

            if (!code) {
                this.createProjectMessageType = 'error';
                this.createProjectMessage = 'Project Code wajib diisi.';
                return;
            }

            if (name.length < 2) {
                this.createProjectMessageType = 'error';
                this.createProjectMessage = 'Project Name wajib diisi minimal 2 karakter.';
                return;
            }

            const duplicate = Object.keys(this.allProjectsData || {})
                .some(key => String(key).trim().toLowerCase() === code.toLowerCase());

            if (duplicate) {
                this.createProjectMessageType = 'error';
                this.createProjectMessage = `Project ${code} sudah ada. Gunakan kode project lain.`;
                return;
            }

            if (this.createProjectForm.initialization === 'import' && !this.createProjectForm.file) {
                this.createProjectMessageType = 'error';
                this.createProjectMessage = 'Pilih file Excel terlebih dahulu untuk opsi Import Excel.';
                return;
            }

            this.isCreatingProject = true;
            this.createProjectMessage = '';

            try {
                const newProject = this.createBlankProject(code, name, description);

                if (this.createProjectForm.initialization === 'empty') {
                    this.allProjectsData[code] = newProject;
                    this.activeSheet = 'Valve';
                } else {
                    const workbook = await this.readExcelWorkbook(this.createProjectForm.file);

                    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                        throw new Error('Workbook tidak memiliki sheet.');
                    }

                    let firstImportedSheet = '';

                    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
                        const worksheet = workbook.Sheets[sheetName];
                        if (!worksheet || !worksheet['!ref']) return;

                        let rows = XLSX.utils.sheet_to_json(worksheet, {
                            defval: '',
                            raw: false,
                            blankrows: false
                        });

                        rows = rows
                            .filter(row => Object.values(row || {}).some(value => String(value).trim() !== ''))
                            .map((row, idx) => this.normalizeRow(row, idx));

                        const cleanSheetName = String(sheetName).trim() || `Sheet ${sheetIndex + 1}`;
                        newProject[cleanSheetName] = rows;

                        if (!firstImportedSheet && rows.length > 0) {
                            firstImportedSheet = cleanSheetName;
                        }
                    });

                    this.allProjectsData[code] = newProject;
                    this.refreshSheetList();
                    this.activeSheet = firstImportedSheet || 'Valve';
                }

                this.activeProject = code;
                this.currentDashboardTab = 'workspace';
                this.currentPage = 1;
                this.globalSearch = '';
                this.saveStorage();

                this.createProjectMessageType = 'success';
                this.createProjectMessage = this.createProjectForm.initialization === 'import'
                    ? `Project ${code} berhasil dibuat dan data Excel berhasil diimport.`
                    : `Project ${code} berhasil dibuat sebagai project kosong.`;

                setTimeout(() => {
                    this.showCreateProjectModal = false;
                    this.createProjectMessage = '';
                    this.isCreatingProject = false;
                }, 700);

            } catch (error) {
                console.error('Create project error:', error);
                this.createProjectMessageType = 'error';
                this.createProjectMessage = error.message || 'Project gagal dibuat.';
                this.isCreatingProject = false;
            }
        },

        readExcelWorkbook(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();

                reader.onload = (event) => {
                    try {
                        const data = new Uint8Array(event.target.result);
                        const workbook = XLSX.read(data, {
                            type: 'array',
                            cellDates: true,
                            raw: false
                        });
                        resolve(workbook);
                    } catch (error) {
                        reject(new Error('File Excel tidak dapat dibaca. Pastikan format file valid.'));
                    }
                };

                reader.onerror = () => reject(new Error('File Excel tidak dapat dibaca oleh browser.'));
                reader.readAsArrayBuffer(file);
            });
        },

        // Backward compatibility untuk kode lama.
        addNewProject() {
            this.openCreateProjectModal();
        },

        openAppModule(appName) {
            this.activeModuleName = appName;
            if (appName === 'Travel') {
                this.moduleDescription = 'Modul Manajemen Perjalanan Dinas, pemesanan tiket pesawat, akomodasi hotel, dan klaim reimbursement.';
            } else if (appName === 'Calendar') {
                this.moduleDescription = 'Kalender Terpadu perusahaan untuk penjadwalan meeting proyek, deadline deliverables, dan milestone engineering.';
            } else if (appName === 'HR Docs') {
                this.moduleDescription = 'Pusat arsip dokumen kepegawaian, surat keputusan (SK), slip gaji, dan sertifikasi keahlian profesional.';
            } else if (appName === 'Policies') {
                this.moduleDescription = 'Direktori Standar Operasional Prosedur (SOP), kebijakan keselamatan kerja (HSE), dan panduan internal perusahaan.';
            } else if (appName === 'Flows') {
                this.moduleDescription = 'Pusat otomatisasi alur kerja digital, formulir persetujuan multi-level, dan pelacakan status dokumen.';
            } else if (appName === 'Perf') {
                this.moduleDescription = 'Modul Penilaian Kinerja Karyawan (Performance Review), penyusunan KPI tahunan, dan feedback berkelanjutan.';
            } else if (appName === 'Attendance') {
                this.moduleDescription = 'Modul Rekapitulasi Kehadiran & Pengajuan Regularisasi Absensi Karyawan.';
            } else if (appName === 'Talent Referral') {
                this.moduleDescription = 'Portal Lowongan Kerja internal dan Program Referensi Talenta Perusahaan.';
            } else {
                this.moduleDescription = 'Modul sistem internal Tripatra BMBQ WebApp.';
            }
            this.showModuleModal = true;
        },

        login() {
            this.loginError = false;
            const email = this.loginForm.user.trim().toLowerCase();
            const pass = this.loginForm.pass.trim();

            if (email === 'engineer@tripatra.com' && pass === 'engineer123') {
                this.loginForm.role = 'Piping Engineer';
                this.isLoggedIn = true;
            } else if (email === 'estimator@tripatra.com' && pass === 'estimator123') {
                this.loginForm.role = 'Estimator Proposal';
                this.isLoggedIn = true;
            } else if (email === 'lead@tripatra.com' && pass === 'lead123') {
                this.loginForm.role = 'Lead Estimator';
                this.isLoggedIn = true;
            } else {
                this.loginError = true;
                this.loginErrorMsg = 'Email atau password salah. Silakan periksa kembali.';
            }
        },

        logout() {
            this.isLoggedIn = false;
            this.loginForm.user = '';
            this.loginForm.pass = '';
        },

        getRoleFocusText() {
            if (this.loginForm.role === 'Piping Engineer') return 'Spesifikasi Teknis & Input Parameter MTO (S3D/E3D/P3D)';
            if (this.loginForm.role === 'Estimator Proposal') return 'Database Harga, Kalkulasi Biaya BOQ & Material (Direct/Indirect Cost)';
            if (this.loginForm.role === 'Lead Estimator') return 'Pengawasan Progres, Approval Flow & Laporan Eksekutif';
            return '';
        },

        getProjectVersion() {
            if (!this.allProjectsData[this.activeProject]?.meta) return 'Rev 0';
            return 'Rev ' + (this.allProjectsData[this.activeProject].meta.version || 0);
        },

        get currentRows() {
            if (!this.allProjectsData[this.activeProject]) {
                this.allProjectsData[this.activeProject] = this.createBlankProject(this.activeProject, 'Piping & Equipment', '');
            }
            if (!Array.isArray(this.allProjectsData[this.activeProject][this.activeSheet])) {
                this.allProjectsData[this.activeProject][this.activeSheet] = [];
            }
            return this.allProjectsData[this.activeProject][this.activeSheet];
        },

        get currentColumns() {
            // PERBAIKAN PERFORMA (delay lama saat pindah ke sheet MTO yang
            // datanya sudah besar, mis. Pipe s.d. Piping and Equipment):
            // Sebelumnya fungsi ini menghitung ulang UNION seluruh header
            // dari SEMUA baris SETIAP kali dipanggil. Karena dipanggil untuk
            // setiap baris yang dirender (bukan cuma sekali per tabel), pada
            // sheet dengan ratusan/ribuan baris hasil import Excel ini jadi
            // sangat berat dan bikin tampilan terasa lambat/nge-lag saat
            // sheet tersebut diklik. Sekarang hasilnya di-cache dan hanya
            // dihitung ulang saat project/sheet aktif berubah atau saat ada
            // perubahan data (tableRenderKey berubah), bukan setiap render.
            const cacheKey = this.activeProject + '::' + this.activeSheet + '::' + this.tableRenderKey;
            if (this._columnsCacheKey === cacheKey) {
                return this._columnsCacheValue;
            }

            const result = this.computeCurrentColumns();
            this._columnsCacheKey = cacheKey;
            this._columnsCacheValue = result;
            return result;
        },

        computeCurrentColumns() {
            const exactSchema = this.getExactSchema(this.activeSheet);
            if (exactSchema) {
                // Number/No tetap tersimpan, tetapi kolom nomor sudah disediakan
                // oleh kolom UI paling kiri.
                return exactSchema.filter(col => col !== 'Number' && col !== 'No');
            }

            const rows = this.currentRows || [];
            if (!rows.length) {
                return ['Material', 'Material Code', 'Spec', 'Size', 'Pressure Class', 'Line Number Tag', 'Status'];
            }

            // Untuk sheet MTO, ambil UNION seluruh header dari semua baris.
            // Urutan tetap mengikuti kemunculan pertama di Excel, sehingga
            // data yang hanya terisi pada baris tertentu tidak pernah hilang
            // hanya karena kolom tersebut kosong pada baris pertama.
            const ordered = [];
            const seen = new Set();
            rows.forEach(row => {
                Object.keys(row || {}).forEach(col => {
                    if (col === 'Number' || col === 'No') return;
                    const key = String(col);
                    const normalized = key.toLowerCase();
                    if (seen.has(normalized)) return;
                    seen.add(normalized);
                    ordered.push(key);
                });
            });
            return ordered;
        },

        get freezeColumns() {
            // FREEZE HARUS BERDASARKAN URUTAN ASLI EXCEL.
            // Tidak pernah memindahkan, mengurutkan, atau menyisipkan kolom.
            // Kolom No selalu berada paling kiri dan ikut freeze.
            const columns = this.currentColumns || [];
            if (!columns.length) return [];

            // MTO: batas freeze tepat di MANUFACTURER.
            // Artinya: semua kolom dari paling kiri sampai Manufacturer diam,
            // sedangkan kolom setelah Manufacturer bergerak bersama scrollbar.
            // Sheet khusus tetap memakai batas yang sebelumnya sudah disepakati.
            const targetBySheet = {
                'SP Items': 'Spec',
                'Support': 'Material',
                'LineList': 'Pipe.Spec'
            };

            const target = targetBySheet[this.activeSheet] || 'Manufacturer';
            const targetIndex = columns.indexOf(target);

            // Jika Manufacturer tidak ditemukan pada suatu MTO sheet, jangan
            // memaksakan sticky yang bisa menyebabkan overlap. Freeze hanya No.
            if (targetIndex < 0) return [];

            return columns.slice(0, targetIndex + 1);
        },

        isFreezeColumn(column) {
            return this.freezeColumns.includes(column);
        },

        getNoFreezeClass() {
            const noClass = 'freeze-col freeze-col-0';
            return this.freezeColumns.length === 0 ? `${noClass} freeze-last` : noClass;
        },

        getFreezeClass(column) {
            const freezeIndex = this.freezeColumns.indexOf(column);
            if (freezeIndex === -1) return '';
            const isLast = freezeIndex === this.freezeColumns.length - 1;
            return `freeze-col freeze-col-${freezeIndex + 1}${isLast ? ' freeze-last' : ''}`;
        },

        get filteredRows() {
            let rows = this.currentRows;

            // Search global tetap bekerja seperti sebelumnya.
            if (this.globalSearch) {
                const q = this.globalSearch.toLowerCase().trim();
                if (q) {
                    rows = rows.filter(row =>
                        Object.values(row || {}).some(val =>
                            String(val ?? '').toLowerCase().includes(q)
                        )
                    );
                }
            }

            // Filter per kolom.
            Object.entries(this.columnFilters || {}).forEach(([column, selectedValues]) => {
                if (!Array.isArray(selectedValues)) return;

                rows = rows.filter(row => {
                    const rawValue = row?.[column];
                    const normalized = this.normalizeFilterValue(rawValue);
                    return selectedValues.includes(normalized);
                });
            });

            return rows;
        },

        get filteredFilterOptions() {
            const q = String(this.filterOptionSearch || '').toLowerCase().trim();
            const source = Array.isArray(this.filterOptions) ? this.filterOptions : [];

            if (!q) return source.slice(0, 300);

            return source
                .filter(value => this.filterDisplayValue(value).toLowerCase().includes(q))
                .slice(0, 300);
        },

        get allFilterValuesSelected() {
            return this.filterOptions.length > 0 &&
                this.filterOptions.every(value => this.filterDraftValues.includes(value));
        },

        get maxPage() {
            return Math.ceil(this.filteredRows.length / this.itemsPerPage) || 1;
        },

        get paginatedRows() {
            const start = (this.currentPage - 1) * this.itemsPerPage;
            return this.filteredRows.slice(start, start + this.itemsPerPage);
        },

        normalizeFilterValue(value) {
            const text = String(value ?? '').trim();
            return text === '' ? '__EMPTY__' : text;
        },

        filterDisplayValue(value) {
            return value === '__EMPTY__' ? '(Kosong)' : String(value);
        },

        isColumnFiltered(column) {
            return Object.prototype.hasOwnProperty.call(this.columnFilters || {}, column);
        },

        getColumnFilterOptions(column) {
            const seen = new Set();
            const values = [];

            (this.currentRows || []).forEach(row => {
                const value = this.normalizeFilterValue(row?.[column]);
                if (!seen.has(value)) {
                    seen.add(value);
                    values.push(value);
                }
            });

            values.sort((a, b) => {
                if (a === '__EMPTY__') return 1;
                if (b === '__EMPTY__') return -1;
                return String(a).localeCompare(String(b), undefined, {
                    numeric: true,
                    sensitivity: 'base'
                });
            });

            return values;
        },

        openColumnFilter(column, event) {
            if (!column) return;

            this.activeFilterColumn = column;
            this.filterOptionSearch = '';
            this.filterOptions = this.getColumnFilterOptions(column);

            // Jika belum pernah difilter, semua nilai dipilih.
            // Jika sudah pernah difilter, popup membuka pilihan terakhir.
            if (this.isColumnFiltered(column)) {
                this.filterDraftValues = [...(this.columnFilters[column] || [])];
            } else {
                this.filterDraftValues = [...this.filterOptions];
            }

            const button = event?.currentTarget || event?.target;
            const rect = button?.getBoundingClientRect?.();

            if (rect) {
                const popupWidth = 300;
                const popupHeight = Math.min(430, window.innerHeight - 24);
                let left = rect.left;
                let top = rect.bottom + 6;

                if (left + popupWidth > window.innerWidth - 12) {
                    left = Math.max(12, window.innerWidth - popupWidth - 12);
                }

                if (top + popupHeight > window.innerHeight - 12) {
                    top = Math.max(12, rect.top - popupHeight - 6);
                }

                this.filterPopupPosition = { left, top };
            }

            this.filterPopupOpen = true;
        },

        closeColumnFilter() {
            this.filterPopupOpen = false;
            this.filterOptionSearch = '';
        },

        toggleAllFilterValues() {
            if (this.allFilterValuesSelected) {
                this.filterDraftValues = [];
            } else {
                this.filterDraftValues = [...this.filterOptions];
            }
        },

        clearColumnFilter(column = this.activeFilterColumn) {
            if (!column) return;

            const next = { ...(this.columnFilters || {}) };
            delete next[column];
            this.columnFilters = next;
            this.currentPage = 1;
            this.filterPopupOpen = false;
            this.tableRenderKey++;

            this.$nextTick(() => this.updateTableScrollbar());
        },

        applyColumnFilter() {
            const column = this.activeFilterColumn;
            if (!column) return;

            const selected = [...this.filterDraftValues];
            const allSelected =
                this.filterOptions.length > 0 &&
                selected.length === this.filterOptions.length;

            const next = { ...(this.columnFilters || {}) };

            if (allSelected) {
                delete next[column];
            } else {
                next[column] = selected;
            }

            this.columnFilters = next;
            this.currentPage = 1;
            this.filterPopupOpen = false;
            this.tableRenderKey++;

            this.$nextTick(() => this.updateTableScrollbar());
        },

        resetAllColumnFilters() {
            this.columnFilters = {};
            this.currentPage = 1;
            this.filterPopupOpen = false;
            this.filterOptionSearch = '';
            this.tableRenderKey++;

            this.$nextTick(() => this.updateTableScrollbar());
        },

                openDataSheet(sheetName) {
        this.saveTableStorage(this.activeProject, this.activeSheet);
        this.activeSheet = sheetName;
        this.currentDashboardTab = 'workspace';
        this.currentPage = 1;
        this.globalSearch = '';
        this.columnFilters = {};
        this.closeColumnFilter();
        this.tableRenderKey++;

        this.$nextTick(() => {
            this.updateTableScrollbar();
        });
    },

    // ==========================================================
    // TASK MANAGEMENT
    // Membuka tugas MTO Valve
    // ==========================================================
    openTaskMTOValve() {
        // Tampilkan informasi tugas terlebih dahulu
        alert('Membuka detail tugas MTO Valve...');

        // Setelah user menekan OK pada alert,
        // otomatis diarahkan ke halaman MTO & Equipment Suite -> Valve
        this.activeSheet = 'Valve';
        this.currentDashboardTab = 'workspace';
        this.currentPage = 1;
        this.globalSearch = '';

        // Pastikan tabel sudah dirender sebelum update scrollbar
        this.$nextTick(() => {
            this.updateTableScrollbar();
        });
    },

        openAddModalForActiveSheet() {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Akses ditolak! Hanya Piping Engineer yang dapat menambah data.');
                return;
            }

            const project = this.allProjectsData?.[this.activeProject];
            if (!project) {
                alert('Project aktif tidak ditemukan.');
                return;
            }

            const sheet = this.activeSheet;
            if (!Array.isArray(project[sheet])) project[sheet] = [];

            /*
             * MODE TAMBAH DATA INLINE:
             * Data lama TIDAK dihapus dari storage.
             * Saat tombol Add data diklik, tabel sementara hanya
             * menampilkan SATU BARIS KOSONG untuk diisi langsung.
             */
            const nextNum = String(project[sheet].length + 1).padStart(2, '0');
            const draft = {
                Number: nextNum,
                Status: 'New',
                __inlineDraft: true
            };

            // Ambil SEMUA kolom sheet yang sedang aktif.
            const columns = this.currentColumns || [];
            columns.forEach(col => {
                if (!(col in draft)) draft[col] = '';
            });

            this.newInlineRow = draft;
            this.addingNewData = true;
            this.globalSearch = '';
            this.currentPage = 1;
            this.tableRenderKey++;

            this.$nextTick(() => {
                this.updateTableScrollbar();

                const firstInput = document.querySelector(
                    '.excel-table tbody tr.inline-new-row input'
                );

                if (firstInput) {
                    firstInput.focus();
                    firstInput.select?.();
                }
            });
        },

        cancelInlineAdd() {
            // Batalkan hanya baris draft. Semua data lama tetap ada.
            this.addingNewData = false;
            this.newInlineRow = {};
            this.tableRenderKey++;
            this.$nextTick(() => this.updateTableScrollbar());
        },

        saveInlineAdd() {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Akses ditolak! Hanya Piping Engineer yang dapat menambah data.');
                return;
            }

            if (!this.addingNewData || !this.newInlineRow) return;

            const draft = { ...this.newInlineRow };
            delete draft.__inlineDraft;
            delete draft.__edited;

            const hasValue = Object.entries(draft).some(([key, value]) => {
                return key !== 'Number' && String(value ?? '').trim() !== '';
            });

            if (!hasValue) {
                alert('Silakan isi minimal satu data pada baris baru.');
                return;
            }

            const rows = this.currentRows;

            // Nomor mengikuti jumlah data terakhir.
            draft.Number = String(rows.length + 1).padStart(2, '0');
            draft.Status = draft.Status || 'New';

            rows.push(draft);

            // Rapikan nomor setelah penambahan.
            rows.forEach((row, idx) => {
                row.Number = String(idx + 1).padStart(2, '0');
            });

            this.saveTableStorage(this.activeProject, this.activeSheet);
            this.saveStorage();

            this.addingNewData = false;
            this.newInlineRow = {};
            this.currentPage = Math.max(1, Math.ceil(rows.length / this.itemsPerPage));
            this.tableRenderKey++;

            this.$nextTick(() => {
                this.updateTableScrollbar();
            });
        },

        markRowEdited(row) {
            if (!row) return;
            row.__edited = true;
            this.saveTableStorage(this.activeProject, this.activeSheet);
            this.saveStorage();
        },

        resetNewRowForm() {
            this.newRowForm = {
                longDesc: '', material: 'CS', spec: 'CS150', size: '2"', pressure: '150',
                wasteFactor: '5', tag: '', qty: '', unit: '', service: '', supportType: ''
            };
        },

        getDisplayRowNumber(row, index) {
            // Untuk sheet equipment (Vessel s.d. Piping and Equipment),
            // kolom NO adalah nomor urut tabel, bukan Equipment/Tag ID dari Excel.
            const equipmentSheets = new Set([
                'Vessel', 'Tank', 'Pump', 'Misc Equipment',
                'Equipment', 'Piping and Equipment'
            ]);

            if (equipmentSheets.has(String(this.activeSheet || '').trim())) {
                return String((this.currentPage - 1) * this.itemsPerPage + index + 1).padStart(2, '0');
            }

            // Sheet lain tetap mempertahankan nomor asli dari Excel.
            if (row?.Number !== undefined && row?.Number !== '') return row.Number;
            if (row?.No !== undefined && row?.No !== '') return row.No;
            return String((this.currentPage - 1) * this.itemsPerPage + index + 1).padStart(2, '0');
        },

        changeSheet(sheetName) {
            const name = String(sheetName || '').replace(/\u00A0/g, ' ').trim();
            if (!name) return;

            const project = this.allProjectsData?.[this.activeProject];

            // Pastikan kategori yang dipilih selalu mempunyai array data.
            if (project && !Array.isArray(project[name])) {
                project[name] = [];
            }

            // Pastikan kategori tetap terdaftar di navigasi.
            if (!this.sheets.includes(name)) {
                this.sheets = [...this.sheets, name];
            }

            this.activeSheet = name;
            this.currentDashboardTab = 'workspace';
            this.currentPage = 1;
            this.globalSearch = '';
            this.columnFilters = {};
            this.closeColumnFilter();
            this.tableRenderKey++;

            this.$nextTick(() => this.updateTableScrollbar());
        },

        updateTableScrollbar() {
            this.$nextTick(() => {
                const container = this.$refs?.tableScroll || document.querySelector('.workspace-table-card .excel-table-container');
                const table = container?.querySelector('.excel-table');
                const bottom = this.$refs?.tableBottomScrollbar || document.querySelector('.workspace-table-card .table-bottom-scrollbar');
                const spacer = this.$refs?.tableBottomScrollbarSpacer || document.querySelector('.workspace-table-card .table-bottom-scrollbar-spacer');
                if (!container || !table || !bottom || !spacer) return;

                // Freeze kolom "No" (index 0) + kolom-kolom dari freezeColumns
                // (index 1, 2, ...), disusun berjejer sesuai lebar kolom
                // sebelumnya supaya tidak saling menumpuk.
                const freezeColumnCount = (this.freezeColumns || []).length;

                for (let i = 0; i <= freezeColumnCount; i++) {
                    const headerCell = table.querySelector(`thead th.freeze-col-${i}`);
                    if (!headerCell) continue;

                    // Offset = total lebar semua kolom freeze sebelumnya.
                    let offset = 0;
                    for (let j = 0; j < i; j++) {
                        const prevHeader = table.querySelector(`thead th.freeze-col-${j}`);
                        if (prevHeader) offset += Math.ceil(prevHeader.getBoundingClientRect().width);
                    }

                    const offsetPx = `${offset}px`;
                    headerCell.style.left = offsetPx;
                    headerCell.style.setProperty('--freeze-left', offsetPx);

                    table.querySelectorAll(`tbody td.freeze-col-${i}`).forEach(cell => {
                        cell.style.left = offsetPx;
                        cell.style.setProperty('--freeze-left', offsetPx);
                    });
                }

                const headerCells = Array.from(table.querySelectorAll('thead tr:first-child > th'));
                let naturalTableWidth = headerCells.reduce((total, cell) => total + Math.ceil(cell.getBoundingClientRect().width), 0);

                Array.from(table.querySelectorAll('tbody tr')).forEach(row => {
                    const rowWidth = Array.from(row.children).reduce((total, cell) => total + Math.ceil(cell.getBoundingClientRect().width), 0);
                    naturalTableWidth = Math.max(naturalTableWidth, rowWidth);
                });

                const tableWidth = Math.max(naturalTableWidth, Math.ceil(table.scrollWidth), Math.ceil(table.offsetWidth), container.clientWidth);
                table.style.width = `${tableWidth}px`;
                table.style.minWidth = `${tableWidth}px`;
                table.style.maxWidth = 'none';
                spacer.style.width = `${tableWidth}px`;
                spacer.style.minWidth = `${tableWidth}px`;
                spacer.style.maxWidth = 'none';
                bottom.style.display = 'block';

                const maxScroll = Math.max(0, tableWidth - container.clientWidth);
                if (container.scrollLeft > maxScroll) container.scrollLeft = maxScroll;
                if (Math.abs(bottom.scrollLeft - container.scrollLeft) > 0.5) bottom.scrollLeft = container.scrollLeft;
            });
        },

        syncTableBottomScroll(event) {
            const container =
                this.$refs?.tableScroll ||
                document.querySelector('.workspace-table-card .excel-table-container');

            const bottom =
                this.$refs?.tableBottomScrollbar ||
                document.querySelector('.workspace-table-card .table-bottom-scrollbar');

            if (!container || !bottom) return;

            const left = event.target.scrollLeft;

            // Tabel -> scrollbar bawah
            if (Math.abs(bottom.scrollLeft - left) > 0.5) {
                bottom.scrollLeft = left;
            }
        },

        syncBottomTableScroll(event) {
            const bottom =
                this.$refs?.tableBottomScrollbar ||
                document.querySelector('.workspace-table-card .table-bottom-scrollbar');

            const container =
                this.$refs?.tableScroll ||
                document.querySelector('.workspace-table-card .excel-table-container');

            if (!bottom || !container) return;

            // Scrollbar bawah -> TABEL
            container.scrollLeft = event.target.scrollLeft;
        },


        importExcelFile(event) {
            const input = event?.target;
            const file = input?.files?.[0];
            if (!file) return;

            const resetInput = () => {
                if (input) input.value = '';
            };

            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const buffer = e.target.result;
                    if (!buffer) {
                        throw new Error('File Excel kosong atau tidak dapat dibaca.');
                    }

                    const workbook = XLSX.read(new Uint8Array(buffer), {
                        type: 'array',
                        cellDates: false,
                        cellNF: false,
                        cellText: false,
                        raw: true
                    });

                    if (!workbook.SheetNames?.length) {
                        throw new Error('Workbook tidak memiliki worksheet.');
                    }

                    const projectKey = this.activeProject;
                    if (!projectKey) {
                        throw new Error('Project aktif belum dipilih.');
                    }

                    if (!this.allProjectsData[projectKey]) {
                        this.allProjectsData[projectKey] =
                            this.createBlankProject(projectKey, 'Piping & Equipment', '');
                    }

                    const targetSheet = this.getImportTargetSheet();

                    /*
                     * ======================================================
                     * ATURAN IMPORT FINAL
                     * ======================================================
                     *
                     * 1. SP Items       -> hanya sheet SP Items
                     * 2. Pipe Support   -> hanya sheet Pipe Support/Support
                     * 3. Line List      -> hanya sheet Line List/LineList
                     * 4. MTO workbook   -> IMPORT SEMUA SHEET MTO
                     *
                     * Ini yang sebelumnya menyebabkan Tee, Pipe, Elbow,
                     * Flange, dll tetap kosong: import hanya mengambil
                     * worksheet yang sedang aktif (Valve).
                     *
                     * Project 3000-Piping and Equipment.xlsx memang merupakan
                     * workbook master yang berisi seluruh kategori MTO.
                     * Saat workbook itu diimport dari area MTO, semua sheet
                     * MTO dimasukkan ke tabelnya masing-masing.
                     *
                     * Sheet Support sengaja tidak ditimpa oleh workbook MTO
                     * karena tabel Pipe Support mempunyai sumber exact
                     * tersendiri: Pipe_Support_Project_3000_EXACT.xlsx.
                     * ======================================================
                     */

                    const dedicatedTargets = new Set([
                        'SP Items',
                        'Support',
                        'LineList'
                    ]);

                    const mtoSheets = new Set([
                        'Valve',
                        'Tee',
                        'Single Branch Fitting',
                        'Pipe',
                        'Nozzle',
                        'Instrument',
                        'Flange',
                        'Elbow',
                        'Coupling',
                        'Pipe Run Component',
                        'Tap Weld',
                        'Socketweld',
                        'Gasket',
                        'Buttweld',
                        'Bolt Set',
                        'Fasteners',
                        'Vessel',
                        'Tank',
                        'Pump',
                        'Misc Equipment',
                        'Equipment',
                        'Piping and Equipment'
                    ]);

                    const normalizeSheetName = (name) =>
                        String(name || '')
                            .replace(/\u00A0/g, ' ')
                            .trim();

                    const lower = (name) =>
                        normalizeSheetName(name).toLowerCase();

                    const findExactWorkbookSheet = (aliases) => {
                        const wanted = aliases.map(lower);
                        return workbook.SheetNames.find(name =>
                            wanted.includes(lower(name))
                        ) || null;
                    };

                    // ------------------------------------------------------
                    // MODE A: dedicated table (SP Items / Pipe Support /
                    // Line List). Hanya satu sheet yang diproses.
                    // ------------------------------------------------------
                    if (dedicatedTargets.has(targetSheet)) {
                        let aliases;

                        if (targetSheet === 'SP Items') {
                            aliases = ['SP Items', 'SP_Items', 'SPItems'];
                        } else if (targetSheet === 'Support') {
                            aliases = ['Pipe Support', 'Support'];
                        } else {
                            aliases = ['Line List', 'LineList'];
                        }

                        let actualSheetName = findExactWorkbookSheet(aliases);

                        if (!actualSheetName && workbook.SheetNames.length === 1) {
                            actualSheetName = workbook.SheetNames[0];
                        }

                        if (!actualSheetName) {
                            throw new Error(
                                `Worksheet untuk "${targetSheet}" tidak ditemukan.\n\n` +
                                `Worksheet tersedia:\n${workbook.SheetNames.join(', ')}`
                            );
                        }

                        const parsed = this.worksheetToExactRows(
                            workbook.Sheets[actualSheetName]
                        );

                        if (!parsed.rows.length) {
                            throw new Error(
                                `Worksheet "${actualSheetName}" tidak mempunyai data.`
                            );
                        }

                        const expected = this.getExactSchema(targetSheet);

                        if (expected) {
                            const normalizedHeaders = parsed.headers.map(h =>
                                String(h)
                                    .replace(/\r\n/g, '\n')
                                    .trim()
                                    .toLowerCase()
                            );

                            const matched = expected.filter(h =>
                                normalizedHeaders.includes(
                                    String(h)
                                        .replace(/\r\n/g, '\n')
                                        .trim()
                                        .toLowerCase()
                                )
                            ).length;

                            const ratio = expected.length
                                ? matched / expected.length
                                : 1;

                            if (ratio < 0.5) {
                                throw new Error(
                                    `Struktur Excel tidak cocok dengan tabel ${targetSheet}.\n\n` +
                                    `Header terbaca:\n${parsed.headers.join(' | ')}`
                                );
                            }
                        }

                        this.allProjectsData[projectKey][targetSheet] =
                            parsed.rows.map(row => ({ ...row }));

                        this.activeSheet = targetSheet;

                        this.currentDashboardTab = 'workspace';
                        this.currentPage = 1;
                        this.globalSearch = '';
                        this.columnFilters = {};
                        this.closeColumnFilter();
                        this.tableRenderKey++;

                        this.saveStorage();

                        this.$nextTick(() => {
                            requestAnimationFrame(() => {
                                this.updateTableScrollbar();
                            });
                        });

                        console.log(
                            `[IMPORT EXACT] ${targetSheet}:`,
                            actualSheetName,
                            parsed.rows.length,
                            'rows',
                            parsed.headers.length,
                            'columns'
                        );

                        alert(
                            `Import berhasil!\n\n` +
                            `Project: ${projectKey}\n` +
                            `Tabel: ${targetSheet}\n` +
                            `Worksheet: ${actualSheetName}\n` +
                            `Kolom: ${parsed.headers.length}\n` +
                            `Data: ${parsed.rows.length} baris`
                        );

                        return;
                    }

                    // ------------------------------------------------------
                    // MODE B: MTO
                    // ------------------------------------------------------
                    //
                    // Jika workbook berisi beberapa sheet MTO, import SEMUA
                    // sheet yang dikenali. Ini membuat navigasi Tee, Pipe,
                    // Elbow, Flange, dst. tidak lagi kosong.
                    //
                    // Jika workbook hanya berisi satu sheet, hanya sheet itu
                    // yang dimasukkan.
                    // ------------------------------------------------------
                    const availableMtoSheets = workbook.SheetNames
                        .map(normalizeSheetName)
                        .filter(name => mtoSheets.has(name));

                    if (availableMtoSheets.length > 0) {
                        const imported = [];
                        const skipped = [];

                        availableMtoSheets.forEach(actualSheetName => {
                            try {
                                const parsed = this.worksheetToExactRows(
                                    workbook.Sheets[actualSheetName]
                                );

                                if (!parsed.rows.length) {
                                    skipped.push(`${actualSheetName} (kosong)`);
                                    return;
                                }

                                // Pastikan target sheet sudah ada di navigasi.
                                if (!this.sheets.includes(actualSheetName)) {
                                    this.sheets = [...this.sheets, actualSheetName];
                                }

                                // Simpan setiap sheet ke storage sheet-nya
                                // sendiri. Tidak menumpuk ke Valve.
                                this.allProjectsData[projectKey][actualSheetName] =
                                    parsed.rows.map(row => ({ ...row }));

                                imported.push({
                                    sheet: actualSheetName,
                                    rows: parsed.rows.length,
                                    columns: parsed.headers.length
                                });
                            } catch (sheetError) {
                                skipped.push(
                                    `${actualSheetName} (${sheetError.message})`
                                );
                            }
                        });

                        if (!imported.length) {
                            throw new Error(
                                'Tidak ada sheet MTO yang berhasil diimport.'
                            );
                        }

                        // Jangan sentuh Support/Pipe Support di sini.
                        // Sumbernya adalah file Pipe Support exact terpisah.

                        // Tampilkan Valve jika ada; kalau tidak, tampilkan
                        // sheet MTO pertama yang berhasil diimport.
                        const valveImported = imported.some(
                            item => item.sheet === 'Valve'
                        );

                        this.activeSheet = valveImported
                            ? 'Valve'
                            : imported[0].sheet;

                        this.currentDashboardTab = 'workspace';
                        this.currentPage = 1;
                        this.globalSearch = '';
                        this.columnFilters = {};
                        this.closeColumnFilter();
                        this.tableRenderKey++;

                        this.saveStorage();

                        this.$nextTick(() => {
                            requestAnimationFrame(() => {
                                this.updateTableScrollbar();
                                setTimeout(() => {
                                    this.updateTableScrollbar();
                                }, 150);
                            });
                        });

                        console.table(imported);

                        const summary = imported
                            .map(item =>
                                `• ${item.sheet}: ${item.rows} baris × ${item.columns} kolom`
                            )
                            .join('\n');

                        alert(
                            `Import MTO berhasil!\n\n` +
                            `Project: ${projectKey}\n\n` +
                            `Sheet yang berhasil dimuat:\n${summary}` +
                            (skipped.length
                                ? `\n\nSheet yang dilewati:\n• ${skipped.join('\n• ')}`
                                : '') +
                            `\n\nSetiap sheet disimpan ke tabelnya masing-masing.`
                        );

                        return;
                    }

                    // ------------------------------------------------------
                    // MODE C: workbook satu sheet MTO non-standar.
                    // ------------------------------------------------------
                    if (workbook.SheetNames.length === 1) {
                        const actualSheetName = workbook.SheetNames[0];
                        const parsed = this.worksheetToExactRows(
                            workbook.Sheets[actualSheetName]
                        );

                        if (!parsed.rows.length) {
                            throw new Error(
                                `Worksheet "${actualSheetName}" tidak mempunyai data.`
                            );
                        }

                        const target = targetSheet || actualSheetName;

                        this.allProjectsData[projectKey][target] =
                            parsed.rows.map(row => ({ ...row }));

                        if (!this.sheets.includes(target)) {
                            this.sheets = [...this.sheets, target];
                        }

                        this.activeSheet = target;
                        this.currentDashboardTab = 'workspace';
                        this.currentPage = 1;
                        this.globalSearch = '';
                        this.columnFilters = {};
                        this.closeColumnFilter();
                        this.tableRenderKey++;

                        this.saveStorage();

                        this.$nextTick(() => {
                            requestAnimationFrame(() => {
                                this.updateTableScrollbar();
                            });
                        });

                        alert(
                            `Import berhasil!\n\n` +
                            `Tabel: ${target}\n` +
                            `Worksheet: ${actualSheetName}\n` +
                            `Data: ${parsed.rows.length} baris`
                        );

                        return;
                    }

                    throw new Error(
                        `Workbook tidak dikenali sebagai workbook MTO atau tabel khusus.\n\n` +
                        `Worksheet tersedia: ${workbook.SheetNames.join(', ')}`
                    );

                } catch (error) {
                    console.error('[IMPORT FIX] Gagal:', error);
                    alert(
                        'Import Excel gagal.\n\n' +
                        (error?.message || 'Format Excel tidak sesuai.')
                    );
                } finally {
                    resetInput();
                }
            };

            reader.onerror = () => {
                resetInput();
                alert('File Excel tidak dapat dibaca oleh browser.');
            };

            reader.readAsArrayBuffer(file);
        },

        exportExcel() {
            const rows = this.currentRows;
            if (!rows || rows.length === 0) return alert(`Tidak ada data pada sheet ${this.activeSheet}.`);
            const worksheet = XLSX.utils.json_to_sheet(rows);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, this.activeSheet.substring(0, 31));
            XLSX.writeFile(workbook, `${this.activeProject}_${this.activeSheet.replace(/\s+/g, '_')}.xlsx`);
        },

        // ==========================================================
        // WORKFLOW MTO -> ESTIMATOR -> LEAD / SPV
        // ==========================================================
        get workflowStorageKey() {
            return `tripatra_workflow_v1::${encodeURIComponent(String(this.activeProject || 'PRJ-3000'))}`;
        },

        loadApprovalHistory() {
            // Riwayat approval sekarang disimpan per project pada workflow storage.
            this.approvalHistory = [];
        },

        loadWorkflowState() {
            try {
                const raw = localStorage.getItem(this.workflowStorageKey);
                if (!raw) {
                    this.workflowStatus = 'DRAFT';
                    this.approvalHistory = [];
                    return;
                }
                const data = JSON.parse(raw) || {};
                this.workflowStatus = data.status || 'DRAFT';
                this.approvalHistory = Array.isArray(data.history) ? data.history : [];
                this.costRows = Array.isArray(data.costRows) ? data.costRows : [];
                this.qtoRows = Array.isArray(data.qtoRows) ? data.qtoRows : [];
                this.qtoSourceSheet = data.qtoSourceSheet || '';
                this.qtoSubmittedAt = data.qtoSubmittedAt || '';
            } catch (e) {
                console.warn('Workflow tidak dapat dibaca:', e);
                this.workflowStatus = 'DRAFT';
                this.approvalHistory = [];
                this.costRows = [];
                this.qtoRows = [];
                this.qtoSourceSheet = '';
                this.qtoSubmittedAt = '';
            }
        },

        saveWorkflowState() {
            const payload = {
                status: this.workflowStatus,
                history: this.approvalHistory || [],
                costRows: this.costRows || [],
                qtoRows: this.qtoRows || [],
                qtoSourceSheet: this.qtoSourceSheet || '',
                qtoSubmittedAt: this.qtoSubmittedAt || '',
                updatedAt: new Date().toISOString()
            };
            localStorage.setItem(this.workflowStorageKey, JSON.stringify(payload));
        },

        get projectApprovalHistory() {
            return Array.isArray(this.approvalHistory) ? this.approvalHistory : [];
        },

        get latestApproval() {
            const h = this.projectApprovalHistory;
            return h.length ? h[h.length - 1] : null;
        },

        get activeTaskCount() {
            const r = this.loginForm.role;
            if (r === 'Piping Engineer') return ['DRAFT', 'REVISION_REQUESTED_TO_ENGINEER'].includes(this.workflowStatus) ? 1 : 0;
            if (r === 'Estimator Proposal') return ['SUBMITTED_TO_ESTIMATOR', 'REVISION_REQUESTED_TO_ESTIMATOR'].includes(this.workflowStatus) ? 1 : 0;
            if (r === 'Lead Estimator') return this.workflowStatus === 'SUBMITTED_TO_LEAD' ? 1 : 0;
            return 0;
        },

        getApprovalStatusText(status) {
            const map = {
                DRAFT: 'Draft',
                SUBMITTED_TO_ESTIMATOR: 'Menunggu Estimator',
                SUBMITTED_TO_LEAD: 'Menunggu Review Lead',
                APPROVED: 'Approved',
                REJECTED: 'Rejected',
                REVISION_REQUESTED: 'Minta Revisi',
                REVISION_REQUESTED_TO_ESTIMATOR: 'Revisi BOQ untuk Estimator',
                REVISION_REQUESTED_TO_ENGINEER: 'Revisi MTO untuk Engineer'
            };
            return map[status] || status || 'Menunggu';
        },

        getApprovalStatusClass(status) {
            if (status === 'APPROVED') return 'bg-emerald-100 text-emerald-700';
            if (status === 'REJECTED') return 'bg-rose-100 text-rose-700';
            if (status === 'REVISION_REQUESTED') return 'bg-amber-100 text-amber-700';
            if (status === 'SUBMITTED_TO_LEAD') return 'bg-sky-100 text-sky-700';
            return 'bg-slate-100 text-slate-700';
        },

        createWorkflowHistory(status, title, notes = '') {
            const item = {
                id: Date.now(),
                title,
                projectId: this.activeProject,
                sheet: this.activeSheet,
                revision: this.allProjectsData?.[this.activeProject]?.meta?.version || 0,
                reviewer: this.loginForm.role,
                status,
                notes,
                timestamp: new Date().toLocaleString('id-ID')
            };
            this.approvalHistory.push(item);
            this.saveWorkflowState();
            return item;
        },

        generateBOQ() {
            if (this.loginForm.role !== 'Estimator Proposal') {
                alert('Fitur Kalkulasi Biaya BOQ hanya dapat dijalankan oleh Estimator.');
                return;
            }
            let groups = Array.isArray(this.qtoRows) ? this.qtoRows : [];
            // Kompatibilitas dengan workflow lama: jika snapshot QTO belum ada tetapi status sudah dikirim,
            // ambil data sheet aktif sekali lalu simpan sebagai snapshot.
            if (!groups.length && this.workflowStatus === 'SUBMITTED_TO_ESTIMATOR') {
                groups = this.getBomGroups();
                if (groups.length) {
                    this.qtoRows = groups.map(g => ({ ...g }));
                    this.qtoSourceSheet = this.activeSheet;
                    this.qtoSubmittedAt = new Date().toISOString();
                    this.saveWorkflowState();
                }
            }
            if (!groups.length) {
                alert('Belum ada Quantity Take-Off dari Engineer. Engineer harus mengirim MTO terlebih dahulu.');
                return;
            }
            this.costRows = groups.map((g, i) => ({
                no: i + 1,
                description: g.description,
                material: g.material,
                size: g.size,
                spec: g.spec,
                quantity: g.quantity,
                unit: 'EA',
                unitPrice: Number((this.costRows[i]?.unitPrice) || 0),
                total: Number((this.costRows[i]?.unitPrice) || 0) * g.quantity
            }));
            this.showCostBoqModal = true;
        },

        updateCostRow(index, value) {
            const row = this.costRows[index];
            if (!row) return;
            row.unitPrice = Math.max(0, Number(value) || 0);
            row.total = row.unitPrice * (Number(row.quantity) || 0);
            this.saveWorkflowState();
        },

        get costGrandTotal() {
            return (this.costRows || []).reduce((sum, r) => sum + (Number(r.total) || 0), 0);
        },

        submitCostToLead() {
            if (this.loginForm.role !== 'Estimator Proposal') return;
            if (!this.costRows.length) return alert('Belum ada data BOQ yang dihitung.');
            this.workflowStatus = 'SUBMITTED_TO_LEAD';
            this.createWorkflowHistory('SUBMITTED_TO_LEAD', 'BOQ & Quantity Take-Off dikirim untuk Review Lead', this.costNotes || 'Estimator mengirim hasil kalkulasi BOQ untuk direview Lead.');
            this.showCostBoqModal = false;
            this.taskView = 'active';
            alert('Kalkulasi BOQ berhasil dikirim ke Lead Estimator untuk review.');
        },

        get leadReviewRows() {
            // Gunakan hasil costing Estimator sebagai sumber utama.
            // Quantity/spec tetap berasal dari snapshot QTO Engineer.
            return Array.isArray(this.costRows) ? this.costRows : [];
        },

        get leadReviewTotalQuantity() {
            return this.leadReviewRows.reduce((sum, row) => sum + (Number(row.quantity) || 0), 0);
        },

        get leadReviewTotalCost() {
            return this.leadReviewRows.reduce((sum, row) => sum + (Number(row.total) || 0), 0);
        },

        get leadReviewCurrencyLabel() {
            return this.boqCurrency === 'IDR' ? 'Rp' : (this.boqCurrency || 'USD');
        },

        formatMoney(value) {
            const n = Number(value) || 0;
            return n.toLocaleString('id-ID');
        },

        approveData() {
            if (this.loginForm.role !== 'Lead Estimator') {
                alert('Approval & Review hanya dapat dijalankan oleh Lead Estimator.');
                return;
            }
            if (this.workflowStatus !== 'SUBMITTED_TO_LEAD') {
                alert('Belum ada BOQ/laporan dari Estimator yang menunggu review.');
                this.currentDashboardTab = 'tasks';
                return;
            }
            if (!Array.isArray(this.costRows) || this.costRows.length === 0) {
                alert('Data BOQ dari Estimator belum tersedia.');
                return;
            }
            this.loadWorkflowState();
            this.approvalNote = '';
            this.showApproveModal = true;
        },

        confirmApprove() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            this.workflowStatus = 'APPROVED';
            const note = this.approvalNote || document.getElementById('reviewer-notes-textarea')?.value || 'Disetujui oleh Lead Estimator.';
            this.createWorkflowHistory(
                'APPROVED',
                'BOQ & Laporan Disetujui',
                `${note} Total BOQ: ${this.formatMoney(this.leadReviewTotalCost)}. Quantity: ${this.leadReviewRows.map(r => `${r.quantity ?? 0} ${r.unit || 'EA'}`).join(', ')}.`
            );
            this.saveWorkflowState();
            this.showApproveModal = false;
            this.taskView = 'history';
            alert('Laporan berhasil di-approve. Status project sekarang APPROVED.');
        },

        rejectData() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            this.workflowStatus = 'REJECTED';
            const note = this.approvalNote || document.getElementById('reviewer-notes-textarea')?.value || 'Laporan ditolak oleh Lead Estimator.';
            this.createWorkflowHistory('REJECTED', 'BOQ & Laporan Ditolak', note);
            this.saveWorkflowState();
            this.showApproveModal = false;
            this.taskView = 'history';
            alert('Laporan ditolak dan tercatat pada riwayat.');
        },

        handleMintaRevisi() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            this.workflowStatus = 'REVISION_REQUESTED_TO_ESTIMATOR';
            const note = this.approvalNote || document.getElementById('reviewer-notes-textarea')?.value || 'Mohon revisi data sebelum diajukan kembali.';
            this.createWorkflowHistory('REVISION_REQUESTED_TO_ESTIMATOR', 'Lead Meminta Revisi BOQ', note);
            this.saveWorkflowState();
            this.showApproveModal = false;
            this.taskView = 'history';
            alert('Permintaan revisi BOQ berhasil dikirim kembali ke Estimator.');
        },

        get finalReportAvailable() {
            return this.workflowStatus === 'APPROVED' && Array.isArray(this.costRows) && this.costRows.length > 0;
        },

        openFinalReport() {
            if (!this.finalReportAvailable) {
                alert('Laporan final belum tersedia. BOQ harus disetujui Lead terlebih dahulu.');
                return;
            }
            this.loadWorkflowState();
            this.showFinalReportModal = true;
        },

        get finalReportTotal() {
            return (this.costRows || []).reduce((sum, row) => sum + (Number(row.total) || 0), 0);
        },

        get finalReportApprovedBy() {
            const approved = [...(this.projectApprovalHistory || [])].reverse().find(item => item.status === 'APPROVED');
            return approved?.reviewer || 'Lead Estimator';
        },

        get finalReportApprovedAt() {
            const approved = [...(this.projectApprovalHistory || [])].reverse().find(item => item.status === 'APPROVED');
            return approved?.timestamp || '-';
        },

        exportFinalReport() {
            if (!this.finalReportAvailable) {
                alert('Laporan final belum tersedia.');
                return;
            }

            const escapeHtml = (v) => String(v ?? '').replace(/[&<>"]/g, ch => ({
                '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;'
            }[ch]));

            const qtoMap = new Map((this.qtoRows || []).map(r => [
                [r.description, r.material, r.size, r.spec].join('|'),
                r
            ]));

            const rowsHtml = (this.costRows || []).map((row, i) => `
                <tr>
                    <td>${i + 1}</td>
                    <td>${escapeHtml(row.description)}</td>
                    <td>${escapeHtml(row.material)}</td>
                    <td>${escapeHtml(row.size)}</td>
                    <td>${escapeHtml(row.spec)}</td>
                    <td class="center">${escapeHtml(row.quantity)} ${escapeHtml(row.unit || 'EA')}</td>
                    <td class="right">${Number(row.unitPrice || 0).toLocaleString('id-ID')}</td>
                    <td class="right">${Number(row.total || 0).toLocaleString('id-ID')}</td>
                </tr>`).join('');

            const approved = [...(this.projectApprovalHistory || [])].reverse().find(item => item.status === 'APPROVED');
            const reportHtml = `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(this.activeProject)} - Final BOQ Report</title>
<style>
body{font-family:Arial,sans-serif;color:#172033;font-size:11px;padding:28px}
h1{font-size:20px;margin:0 0 5px} h2{font-size:13px;margin:22px 0 8px}
.meta{color:#64748b;margin-bottom:18px}.status{display:inline-block;padding:5px 10px;border:1px solid #86efac;background:#f0fdf4;color:#15803d;border-radius:6px;font-weight:700}
table{width:100%;border-collapse:collapse;margin-top:10px} th,td{border:1px solid #cbd5e1;padding:7px} th{background:#e2e8f0;text-align:left}.center{text-align:center}.right{text-align:right}
.total{margin-top:12px;text-align:right;font-size:14px;font-weight:700}.note{margin-top:18px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0}
@page{size:landscape;margin:12mm}
</style></head><body>
<h1>${escapeHtml(this.activeProject)} — Final BOQ Report</h1>
<div class="meta">Sheet: ${escapeHtml(this.qtoSourceSheet || this.activeSheet)} &nbsp;|&nbsp; Generated: ${escapeHtml(new Date().toLocaleString('id-ID'))}</div>
<div class="status">APPROVED</div>
<h2>Quantity Take-Off & Cost Summary</h2>
<table><thead><tr><th>No</th><th>Description</th><th>Material</th><th>Size</th><th>Spec</th><th>Quantity</th><th>Unit Price</th><th>Total</th></tr></thead>
<tbody>${rowsHtml}</tbody></table>
<div class="total">TOTAL BOQ: ${Number(this.finalReportTotal || 0).toLocaleString('id-ID')}</div>
<div class="note"><b>Approved By:</b> ${escapeHtml(approved?.reviewer || 'Lead Estimator')}<br><b>Approved At:</b> ${escapeHtml(approved?.timestamp || '-')}<br><b>Catatan:</b> ${escapeHtml(approved?.notes || 'Laporan disetujui oleh Lead Estimator.')}</div>
</body></html>`;

            const win = window.open('', '_blank', 'width=1200,height=800');
            if (!win) return alert('Popup diblokir browser. Izinkan popup untuk export laporan final.');
            win.document.write(reportHtml);
            win.document.close();
            win.focus();
            setTimeout(() => win.print(), 300);
        },

        openActiveTasks() {
            this.taskView = 'active';
        },

        openApprovalHistory() {
            this.taskView = 'history';
        },

        openTaskMTOValve() {
            this.activeSheet = 'Valve';
            this.currentDashboardTab = 'workspace';
            if (this.loginForm.role === 'Piping Engineer' && this.activeTaskCount > 0) {
                this.alertMessage = '';
            }
        },

        openApprovalDetail(item) {
            this.approvalNote = item?.notes || '';
            // Untuk approval yang sudah final, tampilkan laporan BOQ lengkap, bukan alert browser.
            if (item?.status === 'APPROVED' && this.finalReportAvailable) {
                this.loadWorkflowState();
                this.showFinalReportModal = true;
                return;
            }
            // Untuk status selain APPROVED, tetap gunakan dialog sederhana agar tidak mengubah workflow.
            alert(`${item?.title || 'Detail Approval'}\n\nStatus: ${this.getApprovalStatusText(item?.status)}\nProject: ${item?.projectId || '-'}\nSheet: ${item?.sheet || '-'}\nReviewer: ${item?.reviewer || '-'}\nCatatan: ${item?.notes || '-'}`);
        },

        submitMTOToEstimator() {
            if (this.loginForm.role !== 'Piping Engineer') return;
            if (!this.currentRows.length) return alert('MTO masih kosong.');
            const groups = this.getBomGroups();
            if (!groups.length) return alert('Quantity Take-Off belum dapat dibuat dari data MTO ini.');
            this.qtoRows = groups.map(g => ({ ...g }));
            this.qtoSourceSheet = this.activeSheet;
            this.qtoSubmittedAt = new Date().toISOString();
            this.workflowStatus = 'SUBMITTED_TO_ESTIMATOR';
            this.saveWorkflowState();
            this.createWorkflowHistory('SUBMITTED_TO_ESTIMATOR', 'Quantity Take-Off dikirim ke Estimator', `Engineer mengirim ${this.qtoRows.length} kelompok material dari sheet ${this.activeSheet}. Quantity menjadi sumber BOQ Estimator.`);
            alert('MTO berhasil dikirim ke Estimator.');
        },

        openCalculateBom() {
            this.bomMaterialType = this.getBomMaterialOptions()[0] || 'Pipe';
            this.bomSizeFilter = '';
            this.bomSpecFilter = '';
            this.bomOD = 0;
            this.bomThickness = 0;
            this.bomLength = 6;
            this.bomQuantityInput = 0;
            this.showBoqModal = true;
            this.$nextTick(() => this.refreshBomQuantity());
        },

        refreshBomQuantity() {
            const qty = this.getBomFilteredGroups().reduce((sum, item) => sum + item.quantity, 0);
            this.bomQuantityInput = qty;
            return qty;
        },

        getBomQuantity(row) {
            const keys = Object.keys(row || {});
            const key = keys.find(k => /^(qty|quantity|item count|item_count|count|jumlah)$/i.test(String(k).trim()))
                || keys.find(k => /quantity|qty|item.?count|jumlah/i.test(String(k)));
            const value = key ? Number(String(row[key]).replace(/,/g, '')) : NaN;
            return Number.isFinite(value) && value > 0 ? value : 1;
        },

        getBomField(row, patterns) {
            const key = Object.keys(row || {}).find(k => patterns.some(p => p.test(String(k))));
            return key ? String(row[key] ?? '').trim() : '';
        },

        getBomMaterialOptions() {
            const sheet = String(this.activeSheet || '').trim();
            const map = {
                'Valve': 'Valve', 'SP Items': 'SP Items', 'Support': 'Pipe Support',
                'LineList': 'Pipe', 'Pipe': 'Pipe', 'Single Branch Fitting': 'Fitting',
                'Tee': 'Fitting', 'Flange': 'Fitting', 'Elbow': 'Fitting', 'Nozzle': 'Fitting',
                'Vessel': 'Vessel', 'Tank': 'Tank', 'Pump': 'Pump', 'Misc Equipment': 'Misc Equipment',
                'Equipment': 'Equipment', 'Piping and Equipment': 'Equipment'
            };
            return [map[sheet] || sheet || 'Pipe'];
        },

        getBomSizeOptions() {
            const values = new Set();
            (Array.isArray(this.currentRows) ? this.currentRows : []).forEach(row => {
                const value = this.getBomField(row, [/^size$/i, /line size/i, /available size/i]);
                if (value) values.add(value);
            });
            return Array.from(values);
        },

        getBomSpecOptions() {
            const values = new Set();
            (Array.isArray(this.currentRows) ? this.currentRows : []).forEach(row => {
                const value = this.getBomField(row, [/^spec$/i, /pipe\.?spec/i, /pipe spec/i]);
                if (value) values.add(value);
            });
            return Array.from(values);
        },

        getBomGroups() {
            const rows = Array.isArray(this.currentRows) ? this.currentRows : [];
            const groups = new Map();
            rows.forEach((row) => {
                const material = this.getBomField(row, [/^material$/i, /material/i]) || '-';
                const size = this.getBomField(row, [/^size$/i, /line size/i, /available size/i]) || '-';
                const spec = this.getBomField(row, [/^spec$/i, /pipe\.?spec/i, /pipe spec/i]) || '-';
                const desc = this.getBomField(row, [/long description/i, /short description/i, /description/i]) || this.activeSheet;
                const key = [material, size, spec, desc].join('|');
                if (!groups.has(key)) groups.set(key, { material, size, spec, description: desc, quantity: 0 });
                groups.get(key).quantity += this.getBomQuantity(row);
            });
            return Array.from(groups.values());
        },

        getBomFilteredGroups() {
            const size = String(this.bomSizeFilter || '').toLowerCase();
            const spec = String(this.bomSpecFilter || '').toLowerCase();
            const groups = this.getBomGroups().filter(item =>
                (!size || item.size.toLowerCase() === size || item.size.toLowerCase().includes(size)) &&
                (!spec || item.spec.toLowerCase() === spec || item.spec.toLowerCase().includes(spec))
            );
            return groups;
        },

        syncBomQuantity() {
            const qty = this.getBomFilteredGroups().reduce((sum, item) => sum + item.quantity, 0);
            if (!Number.isFinite(Number(this.bomQuantityInput)) || Number(this.bomQuantityInput) <= 0) {
                this.bomQuantityInput = qty;
            }
            return qty;
        },

        get isPipeBom() {
            return String(this.bomMaterialType || '').trim().toLowerCase() === 'pipe';
        },

        calculatePipeWeight() {
            const od = Number(this.bomOD);
            const t = Number(this.bomThickness);
            const length = Number(this.bomLength);
            const qty = Number(this.bomQuantityInput) || this.getBomFilteredGroups().reduce((sum, item) => sum + item.quantity, 0);
            if (!(od > 0 && t > 0 && length > 0 && qty > 0)) return { kgm: 0, perPipe: 0, total: 0, qty };
            const kgm = (od - t) * t * 0.2466;
            const perPipe = kgm * length;
            return { kgm, perPipe, total: perPipe * qty, qty };
        },

        get qtoTotalQuantity() {
            return (this.qtoRows || []).reduce((sum, r) => sum + (Number(r.quantity) || 0), 0);
        },

        exportBOQ() {
            const groups = this.getBomFilteredGroups();
            if (!groups.length) return alert('Tidak ada data untuk Quantity Take-Off pada sheet ini.');

            const rowsHtml = groups.map((item, i) => `
                <tr>
                    <td>${i + 1}</td>
                    <td>${item.description}</td>
                    <td>${item.material}</td>
                    <td>${item.size}</td>
                    <td>${item.spec}</td>
                    <td>${item.quantity}</td>
                    <td>EA</td>
                </tr>`).join('');

            const htmlReport = `<!doctype html><html><head><meta charset="utf-8"><title>Quantity Take-Off ${this.activeProject}</title>
                <style>
                    body{font-family:Arial,sans-serif;font-size:12px;padding:24px;color:#172033}
                    h1{font-size:18px;margin:0 0 4px} p{color:#64748b}
                    table{width:100%;border-collapse:collapse;margin-top:16px}
                    th{background:#eaf0f6;font-weight:700} th,td{border:1px solid #cbd5e1;padding:7px;text-align:left}
                    td:nth-child(1),td:nth-child(6){text-align:center}
                </style></head><body>
                <h1>${this.activeProject} — Quantity Take-Off</h1>
                <p>Sheet: ${this.activeSheet} • ${new Date().toLocaleString('id-ID')}</p>
                <table><thead><tr><th>No</th><th>Description</th><th>Material</th><th>Size</th><th>Spec</th><th>Quantity</th><th>Unit</th></tr></thead>
                <tbody>${rowsHtml}</tbody></table></body></html>`;

            const win = window.open('', '_blank');
            if (!win) return alert('Popup diblokir browser. Izinkan popup untuk export laporan.');
            win.document.write(htmlReport);
            win.document.close();
            setTimeout(() => win.print(), 300);
        },

        printPDF() {
            const rows = this.filteredRows || [];
            if (!rows.length) return alert(`Tidak ada data pada sheet ${this.activeSheet}.`);
            const columns = this.currentColumns || [];
            const escapeHtml = (v) => String(v ?? '').replace(/[&<>\"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[ch]));
            const headers = ['No', 'Long Description (Family)', ...columns.filter(c => c !== 'Long Description (Family)')];
            const body = rows.map((row, i) => `<tr>${headers.map((h, j) => `<td>${escapeHtml(j === 0 ? String(i + 1).padStart(2,'0') : row[h])}</td>`).join('')}</tr>`).join('');
            const win = window.open('', '_blank', 'width=1200,height=800');
            if (!win) return alert('Popup diblokir browser. Izinkan popup untuk mencetak PDF.');
            win.document.write(`<!doctype html><html><head><meta charset=\"utf-8\"><title>${escapeHtml(this.activeProject)} - ${escapeHtml(this.activeSheet)}</title><style>body{font-family:Arial,sans-serif;padding:24px;color:#111}h1{font-size:18px;margin:0 0 4px}p{font-size:11px;color:#555;margin:0 0 16px}table{border-collapse:collapse;width:100%;font-size:8px}th,td{border:1px solid #999;padding:4px;vertical-align:top}th{background:#e2e8f0;font-weight:700} @page{size:landscape;margin:10mm}</style></head><body><h1>${escapeHtml(this.activeProject)} — ${escapeHtml(this.activeSheet)}</h1><p>Export PDF • ${new Date().toLocaleString('id-ID')} • Total ${rows.length} data</p><table><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></body></html>`);
            win.document.close();
            win.focus();
            setTimeout(() => win.print(), 300);
        }
    }
}