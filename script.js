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
        showBomModal: false,
        bomFilterCategory: 'ALL',
        bomFilterSearch: '',
        // Item BOM yang dipilih Engineer untuk dikirim ke Estimator.
        // Data yang tidak dipilih tetap berada di Detail BOM.
        bomSelectedKeys: [],
        showBoqModal: false,
        showApproveModal: false,
        taskView: 'active',
        approvalHistory: [],
        approvalNote: '',
        showModuleModal: false,

        // Manajemen Team & Roles khusus System Administrator
        showTeamMemberModal: false,
        teamMemberEditingId: null,
        teamMembers: [
            { id: 1, name: 'System Administrator', email: 'admin@tripatra.com', role: 'System Administrator', projects: ['PRJ-3000'], active: true },
            { id: 2, name: 'Piping Engineer', email: 'engineer@tripatra.com', role: 'Piping Engineer', projects: ['PRJ-3000'], active: true },
            { id: 3, name: 'Estimator Proposal', email: 'estimator@tripatra.com', role: 'Estimator Proposal', projects: ['PRJ-3000'], active: true },
            { id: 4, name: 'Lead Estimator', email: 'lead@tripatra.com', role: 'Lead Estimator', projects: ['PRJ-3000'], active: true }
        ],
        teamMemberForm: { name: '', email: '', role: 'Piping Engineer', projects: [], active: true },

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
            files: [],
            fileNames: []
        },

        // Import Excel professional workflow
        showImportPreview: false,
        isImportingExcel: false,
        importPreview: {
            files: [],
            sheets: [],
            totalRows: 0,
            newRows: 0,
            duplicateRows: 0,
            skippedRows: 0,
            errors: [],
            payload: []
        },

        importSuccessVisible: false,
        importSuccessMessage: '',
        activeModuleName: '',
        moduleDescription: '',

        activePriceLevel: 'Menengah',
        boqCurrency: 'IDR',
        boqPriceSearch: '',
        boqPriceFilter: 'all',
        boqPricePage: 1,
        boqPricePageSize: 10,

        get filteredBOQPriceGroups() {
            const groups = this.allProjectsData?.[this.activeProject]?.meta?.boq?.priceGroups || [];
            const q = this.normalizePriceMasterText(this.boqPriceSearch);
            return groups.filter(g => {
                const matchesSearch = !q || [g.sheet, g.component, g.description, g.unit, g.size1, g.size2]
                    .map(v => this.normalizePriceMasterText(v)).some(v => v.includes(q));
                const priced = Number(g.unitPrice || 0) > 0;
                const matchesFilter = this.boqPriceFilter === 'priced' ? priced : this.boqPriceFilter === 'unpriced' ? !priced : true;
                return matchesSearch && matchesFilter;
            });
        },

        get totalBOQPricePages() {
            return Math.max(1, Math.ceil(this.filteredBOQPriceGroups.length / this.boqPricePageSize));
        },

        get pagedBOQPriceGroups() {
            if (this.boqPricePage > this.totalBOQPricePages) this.boqPricePage = this.totalBOQPricePages;
            const start = (this.boqPricePage - 1) * this.boqPricePageSize;
            return this.filteredBOQPriceGroups.slice(start, start + this.boqPricePageSize);
        },

        getBOQPriceGroupIndexByKey(key) {
            const groups = this.allProjectsData?.[this.activeProject]?.meta?.boq?.priceGroups || [];
            return groups.findIndex(g => g.key === key);
        },

        updateBOQPriceGroupByKey(key, value) {
            const index = this.getBOQPriceGroupIndexByKey(key);
            if (index >= 0) this.updateBOQPriceGroup(index, value);
        },

        resetBOQPricePage() { this.boqPricePage = 1; },

        nextBOQPricePage() {
            this.boqPricePage = Math.min(this.totalBOQPricePages, this.boqPricePage + 1);
        },

        prevBOQPricePage() {
            this.boqPricePage = Math.max(1, this.boqPricePage - 1);
        },

        // Penyimpanan hasil BOM/BOQ antar-role.
        // localStorage dipakai sebagai cache cepat; IndexedDB menjadi
        // fallback agar hasil 641 item tidak hilang saat quota localStorage penuh.
        bomBoqStorageReady: false,
        bomBoqStoragePromise: null,

        
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

        // ==========================================================
        // SESSION + NAVIGATION PERSISTENCE
        // Menyimpan halaman terakhir dan sesi login secara aman di browser.
        // Password TIDAK pernah disimpan.
        // ==========================================================
        pageStateStorageKey: 'bmbq_page_state_v2',
        authStorageKey: 'bmbq_auth_state_v2',

        savePageState() {
            try {
                localStorage.setItem(this.pageStateStorageKey, JSON.stringify({
                    currentDashboardTab: this.currentDashboardTab,
                    activeProject: this.activeProject,
                    activeSheet: this.activeSheet,
                    currentPage: this.currentPage,
                    globalSearch: this.globalSearch || ''
                }));
            } catch (e) {
                console.warn('State halaman tidak dapat disimpan:', e);
            }
        },

        loadPageState() {
            try {
                const raw = localStorage.getItem(this.pageStateStorageKey);
                if (!raw) return;
                const state = JSON.parse(raw);
                if (!state || typeof state !== 'object') return;

                const validTabs = ['home', 'workspace', 'tasks', 'team', 'profile'];
                if (validTabs.includes(state.currentDashboardTab)) {
                    this.currentDashboardTab = state.currentDashboardTab;
                }
                if (state.activeProject && this.allProjectsData?.[state.activeProject]) {
                    this.activeProject = state.activeProject;
                }
                if (state.activeSheet && this.sheets.includes(state.activeSheet)) {
                    this.activeSheet = state.activeSheet;
                }
                const page = Number(state.currentPage);
                if (Number.isFinite(page) && page >= 1) {
                    this.currentPage = Math.floor(page);
                }
                if (typeof state.globalSearch === 'string') {
                    this.globalSearch = state.globalSearch;
                }
            } catch (e) {
                console.warn('State halaman tidak dapat dipulihkan:', e);
            }
        },

        saveAuthState() {
            try {
                if (this.isLoggedIn && this.loginForm.user && this.loginForm.role) {
                    localStorage.setItem(this.authStorageKey, JSON.stringify({
                        isLoggedIn: true,
                        user: this.loginForm.user,
                        role: this.loginForm.role
                    }));
                }
            } catch (e) {
                console.warn('Session login tidak dapat disimpan:', e);
            }
        },

        restoreAuthState() {
            try {
                const raw = localStorage.getItem(this.authStorageKey);
                if (!raw) return false;
                const state = JSON.parse(raw);
                if (!state?.isLoggedIn || !state.user || !state.role) return false;

                this.loginForm.user = String(state.user);
                this.loginForm.role = String(state.role);
                // Password sengaja tetap kosong dan tidak disimpan.
                this.loginForm.pass = '';
                this.isLoggedIn = true;
                return true;
            } catch (e) {
                console.warn('Session login tidak dapat dipulihkan:', e);
                return false;
            }
        },

        clearAuthState() {
            try { localStorage.removeItem(this.authStorageKey); } catch (e) {}
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

            // DATA DEMO PROJECT 3000
            // Dijalankan SETELAH membaca localStorage supaya data yang sudah
            // diimport pengguna tetap diprioritaskan. Hanya sheet kosong yang
            // akan diisi dari bundle data awal. Import Excel tetap tersedia.
            this.seedProject3000Data();

            this.normalizeAllProjectsData();
            this.loadBomBoqStorage();

            // BOQ proyek menggunakan IDR sebagai mata uang standar.
            // Data lama yang masih menyimpan currency USD dinormalisasi ke IDR
            // agar tampilan kalkulasi dan hasil laporan konsisten dengan format
            // BOQ/MTO yang digunakan sistem. Nilai numeriknya tidak diubah.
            Object.values(this.allProjectsData || {}).forEach(project => {
                if (project?.meta?.boq) project.meta.boq.currency = 'IDR';
            });
            this.boqCurrency = 'IDR';

            this.refreshSheetList();
            this.loadApprovalHistory();

            // Pulihkan posisi halaman SETELAH seluruh data proyek/sheet siap.
            this.loadPageState();
            const restoredSession = this.restoreAuthState();

            // Simpan perubahan navigasi secara otomatis. Dengan watcher ini,
            // klik menu, ganti project/sheet, pagination, dan pencarian akan
            // tetap berada di posisi terakhir ketika browser di-refresh.
            ['currentDashboardTab', 'activeProject', 'activeSheet', 'currentPage', 'globalSearch'].forEach(prop => {
                this.$watch(prop, () => this.savePageState());
            });
            this.$watch('isLoggedIn', value => {
                if (value) this.saveAuthState();
                else this.clearAuthState();
            });
            this.$watch('loginForm.user', () => {
                if (this.isLoggedIn) this.saveAuthState();
            });
            this.$watch('loginForm.role', () => {
                if (this.isLoggedIn) this.saveAuthState();
            });

            // Jika session berhasil dipulihkan, jangan mengubah tab terakhir.
            // User langsung masuk ke halaman terakhir tanpa melihat login.
            if (restoredSession) {
                this.loginError = false;
            }

            this.savePageState();
            this.saveStorage();
        },

        seedProject3000Data() {
            const seed = window.BMBQ_SEED_DATA;
            if (!seed || typeof seed !== 'object') return;

            const project = this.allProjectsData?.['PRJ-3000'];
            if (!project) return;

            // Hanya mengisi data awal pada project yang benar-benar belum
            // mempunyai data tersimpan. Data hasil Import Excel pengguna
            // tidak ditimpa.
            Object.entries(seed).forEach(([sheetName, rows]) => {
                if (!this.sheets.includes(sheetName)) this.sheets.push(sheetName);

                const existing = Array.isArray(project[sheetName]) ? project[sheetName] : [];
                if (existing.length === 0 && Array.isArray(rows) && rows.length > 0) {
                    project[sheetName] = JSON.parse(JSON.stringify(rows));
                }
            });

            project.meta.projectCode = 'PRJ-3000';
            project.meta.projectName = 'Piping & Equipment';
            project.meta.description = 'Project 3000 - data MTO/Fitting/Valve, SP Items, Line List, dan Pipe Support';
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
                    'Volume Flow\n[m3/h]',
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
            if (!matrix.length) throw new Error('Worksheet kosong.');

            const clean = value => String(value ?? '')
                .replace(/\u00A0/g, ' ')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();

            // ==========================================================
            // LINE LIST: Excel kantor biasanya memakai 2 baris header
            // (header utama + sub-header). Jangan jadikan salah satu
            // baris data sebagai header. Deteksi berdasarkan signature
            // kolom Line Size + Process Fluid + Pipe.Spec + Seq. No.
            // ==========================================================
            const lineListSchema = this.getExactSchema('LineList');
            const isLineListHeader = row => {
                const values = (row || []).map(clean);
                const has = text => values.some(v => v === clean(text) || v.includes(clean(text)));
                return has('Line Size (Inch)') && has('Process Fluid Identifier') &&
                       (has('Pipe.Spec') || has('Seq. No'));
            };

            let lineListHeaderIndex = matrix.findIndex(isLineListHeader);
            if (lineListHeaderIndex >= 0) {
                // Header kedua hanya dipakai untuk menentukan posisi data.
                // Nama kolom tetap memakai schema Line List agar hasil import
                // konsisten dengan tabel aplikasi, bukan nama angka/acak dari
                // baris Excel yang salah terbaca.
                const nextRow = matrix[lineListHeaderIndex + 1] || [];
                const nextText = nextRow.map(clean);
                const hasSubHeader = nextText.some(v =>
                    ['type', 'thickness [mm]', 'operating', 'design', 'medium', 'pressure [barg]']
                        .some(token => v === token)
                );
                const dataStart = lineListHeaderIndex + (hasSubHeader ? 2 : 1);
                const headers = [...lineListSchema];
                const rows = [];

                for (let r = dataStart; r < matrix.length; r++) {
                    const sourceRow = Array.isArray(matrix[r]) ? matrix[r] : [];
                    const hasData = sourceRow.some(v => String(v ?? '').trim() !== '');
                    if (!hasData) continue;

                    // Lewati baris yang ternyata masih merupakan sub-header.
                    const rowText = sourceRow.map(clean).join(' | ');
                    if (rowText.includes('line size (inch)') && rowText.includes('process fluid identifier')) continue;

                    const row = {};
                    headers.forEach((header, c) => {
                        row[header] = sourceRow[c] ?? '';
                    });

                    // Nomor baris tetap mengikuti Excel jika tersedia.
                    // Bila kosong, sistem akan menampilkannya sebagai nomor urut.
                    if (String(row['No'] ?? '').trim() === '') row['No'] = '';
                    rows.push(row);
                }

                if (!rows.length) throw new Error('Header Line List ditemukan, tetapi tidak ada data setelah header.');
                return {
                    headers,
                    rows,
                    headerIndex: lineListHeaderIndex,
                    dataStart,
                    dataType: 'LineList'
                };
            }

            // ==========================================================
            // MASTER/MTO LAIN: tetap fleksibel. Cari header yang paling
            // masuk akal dan pertahankan nama header Excel secara 1:1.
            // ==========================================================
            let bestIndex = -1, bestScore = -Infinity;
            matrix.forEach((row, i) => {
                if (!Array.isArray(row)) return;
                const nonEmpty = row.filter(v => String(v ?? '').trim() !== '').length;
                if (nonEmpty < 2) return;
                const textCount = row.filter(v => typeof v === 'string' && v.trim()).length;
                const next = matrix[i + 1] || [];
                const nextNonEmpty = next.filter(v => String(v ?? '').trim() !== '').length;
                let score = nonEmpty * 5 + textCount * 2;
                if (nextNonEmpty >= Math.max(2, Math.floor(nonEmpty * 0.35))) score += 12;
                if (i === 0) score -= 2;
                if (nonEmpty > 4) score += 5;
                if (score > bestScore) { bestScore = score; bestIndex = i; }
            });

            if (bestIndex < 0) throw new Error('Header Excel tidak ditemukan.');

            const rawHeaders = matrix[bestIndex] || [];
            const headers = [];
            const used = new Set();
            rawHeaders.forEach((value, colIndex) => {
                let header = this.normalizeHeaderExact(value);
                if (!header) header = `Column ${colIndex + 1}`;
                let unique = header;
                let n = 2;
                while (used.has(unique.toLowerCase())) unique = `${header} (${n++})`;
                used.add(unique.toLowerCase());
                headers.push(unique);
            });

            const rows = [];
            for (let r = bestIndex + 1; r < matrix.length; r++) {
                const sourceRow = Array.isArray(matrix[r]) ? matrix[r] : [];
                const hasData = headers.some((_, c) => String(sourceRow[c] ?? '').trim() !== '');
                if (!hasData) continue;
                const row = {};
                headers.forEach((header, c) => { row[header] = sourceRow[c] ?? ''; });
                rows.push(row);
            }
            return { headers, rows, headerIndex: bestIndex, dataStart: bestIndex + 1, dataType: 'Generic' };
        },

        canonicalImportValue(value) {
            if (value === null || value === undefined) return '';
            if (value instanceof Date) return value.toISOString();
            return String(value).replace(/\s+/g, ' ').trim().toLowerCase();
        },

        canonicalImportRow(row) {
            return Object.keys(row || {})
                .sort((a,b) => a.localeCompare(b))
                .map(key => `${this.canonicalImportValue(key)}=${this.canonicalImportValue(row[key])}`)
                .join('\u001F');
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
                        progress: '0%',
                        workflowStatus: 'DRAFT',
                        revisionNotes: '',
                        bom: null,
                        boq: null
                    };
                } else {
                    project.meta.workflowStatus ||= 'DRAFT';
                    project.meta.revisionNotes ||= '';
                    if (!('bom' in project.meta)) project.meta.bom = null;
                    if (!('boq' in project.meta)) project.meta.boq = null;
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

        getBomStorageKey(projectKey) {
            return `tripatra_bom_v4::${encodeURIComponent(String(projectKey || 'default'))}`;
        },

        getBoqStorageKey(projectKey) {
            return `tripatra_boq_v4::${encodeURIComponent(String(projectKey || 'default'))}`;
        },

        // ==========================================================
        // WORKFLOW STORAGE: ENGINEER -> ESTIMATOR -> LEAD
        // ==========================================================
        getWorkflowDB() {
            return new Promise((resolve, reject) => {
                if (!('indexedDB' in window)) {
                    reject(new Error('IndexedDB tidak tersedia pada browser ini.'));
                    return;
                }

                const request = indexedDB.open('tripatra_bmbq_workflow_v1', 1);

                request.onupgradeneeded = (event) => {
                    const db = event.target.result;
                    if (!db.objectStoreNames.contains('snapshots')) {
                        db.createObjectStore('snapshots', { keyPath: 'projectKey' });
                    }
                };

                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error || new Error('Gagal membuka IndexedDB.'));
            });
        },

        async saveWorkflowSnapshot(projectKey, project) {
            if (!projectKey || !project?.meta) return;

            const snapshot = {
                projectKey: String(projectKey),
                bom: project.meta.bom || null,
                boq: project.meta.boq || null,
                workflowStatus: project.meta.workflowStatus || 'DRAFT',
                revision: project.meta.version || 0,
                updatedAt: Date.now()
            };

            try {
                const db = await this.getWorkflowDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('snapshots', 'readwrite');
                    tx.objectStore('snapshots').put(snapshot);
                    tx.oncomplete = resolve;
                    tx.onerror = () => reject(tx.error || new Error('Gagal menyimpan snapshot.'));
                    tx.onabort = () => reject(tx.error || new Error('Transaksi snapshot dibatalkan.'));
                });
                db.close();
            } catch (error) {
                console.warn('IndexedDB workflow storage tidak tersedia:', error);
            }
        },


        async saveLeadBOQSnapshot(projectKey, project) {
            if (!projectKey || !project?.meta?.boq?.items?.length) return;

            const boq = project.meta.boq;
            const submittedAt = Date.now();
            const snapshot = {
                projectKey: String(projectKey),
                revision: Number(project.meta.version || 0),
                workflowStatus: 'SUBMITTED_TO_LEAD',
                submittedAt,
                // Sumber kebenaran Lead: seluruh BOQ hasil Estimator, termasuk
                // priceGroups, item prices, direct/indirect/total cost dan currency.
                // Jangan biarkan data BOQ seed/stale di browser Lead menggantikannya.
                boq: JSON.parse(JSON.stringify({
                    ...boq,
                    handoffTotalCost: Number(boq.totalCost) || 0,
                    handoffDirectCost: Number(boq.directCost) || 0,
                    handoffIndirectCost: Number(boq.indirectCost) || 0,
                    handoffItemCount: Array.isArray(boq.items) ? boq.items.length : 0,
                    handoffGroupCount: Array.isArray(boq.priceGroups) ? boq.priceGroups.length : 0,
                    handoffAt: submittedAt,
                    updatedAtEpoch: submittedAt
                }))
            };

            // Simpan salinan final handoff di localStorage untuk browser
            // yang tidak mengizinkan IndexedDB/file-origin tertentu.
            const localKey = `tripatra_boq_lead_snapshot_v1_${String(projectKey)}`;
            try {
                localStorage.setItem(localKey, JSON.stringify(snapshot));
            } catch (error) {
                console.warn('Snapshot BOQ Lead localStorage penuh:', error);
            }

            try {
                const db = await this.getWorkflowDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('snapshots', 'readwrite');
                    const store = tx.objectStore('snapshots');
                    store.put({
                        projectKey: String(projectKey),
                        bom: project.meta.bom || null,
                        boq: snapshot.boq,
                        workflowStatus: 'SUBMITTED_TO_LEAD',
                        revision: snapshot.revision,
                        updatedAt: snapshot.submittedAt,
                        handoff: 'ESTIMATOR_TO_LEAD'
                    });
                    tx.oncomplete = resolve;
                    tx.onerror = () => reject(tx.error || new Error('Gagal menyimpan snapshot Lead.'));
                    tx.onabort = () => reject(tx.error || new Error('Transaksi snapshot Lead dibatalkan.'));
                });
                db.close();
            } catch (error) {
                console.warn('IndexedDB snapshot Lead tidak tersedia:', error);
            }
        },

        async loadLeadBOQSnapshot(projectKey) {
            const localKey = `tripatra_boq_lead_snapshot_v1_${String(projectKey)}`;

            // Ambil snapshot khusus Lead lebih dulu.
            try {
                const raw = localStorage.getItem(localKey);
                if (raw) {
                    const snapshot = JSON.parse(raw);
                    if (snapshot?.workflowStatus === 'SUBMITTED_TO_LEAD' &&
                        (snapshot?.boq?.items?.length || snapshot?.boq?.priceGroups?.length)) return snapshot;
                }
            } catch (error) {
                console.warn('Snapshot BOQ Lead localStorage tidak dapat dibaca:', error);
            }

            // Fallback ke IndexedDB.
            try {
                const snapshot = await this.loadWorkflowSnapshot(projectKey);
                if (snapshot?.boq?.items?.length) return snapshot;
            } catch (error) {
                console.warn('Snapshot BOQ Lead IndexedDB tidak dapat dibaca:', error);
            }

            return null;
        },

        async loadWorkflowSnapshot(projectKey) {
            try {
                const db = await this.getWorkflowDB();
                const snapshot = await new Promise((resolve, reject) => {
                    const tx = db.transaction('snapshots', 'readonly');
                    const req = tx.objectStore('snapshots').get(String(projectKey));
                    req.onsuccess = () => resolve(req.result || null);
                    req.onerror = () => reject(req.error || new Error('Gagal membaca snapshot.'));
                });
                db.close();
                return snapshot;
            } catch (error) {
                console.warn('IndexedDB workflow snapshot tidak dapat dibaca:', error);
                return null;
            }
        },

        async hydrateBomBoqFromIndexedDB() {
            for (const [projectKey, project] of Object.entries(this.allProjectsData || {})) {
                if (!project?.meta) continue;

                const snapshot = await this.loadWorkflowSnapshot(projectKey);
                if (!snapshot) continue;

                // Snapshot hanya menggantikan data jika lebih lengkap/lebih baru.
                const localBom = project.meta.bom;
                const localBoq = project.meta.boq;

                if (snapshot.bom?.details?.length &&
                    (!localBom?.details?.length ||
                     Number(snapshot.updatedAt || 0) >= Number(localBom.updatedAtEpoch || 0))) {
                    project.meta.bom = {
                        ...snapshot.bom,
                        updatedAtEpoch: snapshot.updatedAt
                    };
                }

                if (snapshot.boq?.items?.length &&
                    (!localBoq?.items?.length ||
                     Number(snapshot.updatedAt || 0) >= Number(localBoq.updatedAtEpoch || 0))) {
                    project.meta.boq = {
                        ...snapshot.boq,
                        updatedAtEpoch: snapshot.updatedAt
                    };
                }

                if (snapshot.workflowStatus &&
                    (!project.meta.workflowStatus || project.meta.workflowStatus === 'DRAFT')) {
                    project.meta.workflowStatus = snapshot.workflowStatus;
                }
            }

            this.bomBoqStorageReady = true;
        },

        ensureBomBoqReady() {
            if (this.bomBoqStorageReady) return Promise.resolve();
            if (!this.bomBoqStoragePromise) {
                this.bomBoqStoragePromise = this.hydrateBomBoqFromIndexedDB()
                    .catch(error => {
                        console.warn('Hydrasi BOM/BOQ gagal:', error);
                        this.bomBoqStorageReady = true;
                    });
            }
            return this.bomBoqStoragePromise;
        },

        loadBomBoqStorage() {
            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                if (!project?.meta) return;
                try {
                    const bomRaw = localStorage.getItem(this.getBomStorageKey(projectKey));
                    const boqRaw = localStorage.getItem(this.getBoqStorageKey(projectKey));

                    if (bomRaw) {
                        const bom = JSON.parse(bomRaw);
                        if (bom && typeof bom === 'object') project.meta.bom = bom;
                    }

                    if (boqRaw) {
                        const boq = JSON.parse(boqRaw);
                        if (boq && typeof boq === 'object') project.meta.boq = boq;
                    }
                } catch (error) {
                    console.warn('BOM/BOQ localStorage tidak dapat dibaca:', error);
                }
            });

            // Baca fallback IndexedDB sebelum Estimator melakukan kalkulasi.
            this.bomBoqStoragePromise = this.hydrateBomBoqFromIndexedDB()
                .catch(error => {
                    console.warn('Hydrasi BOM/BOQ IndexedDB gagal:', error);
                    this.bomBoqStorageReady = true;
                });
        },

        saveBomBoqStorage() {
            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                const bom = project?.meta?.bom;
                const boq = project?.meta?.boq;

                // Tandai snapshot dengan epoch agar versi terbaru dapat dipilih.
                const now = Date.now();
                if (bom) bom.updatedAtEpoch = now;
                if (boq) boq.updatedAtEpoch = now;

                try {
                    if (bom) localStorage.setItem(this.getBomStorageKey(projectKey), JSON.stringify(bom));
                    else localStorage.removeItem(this.getBomStorageKey(projectKey));

                    if (boq) localStorage.setItem(this.getBoqStorageKey(projectKey), JSON.stringify(boq));
                    else localStorage.removeItem(this.getBoqStorageKey(projectKey));
                } catch (error) {
                    // Jangan menghapus data yang sudah ada jika quota localStorage penuh.
                    console.warn('localStorage BOM/BOQ penuh; memakai IndexedDB sebagai penyimpanan workflow.', error);
                }

                // Selalu simpan snapshot workflow ke IndexedDB.
                this.saveWorkflowSnapshot(projectKey, project);
            });
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
            // Metadata ringan; BOM/BOQ besar selalu disimpan terpisah.
            const metas = {};
            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                const meta = project?.meta || {};
                const { bom, boq, ...lightMeta } = meta;
                metas[projectKey] = lightMeta;
            });
            try {
                localStorage.setItem(this.getProjectMetaStorageKey(), JSON.stringify(metas));
            } catch (error) {
                console.error('Metadata project gagal disimpan:', error);
            }
            this.saveBomBoqStorage();

            // Tabel besar tidak disalin ulang seluruhnya ke localStorage setiap
            // saveStorage(). Seed data Project 3000 tetap tersedia dari seed-data.js.
            // Operasi edit/import menyimpan sheet yang benar-benar berubah melalui
            // saveTableStorage(projectKey, sheetName), sehingga data besar tidak
            // memenuhi quota localStorage.
        },

        switchProject() {
            this.loadTableStorage();
            this.loadBomBoqStorage();
            this.currentPage = 1;
            this.globalSearch = '';
            this.columnFilters = {};
            this.closeColumnFilter();
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
                files: [],
                fileNames: []
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
            const files = Array.from(event?.target?.files || []);
            const valid = files.filter(file => /\.(xlsx|xls)$/i.test(file.name));

            if (files.length && valid.length !== files.length) {
                this.createProjectMessageType = 'error';
                this.createProjectMessage = 'Semua file harus berformat .xlsx atau .xls.';
                event.target.value = '';
            }

            this.createProjectForm.files = valid;
            this.createProjectForm.fileNames = valid.map(file => file.name);
            if (!valid.length) {
                this.createProjectForm.files = [];
                this.createProjectForm.fileNames = [];
            }
            this.createProjectMessage = valid.length ? '' : this.createProjectMessage;
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
                    progress: '0%',
                    workflowStatus: 'DRAFT',
                    revisionNotes: '',
                    bom: null,
                    boq: null
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

            if (this.createProjectForm.initialization === 'import' && !(this.createProjectForm.files || []).length) {
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
                    const files = this.createProjectForm.files || [];
                    const pendingKeysBySheet = new Map();
                    let firstImportedSheet = '';

                    for (const file of files) {
                        const workbook = await this.readExcelWorkbook(file);
                        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                            throw new Error(`Workbook ${file.name} tidak memiliki sheet.`);
                        }

                        workbook.SheetNames.forEach((sheetName, sheetIndex) => {
                            const worksheet = workbook.Sheets[sheetName];
                            if (!worksheet || !worksheet['!ref']) return;

                            const parsed = this.worksheetToExactRows(worksheet);
                            if (!parsed.rows.length) return;

                            const cleanSheetName = String(sheetName).trim() || `Sheet ${sheetIndex + 1}`;
                            if (!Array.isArray(newProject[cleanSheetName])) newProject[cleanSheetName] = [];
                            if (!pendingKeysBySheet.has(cleanSheetName)) {
                                pendingKeysBySheet.set(
                                    cleanSheetName,
                                    new Set(newProject[cleanSheetName].map(row => this.canonicalImportRow(row)))
                                );
                            }

                            const keys = pendingKeysBySheet.get(cleanSheetName);
                            parsed.rows.forEach(row => {
                                const copy = { ...row };
                                const key = this.canonicalImportRow(copy);
                                if (keys.has(key)) return;
                                keys.add(key);
                                newProject[cleanSheetName].push(copy);
                            });

                            if (!firstImportedSheet) firstImportedSheet = cleanSheetName;
                        });
                    }

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
                    ? `Project ${code} berhasil dibuat dan ${this.createProjectForm.fileNames.length} file Excel berhasil diimport.`
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

            if (email === 'admin@tripatra.com' && pass === 'admin123') {
                this.loginForm.role = 'System Administrator';
                this.isLoggedIn = true;
                this.currentDashboardTab = 'team';
            } else if (email === 'engineer@tripatra.com' && pass === 'engineer123') {
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
                return;
            }

            this.saveAuthState();
            this.savePageState();
        },

        logout() {
            this.clearAuthState();
            this.isLoggedIn = false;
            this.loginForm.user = '';
            this.loginForm.pass = '';
            this.loginForm.role = '';
            // Setelah logout, halaman berikutnya kembali ke login.
            this.currentDashboardTab = 'home';
            this.currentPage = 1;
            this.savePageState();
        },

        openTeamMemberModal() {
            if (this.loginForm.role !== 'System Administrator') return;
            this.teamMemberEditingId = null;
            this.teamMemberForm = { name: '', email: '', role: 'Piping Engineer', projects: [], active: true };
            this.showTeamMemberModal = true;
        },
        editTeamMember(member) {
            if (this.loginForm.role !== 'System Administrator') return;
            this.teamMemberEditingId = member.id;
            this.teamMemberForm = { ...member, projects: [...(member.projects || [])] };
            this.showTeamMemberModal = true;
        },
        closeTeamMemberModal() { this.showTeamMemberModal = false; this.teamMemberEditingId = null; },
        saveTeamMember() {
            if (this.loginForm.role !== 'System Administrator') return;
            const data = { id: this.teamMemberEditingId || Date.now(), name: this.teamMemberForm.name.trim(), email: this.teamMemberForm.email.trim().toLowerCase(), role: this.teamMemberForm.role, projects: [...(this.teamMemberForm.projects || [])], active: !!this.teamMemberForm.active };
            if (!data.name || !data.email) return;
            if (this.teamMemberEditingId) { const i = this.teamMembers.findIndex(m => m.id === this.teamMemberEditingId); if (i !== -1) this.teamMembers[i] = data; }
            else { if (this.teamMembers.some(m => m.email === data.email)) { alert('Email anggota sudah terdaftar.'); return; } this.teamMembers.push(data); }
            localStorage.setItem('tripatra_team_members_v1', JSON.stringify(this.teamMembers));
            this.closeTeamMemberModal();
        },
        toggleTeamMember(member) {
            if (this.loginForm.role !== 'System Administrator') return;
            member.active = !member.active;
            localStorage.setItem('tripatra_team_members_v1', JSON.stringify(this.teamMembers));
        },
        deleteTeamMember(member) {
            if (this.loginForm.role !== 'System Administrator' || member.email === this.loginForm.user) return;
            if (!confirm(`Hapus anggota ${member.name}?`)) return;
            this.teamMembers = this.teamMembers.filter(m => m.id !== member.id);
            localStorage.setItem('tripatra_team_members_v1', JSON.stringify(this.teamMembers));
        },
        loadTeamMembers() {
            try { const raw = localStorage.getItem('tripatra_team_members_v1'); if (raw) { const parsed = JSON.parse(raw); if (Array.isArray(parsed) && parsed.length) this.teamMembers = parsed; } } catch (e) { console.warn('Data Team & Roles tidak dapat dibaca:', e); }
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
            const rows = this.currentRows || [];

            // IMPORT EXCEL = DATA-DRIVEN:
            // Setelah Excel diimport, kolom tabel selalu dibentuk dari header
            // yang benar-benar ada pada data Excel. Tidak lagi dipaksa mengikuti
            // schema lama sehingga kolom tidak tertukar/bergeser.
            if (rows.length) {
                const columns = [];
                const seen = new Set();

                rows.forEach(row => {
                    Object.keys(row || {}).forEach(column => {
                        if (column === 'Number' || column === 'No') return;
                        const key = String(column);
                        if (seen.has(key)) return;
                        seen.add(key);
                        columns.push(key);
                    });
                });

                if (columns.length) return columns;
            }

            // Saat sheet masih kosong, gunakan schema bawaan yang sudah ada.
            const exactSchema = this.getExactSchema(this.activeSheet);
            if (exactSchema) {
                return exactSchema.filter(col => col !== 'Number' && col !== 'No');
            }

            return ['Material', 'Material Code', 'Spec', 'Size', 'Pressure Class', 'Line Number Tag', 'Status'];
        },

        get freezeColumns() {
            // Freeze dari kiri sampai kolom batas berikut (termasuk kolom
            // batasnya), sisanya tetap bisa digeser. Kolom "No" di paling
            // kiri SELALU ikut freeze secara terpisah (lihat getNoFreezeClass).
            //
            //   Tabel                                  | Freeze dari kiri sampai
            //   ---------------------------------------|-------------------------
            //   MTO Pipe & Fitting Valve                | Material
            //   SP Items                                | Spec
            //   Pipe Support                             | Material
            //   Line List                                | Seq. No
            //   Vessel / Tank / Pump / Misc Equipment /
            //   Equipment / Piping and Equipment         | Long Description (Size)
            const boundaryBySheet = {
                'Valve': 'Material',
                'SP Items': 'Spec',
                'Support': 'Material',
                'LineList': 'Seq. No',
                'Vessel': 'Long Description (Size)',
                'Tank': 'Long Description (Size)',
                'Pump': 'Long Description (Size)',
                'Misc Equipment': 'Long Description (Size)',
                'Equipment': 'Long Description (Size)',
                'Piping and Equipment': 'Long Description (Size)'
            };

            const columns = this.currentColumns || [];
            // Sheet MTO lain (Tee, Flange, Elbow, dst) mengikuti pola yang
            // sama seperti Valve: freeze sampai Material.
            const boundary = boundaryBySheet[this.activeSheet]
                || boundaryBySheet['Valve'];

            const boundaryIndex = columns.indexOf(boundary);
            if (boundaryIndex !== -1) return columns.slice(0, boundaryIndex + 1);

            // Fallback: kalau kolom batas tidak ditemukan pada data sheet ini
            // (mis. hasil import dengan header berbeda), coba freeze sampai
            // "Long Description (Size)" atau "Long Description (Family)"
            // (kolom deskripsi/identitas baris yang paling umum dipakai di
            // seluruh sheet MTO), baru fallback ke kolom pertama.
            const genericBoundary = columns.includes('Long Description (Size)')
                ? 'Long Description (Size)'
                : (columns.includes('Long Description (Family)')
                    ? 'Long Description (Family)'
                    : null);

            if (genericBoundary) {
                const idx = columns.indexOf(genericBoundary);
                return columns.slice(0, idx + 1);
            }

            return columns.length ? [columns[0]] : [];
        },

        isFreezeColumn(column) {
            return this.freezeColumns.includes(column);
        },

        // Sheet keluarga Equipment (Vessel, Tank, Pump, Misc Equipment,
        // Equipment) SELALU pakai nomor urut baris polos (1, 2, 3, ...),
        // tidak memakai nomor tag asli dari Excel (mis. 6001, 3000A),
        // supaya kolom NO konsisten sebagai nomor baris biasa.
        //
        // Sheet lain tetap memakai Number/No asli kalau memang valid;
        // hanya placeholder kosong atau "?" yang diganti nomor urut.
        getRowNumber(row, index) {
            const sequentialOnlySheets = [
                'Vessel', 'Tank', 'Pump', 'Misc Equipment', 'Equipment'
            ];

            const sequentialNumber = () =>
                String((this.currentPage - 1) * this.itemsPerPage + index + 1);

            if (sequentialOnlySheets.includes(this.activeSheet)) {
                return sequentialNumber();
            }

            const raw = row?.Number ?? row?.No ?? '';
            const normalized = String(raw).trim();
            const isPlaceholder = normalized === '' || normalized === '?';

            if (!isPlaceholder) return normalized;

            return sequentialNumber().padStart(2, '0');
        },

        // Kolom "No" di paling kiri selalu ikut dibekukan (index 0),
        // baru diikuti kolom-kolom dari freezeColumns (index 1, 2, dst)
        // supaya semuanya berjejer rapi tanpa saling menumpuk.
        getNoFreezeClass() {
            return 'freeze-col freeze-col-0';
        },

        getFreezeClass(column) {
            const freezeIndex = this.freezeColumns.indexOf(column);
            if (freezeIndex === -1) return '';
            return `freeze-col freeze-col-${freezeIndex + 1}`;
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

        // ==========================================================
        // VALIDASI INPUT MASTER DATA
        // Aturan mengikuti data master yang sudah ada pada seed-data.js.
        // - number  : hanya angka/desimal
        // - digits  : hanya digit, cocok untuk No/Seq. No yang bisa
        //             membutuhkan leading zero
        // - uppercase/lowercase : mengikuti pola kolom
        // - text    : tidak dipaksa mengubah kapitalisasi
        //
        // Nilai lama/import Excel TIDAK diubah. Aturan ini hanya bekerja
        // ketika user mengetik/mengedit cell.
        // ==========================================================
        getSeedColumnValues(sheetName, column) {
            try {
                const seedRows = window.BMBQ_SEED_DATA?.[sheetName];
                if (!Array.isArray(seedRows)) return [];
                return [...new Set(
                    seedRows
                        .map(row => String(row?.[column] ?? '').trim())
                        .filter(Boolean)
                )];
            } catch (e) {
                return [];
            }
        },

        getColumnInputRule(sheetName, column) {
            const sheet = String(sheetName || '').trim();
            const col = String(column || '').trim();

            // Field yang memang merupakan nomor/kode angka tertentu.
            // Seq. No sengaja digits agar 0001 tetap bisa dipertahankan.
            const digitsBySheet = {
                'LineList': new Set(['No', 'Seq. No']),
                'SP Items': new Set(['Item Count']),
                'Nozzle': new Set(['Number']),
                'Vessel': new Set(['Number']),
                'Misc Equipment': new Set(['Number'])
            };
            if (digitsBySheet[sheet]?.has(col)) return col === 'Seq. No' ? 'digits' : 'number';

            // Kolom SIZE bukan angka HTML murni karena data engineering menggunakan
            // format seperti 10", 3/4", dan 1 1/2". Tetap dianggap sebagai field
            // ukuran numerik: hanya digit + pecahan + titik + spasi + tanda inch.
            const sizeColumns = new Set([
                'Size', 'Size 1', 'Size 2', 'Nominal Size', 'Line Size (Inch)',
                'Line Size', 'Available Size'
            ]);
            if (sizeColumns.has(col)) return 'size';

            // Material Code adalah kode engineering: bukan angka murni dan bukan
            // free-text biasa. Nilai yang sudah ada di master dipertahankan
            // kapitalisasinya (contoh: ASTM A216 Gr WPB). Nilai baru yang diketik
            // dengan huruf kecil akan dinormalisasi ke uppercase agar tidak ada
            // kode baru yang tersimpan seluruhnya dalam lowercase.
            if (col === 'Material Code') return 'material-code';

            // Kolom angka yang secara engineering memang berisi nilai numerik.
            const numericColumns = new Set([
                'Area', 'Actuator Height', 'Actuator Width', 'Bottom of Pipe',
                'Branch Angle1', 'Branch Angle2', 'Center of Gravity X',
                'Center of Gravity Y', 'Center of Gravity Z',
                'COP Elevation (Port 1)', 'Curve Radius', 'Cut Length',
                'Cutback Angle', 'Design Pressure Factor', 'Eccentricity',
                'Engagement Length', 'Flange Thickness', 'Fixed Length',
                'Insulation Thickness', 'Length', 'Linear Weight',
                'Matching Pipe OD', 'Max Length', 'Min Length',
                'Minimum Cut Length', 'Nominal Diameter', 'Number In Set',
                'Path Angle', 'PnPID', 'Pressure Class', 'Schedule',
                'Thickness [mm]', 'Top of Pipe', 'Weight',
                'X Coordinate (Port 1)', 'Y Coordinate (Port 1)',
                'Center of Gravity Z', 'Operating Temperature',
                'Design Temperature', 'Mass Flow\\n[kg/h]',
                'Volume Flow\\n[m1.5/h]', 'Density\\n[kg/m3]',
                'Viscosity\\n[cP]', 'Pressure [Barg]', 'Loop Number'
            ]);
            if (numericColumns.has(col)) return 'number';

            // Number pada equipment tertentu adalah identifier alfanumerik
            // (contoh: 2002A, 3000A), jadi tidak dipaksa menjadi angka.

            const values = this.getSeedColumnValues(sheet, col);
            if (!values.length) {
                // Kolom kosong pada seed tetap diberi aturan berdasarkan nama.
                // Yang ambigu (Manufacturer, Remarks, From, To, DWG, Tag, dll.)
                // tetap menjadi text agar tidak merusak kode engineering.
                return 'text';
            }

            const isNumeric = values.filter(v =>
                /^[-+]?\d+(?:\.\d+)?$/.test(v)
            ).length;
            const isUpper = values.filter(v =>
                v === v.toUpperCase() && v !== v.toLowerCase()
            ).length;
            const isLower = values.filter(v =>
                v === v.toLowerCase() && v !== v.toUpperCase()
            ).length;

            const total = values.length;

            // Ambang 90% dipakai supaya satu-dua data legacy yang berbeda
            // tidak mengubah aturan kolom secara keseluruhan.
            if (isNumeric / total >= 0.90) return 'number';
            if (isUpper / total >= 0.90) return 'uppercase';
            if (isLower / total >= 0.90) return 'lowercase';

            // Mixed/Title/campuran tetap text. Contoh penting:
            // Material Code, Size, Long Description (Family) Valve,
            // Short Description, Tag, Part Subtype, Design Std.
            return 'text';
        },

        getCellInputMode(sheetName, column) {
            return this.getColumnInputRule(sheetName, column) === 'number' ? 'decimal' : 'text';
        },

        getCellInputType(sheetName, column) {
            // Tetap text agar format engineering seperti 3/4", CS150,
            // 2002A, 10"x8" tidak dirusak oleh HTML number input.
            return 'text';
        },

        normalizeNumericInput(value) {
            let v = String(value ?? '').replace(/,/g, '.');
            // Pertahankan minus hanya di posisi pertama.
            const negative = v.trim().startsWith('-');
            v = v.replace(/[^0-9.]/g, '');
            const parts = v.split('.');
            if (parts.length > 2) v = parts.shift() + '.' + parts.join('');
            v = v.replace(/^(\d*\.\d*)\..*$/, '$1');
            return negative && v ? '-' + v : v;
        },

        normalizeDigitsInput(value) {
            return String(value ?? '').replace(/\D/g, '');
        },

        normalizeSizeInput(value) {
            let v = String(value ?? '');
            // Hanya izinkan angka, pecahan (/), desimal (.), spasi, dan tanda inch.
            // Huruf seperti abc/xyz langsung dibuang saat user mengetik.
            v = v.replace(/[^0-9\/\.\s"]/g, '');
            v = v.replace(/\s+/g, ' ').trim();
            // Jangan izinkan dua slash atau dua titik berturut-turut.
            v = v.replace(/\/{2,}/g, '/').replace(/\.{2,}/g, '.');
            return v;
        },

        normalizeMaterialCodeInput(sheetName, column, value) {
            const raw = String(value ?? '').trim();
            if (!raw) return '';

            // Prioritaskan format canonical yang memang sudah dipakai oleh master data.
            const canonical = this.findCanonicalCellValue(sheetName, column, raw);
            if (canonical !== null) return canonical;

            // Kode baru tidak boleh tersimpan seluruhnya dalam lowercase.
            return raw.toUpperCase();
        },

        findCanonicalCellValue(sheetName, column, value) {
            const raw = String(value ?? '').trim();
            if (!raw) return '';
            const values = this.getSeedColumnValues(sheetName, column);
            const found = values.find(v => v.toLowerCase() === raw.toLowerCase());
            return found ?? null;
        },

        formatCellInput(row, column, event) {
            if (!row || !column || this.loginForm.role !== 'Piping Engineer') return;

            const raw = String(event?.target?.value ?? row[column] ?? '');
            const rule = this.getColumnInputRule(this.activeSheet, column);
            let formatted = raw;

            if (rule === 'number') {
                formatted = this.normalizeNumericInput(raw);
            } else if (rule === 'digits') {
                formatted = this.normalizeDigitsInput(raw);
            } else if (rule === 'size') {
                formatted = this.normalizeSizeInput(raw);
            } else if (rule === 'material-code') {
                formatted = this.normalizeMaterialCodeInput(this.activeSheet, column, raw);
            } else if (rule === 'uppercase') {
                formatted = raw.toUpperCase();
            } else if (rule === 'lowercase') {
                formatted = raw.toLowerCase();
            } else {
                // Mixed/text: jangan paksa UPPERCASE/title case.
                // Jika user mengetik nilai yang memang sudah ada di master,
                // gunakan kapitalisasi canonical dari master.
                const canonical = this.findCanonicalCellValue(this.activeSheet, column, raw);
                formatted = canonical !== null ? canonical : raw;
            }

            row[column] = formatted;
            if (event?.target && event.target.value !== formatted) {
                event.target.value = formatted;
            }
        },

        validateCellInput(row, column) {
            if (!row || !column || this.loginForm.role !== 'Piping Engineer') return;

            const current = String(row[column] ?? '');
            const rule = this.getColumnInputRule(this.activeSheet, column);
            let formatted = current;

            if (rule === 'number') formatted = this.normalizeNumericInput(current);
            else if (rule === 'digits') formatted = this.normalizeDigitsInput(current);
            else if (rule === 'size') formatted = this.normalizeSizeInput(current);
            else if (rule === 'material-code') formatted = this.normalizeMaterialCodeInput(this.activeSheet, column, current);
            else if (rule === 'uppercase') formatted = current.toUpperCase();
            else if (rule === 'lowercase') formatted = current.toLowerCase();
            else {
                const canonical = this.findCanonicalCellValue(this.activeSheet, column, current);
                if (canonical !== null) formatted = canonical;
            }

            row[column] = formatted;
            this.markRowEdited(row);
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


        makeImportedSheetName(fileName, sourceSheetName, dataType = 'Generic', reservedNames = new Set()) {
            const clean = value => String(value || '')
                .replace(/\u00A0/g, ' ')
                .replace(/[\\/:*?\"<>|]/g, '-')
                .replace(/\s+/g, ' ')
                .trim();

            const baseFile = clean(String(fileName || '').replace(/\.(xlsx|xls)$/i, ''));
            const baseSheet = clean(sourceSheetName) || 'Sheet';
            const prefix = dataType === 'LineList' ? 'IMPORT - Line List' : 'IMPORT';
            let candidate = `${prefix} - ${baseFile} - ${baseSheet}`;
            // Sheet navigation is visual; keep names readable while ensuring uniqueness.
            if (candidate.length > 80) candidate = candidate.slice(0, 80).trim();

            let name = candidate;
            let counter = 2;
            while (this.sheets.some(s => String(s).toLowerCase() === name.toLowerCase()) || [...reservedNames].some(s => String(s).toLowerCase() === name.toLowerCase())) {
                const suffix = ` (${counter++})`;
                name = candidate.slice(0, Math.max(1, 80 - suffix.length)).trim() + suffix;
            }
            return name;
        },

        async importExcelFile(event) {
            const input = event?.target;
            const files = Array.from(input?.files || []);
            if (!files.length) return;

            const resetInput = () => { if (input) input.value = ''; };
            const invalid = files.filter(file => !/\.(xlsx|xls)$/i.test(file.name));
            if (invalid.length) {
                alert('Import dibatalkan. Semua file harus berformat .xlsx atau .xls.');
                resetInput();
                return;
            }

            const projectKey = this.activeProject;
            if (!projectKey) {
                alert('Project aktif belum dipilih.');
                resetInput();
                return;
            }
            if (!this.allProjectsData[projectKey]) {
                this.allProjectsData[projectKey] = this.createBlankProject(projectKey, 'Piping & Equipment', '');
            }

            this.isImportingExcel = true;
            const payload = [];
            const errors = [];
            const fileSummary = [];
            const reservedImportSheets = new Set();
            let totalRows = 0;
            let duplicateRows = 0;
            let skippedRows = 0;

            try {
                for (const file of files) {
                    try {
                        const workbook = await this.readExcelWorkbook(file);
                        let fileRows = 0;
                        let fileSheets = 0;

                        for (const sourceSheetName of workbook.SheetNames || []) {
                            const worksheet = workbook.Sheets[sourceSheetName];
                            if (!worksheet || !worksheet['!ref']) { skippedRows++; continue; }
                            try {
                                const parsed = this.worksheetToExactRows(worksheet);
                                if (!parsed.rows.length) { skippedRows++; continue; }

                                const normalizedSource = String(sourceSheetName).replace(/\u00A0/g, ' ').trim();

                                // PENTING: hasil import SELALU menjadi dataset/sheet baru.
                                // Walaupun struktur Excel sama persis dengan LineList, data
                                // tidak boleh digabung dengan LineList existing. Dengan cara
                                // ini data lama tetap utuh dan setiap file/worksheet punya
                                // ruang data sendiri yang mudah dilacak.
                                const targetSheet = this.makeImportedSheetName(
                                    file.name,
                                    normalizedSource,
                                    parsed.dataType,
                                    reservedImportSheets
                                );
                                reservedImportSheets.add(targetSheet);

                                // Jangan menyentuh dataset/sheet existing saat PREVIEW.
                                // Sheet baru dibuat resmi hanya setelah user menekan Import Data Baru. 

                                // Duplikat hanya dicek DI DALAM dataset import ini, bukan
                                // terhadap data existing pada LineList/Valve/master lain.
                                const seenInThisDataset = new Set();
                                const uniqueNewRows = [];
                                for (const row of parsed.rows) {
                                    const key = this.canonicalImportRow(row);
                                    if (seenInThisDataset.has(key)) {
                                        duplicateRows++;
                                        continue;
                                    }
                                    seenInThisDataset.add(key);
                                    uniqueNewRows.push({ ...row });
                                }

                                payload.push({
                                    fileName: file.name,
                                    sourceSheet: normalizedSource,
                                    targetSheet,
                                    headers: parsed.headers,
                                    dataType: parsed.dataType,
                                    rows: uniqueNewRows
                                });
                                totalRows += parsed.rows.length;
                                fileRows += uniqueNewRows.length;
                                fileSheets++;
                            } catch (err) {
                                errors.push(`${file.name} / ${sourceSheetName}: ${err.message}`);
                            }
                        }
                        fileSummary.push({ name: file.name, sheets: fileSheets, newRows: fileRows });
                    } catch (err) {
                        errors.push(`${file.name}: ${err.message}`);
                    }
                }

                const newRows = payload.reduce((sum, item) => sum + item.rows.length, 0);
                if (!payload.length) throw new Error('Tidak ada data tabel yang dapat dipreview dari file yang dipilih.');

                this.importPreview = {
                    files: fileSummary,
                    sheets: payload.map(p => ({
                        fileName: p.fileName,
                        sourceSheet: p.sourceSheet,
                        targetSheet: p.targetSheet,
                        rows: p.rows.length,
                        columns: p.headers.length,
                        dataType: p.dataType
                    })),
                    totalRows, newRows, duplicateRows, skippedRows, errors, payload
                };
                this.showImportPreview = true;
            } catch (error) {
                alert('Import Excel gagal.\n\n' + (error?.message || 'Format Excel tidak sesuai.'));
            } finally {
                this.isImportingExcel = false;
                resetInput();
            }
        },

        cancelImportPreview() {
            this.showImportPreview = false;
            this.importPreview = { files: [], sheets: [], totalRows: 0, newRows: 0, duplicateRows: 0, skippedRows: 0, errors: [], payload: [] };
        },

        commitImportPreview() {
            if (this.isImportingExcel || !this.importPreview.payload.length) return;
            try {
                const project = this.allProjectsData[this.activeProject];
                if (!project) throw new Error('Project aktif tidak ditemukan.');

                let importedRows = 0;
                for (const item of this.importPreview.payload) {
                    const targetSheet = item.targetSheet;
                    if (!this.sheets.includes(targetSheet)) this.sheets.push(targetSheet);
                    const target = Array.isArray(project[targetSheet]) ? project[targetSheet] : [];
                    const rowsToAdd = item.rows.map(row => ({ ...row }));
                    // Dataset import baru selalu dimulai dari kosong. Tidak pernah concat
                    // dengan master LineList/Valve atau dataset import lain.
                    project[targetSheet] = target.concat(rowsToAdd);
                    importedRows += rowsToAdd.length;
                }

                this.refreshSheetList();

                const targetWithMostRows = [...this.importPreview.payload]
                    .filter(item => item.rows.length > 0)
                    .sort((a, b) => b.rows.length - a.rows.length)[0];
                if (targetWithMostRows) this.activeSheet = targetWithMostRows.targetSheet;

                this.currentPage = 1;
                this.globalSearch = '';
                this.columnFilters = {};
                this.closeColumnFilter();
                this.tableRenderKey++;

                const duplicateRows = this.importPreview.duplicateRows;
                const sheetCount = this.importPreview.sheets.length;

                // Tutup preview sebelum browser melakukan repaint.
                // Tidak memakai alert setelah commit karena alert memblokir repaint
                // sehingga preview lama terlihat seolah-olah belum tertutup.
                this.cancelImportPreview();
                this.saveStorage();

                this.$nextTick(() => {
                    requestAnimationFrame(() => {
                        this.updateTableScrollbar();
                        this.showImportSuccess(
                            `${importedRows} data baru berhasil ditambahkan dari ${sheetCount} worksheet. ` +
                            `${duplicateRows} duplikat dilewati. Data lama tetap aman.`
                        );
                    });
                });
            } catch (error) {
                console.error('[IMPORT COMMIT] Gagal:', error);
                alert('Data belum disimpan karena terjadi kesalahan saat commit import.\n\n' +
                    (error?.message || 'Kesalahan tidak diketahui.'));
            }
        },

        showImportSuccess(message) {
            this.importSuccessMessage = String(message || '');
            this.importSuccessVisible = true;
            clearTimeout(this._importSuccessTimer);
            this._importSuccessTimer = setTimeout(() => {
                this.importSuccessVisible = false;
                this.importSuccessMessage = '';
            }, 4500);
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
        formatCurrency(value) {
            const n = Number(value) || 0;
            // BOQ/MTO menggunakan Rupiah sebagai mata uang default.
            // Tetap menghormati pilihan USD jika user memang memilihnya.
            if (this.boqCurrency === 'IDR') {
                return new Intl.NumberFormat('id-ID', {
                    style: 'currency',
                    currency: 'IDR',
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 0
                }).format(n);
            }
            return new Intl.NumberFormat('en-US', {
                style: 'currency',
                currency: 'USD',
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }).format(n);
        },

        // WORKFLOW ENGINEER -> ESTIMATOR -> LEAD -> ESTIMATOR
        // Jika revisi menyangkut BOM/teknis, ESTIMATOR dapat mengirim kembali
        // ke ENGINEER -> ESTIMATOR -> LEAD.
        // ==========================================================
        get workflowStatus() {
            return this.allProjectsData?.[this.activeProject]?.meta?.workflowStatus || 'DRAFT';
        },

        get workflowStatusText() {
            const map = {
                DRAFT: 'Draft - Pekerjaan Engineer',
                BOM_CALCULATED: 'BOM / BQ Sudah Dihitung',
                SUBMITTED_TO_ESTIMATOR: 'Menunggu Estimator',
                BOQ_CALCULATED: 'BOQ Sudah Dihitung',
                SUBMITTED_TO_LEAD: 'Menunggu Lead Review',
                REVISION_REQUIRED: 'Revisi Diperlukan - Kembali ke Engineer',
                ESTIMATOR_REVISION_REQUIRED: 'Revisi Diperlukan - Kembali ke Estimator',
                ENGINEER_REVISION_REQUIRED: 'Revisi BOM Diperlukan - Kembali ke Engineer',
                APPROVED: 'Approved - Laporan Final'
            };
            return map[this.workflowStatus] || this.workflowStatus;
        },

        get workflowStatusClass() {
            const s = this.workflowStatus;
            if (s === 'APPROVED') return 'bg-emerald-100 text-emerald-700 border-emerald-200';
            if (s === 'REVISION_REQUIRED' || s === 'ESTIMATOR_REVISION_REQUIRED' || s === 'ENGINEER_REVISION_REQUIRED') return 'bg-amber-100 text-amber-700 border-amber-200';
            if (s === 'SUBMITTED_TO_LEAD' || s === 'SUBMITTED_TO_ESTIMATOR') return 'bg-sky-100 text-sky-700 border-sky-200';
            return 'bg-slate-100 text-slate-700 border-slate-200';
        },

        saveProjectMetaOnly() {
            const metas = {};
            Object.entries(this.allProjectsData || {}).forEach(([projectKey, project]) => {
                const meta = project?.meta || {};
                // Metadata hanya menyimpan status/proyek. BOM dan BOQ besar disimpan
                // pada storage terpisah agar hasil 641 item tidak hilang karena quota.
                const { bom, boq, ...lightMeta } = meta;
                metas[projectKey] = lightMeta;
            });
            try {
                localStorage.setItem(this.getProjectMetaStorageKey(), JSON.stringify(metas));
                this.saveBomBoqStorage();
            } catch (error) {
                console.error('Gagal menyimpan metadata project:', error);
                alert('Metadata project gagal disimpan. Data BOM/BOQ tetap dicoba disimpan pada storage terpisah.');
                this.saveBomBoqStorage();
            }
        },

        setWorkflowStatus(status, note = '') {
            const project = this.allProjectsData?.[this.activeProject];
            if (!project) return;
            if (!project.meta) project.meta = {};
            project.meta.workflowStatus = status;
            project.meta.revisionNotes = note || project.meta.revisionNotes || '';
            project.meta.workflowUpdatedAt = new Date().toLocaleString('id-ID');
            // Jangan serialize semua tabel Excel setiap perubahan status.
            this.saveProjectMetaOnly();
        },

        get bomDetailsAll() {
            return this.allProjectsData?.[this.activeProject]?.meta?.bom?.details || [];
        },

        get bomCategoryOptions() {
            const order = [
                'BaseSupport', 'Support', 'DummyLeg', 'Tee', 'Single Branch Fitting',
                'Pipe', 'Instrument', 'Flange', 'Elbow', 'Coupling', 'Pipe Run Component',
                'Socketweld', 'Gasket', 'Buttweld', 'Bolt Set', 'Fasteners'
            ];
            const counts = {};
            this.bomDetailsAll.forEach(r => {
                const c = String(r?.component || '').trim();
                if (c) counts[c] = (counts[c] || 0) + 1;
            });
            return order.filter(c => counts[c] > 0).map(c => ({ name: c, count: counts[c] }));
        },

        get bomFilteredDetails() {
            const q = String(this.bomFilterSearch || '').trim().toLowerCase();
            const cat = String(this.bomFilterCategory || 'ALL');
            return this.bomDetailsAll.filter(r => {
                const component = String(r?.component || '').trim();
                if (cat !== 'ALL' && component !== cat) return false;
                if (!q) return true;
                return [component, r?.description, r?.material, r?.lineNo, r?.size1, r?.size2, r?.unit]
                    .some(v => String(v ?? '').toLowerCase().includes(q));
            });
        },

        getBomSelectionKey(row, index = -1) {
            // Gunakan identitas item BOM, bukan nomor urut tampilan, agar pilihan
            // tetap mengikuti item walaupun tabel difilter atau dihitung ulang.
            return [
                row?.sheet, row?.no, row?.component, row?.description,
                row?.size1, row?.size2, row?.unit
            ].map(v => String(v ?? '').trim()).join('::');
        },

        isBomRowSelected(row, index = -1) {
            return this.bomSelectedKeys.includes(this.getBomSelectionKey(row, index));
        },

        toggleBomRowSelection(row, checked, index = -1) {
            const key = this.getBomSelectionKey(row, index);
            const selected = new Set(this.bomSelectedKeys);
            if (checked) selected.add(key);
            else selected.delete(key);
            this.bomSelectedKeys = Array.from(selected);
            this.persistBomSelection();
        },

        selectAllBomDisplayed() {
            const selected = new Set(this.bomSelectedKeys);
            this.bomFilteredDetails.forEach((row, index) => {
                selected.add(this.getBomSelectionKey(row, index));
            });
            this.bomSelectedKeys = Array.from(selected);
            this.persistBomSelection();
        },

        clearBomSelection() {
            this.bomSelectedKeys = [];
            this.persistBomSelection();
        },

        get selectedBomCount() {
            return this.bomSelectedKeys.length;
        },

        get selectedBomDetails() {
            const selected = new Set(this.bomSelectedKeys);
            return this.bomDetailsAll.filter((row, index) => selected.has(this.getBomSelectionKey(row, index)));
        },

        persistBomSelection() {
            const project = this.allProjectsData?.[this.activeProject];
            if (!project?.meta?.bom) return;
            project.meta.bom.selectedKeys = [...this.bomSelectedKeys];
            this.saveProjectMetaOnly();
        },

        loadBomSelection() {
            const saved = this.allProjectsData?.[this.activeProject]?.meta?.bom?.selectedKeys;
            this.bomSelectedKeys = Array.isArray(saved) ? saved.map(String) : [];
        },

        resetBomFilters() {
            this.bomFilterCategory = 'ALL';
            this.bomFilterSearch = '';
        },

        openCalculateBOM() {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Calculate BOM hanya dapat dilakukan oleh Piping Engineer.');
                return;
            }

            this.loadBomSelection();
            const status = this.workflowStatus;

            // Jika sudah final, Engineer melihat laporan final.
            // Data final tidak dihitung ulang/ditimpa.
            if (status === 'APPROVED') {
                const finalApproval = this.latestApproval || {
                    status: 'APPROVED',
                    notes: 'Laporan final telah disetujui oleh Lead Estimator.'
                };
                this.printFinalReportPDF(finalApproval);
                return;
            }

            // Untuk tahap Estimator, BOM yang sudah dikirim tetap boleh dibuka,
            // tetapi snapshot lama harus dihitung ulang terlebih dahulu agar perubahan
            // Long Description (Size), Size 1, dan rumus Inch-Dia tidak tertinggal.
            // Status workflow dikembalikan ke status semula setelah recalculation.
            if (status === 'SUBMITTED_TO_ESTIMATOR') {
                this.calculateBOM(true);
                const project = this.allProjectsData?.[this.activeProject];
                if (project?.meta) {
                    project.meta.workflowStatus = status;
                    project.meta.workflowUpdatedAt = new Date().toLocaleString('id-ID');
                    this.saveProjectMetaOnly();
                }
                this.showBomModal = true;
                return;
            }

            if (status === 'SUBMITTED_TO_LEAD') {
                const existingBom = this.allProjectsData?.[this.activeProject]?.meta?.bom;
                if (existingBom?.details?.length) {
                    this.showBomModal = true;
                    return;
                }
                alert('BOM sudah dikirim ke Lead dan sedang menunggu proses approval.');
                return;
            }

            // Pada draft/revisi selalu hitung ulang dari sumber MTO terbaru.
            this.calculateBOM(true);
        },

        getNumeric(row, keys) {
            for (const key of keys) {
                const value = row?.[key];
                if (value === undefined || value === null || String(value).trim() === '') continue;

                // Excel dapat menyimpan ukuran sebagai pecahan seperti 3/4.
                // Jangan menghapus slash lalu mengubah 3/4 menjadi 34.
                // Excel sering menyimpan ukuran nominal seperti 3/4" atau 1 1/2".
                // Buang hanya tanda inci di ujung, JANGAN menghapus slash pecahan.
                const raw = String(value).trim().replace(/,/g, '').replace(/[\"″”']+$/g, '').trim();
                const fraction = raw.match(/^(\d+)\s*\/\s*(\d+)$/);
                if (fraction) {
                    const numerator = Number(fraction[1]);
                    const denominator = Number(fraction[2]);
                    if (denominator !== 0) return numerator / denominator;
                }

                const mixedFraction = raw.match(/^(\d+)\s+(\d+)\s*\/\s*(\d+)$/);
                if (mixedFraction) {
                    const whole = Number(mixedFraction[1]);
                    const numerator = Number(mixedFraction[2]);
                    const denominator = Number(mixedFraction[3]);
                    if (denominator !== 0) return whole + numerator / denominator;
                }

                const n = Number(raw.replace(/[^0-9.\-]/g, ''));
                if (Number.isFinite(n)) return n;
            }
            return 0;
        },

        normalizeFieldKey(value) {
            return String(value ?? '')
                .replace(/\u00A0/g, ' ')
                .trim()
                .replace(/\s+/g, ' ')
                .toLowerCase();
        },

        getText(row, keys, fallback = '-') {
            if (!row || typeof row !== 'object') return fallback;

            // 1) Prioritas exact key agar nilai Excel tidak berubah.
            for (const key of keys) {
                const value = row?.[key];
                if (value !== undefined && value !== null && String(value).trim() !== '') {
                    return String(value).trim();
                }
            }

            // 2) Fallback case/whitespace-insensitive untuk header Excel seperti
            // LONG DESCRIPTION (SIZE), Long Description (Size), atau header
            // yang memiliki spasi tersembunyi. Nilai cell tetap dipertahankan.
            const normalized = new Map();
            Object.keys(row).forEach(actualKey => {
                const nk = this.normalizeFieldKey(actualKey);
                if (!normalized.has(nk)) normalized.set(nk, actualKey);
            });
            for (const key of keys) {
                const actualKey = normalized.get(this.normalizeFieldKey(key));
                if (!actualKey) continue;
                const value = row[actualKey];
                if (value !== undefined && value !== null && String(value).trim() !== '') {
                    return String(value).trim();
                }
            }

            return fallback;
        },

        sanitizeLineNo(value) {
            const v = String(value ?? '').trim();
            // Jangan menampilkan placeholder mentah seperti ?->? sebagai Line No.
            if (!v || /^(\?\s*(?:[-–—]|->|→)\s*\?)+$/.test(v) || /^line\s*list[-_ ]?\d+$/i.test(v)) {
                return '-';
            }
            return v;
        },

        recalculateBOM() {
            // Tombol Hitung Ulang selalu memproses ulang sumber MTO terbaru.
            // Jangan bergantung pada selection/filtered rows dan jangan menutup modal.
            try {
                const before = Array.isArray(this.bomDetailsAll) ? this.bomDetailsAll.length : 0;
                this.calculateBOM(true);
                const project = this.allProjectsData?.[this.activeProject];
                const after = Array.isArray(project?.meta?.bom?.details) ? project.meta.bom.details.length : 0;
                // Pastikan modal tetap terbuka setelah perhitungan selesai.
                this.showBomModal = true;
                console.info('[BOM] Hitung Ulang selesai:', { before, after });
            } catch (error) {
                console.error('[BOM] Calculate BOM error:', error);
                alert('Perhitungan BOM gagal dijalankan. Silakan coba lagi.');
            }
        },

        getBomField(row, field, sheetName) {
            const sheet = String(sheetName || row?.__sheet || '').trim().toLowerCase();
            const aliases = {
                lineNo: ['Line No.', 'Line No', 'LINE NO.', 'Complete Line No.', 'Line Number Tag', 'LINE NUMBER TAG', 'Tag', 'TAG', 'LINE NUMBER'],
                component: ['Component', 'COMPONENT', 'Equipment Type', 'TYPE', 'Part Subtype'],
                description: ['Long Description (Size)', 'LONG DESCRIPTION (SIZE)', 'Long Description (Family)', 'LONG DESCRIPTION (FAMILY)', 'Short Description', 'SHORT DESCRIPTION', 'Description', 'DESCRIPTION', 'Fluid Service', 'Process Fluid Identifier'],
                qty: ['Qty', 'Quantity', 'QTY', 'Item Count', 'ITEM COUNT'],
                size1: ['Size 1', 'SIZE 1', 'Size', 'SIZE', 'Nominal Size', 'Line Size (Inch)', 'LINE SIZE (INCH)'],
                size2: ['Size 2', 'SIZE 2'],
                length: ['Length', 'LENGTH', 'Pipe Length', 'PIPE LENGTH', 'Length (m)', 'LENGTH (M)', 'Cut Length', 'CUT LENGTH', 'Fixed Length', 'FIXED LENGTH'],
                unit: ['Unit', 'UNIT', 'UOM', 'Satuan']
            };

            // Komponen BOM mengikuti nama kategori/sheet sumber, bukan digabung
            // dari isi kolom COMPONENT/TYPE. Dengan begitu Tee, Single Branch
            // Fitting, Pipe Run Component, Fasteners, dll tetap menjadi kategori
            // yang jelas dan tidak berubah menjadi teks seperti
            // "Tee / Single Branch Fitting / Piping and Equipment".
            const canonicalBySheet = {
                'basesupport': 'BaseSupport',
                'base support': 'BaseSupport',
                'support': 'Support',
                'pipe support': 'Support',
                'dummyleg': 'DummyLeg',
                'dummy leg': 'DummyLeg',
                'tee': 'Tee',
                'single branch fitting': 'Single Branch Fitting',
                'pipe': 'Pipe',
                'instrument': 'Instrument',
                'flange': 'Flange',
                'elbow': 'Elbow',
                'coupling': 'Coupling',
                'pipe run component': 'Pipe Run Component',
                'socketweld': 'Socketweld',
                'socket weld': 'Socketweld',
                'gasket': 'Gasket',
                'buttweld': 'Buttweld',
                'butt weld': 'Buttweld',
                'bolt set': 'Bolt Set',
                'fasteners': 'Fasteners'
            };
            const canonicalSheet = String(sheet || '').replace(/\u00A0/g, ' ').trim().toLowerCase();
            let canonicalComponent = canonicalBySheet[canonicalSheet];

            // Support bukan satu kategori tunggal. Di data sumber, BaseSupport
            // dan DummyLeg berada di sheet Support sebagai Part Subtype.
            // Pertahankan subtype tersebut agar kategori BOM tetap: BaseSupport,
            // Support, dan DummyLeg (bukan semuanya berubah menjadi Support).
            if (canonicalSheet === 'support' || canonicalSheet === 'pipe support') {
                const subtype = this.getText(row, ['Part Subtype', 'PART SUBTYPE', 'Support Type', 'SUPPORT TYPE'], '').trim().toLowerCase();
                if (subtype === 'basesupport' || subtype === 'base support') canonicalComponent = 'BaseSupport';
                else if (subtype === 'dummyleg' || subtype === 'dummy leg') canonicalComponent = 'DummyLeg';
                else canonicalComponent = 'Support';
            }

            // LineList is a process/line definition table. It has no BOM component
            // quantity or physical pipe length, so do not invent them.
            if (sheet === 'linelist' || sheet === 'line list') {
                if (field === 'lineNo') return this.sanitizeLineNo(this.getText(row, ['Complete Line No.', 'LINE NO.', 'Line No.', 'LINE NUMBER', 'Line Number'], ''));
                if (field === 'component') return 'Pipe';
                if (field === 'description') return this.getText(row, ['Fluid Service', 'Process Fluid Identifier', 'Pipe.Spec'], '-');
                if (field === 'qty') return 1;
                if (field === 'size1') return this.getNumeric(row, ['Line Size (Inch)', 'LINE SIZE (INCH)']);
                if (field === 'size2') return this.getNumeric(row, ['Size 2', 'SIZE 2']);
                if (field === 'length') return this.getNumeric(row, ['Length', 'LENGTH', 'Pipe Length', 'PIPE LENGTH', 'Cut Length', 'CUT LENGTH']);
                if (field === 'unit') return 'LINE';
            }

            // Valve tetap diproses khusus agar aturan RF Valve tetap berlaku.
            if (sheet === 'valve') {
                if (field === 'lineNo') return this.sanitizeLineNo(this.getText(row, ['Line Number Tag', 'LINE NUMBER TAG', 'Tag', 'TAG', 'Complete Line No.'], ''));
                if (field === 'component') return 'Valve';
                if (field === 'description') return this.getText(row, ['Long Description (Size)', 'LONG DESCRIPTION (SIZE)', 'Long Description (Family)', 'LONG DESCRIPTION (FAMILY)', 'Short Description', 'SHORT DESCRIPTION'], 'Valve');
                if (field === 'qty') return this.getNumeric(row, ['Qty', 'Quantity', 'Item Count', 'ITEM COUNT']) || 1;
                if (field === 'size1') return this.getNumeric(row, ['Size', 'SIZE', 'Nominal Diameter', 'Nominal Size']);
                if (field === 'size2') return this.getNumeric(row, ['Size 2', 'SIZE 2']);
                if (field === 'length') return this.getNumeric(row, ['Length', 'LENGTH', 'Engagement Length']);
                if (field === 'unit') return this.getText(row, ['Unit', 'UNIT', 'UOM'], 'EA');
            }

            // Komponen yang berasal dari sheet MTO tertentu sering tidak memiliki
            // kolom Qty/Size 1 dengan nama yang seragam. Ambil dari Size atau
            // Nominal Diameter agar Inch-Dia tetap terhitung.
            const dedicatedMtoSheets = new Set(['gasket', 'buttweld', 'butt weld', 'bolt set']);
            if (dedicatedMtoSheets.has(canonicalSheet)) {
                if (field === 'qty') {
                    return this.getNumeric(row, ['Qty', 'Quantity', 'QTY', 'Item Count', 'ITEM COUNT']) || 1;
                }
                if (field === 'size1') {
                    return this.getNumeric(row, [
                        'Size 1', 'SIZE 1', 'Size', 'SIZE',
                        'Nominal Diameter', 'Nominal Size', 'Line Size (Inch)', 'LINE SIZE (INCH)'
                    ]);
                }
                if (field === 'size2') {
                    return this.getNumeric(row, ['Size 2', 'SIZE 2']);
                }
                if (field === 'unit') {
                    return this.getText(row, ['Unit', 'UNIT', 'UOM', 'Satuan'], 'EA');
                }
            }

            if (field === 'component' && canonicalComponent) return canonicalComponent;

            const keys = aliases[field] || [];
            if (field === 'qty') return this.getNumeric(row, keys) || 1;
            if (field === 'size1' || field === 'size2' || field === 'length') return this.getNumeric(row, keys);
            if (field === 'unit') return this.getText(row, keys, 'EA');
            if (field === 'lineNo') return this.sanitizeLineNo(this.getText(row, keys, ''));
            return this.getText(row, keys, field === 'component' ? (canonicalComponent || row.__sheet || '-') : '-');
        },

        getAllMtoRows() {
            const project = this.allProjectsData?.[this.activeProject] || {};
            // Calculate BOM membaca seluruh sheet MTO yang memang berisi data.
            // LineList hanya menjadi MASTER referensi untuk mencocokkan Line Number,
            // sedangkan SP Items tidak ikut perhitungan. Support tetap dibaca karena
            // masuk ke rekap PIPING SUPPORT. Hanya sheet kosong yang dilewati agar
            // proses tetap ringan dan tidak menyebabkan lag.
            const referenceOnly = new Set(['LineList', 'SP Items']);
            const bomSourceSheets = new Set([
                'Valve', 'Tee', 'Single Branch Fitting', 'Pipe', 'Instrument', 'Flange',
                'Elbow', 'Coupling', 'Pipe Run Component', 'Socketweld', 'Gasket',
                'Buttweld', 'Bolt Set', 'Fasteners', 'Support'
            ]);
            const result = [];
            Object.keys(project).forEach(sheet => {
                // Piping and Equipment, Nozzle, Vessel, Tank, Pump, Equipment,
                // Misc Equipment, Tap Weld, dll tidak dimasukkan ke Detail BOM/BQ
                // versi ini karena bukan kategori yang sudah ditetapkan bersama.
                // Ini juga mencegah sheet agregat Piping and Equipment menggandakan
                // Tee/Pipe/Flange/Elbow dan membuat Component menjadi berantakan.
                if (sheet === 'meta' || referenceOnly.has(sheet) || !bomSourceSheets.has(sheet)) return;
                const rows = Array.isArray(project[sheet]) ? project[sheet] : [];
                if (!rows.length) return;
                rows.forEach((row, index) => {
                    if (!row || typeof row !== 'object') return;
                    const hasValue = Object.values(row).some(v => String(v ?? '').trim() !== '');
                    if (!hasValue) return;
                    result.push({ ...row, __sheet: sheet, __index: index });
                });
            });
            return result;
        },

        calculateBOM(forceRecalculate = false) {
            // forceRecalculate sengaja diterima agar tombol Calculate BOM/Hitung Ulang
            // selalu menghasilkan snapshot baru. Tidak ada early-return berdasarkan BOM lama.
            const rows = this.getAllMtoRows();
            if (!rows.length) {
                alert('Belum ada data MTO. Import atau input data MTO terlebih dahulu.');
                return;
            }

            // Sesuai arahan mentor: item MTO hanya dihitung apabila Line Number
            // ditemukan pada Master Line List. Jika tidak ada pasangan, item tidak masuk BOM.
            const project = this.allProjectsData?.[this.activeProject] || {};
            const lineRows = Array.isArray(project['LineList']) ? project['LineList'] : [];
            const normalizeLine = (v) => String(v ?? '').trim().toUpperCase().replace(/['\"\s]/g, '');
            const lineSet = new Set(lineRows.map(r => normalizeLine(this.getText(r, [
                'Complete Line No.', 'LINE NO.', 'Line No.', 'LINE NUMBER', 'Line Number', 'Line Number Tag', 'LINE NUMBER TAG'
            ], ''))).filter(Boolean));

            const materialOf = (row) => this.getText(row, ['Material', 'MATERIAL', 'Material Code', 'MATERIAL CODE'], '');
            const installationOf = (row) => this.getText(row, ['Installation', 'INSTALLATION', 'Location', 'LOCATION', 'Aboveground / Underground', 'ABOVEGROUND / UNDERGROUND'], '');

            const classify = (component, description, sheet) => {
                const t = `${component} ${description} ${sheet}`.toLowerCase();
                if (/support/.test(t)) return 'SUPPORT';
                if (/hydrostatic/.test(t)) return 'HYDROSTATIC';
                if (/high pressure|air testing|flushing/.test(t)) return 'AIR_TESTING';
                if (/radiographic|radiography/.test(t)) return 'RADIOGRAPHIC';
                if (/pipe\b/.test(t)) return 'PIPE';
                return 'COMPONENT';
            };

            // MASTER PEMETAAN RUMUS BOM / BQ.
            // Setiap jenis item memiliki aturan Inch-Dia sendiri. Jika source
            // tidak menyediakan rumus khusus, item yang mempunyai Size + Qty
            // tetap diberi nilai nominal Size x Qty agar tidak menghasilkan 0
            // hanya karena nama item belum ada di mapping.
            const extractSizesFromDescription = (text) => {
                const src = String(text || '');
                const matches = [];
                const re = /(\d+(?:\s+\d+\/\d+|\/\d+)?)\s*["″]/g;
                let m;
                while ((m = re.exec(src))) {
                    const raw = m[1].trim();
                    const parts = raw.split(/\s+/);
                    let n = 0;
                    if (parts.length === 2 && /\//.test(parts[1])) {
                        const [a,b] = parts[1].split('/').map(Number);
                        n = Number(parts[0]) + (b ? a / b : 0);
                    } else if (/^\d+\/\d+$/.test(raw)) {
                        const [a,b] = raw.split('/').map(Number);
                        n = b ? a / b : 0;
                    } else {
                        n = Number(raw);
                    }
                    if (Number.isFinite(n) && n > 0) matches.push(n);
                }
                return matches;
            };

            const formulaFor = (component, description, sheet, size1, size2, qty, length) => {
                const componentText = String(component || '').trim().toLowerCase();
                const descriptionText = String(description || '').trim().toLowerCase();
                const sheetText = String(sheet || '').trim().toLowerCase();
                const t = `${componentText} ${descriptionText} ${sheetText}`.replace(/[°]/g, '');
                const BF = Number(size1) || 0;
                const descSizes = extractSizesFromDescription(description);
                const BG = Number(size2) || (descSizes.length >= 2 ? descSizes[1] : 0);
                const R = Number(qty) || 0;
                const pipeLength = Number(length) || 0;
                let value = 0;
                let formula = 'Belum ada rumus khusus';

                // ----------------------------------------------------------
                // FITTING / COMPONENT BERDASARKAN MASTER RUMUS ENGINEER
                // ----------------------------------------------------------
                if (/\bcap\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2';
                }
                else if (/\breducing\s+coupling\b/.test(t)) {
                    value = (BF + BG) * R;
                    formula = '(BF2 × R2) + (BG2 × R2)';
                }
                else if (/\bcoupling\b/.test(t)) {
                    value = 2 * BF * R;
                    formula = '2 × BF2 × R2';
                }
                else if (/\b(concentric|eccentric)\s+reducer\b/.test(t)) {
                    value = (BF + BG) * R;
                    formula = '(BF2 × R2) + (BG2 × R2)';
                }
                else if (/\b(elbow|ell)\b/.test(t) && /\b90\b/.test(t)) {
                    value = 2 * BF * R;
                    formula = '2 × BF2 × R2';
                }
                else if (/\b(elbow|ell)\b/.test(t) && /\b45\b/.test(t)) {
                    value = 2 * BF * R;
                    formula = '2 × BF2 × R2';
                }
                else if (/\bflange\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2';
                }
                else if (/\b(strainer|ball\s+valve|check\s+valve|gate\s+valve|globe\s+valve)\b/.test(t)) {
                    value = 2 * BF * R;
                    formula = '(BF2 × R2) + (BF2 × R2)';
                }
                else if (/\bequal\s+tee\b/.test(t) || (/\btee\b/.test(t) && !/\b(reducing|barred)\b/.test(t))) {
                    value = 3 * BF * R;
                    formula = '3 × BF2 × R2';
                }
                else if (/\b(reducing|barred)\s+tee\b/.test(t)) {
                    value = (2 * BF + BG) * R;
                    formula = '(2 × BF2 × R2) + (BG2 × R2)';
                }
                else if (/\bsaddle\s+branch\b|\bsockolet\b/.test(t)) {
                    const branch = BG > 0 ? BG : BF;
                    value = (branch + 1.5 * branch) * R;
                    formula = '(BG2 × R2) + (1.5 × BG2 × R2)';
                }
                else if (/\bweldolet\b|\bthreadolet\b/.test(t)) {
                    const branch = BG > 0 ? BG : BF;
                    value = (branch + 1.5 * branch) * R;
                    formula = '(BG2 × R2) + (1.5 × BG2 × R2)';
                }
                else if (/\b90\s*reinforcing\s+pad\b|\b45\s*reinforcing\s+pad\b/.test(t)) {
                    const branch = BG > 0 ? BG : BF;
                    value = (branch + (8 + branch)) * R;
                    formula = '(BG2 × R2) + ((8+BG2) × R2)';
                }
                else if (/\b45\s*pipe\s+to\s+pipe\s+full\s+encirclement\b/.test(t)) {
                    value = (2 * BF + (2 * 1.5 * BF)) * R;
                    formula = '(2 × BF2 × R2) + (2 × 1.5 × BF2 × R2)';
                }
                // ----------------------------------------------------------
                // WELD / JOINT MARKERS DARI MTO
                // Buttweld, Tapweld dan Socketweld adalah marker joint pada
                // data MTO. Karena satu marker mewakili satu sambungan pada
                // diameter nominalnya, nilai dasarnya = Qty x Size.
                // ----------------------------------------------------------
                else if (componentText === 'buttweld' || /\bbuttweld\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2 (Buttweld joint)';
                }
                else if (/\btap\s*weld\b|\btapweld\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2 (Tapweld joint)';
                }
                else if (/\bsocket\s*weld\b|\bsocketweld\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2 (Socketweld joint)';
                }
                else if (componentText === 'gasket' || /\bgasket\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2 (Gasket)';
                }
                else if (componentText === 'bolt set' || /\bbolt\s*set\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2 (Bolt Set)';
                }
                // ----------------------------------------------------------
                // SUPPORT
                // Harus diuji sebelum PIPE karena "Pipe guide" mengandung
                // kata Pipe tetapi merupakan support.
                // ----------------------------------------------------------
                else if (/\bsupport\b|\bwelded\s+stanchion\b|\bpipe\s+guide\b|\bdummy\s+legs?\b|\bsaddled\s+slide\b|\bsaddled\s+anchor\b|\bwedge\s+support\b/.test(t)) {
                    value = BF * R;
                    formula = 'BF2 × R2';
                }
                // ----------------------------------------------------------
                // PIPE
                // Rumus master: ROUNDDOWN(Length/6,0) x Size x Qty.
                // Bila Length tidak tersedia / < 6 m, tetap gunakan Size x Qty
                // supaya kolom Inch-Dia tidak kosong/0 hanya karena source MTO
                // tidak membawa panjang pipe.
                // ----------------------------------------------------------
                else if (componentText === 'pipe' || /\bpipe\b/.test(descriptionText)) {
                    const pieces = Math.floor(pipeLength / 6);
                    if (pieces > 0) {
                        value = pieces * BF * R;
                        formula = 'ROUNDDOWN(R2/6,0) × BF2 × Qty';
                    } else if (BF > 0 && R > 0) {
                        value = BF * R;
                        formula = 'BF2 × R2 (fallback Length < 6 m/tidak tersedia)';
                    }
                }
                // ----------------------------------------------------------
                // ITEM LAIN
                // Gasket, Bolt Set, Fasteners dan item MTO lain yang memiliki
                // Size tetapi belum punya formula khusus tetap memperoleh nilai
                // nominal Size x Qty. Ini menjaga seluruh detail memiliki angka
                // yang dapat ditelusuri, tanpa mengubah Description sumber.
                // ----------------------------------------------------------
                else if (BF > 0 && R > 0) {
                    value = BF * R;
                    formula = 'BF2 × R2 (fallback item)';
                }

                return { value, formula };
            };

            const details = [];
            let excludedNoLine = 0;
            let excludedRFValve = 0;
            rows.forEach((row, i) => {
                const sheetName = row.__sheet;
                const qty = this.getBomField(row, 'qty', sheetName);
                const size1 = this.getBomField(row, 'size1', sheetName);
                const size2 = this.getBomField(row, 'size2', sheetName);
                const length = this.getBomField(row, 'length', sheetName);
                const unit = this.getBomField(row, 'unit', sheetName);
                const component = this.getBomField(row, 'component', sheetName);
                const description = this.getBomField(row, 'description', sheetName);
                const lineNo = this.sanitizeLineNo(this.getBomField(row, 'lineNo', sheetName));
                const lineKey = normalizeLine(lineNo);

                // Aturan mentor: item MTO hanya dihitung jika Line Number MTO
                // Jika Master Line List tersedia, item MTO hanya dihitung
                // apabila Line Number MTO ditemukan di Line List. Jika Line List
                // belum diisi/kosong, jangan menghapus seluruh BOM saat Hitung Ulang.
                const hasMasterLineList = lineRows.length > 0;
                if (sheetName !== 'LineList' && hasMasterLineList && (!lineKey || lineKey === '-' || !lineSet.has(lineKey))) {
                    excludedNoLine++;
                    return;
                }

                // REVISI SPV: khusus VALVE yang memiliki RF (Raised Face),
                // item tidak boleh masuk ke Detail BOM/BQ dan tidak ikut total.
                // RF pada FLANGE tetap dihitung; jadi pengecualian hanya Valve + RF.
                const componentText = String(component || '').trim().toUpperCase();
                const descriptionText = String(description || '').trim().toUpperCase();
                const isValve = /\bVALVE\b/.test(componentText) || /\bVALVE\b/.test(descriptionText);
                const hasRF = /\bRF\b/.test(descriptionText) || /\bRF\b/.test(componentText);
                if (isValve && hasRF) {
                    excludedRFValve++;
                    return;
                }

                const calc = formulaFor(component, description, sheetName, size1, size2, qty, length);
                details.push({
                    no: details.length + 1,
                    lineNo,
                    sheet: sheetName,
                    component,
                    description,
                    material: materialOf(row),
                    installation: installationOf(row),
                    qty, size1, size2, length, unit,
                    inchDia: calc.value,
                    mentorFormula: calc.formula,
                    category: classify(component, description, sheetName)
                });
            });

            // Gabungkan item yang identik agar Detail BOM/BQ tidak panjang
            // karena baris yang sama berulang. Qty, length, dan Inch-Dia
            // dijumlahkan sehingga nilai total tetap sama.
            const uniqueMap = new Map();
            details.forEach((row) => {
                const key = [
                    row.component,
                    row.description,
                    row.material,
                    row.installation,
                    Number(row.size1) || 0,
                    Number(row.size2) || 0,
                    row.unit,
                    row.category
                ].map(v => String(v ?? '').trim().toUpperCase()).join('|');

                if (!uniqueMap.has(key)) {
                    uniqueMap.set(key, {
                        ...row,
                        lineNo: row.lineNo || '-',
                        lineNos: row.lineNo ? [row.lineNo] : [],
                    });
                } else {
                    const existing = uniqueMap.get(key);
                    existing.qty = (Number(existing.qty) || 0) + (Number(row.qty) || 0);
                    existing.length = (Number(existing.length) || 0) + (Number(row.length) || 0);
                    existing.inchDia = (Number(existing.inchDia) || 0) + (Number(row.inchDia) || 0);
                    if (row.lineNo && !existing.lineNos.includes(row.lineNo)) {
                        existing.lineNos.push(row.lineNo);
                    }
                    if (existing.mentorFormula !== row.mentorFormula) {
                        existing.mentorFormula = 'Gabungan item identik';
                    }
                    existing.lineNo = existing.lineNos.length > 1 ? 'Multiple Lines' : existing.lineNo;
                }
            });

            const categoryOrder = [
                'BaseSupport', 'Support', 'DummyLeg', 'Tee', 'Single Branch Fitting',
                'Pipe', 'Instrument', 'Flange', 'Elbow', 'Coupling', 'Pipe Run Component',
                'Socketweld', 'Gasket', 'Buttweld', 'Bolt Set', 'Fasteners', 'Valve'
            ];
            const categoryRank = new Map(categoryOrder.map((name, index) => [name, index]));

            const compactDetails = Array.from(uniqueMap.values())
                .sort((a, b) => {
                    const ca = categoryRank.has(a.component) ? categoryRank.get(a.component) : 999;
                    const cb = categoryRank.has(b.component) ? categoryRank.get(b.component) : 999;
                    if (ca !== cb) return ca - cb;
                    const da = String(a.description || '').toUpperCase();
                    const db = String(b.description || '').toUpperCase();
                    if (da !== db) return da.localeCompare(db);
                    return (Number(a.size1) || 0) - (Number(b.size1) || 0);
                })
                .map((row, index) => ({
                    ...row,
                    no: index + 1
                }));

            details.length = 0;
            details.push(...compactDetails);

            const totalQty = details.reduce((sum, r) => sum + (Number(r.qty) || 0), 0);
            const pipeQty = details.filter(r => r.category === 'PIPE').reduce((sum, r) => sum + (Number(r.length) || 0), 0);
            const totalInchDia = details.reduce((sum, r) => sum + (Number(r.inchDia) || 0), 0);
            const totalJoint = details.filter(r => r.category === 'COMPONENT').reduce((sum, r) => sum + (Number(r.qty) || 0), 0);

            const bom = {
                revision: project.meta?.version || 0,
                sourceSheet: 'ALL_MTO',
                matchingRule: 'MTO_LINE_LIST_EXACT_MATCH_V2',
                calculatedAt: new Date().toLocaleString('id-ID'),
                totalMtoItems: rows.length,
                totalItems: details.length,
                totalQty,
                pipeQty,
                totalInchDia,
                totalJoint,
                excludedNoLine,
                details
            };
            // Pertahankan pilihan yang masih cocok dengan item hasil hitung ulang.
            // Perhitungan tetap menghasilkan seluruh BOM; pilihan hanya menentukan
            // item mana yang dikirim ke Estimator.
            const previousSelection = new Set(
                Array.isArray(this.bomSelectedKeys) ? this.bomSelectedKeys.map(String) : []
            );
            project.meta.bom = bom;
            this.bomSelectedKeys = bom.details
                .map((row, index) => this.getBomSelectionKey(row, index))
                .filter(key => previousSelection.has(key));
            project.meta.bom.selectedKeys = [...this.bomSelectedKeys];
            project.meta.workflowStatus = 'BOM_CALCULATED';
            project.meta.revisionNotes = '';
            // Jangan menunggu proses storage untuk menyelesaikan perhitungan UI.
            // Snapshot penyimpanan dijalankan setelah hasil BOM sudah tampil.
            this.showBomModal = true;
            try {
                this.saveProjectMetaOnly();
            } catch (storageError) {
                console.warn('[BOM] Penyimpanan snapshot dilewati:', storageError);
            }
        },

        async submitBOMToEstimator() {
            if (this.loginForm.role !== 'Piping Engineer') return alert('Hanya Piping Engineer yang dapat mengirim BOM.');
            const project = this.allProjectsData[this.activeProject];
            const bom = project?.meta?.bom;
            if (!bom?.details?.length) return alert('Hitung BOM terlebih dahulu.');
            if (bom.matchingRule !== 'MTO_LINE_LIST_EXACT_MATCH_V2') {
                return alert('BOM belum menggunakan aturan pencocokan MTO dengan Line List. Silakan Hitung Ulang BOM terlebih dahulu.');
            }

            const selected = this.selectedBomDetails;
            if (!selected.length) {
                return alert('Pilih minimal 1 item BOM yang akan dikirim ke Estimator.');
            }

            // BOM lengkap tetap disimpan. Hanya daftar item terpilih yang menjadi
            // input Estimator, sehingga item yang tidak dipilih tidak hilang.
            bom.selectedKeys = [...this.bomSelectedKeys];
            bom.selectedForEstimator = selected.map(row => ({ ...row }));
            bom.selectedTotalItems = selected.length;
            bom.selectedTotalQty = selected.reduce((sum, row) => sum + (Number(row.qty) || 0), 0);
            bom.selectedTotalInchDia = selected.reduce((sum, row) => sum + (Number(row.inchDia) || 0), 0);

            // Pastikan snapshot BOM benar-benar tersimpan sebelum berpindah role.
            await this.saveWorkflowSnapshot(this.activeProject, project);
            project.meta.revisionTarget = 'ESTIMATOR';
            this.setWorkflowStatus('SUBMITTED_TO_ESTIMATOR', `${selected.length} item BOM dipilih dan dikirim oleh Piping Engineer. Menunggu Estimator melakukan kalkulasi BOQ.`);
            this.showBomModal = false;
            alert(`${selected.length} item BOM / BQ berhasil dikirim ke Estimator.`);
        },

        async generateBOQ() {
            await this.ensureBomBoqReady();
            if (this.loginForm.role !== 'Estimator Proposal') {
                alert('Kalkulasi Biaya BOQ hanya dapat dilakukan oleh Estimator Proposal.');
                return;
            }

            const project = this.allProjectsData?.[this.activeProject];
            const meta = project?.meta;

            // BOQ final tidak dihitung ulang setelah Lead approval.
            if (this.workflowStatus === 'APPROVED') {
                const finalApproval = this.latestApproval || {
                    status: 'APPROVED',
                    notes: 'Laporan final telah disetujui oleh Lead Estimator.'
                };
                this.printFinalReportPDF(finalApproval);
                return;
            }

            if (!['SUBMITTED_TO_ESTIMATOR', 'BOQ_CALCULATED', 'ESTIMATOR_REVISION_REQUIRED'].includes(this.workflowStatus)) {
                alert(`BOQ belum dapat dihitung. Status saat ini: ${this.workflowStatusText}`);
                return;
            }

            // Pastikan BOM yang dipakai Estimator benar-benar mengikuti aturan
            // MTO Line Number ↔ Line List. BOM lama/stale dari versi sebelumnya
            // tidak boleh membuat tombol BOQ macet.
            let currentBom = meta?.bom;
            const bomNeedsRebuild = !currentBom?.details?.length ||
                currentBom.matchingRule !== 'MTO_LINE_LIST_EXACT_MATCH_V2';

            if (bomNeedsRebuild) {
                const keepStatus = this.workflowStatus;
                this.calculateBOM(true);
                currentBom = meta?.bom;

                // calculateBOM membuka modal dan mengubah status menjadi BOM_CALCULATED.
                // Kembalikan status agar Estimator tetap berada pada tahap yang benar.
                if (['SUBMITTED_TO_ESTIMATOR', 'ESTIMATOR_REVISION_REQUIRED'].includes(keepStatus) && currentBom?.details) {
                    this.setWorkflowStatus(keepStatus,
                        keepStatus === 'ESTIMATOR_REVISION_REQUIRED'
                            ? 'Revisi dari Lead diterima. Estimator dapat memperbaiki BOQ lalu mengirim ulang ke Lead.'
                            : 'BOM sudah dikirim oleh Piping Engineer. Menunggu Estimator melakukan kalkulasi BOQ.');
                    this.showBomModal = false;
                }
            }

            if (!currentBom?.details?.length) {
                alert('Tidak ada item BOM yang match antara MTO dan Line List. Tidak ada BOQ yang dapat dihitung.');
                return;
            }

            const previous = meta.boq?.items || [];
            const previousGroups = meta.boq?.priceGroups || [];
            const priceMap = new Map(previous.map(x => [x.key, Number(x.unitPrice) || 0]));
            const previousGroupPriceMap = new Map(previousGroups.map(g => [g.key, Number(g.unitPrice) || 0]));

            const selectedKeys = Array.isArray(currentBom.selectedKeys)
                ? new Set(currentBom.selectedKeys.map(String))
                : new Set();

            // Estimator menggunakan hanya item yang dipilih Engineer.
            // Untuk BOM lama yang belum memiliki selection, jangan mengirim
            // seluruh data secara diam-diam.
            if (!selectedKeys.size) {
                alert('Belum ada item BOM yang dipilih oleh Piping Engineer untuk dikirim ke Estimator.');
                return;
            }

            const selectedDetails = currentBom.details.filter((r, i) =>
                selectedKeys.has(this.getBomSelectionKey(r, i))
            );

            const items = selectedDetails.map((r, i) => ({
                ...r,
                key: `${r.sheet}::${r.no}`,
                unitPrice: priceMap.get(`${r.sheet}::${r.no}`) || 0,
                totalPrice: (priceMap.get(`${r.sheet}::${r.no}`) || 0) * (Number(r.qty) || 0)
            }));

            meta.boq = {
                revision: meta.version || 0,
                calculatedAt: new Date().toLocaleString('id-ID'),
                priceLevel: this.activePriceLevel,
                currency: this.boqCurrency,
                items,
                priceGroups: [],
                directCost: 0,
                indirectCost: Number(meta.boq?.indirectCost) || 0,
                totalCost: 0
            };

            // Estimator tidak perlu memasukkan harga 641 kali.
            // Kelompok harga dibuat dari karakteristik item yang menentukan harga:
            // Sheet + Component + Description + Size 1 + Size 2 + Unit.
            this.rebuildBOQPriceGroups(previousGroupPriceMap);
            this.recalculateBOQTotals();
            this.setWorkflowStatus('BOQ_CALCULATED', 'BOQ sudah dihitung oleh Estimator.');
            this.showBoqModal = true;
        },

        getBOQGroupKey(row) {
            const norm = (value) => String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
            // Harga BOQ dikelompokkan berdasarkan karakteristik harga yang nyata.
            // Ukuran TIDAK lagi diganti menjadi "Semua Ukuran".
            // Setiap kombinasi Component + Description + Size 1 + Size 2 + Unit
            // menjadi kelompok harga tersendiri agar harga ukuran berbeda tidak
            // tercampur (penting untuk material piping Tripatra).
            return [
                norm(row.component),
                norm(row.description),
                norm(row.size1),
                norm(row.size2),
                norm(row.unit)
            ].join('||');
        },
        formatBOQQty(value) {
            const n = Number(value);
            if (!Number.isFinite(n)) return '-';
            if (Number.isInteger(n)) return String(n);
            return String(Number(n.toFixed(2)));
        },

        formatBOQSize(size1, size2) {
            const clean = (value) => {
                if (value === null || value === undefined || String(value).trim() === '') return '';
                const n = Number(value);
                // Size 2 = 0 pada data MTO berarti tidak ada ukuran kedua.
                if (Number.isFinite(n) && n === 0) return '';
                return String(value).trim();
            };
            const s1 = clean(size1);
            const s2 = clean(size2);
            if (s1 && s2) return `${s1} × ${s2}`;
            if (s1) return s1;
            if (s2) return s2;
            return '-';
        },

        getPriceMasterStorageKey() {
            return 'tripatra_boq_price_master_v1';
        },

        normalizePriceMasterText(value) {
            return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
        },

        buildPriceMasterKeyFromFields(component, description, size1, size2, unit, priceLevel = this.activePriceLevel, currency = this.boqCurrency) {
            const norm = (v) => this.normalizePriceMasterText(v);
            const groupKey = [norm(component), norm(description), norm(size1), norm(size2), norm(unit)].join('||');
            return `${priceLevel || 'Menengah'}::${currency || 'IDR'}::${groupKey}`;
        },

        exportBOQPriceMaster() {
            const master = this.loadBOQPriceMaster();
            const rows = [['Price Level', 'Currency', 'Component', 'Description', 'Size 1', 'Size 2', 'Unit', 'Unit Price']];
            Object.entries(master).forEach(([key, value]) => {
                const parts = key.split('::');
                if (parts.length < 5) return;
                const priceLevel = parts[0];
                const currency = parts[1];
                const group = parts.slice(2).join('::').split('||');
                // Mendukung format baru: component | description | size1 | size2 | unit.
                // Format lama tetap diekspor tanpa merusak data yang sudah tersimpan.
                const isNewFormat = group.length >= 5;
                rows.push([
                    priceLevel,
                    currency,
                    group[0] || '',
                    group[1] || '',
                    isNewFormat ? (group[2] || '') : '',
                    isNewFormat ? (group[3] || '') : '',
                    isNewFormat ? (group[4] || '') : (group[2] || ''),
                    Number(value) || 0
                ]);
            });
            const csv = rows.map(row => row.map(v => {
                const s = String(v ?? '');
                return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
            }).join(',')).join('\n');
            const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'BOQ_Price_Master.csv';
            a.click();
            URL.revokeObjectURL(url);
        },

        importBOQPriceMasterFile(event) {
            const file = event?.target?.files?.[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const text = String(reader.result || '').replace(/^\uFEFF/, '');
                    const lines = text.split(/\r?\n/).filter(Boolean);
                    if (!lines.length) throw new Error('File kosong.');

                    const parseCsvLine = (line) => {
                        const out = [];
                        let cur = '', quoted = false;
                        for (let i = 0; i < line.length; i++) {
                            const ch = line[i];
                            if (ch === '"') {
                                if (quoted && line[i + 1] === '"') { cur += '"'; i++; }
                                else quoted = !quoted;
                            } else if (ch === ',' && !quoted) {
                                out.push(cur); cur = '';
                            } else cur += ch;
                        }
                        out.push(cur);
                        return out;
                    };

                    const headers = parseCsvLine(lines[0]).map(h => this.normalizePriceMasterText(h));
                    const find = (...names) => names.map(n => headers.indexOf(this.normalizePriceMasterText(n))).find(i => i >= 0);
                    const iLevel = find('Price Level', 'Level Harga');
                    const iCurrency = find('Currency', 'Mata Uang');
                    const iComponent = find('Component', 'Komponen');
                    const iDescription = find('Description', 'Deskripsi');
                    const iSize1 = find('Size 1', 'Size1', 'Ukuran 1');
                    const iSize2 = find('Size 2', 'Size2', 'Ukuran 2');
                    const iUnit = find('Unit', 'Satuan', 'UOM');
                    const iPrice = find('Unit Price', 'Harga Satuan', 'Harga');

                    if ([iComponent, iDescription, iUnit, iPrice].some(i => i === undefined)) {
                        throw new Error('Kolom wajib: Component, Description, Unit, Unit Price.');
                    }

                    const master = this.loadBOQPriceMaster();
                    let count = 0;
                    lines.slice(1).forEach(line => {
                        const cells = parseCsvLine(line);
                        const price = Number(String(cells[iPrice] ?? '').replace(/[^0-9.-]/g, ''));
                        if (!Number.isFinite(price) || price <= 0) return;
                        const level = iLevel !== undefined && cells[iLevel] ? cells[iLevel] : this.activePriceLevel;
                        const currency = iCurrency !== undefined && cells[iCurrency] ? cells[iCurrency] : this.boqCurrency;
                        const key = this.buildPriceMasterKeyFromFields(
                            cells[iComponent],
                            cells[iDescription],
                            iSize1 !== undefined ? cells[iSize1] : '',
                            iSize2 !== undefined ? cells[iSize2] : '',
                            cells[iUnit],
                            level,
                            currency
                        );
                        master[key] = price;
                        count++;
                    });
                    this.saveBOQPriceMaster(master);
                    this.applyBOQPriceMaster();
                    alert(`${count} harga berhasil dimasukkan ke Price Master. Harga akan otomatis digunakan pada kelompok BOM yang sesuai.`);
                    if (event.target) event.target.value = '';
                } catch (error) {
                    console.error(error);
                    alert(`Import Price Master gagal: ${error.message}`);
                }
            };
            reader.readAsText(file);
        },

        createPriceMasterTemplate() {
            const rows = [
                ['Price Level', 'Currency', 'Component', 'Description', 'Size 1', 'Size 2', 'Unit', 'Unit Price'],
                ['Menengah', 'IDR', 'Valve', 'Gate Valve, Solid Wedge, 150 LB', '6', '-', 'EA', ''],
                ['Menengah', 'IDR', 'Elbow', 'ELL 90 LR', '8', '-', 'EA', ''],
                ['Menengah', 'IDR', 'Flange', 'FLANGE WN', '8', '-', 'EA', '']
            ];
            const csv = rows.map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
            const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'BOQ_Price_Master_Template.csv';
            a.click();
            URL.revokeObjectURL(url);
        },

        loadBOQPriceMaster() {
            try {
                const raw = localStorage.getItem(this.getPriceMasterStorageKey());
                const data = raw ? JSON.parse(raw) : {};
                return data && typeof data === 'object' ? data : {};
            } catch (error) {
                console.warn('Price Master tidak dapat dibaca:', error);
                return {};
            }
        },

        saveBOQPriceMaster(master) {
            try {
                localStorage.setItem(this.getPriceMasterStorageKey(), JSON.stringify(master || {}));
            } catch (error) {
                console.warn('Price Master tidak dapat disimpan:', error);
            }
        },

        getBOQPriceMasterKey(groupKey) {
            return `${this.activePriceLevel || 'Menengah'}::${this.boqCurrency || 'IDR'}::${groupKey}`;
        },

        getBOQFamilyKey(row) {
            const norm = (value) => String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
            return [
                norm(row.component),
                norm(row.unit)
            ].join('||');
        },

        getMasterFamilyPrice(row) {
            const master = this.loadBOQPriceMaster();
            const familyKey = this.getBOQFamilyKey(row);
            const matches = Object.entries(master).filter(([key, value]) =>
                key.includes(`::${familyKey}::`) && Number(value) > 0
            );
            const unique = [...new Set(matches.map(([, value]) => Number(value)))];
            return unique.length === 1 ? unique[0] : 0;
        },

        getMasterPrice(groupKey) {
            const master = this.loadBOQPriceMaster();
            const value = Number(master[this.getBOQPriceMasterKey(groupKey)]);
            return Number.isFinite(value) && value > 0 ? value : 0;
        },

        saveGroupPriceToMaster(group) {
            if (!group || Number(group.unitPrice) <= 0) return;
            const master = this.loadBOQPriceMaster();
            master[this.getBOQPriceMasterKey(group.key)] = Number(group.unitPrice);
            this.saveBOQPriceMaster(master);
        },

        applyBOQPriceMaster() {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq?.priceGroups?.length) return 0;

            const master = this.loadBOQPriceMaster();
            let applied = 0;

            meta.boq.priceGroups.forEach(group => {
                const price = Number(master[this.getBOQPriceMasterKey(group.key)]);
                if (Number.isFinite(price) && price > 0 && Number(group.unitPrice || 0) <= 0) {
                    group.unitPrice = price;
                    meta.boq.items.forEach(row => {
                        if (this.getBOQGroupKey(row) === group.key) {
                            row.unitPrice = price;
                            row.totalPrice = price * (Number(row.qty) || 0);
                        }
                    });
                    applied++;
                }
            });

            this.recalculateBOQTotals();
            this.saveProjectData?.();
            if (typeof this.render === 'function') this.render();
            alert(`${applied} kelompok harga berhasil diambil dari Price Master.`);
            return applied;
        },


        rebuildBOQPriceGroups(previousPriceMap = new Map()) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq?.items) return;

            const groups = new Map();
            meta.boq.items.forEach((row) => {
                const key = this.getBOQGroupKey(row);
                if (!groups.has(key)) {
                    groups.set(key, {
                        key,
                        sheet: row.sheet || '-',
                        component: row.component || '-',
                        description: row.description || '-',
                        size1: row.size1 ?? '-',
                        size2: row.size2 ?? '-',
                        unit: row.unit || '-',
                        qty: 0,
                        itemCount: 0,
                        unitPrice: Number(previousPriceMap.get(key) ?? 0) ||
                            this.getMasterPrice(key) ||
                            this.getMasterFamilyPrice(row) || 0
                    });
                }
                const group = groups.get(key);
                group.qty += Number(row.qty) || 0;
                group.itemCount += 1;
            });

            meta.boq.priceGroups = Array.from(groups.values()).sort((a, b) =>
                `${a.sheet}${a.component}${a.description}${a.size1}`.localeCompare(
                    `${b.sheet}${b.component}${b.description}${b.size1}`
                )
            );

            // Terapkan harga group ke seluruh detail BOM.
            const groupPriceMap = new Map(meta.boq.priceGroups.map(g => [g.key, Number(g.unitPrice) || 0]));
            meta.boq.items.forEach(row => {
                const price = groupPriceMap.get(this.getBOQGroupKey(row)) || 0;
                row.unitPrice = price;
                row.totalPrice = price * (Number(row.qty) || 0);
            });
        },

        updateBOQPriceGroup(index, value) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            const group = meta?.boq?.priceGroups?.[index];
            if (!group) return;

            const n = Number(String(value).replace(/,/g, '')) || 0;
            group.unitPrice = n;
            if (n > 0) this.saveGroupPriceToMaster(group);

            const groupKey = group.key;
            meta.boq.items.forEach(row => {
                if (this.getBOQGroupKey(row) === groupKey) {
                    row.unitPrice = n;
                    row.totalPrice = n * (Number(row.qty) || 0);
                }
            });

            this.recalculateBOQTotals();
        },

        recalculateBOQTotals() {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq?.items) return;

            meta.boq.directCost = meta.boq.items.reduce(
                (sum, row) => sum + (Number(row.totalPrice) || 0), 0
            );
            meta.boq.totalCost =
                meta.boq.directCost + (Number(meta.boq.indirectCost) || 0);

            this.saveStorage();
        },

        getBOQGroupPriceCoverage() {
            const groups = this.allProjectsData?.[this.activeProject]?.meta?.boq?.priceGroups || [];
            const priced = groups.filter(g => Number(g.unitPrice) > 0).length;
            return {
                total: groups.length,
                priced,
                unpriced: groups.length - priced
            };
        },

        // Alias yang dipakai tampilan ringkasan BOQ.
        // Tetap gunakan sumber data priceGroups yang sama agar angka
        // "Harga Sudah Diisi" dan "Belum Diisi" selalu sinkron dengan tabel.
        getBOQPriceCoverage() {
            return this.getBOQGroupPriceCoverage();
        },

        updateBOQPrice(index, value) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq?.items?.[index]) return;
            const n = Number(String(value).replace(/,/g, '')) || 0;
            meta.boq.items[index].unitPrice = n;
            meta.boq.items[index].totalPrice = n * (Number(meta.boq.items[index].qty) || 0);
            meta.boq.directCost = meta.boq.items.reduce((s, x) => s + (Number(x.totalPrice) || 0), 0);
            meta.boq.totalCost = meta.boq.directCost + (Number(meta.boq.indirectCost) || 0);
            this.saveStorage();
        },

        setIndirectCost(value) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq) return;
            meta.boq.indirectCost = Number(String(value).replace(/,/g, '')) || 0;
            meta.boq.totalCost = (Number(meta.boq.directCost) || 0) + meta.boq.indirectCost;
            this.saveStorage();
        },

        async sendBOQBackToEngineer() {
            if (this.loginForm.role !== 'Estimator Proposal') return alert('Hanya Estimator Proposal yang dapat mengirim kembali ke Engineer.');
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.bom?.details?.length) return alert('Data BOM tidak tersedia. Engineer perlu menghitung BOM terlebih dahulu.');

            // Dipakai bila catatan Lead membutuhkan perubahan teknis/BOM, bukan sekadar harga BOQ.
            meta.revisionTarget = 'ENGINEER';
            await this.saveWorkflowSnapshot(this.activeProject, this.allProjectsData[this.activeProject]);
            this.setWorkflowStatus(
                'ENGINEER_REVISION_REQUIRED',
                meta.revisionNotes || 'Estimator meminta Engineer memperbaiki / menghitung ulang BOM berdasarkan catatan revisi Lead.'
            );
            this.showBoqModal = false;
            alert('Project dikirim kembali ke Piping Engineer untuk revisi dan hitung ulang BOM.');
        },

        async submitBOQToLead() {
            if (this.loginForm.role !== 'Estimator Proposal') return alert('Hanya Estimator Proposal yang dapat mengirim BOQ.');
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            if (!meta?.boq?.items?.length) return alert('Hitung BOQ terlebih dahulu.');

            const coverage = this.getBOQGroupPriceCoverage();
            if (coverage.unpriced > 0) {
                return alert(`Masih ada ${coverage.unpriced} kelompok harga yang belum diisi. Lengkapi Unit Price terlebih dahulu sebelum mengirim BOQ ke Lead.`);
            }

            // Pastikan harga terakhir yang sedang tampil sudah dihitung dan
            // disimpan sebelum BOQ berpindah ke Lead.
            this.recalculateBOQTotals();

            // Snapshot ini adalah sumber kebenaran Lead:
            // item, group price, direct cost, indirect cost, currency,
            // dan total cost semuanya berasal dari hasil Estimator terakhir.
            await this.saveLeadBOQSnapshot(
                this.activeProject,
                this.allProjectsData[this.activeProject]
            );

            meta.revisionTarget = 'LEAD';
            this.setWorkflowStatus(
                'SUBMITTED_TO_LEAD',
                'BOQ sudah dikirim oleh Estimator. Menunggu Lead Estimator melakukan review.'
            );
            this.showBoqModal = false;
            alert('BOQ berhasil dikirim ke Lead Estimator untuk review.');
        },

        async approveData() {
            if (this.loginForm.role !== 'Lead Estimator') {
                alert('Approval & Review Laporan hanya dapat dilakukan oleh Lead Estimator.');
                return;
            }

            // Lead HARUS membaca snapshot handoff terakhir dari Estimator.
            // Jangan menggunakan BOQ lokal/seed yang mungkin hanya berisi 12 item.
            // Ini juga menjaga Total BOQ tetap sama persis dengan hasil Estimator.
            await this.ensureBomBoqReady();
            if (this.workflowStatus === 'SUBMITTED_TO_LEAD') {
                const submitted = await this.loadLeadBOQSnapshot(this.activeProject);
                const submittedBoq = submitted?.boq;
                const submittedItems = Array.isArray(submittedBoq?.items) ? submittedBoq.items : [];
                const submittedGroups = Array.isArray(submittedBoq?.priceGroups) ? submittedBoq.priceGroups : [];
                if (submittedItems.length > 0 || submittedGroups.length > 0) {
                    const meta = this.allProjectsData?.[this.activeProject]?.meta;
                    if (meta) {
                        meta.boq = JSON.parse(JSON.stringify(submittedBoq));
                        meta.boq.updatedAtEpoch = Number(submitted?.submittedAt || meta.boq.updatedAtEpoch || Date.now());
                        // Simpan total handoff secara eksplisit agar UI Lead tidak
                        // menghitung ulang dari subset/stale data.
                        meta.boq.totalCost = Number(submittedBoq.handoffTotalCost ?? submittedBoq.totalCost ?? 0);
                        meta.boq.directCost = Number(submittedBoq.handoffDirectCost ?? submittedBoq.directCost ?? 0);
                        meta.boq.indirectCost = Number(submittedBoq.handoffIndirectCost ?? submittedBoq.indirectCost ?? 0);
                        this.saveProjectMetaOnly();
                        this.saveBomBoqStorage();
                    }
                } else {
                    alert('Data BOQ dari Estimator tidak ditemukan. Jangan lanjut review agar total BOQ tidak salah.');
                    return;
                }
            }

            // Setelah approval, tombol review membuka laporan final.
            if (this.workflowStatus === 'APPROVED') {
                const finalApproval = this.latestApproval || {
                    status: 'APPROVED',
                    notes: 'Laporan final telah disetujui oleh Lead Estimator.'
                };
                this.printFinalReportPDF(finalApproval);
                return;
            }

            if (this.workflowStatus !== 'SUBMITTED_TO_LEAD') {
                alert(`Belum ada BOQ yang menunggu review. Status saat ini: ${this.workflowStatusText}`);
                return;
            }

            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            const reviewBoq = meta?.boq;
            if (!reviewBoq?.items?.length) {
                alert('BOQ belum tersedia. Estimator harus menghitung dan mengirim BOQ terlebih dahulu.');
                return;
            }
            const handoffCount = Number(reviewBoq.handoffItemCount || 0);
            if (handoffCount && reviewBoq.items.length !== handoffCount) {
                alert(`Data BOQ Lead tidak sinkron (${reviewBoq.items.length} item terbaca, seharusnya ${handoffCount}). Handoff dari Estimator akan dimuat ulang.`);
                const submitted = await this.loadLeadBOQSnapshot(this.activeProject);
                if (submitted?.boq) {
                    meta.boq = JSON.parse(JSON.stringify(submitted.boq));
                    this.saveProjectMetaOnly();
                    this.saveBomBoqStorage();
                }
            }
            if (meta?.bom?.matchingRule !== 'MTO_LINE_LIST_EXACT_MATCH_V2') {
                alert('BOQ belum menggunakan pencocokan Line Number MTO dengan Line List. Engineer harus menghitung ulang BOM terlebih dahulu.');
                return;
            }

            this.approvalNote = '';
            this.showApproveModal = true;
        },

        getReviewNote() {
            const el = document.getElementById('reviewer-notes-textarea');
            return String(el?.value || this.approvalNote || '').trim();
        },

        recordApproval(status, notes = '') {
            const meta = this.allProjectsData[this.activeProject].meta;
            const item = {
                id: Date.now(),
                projectId: this.activeProject,
                sheet: this.activeSheet,
                revision: meta.version || 0,
                status,
                reviewer: this.loginForm.role || 'Lead Estimator',
                notes,
                timestamp: new Date().toLocaleString('id-ID'),
                title: status === 'APPROVED' ? 'BOQ / Laporan Disetujui' : 'Revisi BOQ / Laporan'
            };
            const historyKey = 'tripatra_approval_history_v2';
            let history = [];
            try { history = JSON.parse(localStorage.getItem(historyKey) || '[]'); } catch (_) {}
            history.push(item);
            localStorage.setItem(historyKey, JSON.stringify(history));
            this.approvalHistory = history;
            return item;
        },

        confirmApprove() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            const meta = this.allProjectsData[this.activeProject].meta;
            const notes = this.getReviewNote() || 'Laporan disetujui oleh Lead Estimator.';
            meta.isApproved = true;
            meta.approvedAt = new Date().toLocaleString('id-ID');
            meta.version = Number(meta.version || 0);
            this.setWorkflowStatus('APPROVED', notes);
            this.recordApproval('APPROVED', notes);
            this.showApproveModal = false;
            alert('Laporan berhasil APPROVED. Workflow selesai dan menjadi laporan final.');
        },

        handleMintaRevisi() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            const meta = this.allProjectsData[this.activeProject].meta;
            const notes = this.getReviewNote();
            if (!notes) return alert('Catatan revisi wajib diisi agar Estimator mengetahui bagian yang harus diperbaiki.');
            meta.isApproved = false;
            meta.version = Number(meta.version || 0) + 1;
            this.setWorkflowStatus('ESTIMATOR_REVISION_REQUIRED', notes);
            meta.revisionTarget = 'ESTIMATOR';
            this.recordApproval('REVISION_REQUIRED', notes);
            this.showApproveModal = false;
            alert(`Revisi diminta. Project kembali ke Estimator sebagai Rev ${meta.version}.`);
        },

        startNewRevision() {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Revisi baru hanya dapat dimulai oleh Piping Engineer.');
                return;
            }

            const project = this.allProjectsData?.[this.activeProject];
            if (!project) return;

            const meta = project.meta || {};
            const nextRevision = Number(meta.version || 0) + 1;

            if (!confirm(`Mulai Rev ${nextRevision} dari laporan final? Data MTO tetap dipertahankan, tetapi BOM/BOQ lama akan menjadi hasil revisi sebelumnya.`)) {
                return;
            }

            meta.version = nextRevision;
            meta.isApproved = false;
            meta.approvedAt = null;
            meta.bom = null;
            meta.boq = null;
            meta.workflowStatus = 'DRAFT';
            meta.revisionNotes = 'Revisi baru dimulai oleh Piping Engineer.';
            meta.workflowUpdatedAt = new Date().toLocaleString('id-ID');

            this.saveProjectMetaOnly();
            this.showBomModal = false;
            this.showBoqModal = false;
            this.showApproveModal = false;

            alert(`Rev ${nextRevision} siap dikerjakan. Alur dimulai kembali dari Engineer.`);
        },

        rejectData() {
            if (this.loginForm.role !== 'Lead Estimator') return;
            const notes = this.getReviewNote() || 'Laporan ditolak dan perlu diperbaiki oleh Estimator.';
            const meta = this.allProjectsData[this.activeProject].meta;
            meta.isApproved = false;
            meta.version = Number(meta.version || 0) + 1;
            this.setWorkflowStatus('ESTIMATOR_REVISION_REQUIRED', notes);
            meta.revisionTarget = 'ESTIMATOR';
            this.recordApproval('REJECTED', notes);
            this.showApproveModal = false;
            alert(`Laporan ditolak. Project dikembalikan ke Estimator sebagai Rev ${meta.version}.`);
        },

        loadApprovalHistory() {
            try { this.approvalHistory = JSON.parse(localStorage.getItem('tripatra_approval_history_v2') || '[]'); }
            catch (_) { this.approvalHistory = []; }
        },

        get projectApprovalHistory() {
            return (this.approvalHistory || []).filter(x => x.projectId === this.activeProject);
        },

        get latestApproval() {
            const arr = this.projectApprovalHistory;
            return arr.length ? arr[arr.length - 1] : null;
        },

        get activeTaskCount() {
            const s = this.workflowStatus;
            if (this.loginForm.role === 'Piping Engineer') return ['DRAFT', 'REVISION_REQUIRED', 'ENGINEER_REVISION_REQUIRED'].includes(s) ? 1 : 0;
            if (this.loginForm.role === 'Estimator Proposal') return ['SUBMITTED_TO_ESTIMATOR', 'ESTIMATOR_REVISION_REQUIRED'].includes(s) ? 1 : 0;
            if (this.loginForm.role === 'Lead Estimator') return s === 'SUBMITTED_TO_LEAD' ? 1 : 0;
            return 0;
        },

        getApprovalStatusText(status) {
            const map = { APPROVED: 'Approved', REVISION_REQUIRED: 'Minta Revisi', REJECTED: 'Rejected' };
            return map[status] || status || 'Menunggu';
        },

        getApprovalStatusClass(status) {
            if (status === 'APPROVED') return 'bg-emerald-100 text-emerald-700';
            if (status === 'REVISION_REQUIRED') return 'bg-amber-100 text-amber-700';
            return 'bg-rose-100 text-rose-700';
        },

        openActiveTasks() { this.taskView = 'active'; },
        openApprovalHistory() { this.taskView = 'history'; },
        openTaskMTOValve() { this.currentDashboardTab = 'workspace'; this.activeSheet = 'Valve'; this.currentPage = 1; },
        openApprovalDetail(item) { this.printFinalReportPDF(item); },

        printFinalReportPDF(approvalItem = null) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta || {};
            const bom = meta.bom || { details: [] };
            const boq = meta.boq || { items: [], directCost: 0, indirectCost: 0, totalCost: 0 };
            const esc = (v) => String(v ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
            const num = (v) => Number(v || 0).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
            const money = (v) => {
                const n = Number(v) || 0;
                if (boq.currency === 'IDR') {
                    return new Intl.NumberFormat('id-ID', { style:'currency', currency:'IDR', minimumFractionDigits:0, maximumFractionDigits:0 }).format(n);
                }
                return new Intl.NumberFormat('en-US', { style:'currency', currency:'USD', minimumFractionDigits:2, maximumFractionDigits:2 }).format(n);
            };

            const rows = Array.isArray(bom.details) ? bom.details : [];
            const materialGroups = ['CS PIPE','SS PIPE','LTCS PIPE','LOW ALLOY PIPE','HIGH ALLOY PIPE','HDPE PIPE','RTRP PIPE'];
            const materialValue = (r) => `${r.material || ''} ${r.description || ''}`.toUpperCase();
            const pipeRows = rows.filter(r => r.category === 'PIPE' || /\bPIPE\b/i.test(`${r.component} ${r.description}`));
            const pipeAmount = (label) => pipeRows.filter(r => {
                const t = materialValue(r);
                if (label === 'CS PIPE') return /\bCS\b|CARBON STEEL/.test(t) && !/LTCS|SS|STAINLESS/.test(t);
                if (label === 'SS PIPE') return /\bSS\b|STAINLESS/.test(t);
                if (label === 'LTCS PIPE') return /LTCS/.test(t);
                if (label === 'LOW ALLOY PIPE') return /LOW ALLOY/.test(t);
                if (label === 'HIGH ALLOY PIPE') return /HIGH ALLOY/.test(t);
                if (label === 'HDPE PIPE') return /HDPE/.test(t);
                if (label === 'RTRP PIPE') return /RTRP/.test(t);
                return false;
            }).reduce((s,r) => s + (Number(r.inchDia)||0), 0);

            const supportQty = rows.filter(r => r.category === 'SUPPORT' || /support/i.test(`${r.sheet} ${r.component} ${r.description}`)).reduce((s,r)=>s+(Number(r.qty)||0),0);
            const testRows = rows.filter(r => ['HYDROSTATIC','AIR_TESTING','RADIOGRAPHIC'].includes(r.category));
            const testAmount = cat => testRows.filter(r=>r.category===cat).reduce((s,r)=>s+(Number(r.inchDia)||Number(r.qty)||0),0);
            const itemRow = (no, desc, qty, unit='', cls='') => `<tr class="${cls}"><td class="no">${no}</td><td>${esc(desc)}</td><td class="qty">${typeof qty === 'number' ? num(qty) : esc(qty)}</td><td>${esc(unit)}</td><td></td></tr>`;
            const pipeSection = (title, no) => {
                let h = itemRow(no, title, '', '', 'section');
                materialGroups.forEach(m => { h += itemRow('', '- ' + m, pipeAmount(m), 'DIA.INCH'); });
                return h;
            };
            const approval = approvalItem || this.latestApproval || {};
            const approved = approval.status === 'APPROVED' || meta.isApproved;
            const status = approved ? 'APPROVED' : this.getApprovalStatusText(approval.status || meta.workflowStatus);

            const reportRows = pipeSection('1  INSTALLATION PIPE', '1') +
                itemRow('2', 'PIPING SUPPORT', supportQty, 'TON') +
                itemRow('3', 'TESTING', '', '', 'section') +
                itemRow('', 'Hydrostatic testing and cleaning', testAmount('HYDROSTATIC'), 'LM', testAmount('HYDROSTATIC') ? '' : 'zero') +
                itemRow('', 'High Pressure air testing (flushing) and cleaning', testAmount('AIR_TESTING'), 'LM', testAmount('AIR_TESTING') ? '' : 'zero') +
                itemRow('', 'Radiographic Testing', testAmount('RADIOGRAPHIC'), 'DIA.INCH', testAmount('RADIOGRAPHIC') ? '' : 'zero');

            const html = `<!doctype html><html><head><meta charset="utf-8"><title>Final Total BOQ KPI Format - ${esc(this.activeProject)}</title>
            <style>
            *{box-sizing:border-box}body{font-family:Arial,Helvetica,sans-serif;color:#111;margin:0;padding:14mm 12mm;font-size:10px}h1{font-size:15px;text-align:center;margin:0 0 2px;font-weight:700}h2{font-size:12px;text-align:center;margin:0 0 16px;font-weight:700}.meta{display:grid;grid-template-columns:130px 1fr 130px 1fr;margin-bottom:10px}.meta div{border:1px solid #333;padding:5px}.meta .label{font-weight:700;background:#f1f1f1}.status{text-align:center;font-weight:700;margin:7px 0 12px}.main{width:100%;border-collapse:collapse}.main th,.main td{border:1px solid #222;padding:5px 6px;vertical-align:middle}.main th{background:#bdbdbd;text-align:center;font-weight:700}.main .no{width:45px;text-align:center}.main .qty{width:110px;text-align:right;font-variant-numeric:tabular-nums}.main .section td{font-weight:700}.main .zero td{color:#333}.summary{margin-top:16px;width:100%;border-collapse:collapse}.summary th,.summary td{border:1px solid #222;padding:5px}.summary th{background:#d9ead3}.detail{page-break-before:always}.detail table{width:100%;border-collapse:collapse;font-size:8px}.detail th,.detail td{border:1px solid #777;padding:4px}.detail th{background:#e5e7eb}.right{text-align:right}.footer{margin-top:18px;display:grid;grid-template-columns:1fr 1fr;gap:25px}.sign{border-top:1px solid #555;padding-top:6px;text-align:center}.note{margin-top:10px;border:1px solid #999;padding:7px}.muted{color:#666}@page{size:A4 portrait;margin:8mm}@media print{body{padding:0}.no-print{display:none}}
            </style></head><body>
            <h1>PIPING WORK MATERIAL TAKE OFF (MTO) FOR UTILITY &amp; OFFSITE FACILITIES</h1>
            <h2>${esc(meta.projectName || this.activeProject)}</h2>
            <div class="meta"><div class="label">CONSULTANT NAME</div><div>TTS Consortium</div><div class="label">DATE</div><div>${esc(new Date().toLocaleDateString('en-GB'))}</div>
            <div class="label">CLIENT</div><div>PT. KILANG PERTAMINA INTERNASIONAL</div><div class="label">REVISION</div><div>Rev ${esc(meta.version || 0)}</div>
            <div class="label">SITE</div><div>${esc(meta.projectName || this.activeProject)}</div><div class="label">MADE / CHECKED BY</div><div>${approved ? 'LEAD ESTIMATOR' : '-'}</div></div>
            <div class="status">STATUS: ${esc(status)}</div>
            <table class="main"><thead><tr><th>NO</th><th>DESCRIPTION</th><th>QTY</th><th>UNIT</th><th>REMARKS</th></tr></thead><tbody>${reportRows}</tbody></table>
            <table class="summary"><tr><th>BOQ SUMMARY</th><th class="right">VALUE</th></tr><tr><td>Direct Cost</td><td class="right">${money(boq.directCost)}</td></tr><tr><td>Indirect Cost</td><td class="right">${money(boq.indirectCost)}</td></tr><tr><th>Total BOQ</th><th class="right">${money(boq.totalCost)}</th></tr></table>
            <div class="note"><b>Approval Note:</b> ${esc(approval.notes || meta.revisionNotes || '-')}</div>
            <div class="footer"><div class="sign">Prepared by<br><b>Estimator Proposal</b></div><div class="sign">Reviewed &amp; Approved by<br><b>Lead Estimator</b></div></div>
            </body></html>`;

            const win = window.open('', '_blank', 'width=1100,height=850');
            if (!win) return alert('Popup diblokir browser. Izinkan popup untuk membuka laporan PDF.');
            win.document.open(); win.document.write(html); win.document.close(); win.focus();
            setTimeout(() => win.print(), 500);
        },

        exportBOQ() {
            const meta = this.allProjectsData?.[this.activeProject]?.meta;
            const boq = meta?.boq;
            if (!boq?.items?.length) return alert('BOQ belum tersedia. Estimator harus menghitung BOM terlebih dahulu.');

            const format = (num) => {
                const n = Number(num) || 0;
                if (boq.currency === 'IDR') {
                    return new Intl.NumberFormat('id-ID', { style:'currency', currency:'IDR', minimumFractionDigits:0, maximumFractionDigits:0 }).format(n);
                }
                return new Intl.NumberFormat('en-US', { style:'currency', currency:'USD', minimumFractionDigits:2, maximumFractionDigits:2 }).format(n);
            };
            const rows = boq.items.map((item, index) => `
                <tr>
                    <td>${index + 1}</td><td>${item.lineNo || '-'}</td><td>${item.sheet || '-'}</td><td>${item.component || '-'}</td>
                    <td>${Number(item.qty || 0).toFixed(2)}</td><td>${item.unit || '-'}</td><td>${format(item.unitPrice)}</td><td>${format(item.totalPrice)}</td>
                </tr>`).join('');

            const html = `<!doctype html><html><head><meta charset="utf-8"><title>BOQ ${this.activeProject}</title><style>
                body{font-family:Arial,sans-serif;font-size:10pt;color:#111;padding:25px}h1{font-size:16pt;margin-bottom:4px}h2{font-size:12pt;margin-top:25px}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{border:1px solid #777;padding:6px}th{background:#e2e8f0;text-align:center}td:nth-child(1),td:nth-child(5){text-align:center}td:nth-child(n+7){text-align:right}.summary{margin-top:18px;width:420px;margin-left:auto}.summary td{border:none;padding:4px}.total{font-weight:bold;border-top:1px solid #111!important}.muted{color:#64748b;font-size:9pt}
            </style></head><body>
                <h1>PT. TRIPATRA ENGINEERING</h1><div class="muted">Laporan Kalkulasi Biaya BOQ • ${this.activeProject} • Rev ${meta.version || 0}</div>
                <h2>Ringkasan BOQ</h2><div>Level Harga: <b>${boq.priceLevel || '-'}</b> &nbsp; | &nbsp; Mata Uang: <b>${boq.currency || '-'}</b></div>
                <table><thead><tr><th>No</th><th>Line No.</th><th>Sheet</th><th>Component</th><th>Qty</th><th>Unit</th><th>Unit Price</th><th>Total</th></tr></thead><tbody>${rows}</tbody></table>
                <table class="summary"><tr><td>Direct Cost</td><td>${format(boq.directCost)}</td></tr><tr><td>Indirect Cost</td><td>${format(boq.indirectCost)}</td></tr><tr class="total"><td>Total BOQ</td><td>${format(boq.totalCost)}</td></tr></table>
            </body></html>`;
            const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
            const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `Laporan_BOQ_${this.activeProject}_Rev${meta.version || 0}.doc`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
        },

        printPDF() {
            const rows = this.filteredRows || [];
            if (!rows.length) return alert(`Tidak ada data pada sheet ${this.activeSheet}.`);

            // PDF Export khusus: tidak mengubah tampilan tabel utama.
            const columns = this.currentColumns || [];
            const escapeHtml = (v) => String(v ?? '').replace(/[&<>"']/g, ch => ({
                '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
            }[ch]));

            const paperOptions = [
                { value: 'A3', label: 'A3 (Default - Landscape)' },
                { value: 'A4', label: 'A4 (Landscape)' },
                { value: 'Letter', label: 'Letter (Landscape)' }
            ];

            // Modal pilihan kertas agar export tetap konsisten dengan kebutuhan kantor.
            const old = document.getElementById('pdf-export-settings-modal');
            if (old) old.remove();
            const modal = document.createElement('div');
            modal.id = 'pdf-export-settings-modal';
            modal.innerHTML = `
                <div style="position:fixed;inset:0;background:rgba(15,23,42,.42);z-index:99999;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Arial,sans-serif">
                    <div style="width:min(420px,94vw);background:#fff;border:1px solid #dbe4ef;border-radius:14px;box-shadow:0 20px 60px rgba(15,23,42,.22);overflow:hidden">
                        <div style="padding:16px 18px;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;gap:10px">
                            <div style="width:30px;height:30px;border-radius:8px;background:#eef8e7;display:flex;align-items:center;justify-content:center;color:#63ad20;font-size:14px">▣</div>
                            <div>
                                <div style="font-size:14px;font-weight:700;color:#1e293b">Pengaturan Ekspor PDF</div>
                                <div style="font-size:10px;color:#64748b;margin-top:2px">${escapeHtml(this.activeProject)} — ${escapeHtml(this.activeSheet)}</div>
                            </div>
                        </div>
                        <div style="padding:18px">
                            <label style="display:block;font-size:11px;font-weight:700;color:#334155;margin-bottom:7px">Pilih Ukuran Kertas</label>
                            <select id="pdf-paper-size" style="width:100%;height:38px;border:1px solid #94b8df;border-radius:8px;padding:0 10px;font-size:11px;color:#1e293b;background:#fff;outline:none">
                                ${paperOptions.map((o,i)=>`<option value="${o.value}" ${i===0?'selected':''}>${o.label}</option>`).join('')}
                            </select>
                            <div style="margin-top:10px;padding:10px 11px;border-radius:8px;background:#f8fafc;border:1px solid #e2e8f0;color:#64748b;font-size:10px;line-height:1.45">
                                PDF akan dibuat <b style="color:#334155">Landscape</b>, memakai header perusahaan, logo, informasi project, dan tabel yang dioptimalkan untuk kertas terpilih.
                            </div>
                        </div>
                        <div style="padding:12px 18px;border-top:1px solid #e5e7eb;display:flex;justify-content:flex-end;gap:8px">
                            <button id="pdf-cancel" style="height:34px;padding:0 14px;border:1px solid #cbd5e1;background:#fff;color:#475569;border-radius:8px;font-size:11px;cursor:pointer">Batal</button>
                            <button id="pdf-generate" style="height:34px;padding:0 16px;border:0;background:#82C341;color:#fff;border-radius:8px;font-size:11px;font-weight:700;cursor:pointer">Lanjutkan Export</button>
                        </div>
                    </div>
                </div>`;
            document.body.appendChild(modal);

            const closeModal = () => modal.remove();
            modal.querySelector('#pdf-cancel').onclick = closeModal;
            modal.querySelector('#pdf-generate').onclick = () => {
                const paper = modal.querySelector('#pdf-paper-size').value;
                closeModal();
                this._openOptimizedPDF(rows, columns, paper, escapeHtml);
            };
        },

        _openOptimizedPDF(rows, columns, paper = 'A3', escapeHtml = (v) => String(v ?? '')) {
            const meta = this.allProjectsData?.[this.activeProject]?.meta || {};
            const logoUrl = (() => {
                try { return new URL('logo tripatra.png', window.location.href).href; }
                catch(e) { return 'logo tripatra.png'; }
            })();
            const now = new Date().toLocaleString('id-ID', { dateStyle:'short', timeStyle:'medium' });
            const safeProject = escapeHtml(this.activeProject || meta.projectName || '-');
            const safeSheet = escapeHtml(this.activeSheet || '-');
            const safeRevision = escapeHtml(meta.version ?? 0);
            const safeStatus = escapeHtml(meta.status || 'Draft - Pekerjaan Engineer');

            /*
             * Jangan mengecilkan 50-100 kolom menjadi satu halaman.
             * PDF dibagi menjadi beberapa bagian horizontal yang mudah dibaca.
             * Setiap bagian mengulang No + Long Description sehingga baris tetap
             * mudah dicocokkan walaupun berpindah bagian.
             */
            const allCols = Array.from(new Set(columns.filter(Boolean)));
            const descriptionCol = allCols.includes('Long Description (Family)')
                ? 'Long Description (Family)'
                : (allCols.find(c => /description/i.test(c)) || allCols[0] || 'Description');
            const remaining = allCols.filter(c => c !== descriptionCol);

            const preferred = [
                'Compatible Standard','Manufacturer','Material','Material Code','Long Description (Size)',
                'Short Description','Spec','Size','Line Number Tag','Design Std','Design Type',
                'End Type','Engagement Length','Facing','Flange Std','Gasket Std','Port Unit','Nominal Diameter'
            ];
            const ordered = [
                ...preferred.filter(c => remaining.includes(c)),
                ...remaining.filter(c => !preferred.includes(c))
            ];

            const maxPerSection = paper === 'A3' ? 8 : 6;
            const sections = [];
            for (let i = 0; i < ordered.length; i += maxPerSection) {
                sections.push(ordered.slice(i, i + maxPerSection));
            }
            if (!sections.length) sections.push([]);

            const paperConfig = {
                A3: { font:'8.2px', head:'8.0px', pad:'4.2px', title:'15px', desc:'30%', sectionCols:8 },
                A4: { font:'7.6px', head:'7.3px', pad:'3.8px', title:'13px', desc:'30%', sectionCols:6 },
                Letter: { font:'7.4px', head:'7.1px', pad:'3.6px', title:'13px', desc:'30%', sectionCols:6 }
            }[paper] || { font:'7.1px', head:'6.7px', pad:'3.4px', title:'14px', desc:'30%', sectionCols:6 };

            const sectionName = (cols, index) => {
                const text = cols.join(' ').toLowerCase();
                if (/line|description|standard|material|spec|size|manufacturer/.test(text) && index === 0) return 'Identification & Specification';
                if (/design|end|engagement|facing|flange|gasket|port|pressure|temperature/.test(text)) return 'Design & Connection';
                if (/coordinate|elevation|length|thickness|weight|volume|area|dimension/.test(text)) return 'Physical & Location Data';
                if (/status|revision|remark|approval|inspection|paint|coating|insulation/.test(text)) return 'Status & Engineering Data';
                return `Data Section ${index + 1}`;
            };

            const makeTable = (sectionCols) => {
                const headers = ['No', descriptionCol, ...sectionCols.filter(c => c !== descriptionCol)];

                // Lebar kolom mengikuti panjang header + isi aktual.
                // Bobot memakai panjang teks maksimum yang benar-benar ada di data,
                // lalu dinormalisasi ke lebar halaman PDF agar kolom pendek tidak
                // memakan ruang berlebihan dan kolom panjang tetap mendapat ruang.
                const getCellText = (h, row) => String(row?.[h] ?? '').trim();
                const naturalLengths = headers.map((h, idx) => {
                    if (idx === 0) return 4;
                    const values = rows.map(row => getCellText(h, row));
                    const maxValueLength = values.reduce((m, v) => Math.max(m, v.length), 0);
                    const avgValueLength = values.length
                        ? values.reduce((sum, v) => sum + Math.min(v.length, 80), 0) / values.length
                        : 0;
                    const headerLength = String(h).length;
                    // Gabungkan header, isi terpanjang, dan rata-rata isi.
                    // Description mendapat prioritas sedikit lebih tinggi karena biasanya paling panjang.
                    const score = Math.max(
                        headerLength,
                        maxValueLength * 0.58 + avgValueLength * 0.42
                    );
                    return h === descriptionCol ? score * 1.12 : score;
                });

                // Batasi ekstrem agar satu cell tidak mengambil seluruh halaman.
                const clipped = naturalLengths.map((v, idx) => {
                    if (idx === 0) return 4;
                    return Math.max(6, Math.min(34, v));
                });
                const weightTotal = clipped.reduce((a,b) => a + b, 0) || 1;
                const widths = clipped.map(w => (w / weightTotal) * 100);
                const colgroup = widths.map((w, i) => `<col style="width:${w.toFixed(2)}%">`).join('');

                const head = headers.map(h => `<th class="${h === descriptionCol ? 'desc-head' : ''}">${escapeHtml(h)}</th>`).join('');
                const body = rows.map((row, i) => {
                    const cells = headers.map((h, j) => {
                        const value = j === 0 ? String(i + 1).padStart(2,'0') : row[h];
                        return `<td class="${j === 0 ? 'no-cell' : ''} ${h === descriptionCol ? 'desc-cell' : ''}">${escapeHtml(value)}</td>`;
                    }).join('');
                    return `<tr>${cells}</tr>`;
                }).join('');
                return `<table><colgroup>${colgroup}</colgroup><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
            };

            const sectionsHtml = sections.map((sectionCols, index) => {
                const label = sectionName(sectionCols, index);
                return `<section class="pdf-section ${index ? 'new-section' : ''}">
                    <header class="pdf-header">
                        <div class="logo-wrap"><img class="logo" src="${escapeHtml(logoUrl)}" alt="Tripatra"></div>
                        <div class="title">
                            <h1>MASTER LINE LIST WEBAPP</h1>
                            <div class="sub">PROJECT: ${safeProject} &nbsp; • &nbsp; SHEET: ${safeSheet}</div>
                            <div class="section-title">${escapeHtml(label)} &nbsp; • &nbsp; Part ${index + 1} of ${sections.length}</div>
                        </div>
                        <div class="meta">
                            <div><b>REV:</b> ${safeRevision}</div>
                            <div class="status"><b>STATUS:</b> ${safeStatus}</div>
                            <div><b>EXPORTED:</b> ${escapeHtml(now)}</div>
                        </div>
                    </header>
                    <section class="report-info">
                        <div class="info-box"><div class="info-label">Project</div><div class="info-value">${safeProject}</div></div>
                        <div class="info-box"><div class="info-label">Document</div><div class="info-value">${escapeHtml(meta.documentNumber || meta.documentNo || '-')}</div></div>
                        <div class="info-box"><div class="info-label">Sheet</div><div class="info-value">${safeSheet}</div></div>
                        <div class="info-box"><div class="info-label">Total Data</div><div class="info-value">${rows.length} data</div></div>
                    </section>
                    <div class="table-wrap">${makeTable(sectionCols)}</div>
                    <div class="footer">MASTER LINE LIST WEBAPP • Tripatra Engineering Studio • Part ${index + 1}/${sections.length}</div>
                </section>`;
            }).join('');

            const html = `<!doctype html>
<html><head><meta charset="utf-8"><title>${safeProject} — ${safeSheet}</title>
<style>
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:#fff;color:#142033;font-family:Arial,Helvetica,sans-serif}
body{font-size:${paperConfig.font};-webkit-print-color-adjust:exact;print-color-adjust:exact}
.pdf-section{padding:8mm 9mm 7mm;min-height:100vh;display:flex;flex-direction:column}
.new-section{break-before:page;page-break-before:always}
.pdf-header{display:grid;grid-template-columns:125px 1fr 190px;align-items:center;gap:16px;margin-bottom:4mm;border-bottom:2px solid #173b63;padding-bottom:3mm}
.logo-wrap{display:flex;align-items:center;justify-content:flex-start;height:22mm}.logo{max-width:105px;max-height:21mm;object-fit:contain}
.title{text-align:center}.title h1{font-size:${paperConfig.title};margin:0 0 2px;font-weight:800;letter-spacing:.2px;color:#0f172a}.title .sub{font-size:7.5px;color:#64748b;margin-top:2px}.section-title{font-size:8.5px;color:#2d7964;font-weight:800;margin-top:4px}
.meta{font-size:7.2px;line-height:1.5;text-align:right;color:#334155}.meta b{color:#0f172a}.meta .status{color:#4e9f17;font-weight:700}
.report-info{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;margin-bottom:4mm}.info-box{border:1px solid #d7e0ea;border-radius:4px;padding:3px 5px;background:#f8fafc;min-height:9mm}.info-label{font-size:6.2px;text-transform:uppercase;color:#64748b;font-weight:700}.info-value{font-size:7.5px;color:#0f172a;font-weight:700;margin-top:1.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.table-wrap{width:100%;overflow:visible}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:${paperConfig.font};margin:0}
thead{display:table-header-group}tr{page-break-inside:avoid;break-inside:avoid}th,td{border:1px solid #aebed0;padding:${paperConfig.pad};vertical-align:middle;overflow-wrap:anywhere;word-break:normal;line-height:1.18;white-space:normal;text-align:center}th{background:#2d7964;color:#fff;text-align:center;font-weight:700;font-size:${paperConfig.head};white-space:normal}td{color:#172033;background:#fff}tbody tr:nth-child(even) td{background:#f7fafc}
th:first-child,.no-cell{min-width:22px}.no-cell{text-align:center;font-weight:700}.desc-head,.desc-cell{min-width:70px}.desc-cell{font-weight:600;line-height:1.25;text-align:center}
.footer{margin-top:4mm;padding-top:1.5mm;border-top:1px solid #d7e0ea;text-align:right;font-size:6px;color:#94a3b8}
@page{size:${paper} landscape;margin:0}
@media print{.pdf-section{min-height:auto}}
</style></head><body>${sectionsHtml}</body></html>`;

            const win = window.open('', '_blank', 'width=1400,height=900');
            if (!win) return alert('Popup diblokir browser. Izinkan popup untuk mencetak PDF.');
            win.document.open();
            win.document.write(html);
            win.document.close();
            win.focus();
            setTimeout(() => win.print(), 800);
        }
    }
}