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

        showAddModal: false,
        showBoqModal: false,
        showApproveModal: false,
        taskView: 'active',
        approvalHistory: [],
        approvalNote: '',
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

        activePriceLevel: 'Menengah',
        boqCurrency: 'USD',
        
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

            const fullValveData = [
                {"Number": "01", "Long Description (Family)": "Check Valve, Swing, 150 LB, RF, ASME B16.10", "Material": "CS", "Material Code": "ASTM A216 Gr WPB", "Short Description": "Check Valve", "Spec": "CS150", "Size": "10\"", "End Type": "FL", "Facing": "RF", "Pressure Class": "150", "Status": "New", "Weight": "440.92 LB", "Shop/Field": "SHOP", "PnPID": "3457"},
                {"Number": "02", "Long Description (Family)": "Gate Valve, Solid Wedge, 150 LB, RF, ASME B16.10", "Material": "CS", "Material Code": "ASTM A216 Gr WPB", "Short Description": "Gate Valve", "Spec": "CS150", "Size": "3/4\"", "End Type": "FL", "Facing": "RF", "Pressure Class": "150", "Status": "New", "Weight": "", "Shop/Field": "SHOP", "PnPID": "3690"}
            ];

            const defaultPRJ3000 = {
                'meta': {
                    isApproved: false,
                    approvedAt: '',
                    version: 0,
                    progress: '30%',
                    projectCode: 'PRJ-3000',
                    projectName: 'Piping & Equipment',
                    description: 'Project utama BMBQ WebApp'
                },
                'Valve': fullValveData,
                'SP Items': [],
                'Support': [],
                'LineList': [],
                'Tee': []
            };

            this.sheets.forEach(sh => {
                if (!defaultPRJ3000[sh]) defaultPRJ3000[sh] = [];
            });

            const defaultPRJ4000 = JSON.parse(JSON.stringify(defaultPRJ3000));
            defaultPRJ4000.meta.projectCode = 'PRJ-4000';
            defaultPRJ4000.meta.projectName = 'Piping & Equipment';

            const defaultPRJ5000 = JSON.parse(JSON.stringify(defaultPRJ3000));
            defaultPRJ5000.meta.projectCode = 'PRJ-5000';
            defaultPRJ5000.meta.projectName = 'Piping & Equipment';

            this.allProjectsData = {
                'PRJ-3000': defaultPRJ3000,
                'PRJ-4000': defaultPRJ4000,
                'PRJ-5000': defaultPRJ5000
            };

            const saved = localStorage.getItem('tripatra_multiproject_db');
            if (saved) {
                try {
                    const parsed = JSON.parse(saved);
                    if (parsed && typeof parsed === 'object' && Object.keys(parsed).length > 0) {
                        this.allProjectsData = parsed;
                    }
                } catch (e) {
                    console.error('Gagal membaca localStorage:', e);
                }
            }

            this.normalizeAllProjectsData();
            this.refreshSheetList();
            this.loadApprovalHistory();
            this.saveStorage();
        },

        normalizeRow(row, idx = 0) {
            const normalized = {};
            Object.entries(row || {}).forEach(([key, value]) => {
                const cleanKey = String(key).replace(/\u00a0/g, ' ').trim();
                normalized[cleanKey] = value;
            });

            const aliases = {
                'No': 'Number',
                'NO': 'Number',
                'No.': 'Number',
                'Long Description': 'Long Description (Family)',
                'Compatible standard': 'Compatible Standard',
                'MFG': 'Manufacturer',
                'Mfr': 'Manufacturer'
            };

            Object.entries(aliases).forEach(([from, to]) => {
                if (Object.prototype.hasOwnProperty.call(normalized, from) &&
                    !Object.prototype.hasOwnProperty.call(normalized, to)) {
                    normalized[to] = normalized[from];
                    delete normalized[from];
                }
            });

            delete normalized['Long Description (Family) '];

            if (!normalized.Number) normalized.Number = String(idx + 1).padStart(2, '0');
            return normalized;
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
                        progress: '30%'
                    };
                } else {
                    if (!project.meta.projectCode) project.meta.projectCode = projectKey;
                    if (!project.meta.projectName) project.meta.projectName = 'Piping & Equipment';
                    if (!project.meta.description) project.meta.description = '';
                    if (!project.meta.progress) project.meta.progress = '30%';
                }

                Object.keys(project).forEach(sheetName => {
                    if (sheetName === 'meta') return;
                    if (!Array.isArray(project[sheetName])) project[sheetName] = [];
                    project[sheetName] = project[sheetName]
                        .filter(row => row && typeof row === 'object')
                        .map((row, idx) => this.normalizeRow(row, idx));
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

        saveStorage() {
            localStorage.setItem('tripatra_multiproject_db', JSON.stringify(this.allProjectsData));
        },

        switchProject() {
            this.currentPage = 1;
            this.globalSearch = '';
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
            if (!this.allProjectsData[this.activeProject]) this.allProjectsData[this.activeProject] = {};
            if (!this.allProjectsData[this.activeProject][this.activeSheet]) this.allProjectsData[this.activeProject][this.activeSheet] = [];
            return this.allProjectsData[this.activeProject][this.activeSheet];
        },

        get currentColumns() {
            if (this.currentRows.length === 0) {
                if (this.activeSheet === 'SP Items') return ['Material', 'Material Code', 'Spec', 'Size', 'Qty', 'Unit', 'Line Number Tag', 'Status'];
                if (this.activeSheet === 'Support') return ['Support Type', 'Material', 'Spec', 'Size', 'Qty', 'Unit', 'Line Number Tag', 'Status'];
                if (this.activeSheet === 'LineList') return ['Line Number Tag', 'Material', 'Spec', 'Size', 'Pressure Class', 'Fluid Service', 'Status'];
                return ['Material', 'Spec', 'Size', 'Pressure Class', 'Line Number Tag', 'Status', 'Shop/Field', 'Waste Factor'];
            }

            const seen = new Set();
            const keys = [];

            this.currentRows.forEach(row => {
                Object.keys(row || {}).forEach(col => {
                    const clean = String(col).trim();
                    const lower = clean.toLowerCase();

                    if (
                        clean &&
                        !seen.has(lower) &&
                        lower !== 'number' &&
                        lower !== 'no' &&
                        lower !== 'long description (family)' &&
                        lower !== 'long description'
                    ) {
                        seen.add(lower);
                        keys.push(clean);
                    }
                });
            });

            return keys;
        },

        get filteredRows() {
            let rows = this.currentRows;
            if (this.globalSearch) {
                const q = this.globalSearch.toLowerCase();
                rows = rows.filter(row => 
                    Object.values(row).some(val => String(val).toLowerCase().includes(q))
                );
            }
            return rows;
        },

        get maxPage() {
            return Math.ceil(this.filteredRows.length / this.itemsPerPage) || 1;
        },

        get paginatedRows() {
            const start = (this.currentPage - 1) * this.itemsPerPage;
            return this.filteredRows.slice(start, start + this.itemsPerPage);
        },

        openDataSheet(sheetName) {
        this.activeSheet = sheetName;
        this.currentDashboardTab = 'workspace';
        this.currentPage = 1;
        this.globalSearch = '';

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

            const rows = project[sheet];
            const nextNum = String(rows.length + 1).padStart(2, '0');

            const newRow = {
                Number: nextNum,
                'Long Description (Family)': '',
                'Compatible Standard': '',
                Manufacturer: '',
                Material: '',
                'Material Code': '',
                Spec: '',
                Size: '',
                'Pressure Class': '',
                'Fluid Service': '',
                'Line Number Tag': '',
                'Support Type': '',
                Qty: '',
                Unit: '',
                'Waste Factor': '',
                'Shop/Field': '',
                Status: 'New'
            };

            (this.currentColumns || []).forEach(col => {
                if (!(col in newRow)) newRow[col] = '';
            });

            rows.push(newRow);
            this.globalSearch = '';
            this.currentPage = Math.max(1, Math.ceil(rows.length / this.itemsPerPage));
            this.saveStorage();

            this.$nextTick(() => {
                this.updateTableScrollbar();
                const tableRows = document.querySelectorAll('.excel-table tbody tr');
                const lastRow = tableRows[tableRows.length - 1];
                const firstInput = lastRow?.querySelector('input');
                if (firstInput) firstInput.focus();
            });
        },
        resetNewRowForm() {
            this.newRowForm = {
                longDesc: '', material: 'CS', spec: 'CS150', size: '2"', pressure: '150',
                wasteFactor: '5', tag: '', qty: '', unit: '', service: '', supportType: ''
            };
        },

        changeSheet(sheetName) {
            this.activeSheet = sheetName;
            this.currentPage = 1;
            this.globalSearch = '';
            this.$nextTick(() => this.updateTableScrollbar());
        },

        updateTableScrollbar() {
            this.$nextTick(() => {
                const container = this.$refs?.tableScroll;
                const table = container?.querySelector('.excel-table');
                const spacer = this.$refs?.tableBottomScrollbarSpacer;
                const bottom = this.$refs?.tableBottomScrollbar;

                if (!container || !table || !spacer || !bottom) return;

                const tableWidth = Math.max(
                    table.scrollWidth,
                    table.offsetWidth,
                    table.getBoundingClientRect().width,
                    1800
                );

                spacer.style.width = tableWidth + 'px';

                // Tampilkan scrollbar hanya jika tabel lebih lebar dari viewport.
                bottom.style.display =
                    tableWidth > container.clientWidth + 1 ? 'block' : 'none';

                bottom.scrollLeft = container.scrollLeft || 0;
            });
        },

        syncTableBottomScroll(event) {
            const container = this.$refs?.tableScroll;
            if (!container) return;

            if (container.scrollLeft !== event.target.scrollLeft) {
                container.scrollLeft = event.target.scrollLeft;
            }
        },

        syncBottomTableScroll(event) {
            const bottom = this.$refs?.tableBottomScrollbar;
            if (!bottom) return;

            if (bottom.scrollLeft !== event.target.scrollLeft) {
                bottom.scrollLeft = event.target.scrollLeft;
            }
        },

        importExcelFile(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellDates: true,
                        raw: false
                    });

                    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                        throw new Error('Workbook tidak memiliki sheet.');
                    }

                    if (!this.allProjectsData[this.activeProject]) {
                        this.allProjectsData[this.activeProject] = {};
                    }

                    if (!this.allProjectsData[this.activeProject].meta) {
                        this.allProjectsData[this.activeProject].meta = {
                            isApproved: false,
                            approvedAt: '',
                            version: 0,
                            progress: '30%'
                        };
                    }

                    let importedCount = 0;
                    let firstImportedSheet = '';

                    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
                        const worksheet = workbook.Sheets[sheetName];
                        if (!worksheet || !worksheet['!ref']) return;

                        let jsonRows = XLSX.utils.sheet_to_json(worksheet, {
                            defval: '',
                            raw: false,
                            blankrows: false
                        });

                        jsonRows = jsonRows
                            .filter(row => Object.values(row || {}).some(value => String(value).trim() !== ''))
                            .map((row, idx) => this.normalizeRow(row, idx));

                        const cleanSheetName = String(sheetName).trim() || `Sheet ${sheetIndex + 1}`;
                        this.allProjectsData[this.activeProject][cleanSheetName] = jsonRows;
                        importedCount += jsonRows.length;

                        if (!firstImportedSheet) firstImportedSheet = cleanSheetName;
                    });

                    this.refreshSheetList();
                    if (firstImportedSheet) this.activeSheet = firstImportedSheet;

                    this.currentPage = 1;
                    this.globalSearch = '';
                    this.saveStorage();

                    alert(
                        `Import berhasil!\n\n` +
                        `Proyek: ${this.activeProject}\n` +
                        `Sheet: ${workbook.SheetNames.length}\n` +
                        `Total data: ${importedCount} baris\n\n` +
                        `Sheet aktif: ${this.activeSheet}`
                    );
                } catch (error) {
                    console.error('Import Excel error:', error);
                    alert('Gagal import Excel: ' + (error.message || 'format file tidak valid.'));
                } finally {
                    event.target.value = '';
                }
            };

            reader.onerror = () => {
                alert('File Excel tidak dapat dibaca oleh browser.');
                event.target.value = '';
            };

            reader.readAsArrayBuffer(file);
        },

        saveNewRowData() {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Akses ditolak! Hanya Piping Engineer yang dapat menambah data.');
                return;
            }

            const sheet = this.activeSheet;
            const rows = this.currentRows;
            const nextNum = String(rows.length + 1).padStart(2, '0');
            let row = { Number: nextNum, Status: 'New' };

            if (sheet === 'LineList') {
                row = {
                    Number: nextNum,
                    'Line Number Tag': this.newRowForm.tag,
                    'Long Description (Family)': this.newRowForm.longDesc,
                    Material: this.newRowForm.material,
                    Spec: this.newRowForm.spec,
                    Size: this.newRowForm.size,
                    'Pressure Class': this.newRowForm.pressure,
                    'Fluid Service': this.newRowForm.service,
                    'Status': 'New'
                };
            } else if (sheet === 'Support') {
                row = {
                    Number: nextNum,
                    'Long Description (Family)': this.newRowForm.longDesc,
                    'Support Type': this.newRowForm.supportType,
                    Material: this.newRowForm.material,
                    Spec: this.newRowForm.spec,
                    Size: this.newRowForm.size,
                    'Line Number Tag': this.newRowForm.tag,
                    Qty: this.newRowForm.qty,
                    Unit: this.newRowForm.unit || 'EA',
                    Status: 'New'
                };
            } else if (sheet === 'SP Items') {
                row = {
                    Number: nextNum,
                    'Long Description (Family)': this.newRowForm.longDesc,
                    Material: this.newRowForm.material,
                    'Material Code': '',
                    Spec: this.newRowForm.spec,
                    Size: this.newRowForm.size,
                    Qty: this.newRowForm.qty,
                    Unit: this.newRowForm.unit || 'EA',
                    'Line Number Tag': this.newRowForm.tag,
                    Status: 'New'
                };
            } else {
                row = {
                    Number: nextNum,
                    'Long Description (Family)': this.newRowForm.longDesc,
                    Material: this.newRowForm.material,
                    Spec: this.newRowForm.spec,
                    Size: this.newRowForm.size,
                    'Pressure Class': this.newRowForm.pressure,
                    'Waste Factor': this.newRowForm.wasteFactor ? this.newRowForm.wasteFactor + '%' : '',
                    'Line Number Tag': this.newRowForm.tag,
                    Status: 'New',
                    'Shop/Field': 'SHOP',
                    'Compatible Standard': '',
                    Manufacturer: ''
                };
            }

            rows.push(row);
            rows.forEach((item, idx) => item.Number = String(idx + 1).padStart(2, '0'));
            this.saveStorage();
            this.showAddModal = false;
            this.resetNewRowForm();
            alert(`Data ${sheet} berhasil ditambahkan.`);
        },

        deleteRow(index) {
            if (this.loginForm.role !== 'Piping Engineer') {
                alert('Akses ditolak.');
                return;
            }
            const actualIndex = (this.currentPage - 1) * this.itemsPerPage + index;
            if (confirm('Hapus baris data ini?')) {
                this.currentRows.splice(actualIndex, 1);
                
                // Re-index nomor urut data setelah aksi hapus[cite: 7]
                this.currentRows.forEach((row, idx) => {
                    row['Number'] = String(idx + 1).padStart(2, '0');
                });

                this.saveStorage();
                alert('Data dihapus.');
            }
        },

        loadApprovalHistory() {
            try {
                const saved = localStorage.getItem('tripatra_approval_history');
                this.approvalHistory = saved ? JSON.parse(saved) : [];
                if (!Array.isArray(this.approvalHistory)) this.approvalHistory = [];
            } catch (e) {
                console.error('Gagal membaca riwayat approval:', e);
                this.approvalHistory = [];
            }
        },

        saveApprovalHistory() {
            localStorage.setItem('tripatra_approval_history', JSON.stringify(this.approvalHistory));
        },

        get projectApprovalHistory() {
            return (this.approvalHistory || []).filter(item => item.projectId === this.activeProject);
        },

        get activeTaskCount() {
            // Tugas aktif tetap ditampilkan setelah ada keputusan.
            // Riwayat keputusan dicatat terpisah di tab "Riwayat Selesai".
            return 2;
        },

        get latestApproval() {
            const list = this.projectApprovalHistory;
            return list.length ? list[list.length - 1] : null;
        },

        openApprovalHistory() {
            this.taskView = 'history';
        },

        openActiveTasks() {
            this.taskView = 'active';
        },

        getApprovalStatusText(status) {
            const map = {
                Approved: 'Approved',
                Rejected: 'Rejected',
                Revision: 'Minta Revisi'
            };
            return map[status] || status || 'Menunggu';
        },

        getApprovalStatusClass(status) {
            if (status === 'Approved') return 'bg-emerald-100 text-emerald-700';
            if (status === 'Rejected') return 'bg-rose-100 text-rose-700';
            if (status === 'Revision') return 'bg-amber-100 text-amber-700';
            return 'bg-slate-100 text-slate-700';
        },

        openApprovalDetail(item) {
            const note = item.notes ? `\n\nCatatan: ${item.notes}` : '';
            alert(`Detail Approval\n\nProject: ${item.projectId}\nSheet: ${item.sheet}\nStatus: ${this.getApprovalStatusText(item.status)}\nRevision: Rev ${item.revision}\nReviewer: ${item.reviewer}\nWaktu: ${item.timestamp}${note}`);
        },

        generateBOQ() { this.showBoqModal = true; },
        approveData() {
            this.approvalNote = '';
            const textarea = document.getElementById('reviewer-notes-textarea');
            if (textarea) textarea.value = '';
            this.showApproveModal = true;
        },

        confirmApprove() {
            const project = this.allProjectsData[this.activeProject];
            if (!project) return;
            if (!project.meta) project.meta = {};

            const currentRev = Number(project.meta.version || 0);
            const nextRev = currentRev + 1;
            const textarea = document.getElementById('reviewer-notes-textarea');
            const notes = textarea ? textarea.value.trim() : '';
            const now = new Date().toLocaleString('id-ID');

            project.meta.version = nextRev;
            project.meta.isApproved = true;
            project.meta.approvedAt = now;
            project.meta.approvalStatus = 'Approved';

            this.approvalHistory.push({
                id: Date.now(),
                projectId: this.activeProject,
                projectName: project.meta.projectName || 'Piping & Equipment',
                sheet: this.activeSheet || 'Valve',
                title: 'Review MTO & Spesifikasi Valve Proyek Saat Ini',
                status: 'Approved',
                revision: nextRev,
                reviewer: this.loginForm.role || 'Lead Estimator',
                notes,
                timestamp: now
            });

            this.saveApprovalHistory();
            this.saveStorage();
            this.showApproveModal = false;
            this.taskView = 'active';
            alert(`Proyek disetujui. Revision berubah menjadi Rev ${nextRev}.`);
        },

        rejectData() {
            const project = this.allProjectsData[this.activeProject];
            if (!project) return;
            if (!project.meta) project.meta = { version: 0, isApproved: false };

            const textarea = document.getElementById('reviewer-notes-textarea');
            const notes = textarea ? textarea.value.trim() : '';
            const now = new Date().toLocaleString('id-ID');
            const currentRev = Number(project.meta.version || 0);

            project.meta.isApproved = false;
            project.meta.approvalStatus = 'Rejected';
            project.meta.approvedAt = '';

            this.approvalHistory.push({
                id: Date.now(),
                projectId: this.activeProject,
                projectName: project.meta.projectName || 'Piping & Equipment',
                sheet: this.activeSheet || 'Valve',
                title: 'Review MTO & Spesifikasi Valve Proyek Saat Ini',
                status: 'Rejected',
                revision: currentRev,
                reviewer: this.loginForm.role || 'Lead Estimator',
                notes,
                timestamp: now
            });

            this.saveApprovalHistory();
            this.saveStorage();
            this.showApproveModal = false;
            this.taskView = 'active';
            alert(`Proyek ditolak pada Rev ${currentRev}.`);
        },

        handleMintaRevisi() {
            const project = this.allProjectsData[this.activeProject];
            if (!project) return;
            if (!project.meta) project.meta = { version: 0, isApproved: false };

            const textarea = document.getElementById('reviewer-notes-textarea');
            const notes = textarea ? textarea.value.trim() : '';
            if (!notes) {
                alert('Silakan isi catatan atau alasan revisi terlebih dahulu.');
                return;
            }

            const now = new Date().toLocaleString('id-ID');
            const currentRev = Number(project.meta.version || 0);
            project.meta.isApproved = false;
            project.meta.approvalStatus = 'Revision';
            project.meta.revisionNotes = notes;

            this.approvalHistory.push({
                id: Date.now(),
                projectId: this.activeProject,
                projectName: project.meta.projectName || 'Piping & Equipment',
                sheet: this.activeSheet || 'Valve',
                title: 'Review MTO & Spesifikasi Valve Proyek Saat Ini',
                status: 'Revision',
                revision: currentRev,
                reviewer: this.loginForm.role || 'Lead Estimator',
                notes,
                timestamp: now
            });

            this.saveApprovalHistory();
            this.saveStorage();
            this.showApproveModal = false;
            this.taskView = 'active';
            alert(`Permintaan revisi berhasil disimpan pada Rev ${currentRev}.`);
        },

        formatCurrency(amount) {
            if (this.boqCurrency === 'IDR') {
                return 'Rp ' + (amount * 16000).toLocaleString('id-ID');
            }
            return '$ ' + amount.toLocaleString('en-US');
        },

        exportExcel() {
            const rows = this.currentRows;
            if (!rows || rows.length === 0) return alert(`Tidak ada data pada sheet ${this.activeSheet}.`);
            const worksheet = XLSX.utils.json_to_sheet(rows);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, this.activeSheet.substring(0, 31));
            XLSX.writeFile(workbook, `${this.activeProject}_${this.activeSheet.replace(/\s+/g, '_')}.xlsx`);
        },

        exportBOQ() {
            const totalItems = this.currentRows.length;
            
            const directCostIDR = this.activePriceLevel === 'Bawah' ? 2400000000 : (this.activePriceLevel === 'Menengah' ? 2960000000 : 3500000000);
            const weldingCostIDR = 400000000;
            const totalAnggaranIDR = directCostIDR + weldingCostIDR;
            const formatRp = (num) => "Rp " + (num || 0).toLocaleString('id-ID');

            let tableRowsHTML = "";
            let totalBiayaMTO = 0;

            this.currentRows.forEach((item, index) => {
                const harga = item.hargaSatuan || 120000;
                const qty = item.qty || 10;
                const total = harga * qty;
                totalBiayaMTO += total;

                tableRowsHTML += `
                    <tr>
                        <td style="text-align: center; border: 1px solid #000; padding: 6px;">${index + 1}</td>
                        <td style="border: 1px solid #000; padding: 6px;">${item.kode || item.tag || 'PIP-0' + (index+1)}</td>
                        <td style="border: 1px solid #000; padding: 6px;">${item.deskripsi || item.nama || 'Carbon Steel Pipe ASTM A106'}</td>
                        <td style="text-align: center; border: 1px solid #000; padding: 6px;">${item.ukuran || '4 Inch'}</td>
                        <td style="text-align: center; border: 1px solid #000; padding: 6px;">${qty}</td>
                        <td style="text-align: center; border: 1px solid #000; padding: 6px;">${item.satuan || 'M'}</td>
                        <td style="text-align: right; border: 1px solid #000; padding: 6px;">${formatRp(harga)}</td>
                        <td style="text-align: right; border: 1px solid #000; padding: 6px;">${formatRp(total)}</td>
                    </tr>
                `;
            });

            const docHTML = `
                <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                <head>
                    <meta charset="utf-8">
                    <title>Laporan BOQ</title>
                    <style>
                        body { font-family: 'Calibri', Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.15; }
                        .header-title { text-align: center; font-weight: bold; font-size: 14pt; margin-bottom: 2px; }
                        .header-sub { text-align: center; font-size: 11pt; margin-bottom: 20px; }
                        .section-title { font-weight: bold; font-size: 11pt; background-color: #f2f2f2; padding: 5px; margin-top: 15px; border-bottom: 1px solid #000; }
                        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10pt; }
                        th { background-color: #002060; color: #ffffff; border: 1px solid #000; padding: 8px; text-align: center; }
                    </style>
                </head>
                <body>
                    <div class="header-title">PT. TRIPATRA ENGINEERING</div>
                    <div class="header-sub">Proyek BMBQ - Gedung Utama Lt. 4, Jakarta</div>
                    <hr style="border: 1px solid #000;">
                    <div style="text-align: center; font-weight: bold; font-size: 12pt; margin: 15px 0;">LAPORAN REKAPITULASI BIAYA & ESTIMASI BOQ</div>
                    
                    <table style="border: none; margin-bottom: 20px;">
                        <tr><td style="width: 180px; border:none;"><b>ID Proyek</b></td><td style="width: 10px; border:none;">:</td><td style="border: none;">${this.activeProject}</td></tr>
                        <tr><td style="border: none;"><b>Tanggal Cetak</b></td><td style="border: none;">:</td><td style="border: none;">04 Agustus 2026</td></tr>
                        <tr><td style="border: none;"><b>Level Harga Satuan</b></td><td style="border: none;">:</td><td style="border: none;">${this.activePriceLevel}</td></tr>
                        <tr><td style="border: none;"><b>Mata Uang</b></td><td style="border: none;">:</td><td style="border: none;">IDR (Rp)</td></tr>
                    </table>

                    <div class="section-title">1. RINGKASAN EKSEKUTIF (EXECUTIVE SUMMARY)</div>
                    <p>Laporan ini merangkum estimasi anggaran biaya langsung (Direct Cost) dan biaya fabrikasi berdasarkan data parameter yang telah di-input ke dalam sistem.</p>
                    
                    <table style="border: none; width: 80%;">
                        <tr><td style="border:none; padding: 3px;">* Total Item MTO</td><td style="border:none; padding: 3px;">: ${totalItems} Item</td></tr>
                        <tr><td style="border:none; padding: 3px;">* Estimasi Direct Cost</td><td style="border:none; padding: 3px;">: ${formatRp(directCostIDR)}</td></tr>
                        <tr><td style="border:none; padding: 3px;">* Estimasi Biaya Las (Inch-Dia)</td><td style="border:none; padding: 3px;">: ${formatRp(weldingCostIDR)}</td></tr>
                        <tr><td colspan="2" style="border-top: 1px solid #000; padding: 3px;"></td></tr>
                        <tr><td style="border:none; padding: 3px;"><b>* TOTAL ESTIMASI ANGGARAN</b></td><td style="border:none; padding: 3px;"><b>: ${formatRp(totalAnggaranIDR)}</b></td></tr>
                    </table>

                    <div class="section-title" style="margin-top: 25px;">2. RINCIAN DATA MATERIAL (MTO BREAKDOWN)</div>
                    <table>
                        <thead>
                            <tr>
                                <th>No</th>
                                <th>Kode / Tag</th>
                                <th>Deskripsi Material</th>
                                <th>Ukuran</th>
                                <th>Qty</th>
                                <th>Satuan</th>
                                <th>Harga Satuan</th>
                                <th>Total Harga</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRowsHTML}
                            <tr>
                                <td colspan="7" style="text-align: right; font-weight: bold; border: 1px solid #000; padding: 6px;">TOTAL BIAYA MTO</td>
                                <td style="text-align: right; font-weight: bold; border: 1px solid #000; padding: 6px;">${formatRp(totalBiayaMTO)}</td>
                            </tr>
                        </tbody>
                    </table>

                    <br><br>
                    <div style="float: right; text-align: right; width: 250px;">
                        <p>Disetujui Oleh,</p>
                        <br><br><br>
                        <p><b>[ Tanda Tangan ]</b><br>Cost Control / Estimator</p>
                    </div>
                </body>
                </html>
            `;

            const blob = new Blob(['\ufeff' + docHTML], {
                type: 'application/msword'
            });
            
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Laporan_BOQ_${this.activeProject}.doc`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
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