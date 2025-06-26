// Variáveis globais
let workbook = null;
let credentials = [];
let notebooks = [];
let printers = [];
let currentFileName = 'gerenciador_ti.xlsx';
let isEditingIndex = null;
let editingType = null;
let selectedUsers = [];
let companyLogo = null;

// Elementos do DOM
const fileInput = document.getElementById('file-input');
const saveBtn = document.getElementById('save-btn');
const newBtn = document.getElementById('new-btn');
const credentialsList = document.getElementById('credentials-list');
const notebooksList = document.getElementById('notebooks-list');
const printersList = document.getElementById('printers-list');
const tabBtns = document.querySelectorAll('.tab-btn');
const subtabBtns = document.querySelectorAll('.subtab-btn');
const usersInput = document.getElementById('printer-users-input');
const userSuggestions = document.getElementById('user-suggestions');

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    updateUI();
});

// Configuração dos event listeners
function setupEventListeners() {
    // Botões principais
    newBtn.addEventListener('click', createNewWorkbook);
    fileInput.addEventListener('change', handleFileLoad);
    saveBtn.addEventListener('click', handleSave);
    
    // Formulário de credenciais
    document.getElementById('save-credential-btn').addEventListener('click', saveCredential);
    document.getElementById('clear-form-btn').addEventListener('click', clearCredentialForm);
    document.getElementById('url-input').addEventListener('blur', formatUrl);
    document.getElementById('generate-password-btn').addEventListener('click', openPasswordGenerator);
    document.getElementById('toggle-password').addEventListener('click', togglePasswordVisibility);
    document.getElementById('password-input').addEventListener('input', () => 
        updatePasswordStrength(document.getElementById('password-input').value));
    
    // Upload de logo
    document.getElementById('company-logo').addEventListener('change', handleLogoUpload);
    
    // Modal de senha
    document.getElementById('use-password').addEventListener('click', useGeneratedPassword);
    document.getElementById('close-modal').addEventListener('click', closePasswordModal);
    document.getElementById('regenerate-password').addEventListener('click', generateNewPassword);
    document.getElementById('copy-password').addEventListener('click', copyPasswordToClipboard);
    
    // Tabelas (event delegation)
    credentialsList.addEventListener('click', handleTableActions);
    notebooksList.addEventListener('click', handleTableActions);
    printersList.addEventListener('click', handleTableActions);
    
    // Abas
    tabBtns.forEach(btn => btn.addEventListener('click', switchTab));
    subtabBtns.forEach(btn => btn.addEventListener('click', switchSubtab));
    
    // Notebooks
    document.getElementById('save-notebook').addEventListener('click', saveNotebook);
    
    // Impressoras - Sistema de tags de usuários
    usersInput.addEventListener('input', handleUserInput);
    usersInput.addEventListener('keydown', handleUserKeyDown);
    usersInput.addEventListener('focus', showUserSuggestions);
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.tags-input-container')) {
            userSuggestions.style.display = 'none';
        }
    });
    
    // Impressoras
    document.getElementById('save-printer').addEventListener('click', savePrinter);
}

// Funções para navegação por abas
function switchTab(e) {
    const tabId = e.target.getAttribute('data-tab');
    
    tabBtns.forEach(btn => btn.classList.remove('active'));
    e.target.classList.add('active');
    
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${tabId}-tab`).classList.add('active');
}

function switchSubtab(e) {
    const subtabId = e.target.getAttribute('data-subtab');
    
    subtabBtns.forEach(btn => btn.classList.remove('active'));
    e.target.classList.add('active');
    
    document.querySelectorAll('.subtab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`${subtabId}-subtab`).classList.add('active');
}

// Funções do Workbook
function createNewWorkbook() {
    if ((credentials.length > 0 || notebooks.length > 0 || printers.length > 0) && 
        !confirm('Isso criará uma nova planilha em branco. Deseja continuar?')) {
        return;
    }

    workbook = XLSX.utils.book_new();
    
    // Criar abas vazias
    const companyWorksheet = XLSX.utils.aoa_to_sheet([["Nome da Empresa", ""], ["Endereço", ""], ["Contato", ""], ["Logo", ""]]);
    const credentialsWorksheet = XLSX.utils.aoa_to_sheet([["Serviço", "URL", "Usuário", "Senha", "MFA", "Notas", "Data Atualização"]]);
    const notebooksWorksheet = XLSX.utils.aoa_to_sheet([["Usuário", "Marca/Modelo", "Número Série", "Sistema Operacional", "Conexões", "Data Cadastro"]]);
    const printersWorksheet = XLSX.utils.aoa_to_sheet([["Modelo", "Número Série", "Localização", "Tipo Conexão", "Usuários", "Data Cadastro"]]);
    
    XLSX.utils.book_append_sheet(workbook, companyWorksheet, "Empresa");
    XLSX.utils.book_append_sheet(workbook, credentialsWorksheet, "Credenciais");
    XLSX.utils.book_append_sheet(workbook, notebooksWorksheet, "Notebooks");
    XLSX.utils.book_append_sheet(workbook, printersWorksheet, "Impressoras");
    
    // Resetar dados
    credentials = [];
    notebooks = [];
    printers = [];
    selectedUsers = [];
    isEditingIndex = null;
    editingType = null;
    currentFileName = 'gerenciador_ti.xlsx';
    companyLogo = null;
    
    // Limpar formulários
    document.getElementById('company-name').value = '';
    document.getElementById('company-address').value = '';
    document.getElementById('company-contact').value = '';
    document.getElementById('company-logo').value = '';
    const logoPreview = document.getElementById('logo-preview');
    logoPreview.src = '#';
    logoPreview.style.display = 'none';
    clearCredentialForm();
    clearNotebookForm();
    clearPrinterForm();
    
    // Atualizar UI
    renderCredentialsTable();
    renderNotebooksTable();
    renderPrintersTable();
    updateUI();
    
    showAlert('Nova planilha criada!', 'success');
}

async function handleFileLoad() {
    const file = fileInput.files[0];
    if (!file) return;

    try {
        const data = await readFileAsArrayBuffer(file);
        workbook = XLSX.read(data, { type: 'array' });
        currentFileName = file.name;
        
        // Carregar dados de cada aba se existir
        if (workbook.SheetNames.includes("Empresa")) {
            const companyWorksheet = workbook.Sheets["Empresa"];
            const companyData = XLSX.utils.sheet_to_json(companyWorksheet, { header: 1 });
            
            if (companyData.length > 0) {
                document.getElementById('company-name').value = companyData[0][1] || '';
                document.getElementById('company-address').value = companyData[1][1] || '';
                document.getElementById('company-contact').value = companyData[2][1] || '';
                
                // Carregar logo se existir
                if (companyData[3] && companyData[3][1]) {
                    companyLogo = companyData[3][1];
                    const logoPreview = document.getElementById('logo-preview');
                    logoPreview.src = companyLogo;
                    logoPreview.style.display = 'block';
                }
            }
        }
        
        if (workbook.SheetNames.includes("Credenciais")) {
            loadSheetData("Credenciais", 'credentials');
        }
        
        if (workbook.SheetNames.includes("Notebooks")) {
            loadSheetData("Notebooks", 'notebooks');
        }
        
        if (workbook.SheetNames.includes("Impressoras")) {
            loadSheetData("Impressoras", 'printers');
        }
        
        updateUI();
        showAlert('Planilha carregada com sucesso!', 'success');
    } catch (error) {
        console.error("Erro ao carregar arquivo:", error);
        showAlert('Erro ao carregar arquivo. Formato inválido.', 'error');
    }
}

function handleSave() {
    if (!workbook) {
        workbook = XLSX.utils.book_new();
    }

    try {
        // Aba Empresa
        const companyData = [
            ["Nome da Empresa", document.getElementById('company-name').value],
            ["Endereço", document.getElementById('company-address').value],
            ["Contato", document.getElementById('company-contact').value],
            ["Logo", companyLogo || ""],
            ["", ""],
            ["Data Exportação", new Date().toLocaleString()]
        ];
        
        const companyWorksheet = XLSX.utils.aoa_to_sheet(companyData);
        XLSX.utils.book_append_sheet(workbook, companyWorksheet, "Empresa");

        // Aba Credenciais
        const credentialsData = [["Serviço", "URL", "Usuário", "Senha", "MFA", "Notas", "Data Atualização"]];
        credentials.forEach(cred => {
            credentialsData.push([
                cred.service,
                cred.url,
                cred.user,
                cred.password,
                cred.mfa ? 'Sim' : 'Não',
                cred.notes,
                cred.date
            ]);
        });
        
        const credentialsWorksheet = XLSX.utils.aoa_to_sheet(credentialsData);
        XLSX.utils.book_append_sheet(workbook, credentialsWorksheet, "Credenciais");

        // Aba Notebooks
        const notebooksData = [["Usuário", "Marca/Modelo", "Número Série", "Sistema Operacional", "Conexões", "Data Cadastro"]];
        notebooks.forEach(note => {
            notebooksData.push([
                note.user,
                note.brand,
                note.serial,
                note.os,
                note.connections,
                note.date
            ]);
        });
        
        const notebooksWorksheet = XLSX.utils.aoa_to_sheet(notebooksData);
        XLSX.utils.book_append_sheet(workbook, notebooksWorksheet, "Notebooks");

        // Aba Impressoras
        const printersData = [["Modelo", "Número Série", "Localização", "Tipo Conexão", "Usuários", "Data Cadastro"]];
        printers.forEach(printer => {
            printersData.push([
                printer.name,
                printer.serial,
                printer.location,
                printer.connection,
                printer.users,
                printer.date
            ]);
        });
        
        const printersWorksheet = XLSX.utils.aoa_to_sheet(printersData);
        XLSX.utils.book_append_sheet(workbook, printersWorksheet, "Impressoras");

        // Salvar arquivo
        XLSX.writeFile(workbook, currentFileName);
        
        showAlert(`Planilha salva como ${currentFileName}`, 'success');
    } catch (error) {
        console.error("Erro ao salvar planilha:", error);
        showAlert('Erro ao salvar planilha: ' + error.message, 'error');
    }
}

// Funções para manipulação de dados
function loadSheetData(sheetName, type) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Pular cabeçalho
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        
        if (type === 'credentials' && row.length >= 4) {
            credentials.push({
                service: row[0] || '',
                url: row[1] || '',
                user: row[2] || '',
                password: row[3] || '',
                mfa: (row[4] && (row[4].toString().toLowerCase() === 'sim' || row[4] === true)) || false,
                notes: row[5] || '',
                date: row[6] || new Date().toLocaleString()
            });
        } else if (type === 'notebooks' && row.length >= 4) {
            notebooks.push({
                user: row[0] || '',
                brand: row[1] || '',
                serial: row[2] || '',
                os: row[3] || '',
                connections: row[4] || '',
                date: row[5] || new Date().toLocaleString()
            });
        } else if (type === 'printers' && row.length >= 4) {
            printers.push({
                name: row[0] || '',
                serial: row[1] || '',
                location: row[2] || '',
                connection: row[3] || '',
                users: row[4] || '',
                date: row[5] || new Date().toLocaleString()
            });
        }
    }
    
    // Atualizar a tabela correspondente
    if (type === 'credentials') {
        renderCredentialsTable();
    } else if (type === 'notebooks') {
        renderNotebooksTable();
    } else if (type === 'printers') {
        renderPrintersTable();
    }
}

// Funções para manipulação do logo da empresa
function handleLogoUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        companyLogo = event.target.result;
        const logoPreview = document.getElementById('logo-preview');
        logoPreview.src = companyLogo;
        logoPreview.style.display = 'block';
    };
    reader.readAsDataURL(file);
}

// Funções para Credenciais
function saveCredential() {
    const service = document.getElementById('service-input').value.trim();
    const url = document.getElementById('url-input').value.trim();
    const user = document.getElementById('user-input').value.trim();
    const password = document.getElementById('password-input').value.trim();
    
    if (!service || !url || !user || !password) {
        showAlert('Preencha todos os campos obrigatórios', 'warning');
        return;
    }
    
    const credentialData = {
        service,
        url,
        user,
        password,
        mfa: document.getElementById('mfa-checkbox').checked,
        notes: document.getElementById('notes-input').value.trim(),
        date: new Date().toLocaleString()
    };
    
    if (isEditingIndex !== null && editingType === 'credential') {
        credentials[isEditingIndex] = credentialData;
        showAlert('Credencial atualizada!', 'success');
    } else {
        credentials.push(credentialData);
        showAlert('Credencial adicionada!', 'success');
    }
    
    renderCredentialsTable();
    clearCredentialForm();
    updateUI();
}

function clearCredentialForm() {
    document.getElementById('service-input').value = '';
    document.getElementById('url-input').value = '';
    document.getElementById('user-input').value = '';
    document.getElementById('password-input').value = '';
    document.getElementById('notes-input').value = '';
    document.getElementById('mfa-checkbox').checked = false;
    isEditingIndex = null;
    editingType = null;
    document.getElementById('save-credential-btn').textContent = 'Salvar Credencial';
    updatePasswordStrength('');
}

function formatUrl() {
    const urlInput = document.getElementById('url-input');
    const url = urlInput.value.trim();
    if (url && !url.match(/^https?:\/\//) && !url.match(/^mailto:/) && !url.match(/^tel:/)) {
        urlInput.value = 'https://' + url;
    }
}

// Funções para Notebooks
function saveNotebook() {
    const user = document.getElementById('notebook-user').value.trim();
    const brand = document.getElementById('notebook-brand').value.trim();
    const serial = document.getElementById('notebook-serial').value.trim();
    const os = document.getElementById('notebook-os').value;
    const connections = Array.from(document.querySelectorAll('input[name="notebook-connection"]:checked'))
                            .map(el => el.value);
    
    if (!user || !brand || !os || connections.length === 0) {
        showAlert('Preencha todos os campos obrigatórios', 'warning');
        return;
    }
    
    const notebookData = {
        user,
        brand,
        serial,
        os,
        connections: connections.join(', '),
        date: new Date().toLocaleString()
    };
    
    if (isEditingIndex !== null && editingType === 'notebook') {
        notebooks[isEditingIndex] = notebookData;
        showAlert('Notebook atualizado!', 'success');
    } else {
        notebooks.push(notebookData);
        showAlert('Notebook adicionado!', 'success');
    }
    
    renderNotebooksTable();
    clearNotebookForm();
    updateUI();
}

function clearNotebookForm() {
    document.getElementById('notebook-user').value = '';
    document.getElementById('notebook-brand').value = '';
    document.getElementById('notebook-serial').value = '';
    document.getElementById('notebook-os').value = '';
    document.querySelectorAll('input[name="notebook-connection"]').forEach(checkbox => {
        checkbox.checked = false;
    });
    isEditingIndex = null;
    editingType = null;
    document.getElementById('save-notebook').textContent = 'Salvar Notebook';
}

// Funções para Impressoras (com sistema de tags de usuários)
function savePrinter() {
    const name = document.getElementById('printer-name').value.trim();
    const serial = document.getElementById('printer-serial').value.trim();
    const ip = document.getElementById('printer-ip').value.trim();
    const location = document.getElementById('printer-location').value.trim();
    const connection = document.querySelector('input[name="printer-connection"]:checked')?.value || '';
    
    if (!name || !location || !connection) {
        showAlert('Preencha todos os campos obrigatórios', 'warning');
        return;
    }
    
    const printerData = {
        name,
        serial,
        ip,
        location,
        connection,
        users: selectedUsers.join(', '),
        date: new Date().toLocaleString()
    };
    
    if (isEditingIndex !== null && editingType === 'printer') {
        printers[isEditingIndex] = printerData;
        showAlert('Impressora atualizada!', 'success');
    } else {
        printers.push(printerData);
        showAlert('Impressora adicionada!', 'success');
    }
    
    renderPrintersTable();
    clearPrinterForm();
    updateUI();
}

function clearPrinterForm() {
    document.getElementById('printer-name').value = '';
    document.getElementById('printer-serial').value = '';
    document.getElementById('printer-ip').value = '';
    document.getElementById('printer-location').value = '';
    document.querySelector('input[name="printer-connection"][value="Wi-Fi"]').checked = true;
    selectedUsers = [];
    renderUserTags();
    userSuggestions.style.display = 'none';
    isEditingIndex = null;
    editingType = null;
    document.getElementById('save-printer').textContent = 'Salvar Impressora';
}

// Sistema de tags de usuários para impressoras
function handleUserInput(e) {
    const input = e.target.value.toLowerCase();
    const allUsers = getAllUsers();
    
    if (input.length === 0) {
        showUserSuggestions();
        return;
    }
    
    const filteredUsers = allUsers.filter(user => 
        user.toLowerCase().includes(input) && !selectedUsers.includes(user)
    );
    
    renderUserSuggestions(filteredUsers);
}

function handleUserKeyDown(e) {
    if (e.key === 'Enter' && usersInput.value.trim()) {
        addUserTag(usersInput.value.trim());
        usersInput.value = '';
        e.preventDefault();
    } else if (e.key === 'Backspace' && usersInput.value === '') {
        selectedUsers.pop();
        renderUserTags();
    }
}

function showUserSuggestions() {
    const allUsers = getAllUsers().filter(user => !selectedUsers.includes(user));
    renderUserSuggestions(allUsers);
    userSuggestions.style.display = allUsers.length ? 'block' : 'none';
}

function renderUserSuggestions(users) {
    userSuggestions.innerHTML = '';
    
    if (users.length === 0) {
        userSuggestions.style.display = 'none';
        return;
    }
    
    users.forEach(user => {
        const div = document.createElement('div');
        div.className = 'user-suggestion';
        div.textContent = user;
        div.addEventListener('click', () => {
            addUserTag(user);
            usersInput.value = '';
            userSuggestions.style.display = 'none';
        });
        userSuggestions.appendChild(div);
    });
}

function addUserTag(user) {
    if (!user || selectedUsers.includes(user)) return;
    
    selectedUsers.push(user);
    renderUserTags();
    usersInput.value = '';
    userSuggestions.style.display = 'none';
    usersInput.focus();
}

function renderUserTags() {
    const container = document.getElementById('printer-users-tags');
    container.innerHTML = '';
    
    selectedUsers.forEach(user => {
        const tag = document.createElement('div');
        tag.className = 'tag';
        tag.innerHTML = `
            ${user}
            <span class="tag-remove" data-user="${user}">&times;</span>
        `;
        container.appendChild(tag);
    });
    
    document.querySelectorAll('.tag-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const userToRemove = e.target.getAttribute('data-user');
            selectedUsers = selectedUsers.filter(u => u !== userToRemove);
            renderUserTags();
            showUserSuggestions();
        });
    });
}

function getAllUsers() {
    const users = new Set();
    
    notebooks.forEach(notebook => {
        if (notebook.user) {
            users.add(notebook.user.trim());
        }
    });
    
    return Array.from(users).sort();
}

// Renderização das tabelas
function renderCredentialsTable() {
    credentialsList.innerHTML = '';
    
    if (credentials.length === 0) {
        credentialsList.innerHTML = '<tr><td colspan="7" class="text-center">Nenhuma credencial cadastrada</td></tr>';
        return;
    }
    
    credentials.forEach((cred, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${cred.service}</td>
            <td><a href="${cred.url}" target="_blank" rel="noopener">${cred.url}</a></td>
            <td>${cred.user}</td>
            <td class="password-cell">
                <span class="password-placeholder">••••••••</span>
                <button class="action-btn show" data-type="credential" data-index="${index}" title="Mostrar senha">
                    <i class="fas fa-eye"></i>
                </button>
            </td>
            <td>${cred.mfa ? '<i class="fas fa-check text-success"></i>' : '<i class="fas fa-times text-danger"></i>'}</td>
            <td>${cred.notes}</td>
            <td class="actions">
                <button class="action-btn edit" data-type="credential" data-index="${index}" title="Editar">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="action-btn delete" data-type="credential" data-index="${index}" title="Excluir">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        credentialsList.appendChild(row);
    });
}

function renderNotebooksTable() {
    notebooksList.innerHTML = '';
    
    if (notebooks.length === 0) {
        notebooksList.innerHTML = '<tr><td colspan="6" class="text-center">Nenhum notebook cadastrado</td></tr>';
        return;
    }
    
    notebooks.forEach((notebook, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${notebook.user}</td>
            <td>${notebook.brand}</td>
            <td>${notebook.serial || '-'}</td>
            <td>${notebook.os}</td>
            <td>${notebook.connections}</td>
            <td class="actions">
                <button class="action-btn edit" data-type="notebook" data-index="${index}" title="Editar">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="action-btn delete" data-type="notebook" data-index="${index}" title="Excluir">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        notebooksList.appendChild(row);
    });
}

function renderPrintersTable() {
    printersList.innerHTML = '';
    
    if (printers.length === 0) {
        printersList.innerHTML = '<tr><td colspan="6" class="text-center">Nenhuma impressora cadastrada</td></tr>';
        return;
    }
    
    printers.forEach((printer, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${printer.name}</td>
            <td>${printer.serial || '-'}</td>
            <td>${printer.location}</td>
            <td>${printer.connection}</td>
            <td>${printer.users || '-'}</td>
            <td class="actions">
                <button class="action-btn edit" data-type="printer" data-index="${index}" title="Editar">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="action-btn delete" data-type="printer" data-index="${index}" title="Excluir">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        printersList.appendChild(row);
    });
}

// Manipulação de ações nas tabelas
function handleTableActions(e) {
    const btn = e.target.closest('button');
    if (!btn) return;

    const type = btn.getAttribute('data-type');
    const index = parseInt(btn.getAttribute('data-index'));
    if (isNaN(index)) return;

    if (btn.classList.contains('show')) {
        togglePasswordInTable(index, type);
    } else if (btn.classList.contains('edit')) {
        editItem(index, type);
    } else if (btn.classList.contains('delete')) {
        deleteItem(index, type);
    }
}

function togglePasswordInTable(index, type) {
    if (type !== 'credential') return;
    
    const row = credentialsList.children[index];
    const placeholder = row.querySelector('.password-placeholder');
    const icon = row.querySelector('.show i');
    
    if (placeholder.textContent === '••••••••') {
        placeholder.textContent = credentials[index].password;
        icon.classList.replace('fa-eye', 'fa-eye-slash');
    } else {
        placeholder.textContent = '••••••••';
        icon.classList.replace('fa-eye-slash', 'fa-eye');
    }
}

function editItem(index, type) {
    isEditingIndex = index;
    editingType = type;
    
    if (type === 'credential') {
        const cred = credentials[index];
        document.getElementById('service-input').value = cred.service;
        document.getElementById('url-input').value = cred.url;
        document.getElementById('user-input').value = cred.user;
        document.getElementById('password-input').value = cred.password;
        document.getElementById('mfa-checkbox').checked = cred.mfa;
        document.getElementById('notes-input').value = cred.notes;
        
        document.getElementById('save-credential-btn').textContent = 'Atualizar Credencial';
        updatePasswordStrength(cred.password);
        switchTab({ target: document.querySelector('[data-tab="credentials"]') });
    } else if (type === 'notebook') {
        const notebook = notebooks[index];
        document.getElementById('notebook-user').value = notebook.user;
        document.getElementById('notebook-brand').value = notebook.brand;
        document.getElementById('notebook-serial').value = notebook.serial;
        document.getElementById('notebook-os').value = notebook.os;
        
        // Marcar conexões
        const connections = notebook.connections.split(', ');
        document.querySelectorAll('input[name="notebook-connection"]').forEach(checkbox => {
            checkbox.checked = connections.includes(checkbox.value);
        });
        
        document.getElementById('save-notebook').textContent = 'Atualizar Notebook';
        switchTab({ target: document.querySelector('[data-tab="equipment"]') });
        switchSubtab({ target: document.querySelector('[data-subtab="notebooks"]') });
    } else if (type === 'printer') {
        const printer = printers[index];
        document.getElementById('printer-name').value = printer.name;
        document.getElementById('printer-serial').value = printer.serial;
        document.getElementById('printer-ip').value = printer.ip;
        document.getElementById('printer-location').value = printer.location;
        document.querySelector(`input[name="printer-connection"][value="${printer.connection}"]`).checked = true;
        
        // Carregar usuários
        selectedUsers = printer.users ? printer.users.split(',').map(u => u.trim()) : [];
        renderUserTags();
        
        document.getElementById('save-printer').textContent = 'Atualizar Impressora';
        switchTab({ target: document.querySelector('[data-tab="equipment"]') });
        switchSubtab({ target: document.querySelector('[data-subtab="printers"]') });
    }
}

function deleteItem(index, type) {
    let itemName = '';
    let successMessage = '';
    
    if (type === 'credential') {
        itemName = credentials[index].service;
        successMessage = 'Credencial excluída!';
    } else if (type === 'notebook') {
        itemName = notebooks[index].brand + ' (' + notebooks[index].user + ')';
        successMessage = 'Notebook excluído!';
    } else if (type === 'printer') {
        itemName = printers[index].name + ' (' + printers[index].location + ')';
        successMessage = 'Impressora excluída!';
    }
    
    if (confirm(`Tem certeza que deseja excluir "${itemName}"?`)) {
        if (type === 'credential') {
            credentials.splice(index, 1);
            renderCredentialsTable();
            if (editingType === 'credential' && isEditingIndex === index) {
                clearCredentialForm();
            }
        } else if (type === 'notebook') {
            notebooks.splice(index, 1);
            renderNotebooksTable();
            if (editingType === 'notebook' && isEditingIndex === index) {
                clearNotebookForm();
            }
        } else if (type === 'printer') {
            printers.splice(index, 1);
            renderPrintersTable();
            if (editingType === 'printer' && isEditingIndex === index) {
                clearPrinterForm();
            }
        }
        
        showAlert(successMessage, 'success');
        updateUI();
    }
}

// Gerador de senhas
function openPasswordGenerator() {
    document.getElementById('password-modal').style.display = 'flex';
    generateNewPassword();
}

function closePasswordModal() {
    document.getElementById('password-modal').style.display = 'none';
}

function generateNewPassword() {
    const length = parseInt(document.querySelector('input[name="length"]:checked').value);
    const uppercase = document.querySelector('input[name="uppercase"]').checked;
    const lowercase = document.querySelector('input[name="lowercase"]').checked;
    const numbers = document.querySelector('input[name="numbers"]').checked;
    const symbols = document.querySelector('input[name="symbols"]').checked;
    
    document.getElementById('generated-password').value = 
        generateRandomPassword(length, uppercase, lowercase, numbers, symbols);
    updatePasswordStrength(
        document.getElementById('generated-password').value, 
        'modal-strength-bar', 
        'modal-strength-text'
    );
}

function generateRandomPassword(length = 12, uppercase = true, lowercase = true, numbers = true, symbols = true) {
    let chars = '';
    if (uppercase) chars += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    if (lowercase) chars += 'abcdefghijklmnopqrstuvwxyz';
    if (numbers) chars += '0123456789';
    if (symbols) chars += '!@#$%^&*()_+-=[]{}|;:,.<>?';
    
    if (!chars) chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    
    let password = '';
    for (let i = 0; i < length; i++) {
        password += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    
    return password;
}

function useGeneratedPassword() {
    document.getElementById('password-input').value = 
        document.getElementById('generated-password').value;
    updatePasswordStrength(document.getElementById('password-input').value);
    closePasswordModal();
}

function copyPasswordToClipboard() {
    const generatedPassword = document.getElementById('generated-password');
    generatedPassword.select();
    document.execCommand('copy');
    showAlert('Senha copiada!', 'success');
}

function togglePasswordVisibility() {
    const passwordInput = document.getElementById('password-input');
    const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
    passwordInput.setAttribute('type', type);
    document.getElementById('toggle-password').innerHTML = 
        type === 'password' ? '<i class="fas fa-eye"></i>' : '<i class="fas fa-eye-slash"></i>';
}

function updatePasswordStrength(password = '', barId = 'strength-bar', textId = 'strength-text') {
    const strengthBar = document.getElementById(barId);
    const strengthText = document.getElementById(textId);
    
    if (!password) {
        strengthBar.style.width = '0%';
        strengthBar.style.backgroundColor = '#e9ecef';
        strengthText.textContent = 'Força: ';
        return;
    }
    
    let strength = 0;
    
    // Comprimento
    if (password.length >= 12) strength += 2;
    else if (password.length >= 8) strength += 1;
    
    // Complexidade
    if (/[A-Z]/.test(password)) strength += 1;
    if (/[a-z]/.test(password)) strength += 1;
    if (/[0-9]/.test(password)) strength += 1;
    if (/[^A-Za-z0-9]/.test(password)) strength += 1;
    
    // Atualizar UI
    let width, color, text;
    if (strength <= 2) {
        width = '25%'; color = '#dc3545'; text = 'Fraca';
    } else if (strength <= 4) {
        width = '50%'; color = '#ffc107'; text = 'Média';
    } else if (strength <= 6) {
        width = '75%'; color = '#28a745'; text = 'Forte';
    } else {
        width = '100%'; color = '#20c997'; text = 'Excelente';
    }
    
    strengthBar.style.width = width;
    strengthBar.style.backgroundColor = color;
    strengthText.textContent = `Força: ${text}`;
}

// Funções auxiliares
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function updateUI() {
    const hasCompanyInfo = document.getElementById('company-name').value.trim() !== '';
    document.getElementById('save-btn').enabled = !workbook || !hasCompanyInfo;
}

function showAlert(message, type = 'info') {
    const alertBox = document.createElement('div');
    alertBox.className = `alert alert-${type}`;
    alertBox.innerHTML = `
        <span>${message}</span>
        <button class="close-alert">&times;</button>
    `;
    
    document.body.appendChild(alertBox);
    
    // Fechar automaticamente após 3 segundos
    setTimeout(() => {
        alertBox.classList.add('fade-out');
        setTimeout(() => alertBox.remove(), 300);
    }, 3000);
    
    // Fechar manualmente
    alertBox.querySelector('.close-alert').addEventListener('click', () => {
        alertBox.remove();
    });
}
