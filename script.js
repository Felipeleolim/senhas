// Variáveis globais
let workbook = null;
let credentials = [];
let notebooks = [];
let printers = [];
let currentFileName = 'gerenciador_ti.xlsx';
let isEditingIndex = null;
let editingType = null;
let selectedUsers = [];
let companyDataSaved = false;
let companyLogo = null;
let currentNotebookIndex = null;

// Elementos do DOM
const fileInput = document.getElementById('file-input');
const saveBtn = document.getElementById('save-btn');
const newBtn = document.getElementById('new-btn');
const credentialsList = document.getElementById('credentials-list');
const equipmentGrid = document.getElementById('equipment-grid');
const tabBtns = document.querySelectorAll('.tab-btn');
const subtabBtns = document.querySelectorAll('.subtab-btn');
const usersInput = document.getElementById('printer-users-input');
const userSuggestions = document.getElementById('user-suggestions');

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    updateUI();
    renderEquipmentGrid();
});

// Configuração dos event listeners
function setupEventListeners() {
    // Botões principais
    newBtn.addEventListener('click', createNewWorkbook);
    fileInput.addEventListener('change', handleFileLoad);
    saveBtn.addEventListener('click', handleSave);
    
    // Formulário de empresa
    document.getElementById('save-company-btn').addEventListener('click', saveCompanyData);
    document.getElementById('company-logo').addEventListener('change', handleLogoUpload);
    
    // Formulário de credenciais
    document.getElementById('save-credential-btn').addEventListener('click', saveCredential);
    document.getElementById('clear-form-btn').addEventListener('click', clearCredentialForm);
    document.getElementById('url-input').addEventListener('blur', formatUrl);
    document.getElementById('generate-password-btn').addEventListener('click', openPasswordGenerator);
    document.getElementById('toggle-password').addEventListener('click', togglePasswordVisibility);
    document.getElementById('password-input').addEventListener('input', () => 
        updatePasswordStrength(document.getElementById('password-input').value));
    
    // Modal de senha
    document.getElementById('use-password').addEventListener('click', useGeneratedPassword);
    document.getElementById('close-modal').addEventListener('click', closePasswordModal);
    document.getElementById('regenerate-password').addEventListener('click', generateNewPassword);
    document.getElementById('copy-password').addEventListener('click', copyPasswordToClipboard);
    
    // Tabelas (event delegation)
    document.addEventListener('click', handleTableActions);
    
    // Abas
    tabBtns.forEach(btn => btn.addEventListener('click', switchTab));
    subtabBtns.forEach(btn => btn.addEventListener('click', switchSubtab));
    
    // Notebooks
    document.getElementById('save-notebook').addEventListener('click', saveNotebook);
    document.getElementById('cancel-notebook').addEventListener('click', cancelEquipmentForm);
    document.getElementById('notebook-image').addEventListener('change', handleNotebookImageUpload);
    
    // Impressoras
    document.getElementById('save-printer').addEventListener('click', savePrinter);
    document.getElementById('cancel-printer').addEventListener('click', cancelEquipmentForm);
    document.getElementById('printer-image').addEventListener('change', handlePrinterImageUpload);
    
    // Impressoras - Sistema de tags de usuários
    usersInput.addEventListener('input', handleUserInput);
    usersInput.addEventListener('keydown', handleUserKeyDown);
    usersInput.addEventListener('focus', showUserSuggestions);
    
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.tags-input-container')) {
            userSuggestions.style.display = 'none';
        }
    });
    
    // Equipamentos
    document.getElementById('add-equipment-btn').addEventListener('click', showEquipmentForm);
    
    // Notebook modal
    document.getElementById('add-note-btn').addEventListener('click', toggleNoteField);
    document.getElementById('save-note-btn').addEventListener('click', saveNote);
    document.getElementById('cancel-note-btn').addEventListener('click', toggleNoteField);
    document.getElementById('close-notebook-modal').addEventListener('click', () => {
        document.getElementById('notebook-modal').style.display = 'none';
    });
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

// Funções para gerenciamento da empresa
function saveCompanyData() {
    const companyName = document.getElementById('company-name').value.trim();
    
    if (!companyName) {
        showAlert('O nome da empresa é obrigatório', 'warning');
        document.getElementById('company-name').focus();
        return;
    }

    document.getElementById('company-name').disabled = true;
    document.getElementById('company-address').disabled = true;
    document.getElementById('company-contact').disabled = true;
    document.getElementById('company-logo').disabled = true;

    companyDataSaved = true;
    updateUI();
    showAlert('Informações da empresa salvas com sucesso!', 'success');
}

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

// Funções para manipulação de imagens
function handleNotebookImageUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const notebookPreview = document.getElementById('notebook-preview');
        notebookPreview.src = event.target.result;
        notebookPreview.style.display = 'block';
    };
    reader.readAsDataURL(file);
}

function handlePrinterImageUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const printerPreview = document.getElementById('printer-preview');
        printerPreview.src = event.target.result;
        printerPreview.style.display = 'block';
    };
    reader.readAsDataURL(file);
}

// Funções do Workbook
function createNewWorkbook() {
    if ((credentials.length > 0 || notebooks.length > 0 || printers.length > 0) && 
        !confirm('Isso criará uma nova planilha em branco. Deseja continuar?')) {
        return;
    }

    workbook = XLSX.utils.book_new();
    
    const companyWorksheet = XLSX.utils.aoa_to_sheet([["Nome da Empresa", ""], ["Endereço", ""], ["Contato", ""], ["Logo", ""]]);
    const credentialsWorksheet = XLSX.utils.aoa_to_sheet([["Serviço", "URL", "Usuário", "Senha", "MFA", "Notas", "Data Atualização"]]);
    const notebooksWorksheet = XLSX.utils.aoa_to_sheet([["Usuário", "Marca/Modelo", "Número Série", "Sistema Operacional", "Conexões", "Notas", "Imagem", "Data Cadastro"]]);
    const printersWorksheet = XLSX.utils.aoa_to_sheet([["Modelo", "Número Série", "Localização", "Tipo Conexão", "Usuários", "Imagem", "Data Cadastro"]]);
    
    XLSX.utils.book_append_sheet(workbook, companyWorksheet, "Empresa");
    XLSX.utils.book_append_sheet(workbook, credentialsWorksheet, "Credenciais");
    XLSX.utils.book_append_sheet(workbook, notebooksWorksheet, "Notebooks");
    XLSX.utils.book_append_sheet(workbook, printersWorksheet, "Impressoras");
    
    credentials = [];
    notebooks = [];
    printers = [];
    selectedUsers = [];
    isEditingIndex = null;
    editingType = null;
    companyLogo = null;
    companyDataSaved = false;
    currentFileName = 'gerenciador_ti.xlsx';
    currentNotebookIndex = null;
    
    document.getElementById('company-name').value = '';
    document.getElementById('company-address').value = '';
    document.getElementById('company-contact').value = '';
    document.getElementById('company-logo').value = '';
    document.getElementById('logo-preview').style.display = 'none';
    
    document.getElementById('company-name').disabled = false;
    document.getElementById('company-address').disabled = false;
    document.getElementById('company-contact').disabled = false;
    document.getElementById('company-logo').disabled = false;
    
    clearCredentialForm();
    clearNotebookForm();
    clearPrinterForm();
    
    renderCredentialsTable();
    renderEquipmentGrid();
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
        
        if (workbook.SheetNames.includes("Empresa")) {
            const companyWorksheet = workbook.Sheets["Empresa"];
            const companyData = XLSX.utils.sheet_to_json(companyWorksheet, { header: 1 });
            
            if (companyData.length > 0) {
                document.getElementById('company-name').value = companyData[0][1] || '';
                document.getElementById('company-address').value = companyData[1][1] || '';
                document.getElementById('company-contact').value = companyData[2][1] || '';
                
                if (companyData[3] && companyData[3][1]) {
                    companyLogo = companyData[3][1];
                    document.getElementById('logo-preview').src = companyLogo;
                    document.getElementById('logo-preview').style.display = 'block';
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
        
        if (document.getElementById('company-name').value) {
            document.getElementById('company-name').disabled = true;
            document.getElementById('company-address').disabled = true;
            document.getElementById('company-contact').disabled = true;
            document.getElementById('company-logo').disabled = true;
            companyDataSaved = true;
        } else {
            document.getElementById('company-name').disabled = false;
            document.getElementById('company-address').disabled = false;
            document.getElementById('company-contact').disabled = false;
            document.getElementById('company-logo').disabled = false;
            companyDataSaved = false;
        }
        
        updateUI();
        renderEquipmentGrid();
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
        // Salvar dados da empresa
        if (workbook.SheetNames.includes("Empresa")) {
            workbook.SheetNames.splice(workbook.SheetNames.indexOf("Empresa"), 1);
            delete workbook.Sheets["Empresa"];
        }
        
        const companyData = [
            ["Nome da Empresa", document.getElementById('company-name').value],
            ["Endereço", document.getElementById('company-address').value],
            ["Contato", document.getElementById('company-contact').value],
            ["Logo", companyLogo || ""],
            ["", ""],
            ["Data Exportação", new Date().toLocaleString()]
        ];
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(companyData), "Empresa");

        // Salvar credenciais
        if (workbook.SheetNames.includes("Credenciais")) {
            workbook.SheetNames.splice(workbook.SheetNames.indexOf("Credenciais"), 1);
            delete workbook.Sheets["Credenciais"];
        }
        
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
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(credentialsData), "Credenciais");

        // Salvar notebooks
        if (workbook.SheetNames.includes("Notebooks")) {
            workbook.SheetNames.splice(workbook.SheetNames.indexOf("Notebooks"), 1);
            delete workbook.Sheets["Notebooks"];
        }
        
        const notebooksData = [["Usuário", "Marca/Modelo", "Número Série", "Sistema Operacional", "Conexões", "Notas", "Imagem", "Data Cadastro"]];
        notebooks.forEach(note => {
            const notesText = note.notes && note.notes.length > 0 ? 
                note.notes.map(n => `${n.date}: ${n.text}`).join('\n\n') : 
                '';
            
            notebooksData.push([
                note.user,
                note.brand,
                note.serial,
                note.os,
                note.connections,
                notesText,
                note.image ? "Sim" : "Não",
                note.date
            ]);
        });
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(notebooksData), "Notebooks");

        // Salvar impressoras
        if (workbook.SheetNames.includes("Impressoras")) {
            workbook.SheetNames.splice(workbook.SheetNames.indexOf("Impressoras"), 1);
            delete workbook.Sheets["Impressoras"];
        }
        
        const printersData = [["Modelo", "Número Série", "Localização", "Tipo Conexão", "Usuários", "Imagem", "Data Cadastro"]];
        printers.forEach(printer => {
            printersData.push([
                printer.name,
                printer.serial,
                printer.location,
                printer.connection,
                printer.users,
                printer.image ? "Sim" : "Não",
                printer.date
            ]);
        });
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(printersData), "Impressoras");

        XLSX.writeFile(workbook, currentFileName);
        showAlert(`Planilha salva como ${currentFileName}`, 'success');
    } catch (error) {
        console.error("Erro ao salvar planilha:", error);
        showAlert('Erro ao salvar planilha: ' + error.message, 'error');
    }
}

// Funções para equipamentos
function showEquipmentForm() {
    document.getElementById('equipment-form-container').style.display = 'block';
    document.getElementById('add-equipment-btn').style.display = 'none';
    document.querySelector('.subtab-btn[data-subtab="notebooks"]').click();
    clearNotebookForm();
    clearPrinterForm();
}

function cancelEquipmentForm() {
    document.getElementById('equipment-form-container').style.display = 'none';
    document.getElementById('add-equipment-btn').style.display = 'flex';
    isEditingIndex = null;
    editingType = null;
}

// Funções para notebooks
function saveNotebook() {
    const user = document.getElementById('notebook-user').value.trim();
    const brand = document.getElementById('notebook-brand').value.trim();
    const serial = document.getElementById('notebook-serial').value.trim();
    const os = document.getElementById('notebook-os').value.trim();
    const notes = document.getElementById('notebook-notes').value.trim();
    const imageInput = document.getElementById('notebook-image');
    
    const connectionCheckboxes = document.querySelectorAll('input[name="notebook-connection"]:checked');
    const connections = Array.from(connectionCheckboxes).map(cb => cb.value).join(', ');
    
    if (!user || !brand || !os) {
        showAlert('Usuário, marca e sistema operacional são obrigatórios!', 'warning');
        return;
    }
    
    let image = null;
    if (imageInput.files.length > 0) {
        const reader = new FileReader();
        reader.onload = function(e) {
            image = e.target.result;
            saveNotebookData(user, brand, serial, os, connections, notes, image);
        };
        reader.readAsDataURL(imageInput.files[0]);
    } else {
        saveNotebookData(user, brand, serial, os, connections, notes, null);
    }
}

function saveNotebookData(user, brand, serial, os, connections, notes, image) {
    const existingNotebook = isEditingIndex !== null && editingType === 'notebook' ? notebooks[isEditingIndex] : null;
    
    const notebook = {
        user,
        brand,
        serial,
        os,
        connections,
        notes: existingNotebook && existingNotebook.notes ? existingNotebook.notes : (notes ? [{ text: notes, date: new Date().toLocaleString() }] : []),
        image,
        date: new Date().toLocaleDateString()
    };
    
    if (isEditingIndex !== null && editingType === 'notebook') {
        notebooks[isEditingIndex] = notebook;
        showAlert('Notebook atualizado com sucesso!', 'success');
    } else {
        notebooks.push(notebook);
        showAlert('Notebook adicionado com sucesso!', 'success');
    }
    
    renderEquipmentGrid();
    cancelEquipmentForm();
    updateUI();
}

// Funções para impressoras
function savePrinter() {
    const name = document.getElementById('printer-name').value.trim();
    const serial = document.getElementById('printer-serial').value.trim();
    const location = document.getElementById('printer-location').value.trim();
    const ip = document.getElementById('printer-ip').value.trim();
    const imageInput = document.getElementById('printer-image');
    
    const connection = document.querySelector('input[name="printer-connection"]:checked').value;
    const users = selectedUsers.join(', ');
    
    if (!name || !location) {
        showAlert('Modelo e localização são obrigatórios!', 'warning');
        return;
    }
    
    let image = null;
    if (imageInput.files.length > 0) {
        const reader = new FileReader();
        reader.onload = function(e) {
            image = e.target.result;
            savePrinterData(name, serial, ip, location, connection, users, image);
        };
        reader.readAsDataURL(imageInput.files[0]);
    } else {
        savePrinterData(name, serial, ip, location, connection, users, null);
    }
}

function savePrinterData(name, serial, ip, location, connection, users, image) {
    const printer = {
        name,
        serial,
        ip,
        location,
        connection,
        users,
        image,
        date: new Date().toLocaleDateString()
    };
    
    if (isEditingIndex !== null && editingType === 'printer') {
        printers[isEditingIndex] = printer;
        showAlert('Impressora atualizada com sucesso!', 'success');
    } else {
        printers.push(printer);
        showAlert('Impressora adicionada com sucesso!', 'success');
    }
    
    renderEquipmentGrid();
    cancelEquipmentForm();
}

// Função para renderizar os equipamentos em formato de grid
function renderEquipmentGrid() {
    const grid = document.getElementById('equipment-grid');
    grid.innerHTML = '';

    // Combinar notebooks e impressoras
    const allEquipment = [
        ...notebooks.map(item => ({ ...item, type: 'notebook' })),
        ...printers.map(item => ({ ...item, type: 'printer' }))
    ];

    // Ordenar por data (mais recente primeiro)
    allEquipment.sort((a, b) => new Date(b.date) - new Date(a.date));

    if (allEquipment.length === 0) {
        grid.innerHTML = `
            <div class="no-equipment" style="grid-column: 1/-1; text-align: center; padding: 40px;">
                <i class="fas fa-laptop" style="font-size: 3rem; color: var(--gray-light); margin-bottom: 15px;"></i>
                <h3 style="color: var(--gray-dark);">Nenhum equipamento cadastrado</h3>
                <p>Clique em "Adicionar Equipamento" para começar</p>
            </div>
        `;
        return;
    }

    allEquipment.forEach((equip, index) => {
        const card = document.createElement('div');
        card.className = 'equipment-card';
        card.innerHTML = `
            <div class="equipment-image" style="position: relative;">
                ${equip.image ? 
                    `<img src="${equip.image}" alt="${equip.type === 'notebook' ? equip.brand : equip.name}" style="width: 100%; height: 180px; object-fit: cover;">` : 
                    `<i class="fas ${equip.type === 'notebook' ? 'fa-laptop' : 'fa-print'}" style="font-size: 3rem; color: var(--gray-dark);"></i>`
                }
                <span class="equipment-type">${equip.type === 'notebook' ? 'Notebook' : 'Impressora'}</span>
            </div>
            <div class="equipment-info">
                <div class="equipment-name">${equip.type === 'notebook' ? equip.brand : equip.name}</div>
                <div class="equipment-user"><i class="fas fa-user"></i> ${equip.user || equip.location || 'Sem responsável'}</div>
                <div class="equipment-date"><i class="far fa-calendar-alt"></i> ${equip.date}</div>
                <div class="equipment-actions">
                    <button class="btn-view-equipment" data-index="${equip.type === 'notebook' ? notebooks.indexOf(equip) : printers.indexOf(equip)}" data-type="${equip.type}">
                        <i class="fas fa-eye"></i> Ver
                    </button>
                    <button class="btn-edit-equipment" data-index="${equip.type === 'notebook' ? notebooks.indexOf(equip) : printers.indexOf(equip)}" data-type="${equip.type}">
                        <i class="fas fa-edit"></i> Editar
                    </button>
                </div>
            </div>
        `;
        grid.appendChild(card);
    });
}

// Funções para visualização de notebooks
function viewNotebookDetails(index) {
    currentNotebookIndex = index;
    const notebook = notebooks[index];
    
    document.getElementById('modal-notebook-user').textContent = notebook.user || 'Não informado';
    document.getElementById('modal-notebook-brand').textContent = notebook.brand || 'Não informado';
    document.getElementById('modal-notebook-serial').textContent = notebook.serial || 'Não informado';
    document.getElementById('modal-notebook-os').textContent = notebook.os || 'Não informado';
    document.getElementById('modal-notebook-connections').textContent = notebook.connections || 'Não informado';
    document.getElementById('modal-notebook-date').textContent = notebook.date || 'Data não registrada';
    
    const notebookImage = document.getElementById('modal-notebook-image');
    if (notebook.image) {
        notebookImage.src = notebook.image;
        notebookImage.style.display = 'block';
    } else {
        notebookImage.style.display = 'none';
    }
    
    loadNotebookNotes();
    document.getElementById('notebook-modal').style.display = 'flex';
}

function loadNotebookNotes() {
    const notesContainer = document.getElementById('modal-notebook-notes');
    notesContainer.innerHTML = '';
    
    const notebook = notebooks[currentNotebookIndex];
    
    if (!notebook.notes || notebook.notes.length === 0) {
        notesContainer.innerHTML = '<p class="no-notes">Nenhum atendimento registrado ainda.</p>';
        return;
    }
    
    const sortedNotes = [...notebook.notes].sort((a, b) => {
        if (typeof a === 'string' || typeof b === 'string') return 0;
        return new Date(b.date) - new Date(a.date);
    });
    
    sortedNotes.forEach(note => {
        const noteElement = document.createElement('div');
        noteElement.className = 'note-item';
        
        if (typeof note === 'string') {
            noteElement.innerHTML = `
                <div class="note-content">${note}</div>
                <div class="note-date">
                    <i class="fas fa-clock"></i>
                    Data não registrada
                </div>
            `;
        } else {
            noteElement.innerHTML = `
                <div class="note-content">${note.text}</div>
                <div class="note-date">
                    <i class="fas fa-clock"></i>
                    ${note.date}
                </div>
            `;
        }
        
        notesContainer.appendChild(noteElement);
    });
}

function toggleNoteField() {
    const container = document.getElementById('new-note-container');
    container.style.display = container.style.display === 'none' ? 'block' : 'none';
    
    if (container.style.display === 'block') {
        document.getElementById('new-note-text').focus();
    }
}

function saveNote() {
    const noteText = document.getElementById('new-note-text').value.trim();
    if (!noteText) {
        showAlert('Por favor, descreva o atendimento antes de salvar', 'warning');
        return;
    }
    
    const newNote = {
        text: noteText,
        date: new Date().toLocaleString('pt-BR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        })
    };
    
    if (!notebooks[currentNotebookIndex].notes) {
        notebooks[currentNotebookIndex].notes = [];
    }
    
    notebooks[currentNotebookIndex].notes.push(newNote);
    
    document.getElementById('new-note-text').value = '';
    document.getElementById('new-note-container').style.display = 'none';
    
    loadNotebookNotes();
    showAlert('Atendimento registrado com sucesso!', 'success');
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
    document.getElementById('save-btn').disabled = !workbook && !companyDataSaved;
}

function showAlert(message, type = 'info') {
    const alertBox = document.createElement('div');
    alertBox.className = `alert alert-${type}`;
    alertBox.innerHTML = `
        <span>${message}</span>
        <button class="close-alert">&times;</button>
    `;
    
    document.body.appendChild(alertBox);
    
    setTimeout(() => {
        alertBox.classList.add('fade-out');
        setTimeout(() => alertBox.remove(), 300);
    }, 3000);
    
    alertBox.querySelector('.close-alert').addEventListener('click', () => {
        alertBox.remove();
    });
}

// Funções para renderizar tabelas
function renderCredentialsTable() {
    const credentialsList = document.getElementById('credentials-list');
    credentialsList.innerHTML = '';
    credentials.forEach((cred, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${cred.service}</td>
            <td>${cred.url}</td>
            <td>${cred.user}</td>
            <td>••••••••</td>
            <td>${cred.mfa ? 'Sim' : 'Não'}</td>
            <td>${cred.notes}</td>
            <td>
                <button class="btn-view" data-index="${index}" data-type="credential">👁️</button>
                <button class="btn-edit" data-index="${index}" data-type="credential">✏️</button>
                <button class="btn-delete" data-index="${index}" data-type="credential">🗑️</button>
            </td>
        `;
        credentialsList.appendChild(row);
    });
}

// Funções para formulários
function clearCredentialForm() {
    document.getElementById('service-input').value = '';
    document.getElementById('url-input').value = '';
    document.getElementById('user-input').value = '';
    document.getElementById('password-input').value = '';
    document.getElementById('mfa-checkbox').checked = false;
    document.getElementById('notes-input').value = '';
    isEditingIndex = null;
    editingType = null;
}

function clearNotebookForm() {
    document.getElementById('notebook-user').value = '';
    document.getElementById('notebook-brand').value = '';
    document.getElementById('notebook-serial').value = '';
    document.getElementById('notebook-os').value = '';
    document.getElementById('notebook-notes').value = '';
    document.getElementById('notebook-image').value = '';
    document.getElementById('notebook-preview').style.display = 'none';
    document.querySelectorAll('input[name="notebook-connection"]').forEach(cb => cb.checked = false);
    isEditingIndex = null;
    editingType = null;
}

function clearPrinterForm() {
    document.getElementById('printer-name').value = '';
    document.getElementById('printer-serial').value = '';
    document.getElementById('printer-ip').value = '';
    document.getElementById('printer-location').value = '';
    document.getElementById('printer-image').value = '';
    document.getElementById('printer-preview').style.display = 'none';
    document.querySelector('input[name="printer-connection"][value="Wi-Fi"]').checked = true;
    document.getElementById('printer-users-input').value = '';
    selectedUsers = [];
    updateUserTags();
    isEditingIndex = null;
    editingType = null;
}

// Funções para manipulação de dados
function handleTableActions(e) {
    if (e.target.classList.contains('btn-view') || e.target.closest('.btn-view')) {
        const btn = e.target.classList.contains('btn-view') ? e.target : e.target.closest('.btn-view');
        const index = btn.getAttribute('data-index');
        const type = btn.getAttribute('data-type');
        viewItem(index, type);
    } else if (e.target.classList.contains('btn-view-equipment') || e.target.closest('.btn-view-equipment')) {
        const btn = e.target.classList.contains('btn-view-equipment') ? e.target : e.target.closest('.btn-view-equipment');
        const index = btn.getAttribute('data-index');
        const type = btn.getAttribute('data-type');
        
        if (type === 'notebook') {
            viewNotebookDetails(index);
        } else {
            const printer = printers[index];
            alert(`Impressora: ${printer.name}\nLocal: ${printer.location}\nConexão: ${printer.connection}`);
        }
    } else if (e.target.classList.contains('btn-edit-equipment') || e.target.closest('.btn-edit-equipment')) {
        const btn = e.target.classList.contains('btn-edit-equipment') ? e.target : e.target.closest('.btn-edit-equipment');
        const index = btn.getAttribute('data-index');
        const type = btn.getAttribute('data-type');
        
        document.getElementById('equipment-form-container').style.display = 'block';
        document.getElementById('add-equipment-btn').style.display = 'none';
        
        if (type === 'notebook') {
            document.querySelector('.subtab-btn[data-subtab="notebooks"]').click();
            editItem(index, 'notebook');
        } else {
            document.querySelector('.subtab-btn[data-subtab="printers"]').click();
            editItem(index, 'printer');
        }
    } else if (e.target.classList.contains('btn-edit') || e.target.closest('.btn-edit')) {
        const btn = e.target.classList.contains('btn-edit') ? e.target : e.target.closest('.btn-edit');
        const index = btn.getAttribute('data-index');
        const type = btn.getAttribute('data-type');
        editItem(index, type);
    } else if (e.target.classList.contains('btn-delete') || e.target.closest('.btn-delete')) {
        const btn = e.target.classList.contains('btn-delete') ? e.target : e.target.closest('.btn-delete');
        const index = btn.getAttribute('data-index');
        const type = btn.getAttribute('data-type');
        deleteItem(index, type);
    }
}

function viewItem(index, type) {
    if (type === 'credential') {
        const cred = credentials[index];
        alert(`Detalhes da Credencial:\nServiço: ${cred.service}\nURL: ${cred.url}\nUsuário: ${cred.user}\nSenha: ${cred.password}\nMFA: ${cred.mfa ? 'Sim' : 'Não'}\nNotas: ${cred.notes}`);
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
        
        document.querySelector('.tab-btn[data-tab="credentials"]').click();
    } else if (type === 'notebook') {
        const note = notebooks[index];
        document.getElementById('notebook-user').value = note.user;
        document.getElementById('notebook-brand').value = note.brand;
        document.getElementById('notebook-serial').value = note.serial;
        document.getElementById('notebook-os').value = note.os;
        
        const initialNote = note.notes && note.notes.length > 0 ? 
            (typeof note.notes[0] === 'string' ? note.notes[0] : note.notes[0].text) : 
            '';
        document.getElementById('notebook-notes').value = initialNote;
        
        if (note.image) {
            document.getElementById('notebook-preview').src = note.image;
            document.getElementById('notebook-preview').style.display = 'block';
        }
        
        document.querySelectorAll('input[name="notebook-connection"]').forEach(cb => {
            cb.checked = note.connections.includes(cb.value);
        });
        
        document.querySelector('.tab-btn[data-tab="equipment"]').click();
        document.getElementById('equipment-form-container').style.display = 'block';
        document.getElementById('add-equipment-btn').style.display = 'none';
        document.querySelector('.subtab-btn[data-subtab="notebooks"]').click();
    } else if (type === 'printer') {
        const printer = printers[index];
        document.getElementById('printer-name').value = printer.name;
        document.getElementById('printer-serial').value = printer.serial;
        document.getElementById('printer-ip').value = printer.ip;
        document.getElementById('printer-location').value = printer.location;
        
        if (printer.image) {
            document.getElementById('printer-preview').src = printer.image;
            document.getElementById('printer-preview').style.display = 'block';
        }
        
        document.querySelector(`input[name="printer-connection"][value="${printer.connection}"]`).checked = true;
        
        selectedUsers = printer.users.split(',').map(u => u.trim()).filter(u => u);
        updateUserTags();
        
        document.querySelector('.tab-btn[data-tab="equipment"]').click();
        document.getElementById('equipment-form-container').style.display = 'block';
        document.getElementById('add-equipment-btn').style.display = 'none';
        document.querySelector('.subtab-btn[data-subtab="printers"]').click();
    }
}

function deleteItem(index, type) {
    if (!confirm('Tem certeza que deseja excluir este item?')) return;
    
    if (type === 'credential') {
        credentials.splice(index, 1);
        renderCredentialsTable();
    } else if (type === 'notebook') {
        notebooks.splice(index, 1);
        renderEquipmentGrid();
    } else if (type === 'printer') {
        printers.splice(index, 1);
        renderEquipmentGrid();
    }
    
    showAlert('Item excluído com sucesso!', 'success');
}

// Funções para credenciais
function saveCredential() {
    const service = document.getElementById('service-input').value.trim();
    const url = document.getElementById('url-input').value.trim();
    const user = document.getElementById('user-input').value.trim();
    const password = document.getElementById('password-input').value;
    const mfa = document.getElementById('mfa-checkbox').checked;
    const notes = document.getElementById('notes-input').value.trim();
    
    if (!service || !user || !password) {
        showAlert('Serviço, usuário e senha são obrigatórios!', 'warning');
        return;
    }
    
    const credential = {
        service,
        url,
        user,
        password,
        mfa,
        notes,
        date: new Date().toLocaleDateString()
    };
    
    if (isEditingIndex !== null && editingType === 'credential') {
        credentials[isEditingIndex] = credential;
        showAlert('Credencial atualizada com sucesso!', 'success');
    } else {
        credentials.push(credential);
        showAlert('Credencial adicionada com sucesso!', 'success');
    }
    
    renderCredentialsTable();
    clearCredentialForm();
}

function formatUrl() {
    const urlInput = document.getElementById('url-input');
    let url = urlInput.value.trim();
    
    if (url && !url.startsWith('http://') && !url.startsWith('https://')) {
        url = 'https://' + url;
        urlInput.value = url;
    }
}

// Funções para gerenciamento de usuários (tags)
function handleUserInput(e) {
    const input = e.target;
    const value = input.value.trim();
    
    if (value.includes(',')) {
        addUserTag(value.replace(',', '').trim());
        input.value = '';
    }
}

function handleUserKeyDown(e) {
    const input = e.target;
    const value = input.value.trim();
    
    if (e.key === 'Enter' && value) {
        addUserTag(value);
        input.value = '';
        e.preventDefault();
    } else if (e.key === 'Backspace' && !value && selectedUsers.length > 0) {
        removeUserTag(selectedUsers.length - 1);
    }
}

function showUserSuggestions() {
    userSuggestions.innerHTML = `
        <div class="suggestion" data-user="admin">admin</div>
        <div class="suggestion" data-user="ti">ti</div>
        <div class="suggestion" data-user="rh">rh</div>
    `;
    userSuggestions.style.display = 'block';
    
    userSuggestions.querySelectorAll('.suggestion').forEach(item => {
        item.addEventListener('click', () => {
            addUserTag(item.getAttribute('data-user'));
            userSuggestions.style.display = 'none';
        });
    });
}

function addUserTag(user) {
    if (user && !selectedUsers.includes(user)) {
        selectedUsers.push(user);
        updateUserTags();
    }
}

function removeUserTag(index) {
    selectedUsers.splice(index, 1);
    updateUserTags();
}

function updateUserTags() {
    const tagsContainer = document.getElementById('printer-users-tags');
    tagsContainer.innerHTML = '';
    
    selectedUsers.forEach((user, index) => {
        const tag = document.createElement('div');
        tag.className = 'tag';
        tag.innerHTML = `
            ${user}
            <span class="remove-tag" data-index="${index}">&times;</span>
        `;
        tagsContainer.appendChild(tag);
    });
    
    document.querySelectorAll('.remove-tag').forEach(btn => {
        btn.addEventListener('click', (e) => {
            removeUserTag(parseInt(e.target.getAttribute('data-index')));
            e.stopPropagation();
        });
    });
}

// Funções para geração de senhas
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
    
    let charset = '';
    if (uppercase) charset += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    if (lowercase) charset += 'abcdefghijklmnopqrstuvwxyz';
    if (numbers) charset += '0123456789';
    if (symbols) charset += '!@#$%^&*()_+-=[]{}|;:,.<>?';
    
    if (!charset) {
        document.getElementById('generated-password').value = 'Selecione pelo menos um tipo de caractere';
        return;
    }
    
    let password = '';
    for (let i = 0; i < length; i++) {
        password += charset.charAt(Math.floor(Math.random() * charset.length));
    }
    
    document.getElementById('generated-password').value = password;
    updatePasswordStrength(password, true);
}

function useGeneratedPassword() {
    document.getElementById('password-input').value = document.getElementById('generated-password').value;
    closePasswordModal();
}

function copyPasswordToClipboard() {
    const passwordField = document.getElementById('generated-password');
    passwordField.select();
    document.execCommand('copy');
    showAlert('Senha copiada para a área de transferência!', 'success');
}

function togglePasswordVisibility() {
    const passwordInput = document.getElementById('password-input');
    const toggleBtn = document.getElementById('toggle-password');
    
    if (passwordInput.type === 'password') {
        passwordInput.type = 'text';
        toggleBtn.innerHTML = '<i class="fas fa-eye-slash"></i>';
    } else {
        passwordInput.type = 'password';
        toggleBtn.innerHTML = '<i class="fas fa-eye"></i>';
    }
}

function updatePasswordStrength(password, isGenerated = false) {
    const strengthBar = document.getElementById('strength-bar');
    let strength = 0;
    
    if (!password) {
        strengthBar.style.width = '0%';
        strengthBar.className = 'strength-bar';
        return;
    }
    
    if (password.length >= 8) strength++;
    if (password.length >= 12) strength++;
    if (/[A-Z]/.test(password)) strength++;
    if (/[0-9]/.test(password)) strength++;
    if (/[^A-Za-z0-9]/.test(password)) strength++;
    
    const width = strength * 20;
    strengthBar.style.width = `${width}%`;
    
    let colorClass;
    if (strength <= 1) colorClass = 'weak';
    else if (strength <= 3) colorClass = 'medium';
    else colorClass = 'strong';
    
    strengthBar.className = `strength-bar ${colorClass}`;
    
    if (isGenerated) {
        let feedback = '';
        if (strength <= 1) feedback = 'Senha fraca';
        else if (strength <= 3) feedback = 'Senha média';
        else feedback = 'Senha forte';
        
        document.getElementById('strength-text').textContent = `Força: ${feedback}`;
    }
}

// Função para carregar dados da planilha
function loadSheetData(sheetName, dataType) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    data.shift();
    
    if (dataType === 'credentials') {
        credentials = data.map(row => ({
            service: row[0] || '',
            url: row[1] || '',
            user: row[2] || '',
            password: row[3] || '',
            mfa: (row[4] || '').toLowerCase() === 'sim',
            notes: row[5] || '',
            date: row[6] || new Date().toLocaleDateString()
        }));
        renderCredentialsTable();
    } else if (dataType === 'notebooks') {
        notebooks = data.map(row => {
            let notes = [];
            if (row[5]) {
                if (row[5].includes('\n\n')) {
                    const noteTexts = row[5].split('\n\n');
                    notes = noteTexts.map(text => ({
                        text: text,
                        date: 'Data não registrada'
                    }));
                } else {
                    notes = [{
                        text: row[5],
                        date: 'Data não registrada'
                    }];
                }
            }
            
            return {
                user: row[0] || '',
                brand: row[1] || '',
                serial: row[2] || '',
                os: row[3] || '',
                connections: row[4] || '',
                notes: notes,
                image: null,
                date: row[7] || new Date().toLocaleDateString()
            };
        });
    } else if (dataType === 'printers') {
        printers = data.map(row => ({
            name: row[0] || '',
            serial: row[1] || '',
            ip: row[2] || '',
            location: row[3] || '',
            connection: row[4] || '',
            users: row[5] || '',
            image: null,
            date: row[7] || new Date().toLocaleDateString()
        }));
    }
    renderEquipmentGrid();
}
