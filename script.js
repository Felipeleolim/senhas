// Variáveis globais
let workbook = null;
let currentSheetName = 'Credenciais';
let credentials = [];
let currentFileName = 'credenciais.xlsx';
let isEditingIndex = null;

// Elementos do DOM
const fileInput = document.getElementById('file-input');
const saveBtn = document.getElementById('save-btn');
const newBtn = document.getElementById('new-btn');
const credentialsList = document.getElementById('credentials-list');
const credentialForm = document.getElementById('credential-form');
const passwordGeneratorBtn = document.getElementById('generate-password-btn');
const togglePasswordBtn = document.getElementById('toggle-password');
const passwordModal = document.getElementById('password-modal');
const generatedPasswordInput = document.getElementById('generated-password');
const usePasswordBtn = document.getElementById('use-password');

// Elementos do formulário
const serviceInput = document.getElementById('service-input');
const urlInput = document.getElementById('url-input');
const userInput = document.getElementById('user-input');
const passwordInput = document.getElementById('password-input');
const notesInput = document.getElementById('notes-input');
const mfaCheckbox = document.getElementById('mfa-checkbox');
const saveCredentialBtn = document.getElementById('save-credential-btn');
const clearFormBtn = document.getElementById('clear-form-btn');

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    createNewWorkbook();
    setupEventListeners();
    updateUI();
});

// Configuração dos event listeners
function setupEventListeners() {
    // Botões principais
    newBtn.addEventListener('click', createNewWorkbook);
    fileInput.addEventListener('change', handleFileLoad);
    saveBtn.addEventListener('click', handleSave);
    
    // Formulário
    saveCredentialBtn.addEventListener('click', saveCredential);
    clearFormBtn.addEventListener('click', clearCredentialForm);
    urlInput.addEventListener('input', formatUrl);
    urlInput.addEventListener('blur', formatUrl);
    passwordGeneratorBtn.addEventListener('click', openPasswordGenerator);
    togglePasswordBtn.addEventListener('click', togglePasswordVisibility);
    passwordInput.addEventListener('input', () => updatePasswordStrength(passwordInput.value));
    
    // Modal de senha
    usePasswordBtn.addEventListener('click', useGeneratedPassword);
    document.getElementById('close-modal').addEventListener('click', closePasswordModal);
    document.getElementById('regenerate-password').addEventListener('click', generateNewPassword);
    document.getElementById('copy-password').addEventListener('click', copyPasswordToClipboard);
    
    // Tabela (event delegation)
    credentialsList.addEventListener('click', handleTableActions);
}

// Funções do Workbook
function createNewWorkbook() {
    if (credentials.length > 0 && !confirm('Isso criará uma nova planilha em branco. Deseja continuar?')) {
        return;
    }

    workbook = XLSX.utils.book_new();
    const headers = ["Serviço", "URL", "Usuário", "Senha", "MFA", "Notas", "Data Atualização"];
    const worksheet = XLSX.utils.aoa_to_sheet([headers]);
    XLSX.utils.book_append_sheet(workbook, worksheet, currentSheetName);
    
    credentials = [];
    currentFileName = 'credenciais.xlsx';
    isEditingIndex = null;
    renderCredentialsTable();
    clearCredentialForm();
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
        currentSheetName = workbook.SheetNames[0];
        loadSheetData(currentSheetName);
        updateUI();
        showAlert('Planilha carregada com sucesso!', 'success');
    } catch (error) {
        console.error("Erro ao carregar arquivo:", error);
        showAlert('Erro ao carregar arquivo. Formato inválido.', 'error');
    }
}

function handleSave() {
    if (!workbook || credentials.length === 0) {
        showAlert('Nenhuma credencial para salvar!', 'warning');
        return;
    }

    try {
        // Prepara os dados para salvar
        const dataToSave = [["Serviço", "URL", "Usuário", "Senha", "MFA", "Notas", "Data Atualização"]];
        
        credentials.forEach(cred => {
            dataToSave.push([
                cred.service,
                cred.url,
                cred.user,
                cred.password,
                cred.mfa ? 'Sim' : 'Não',
                cred.notes,
                cred.date || new Date().toLocaleString()
            ]);
        });

        // Cria uma nova worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(dataToSave);
        
        // Cria um novo workbook
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, currentSheetName);

        // Salva o arquivo
        XLSX.writeFile(newWorkbook, currentFileName);
        
        showAlert(`Planilha salva como ${currentFileName}`, 'success');
    } catch (error) {
        console.error("Erro ao salvar planilha:", error);
        showAlert('Erro ao salvar planilha: ' + error.message, 'error');
    }
}

// Funções para manipulação de dados
function loadSheetData(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    credentials = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row && row.length >= 4) {
            credentials.push({
                service: row[0] || '',
                url: row[1] || '',
                user: row[2] || '',
                password: row[3] || '',
                mfa: (row[4] && (row[4].toString().toLowerCase() === 'sim' || row[4] === true)) || false,
                notes: row[5] || '',
                date: row[6] || new Date().toLocaleString()
            });
        }
    }
    renderCredentialsTable();
}

// Funções do formulário
function showCredentialForm() {
    credentialForm.scrollIntoView({ behavior: 'smooth' });
    serviceInput.focus();
}

function clearCredentialForm() {
    serviceInput.value = '';
    urlInput.value = '';
    userInput.value = '';
    passwordInput.value = '';
    notesInput.value = '';
    mfaCheckbox.checked = false;
    isEditingIndex = null;
    saveCredentialBtn.textContent = 'Salvar Credencial';
    updatePasswordStrength('');
}

function formatUrl() {
    const url = urlInput.value.trim();
    if (url && !url.match(/^https?:\/\//) && !url.match(/^mailto:/) && !url.match(/^tel:/)) {
        urlInput.value = 'https://' + url;
    }
}

function saveCredential() {
    const service = serviceInput.value.trim();
    const url = urlInput.value.trim();
    const user = userInput.value.trim();
    const password = passwordInput.value.trim();
    
    if (!service || !url || !user || !password) {
        showAlert('Preencha todos os campos obrigatórios', 'warning');
        return;
    }
    
    const credentialData = {
        service,
        url,
        user,
        password,
        mfa: mfaCheckbox.checked,
        notes: notesInput.value.trim(),
        date: new Date().toLocaleString()
    };
    
    if (isEditingIndex !== null) {
        // Edição existente
        credentials[isEditingIndex] = credentialData;
        showAlert('Credencial atualizada!', 'success');
    } else {
        // Nova credencial
        credentials.push(credentialData);
        showAlert('Credencial adicionada!', 'success');
    }
    
    renderCredentialsTable();
    clearCredentialForm();
    updateUI();
}

// Funções da tabela
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
                <button class="action-btn show" data-index="${index}" title="Mostrar senha">
                    <i class="fas fa-eye"></i>
                </button>
            </td>
            <td>${cred.mfa ? '<i class="fas fa-check text-success"></i>' : '<i class="fas fa-times text-danger"></i>'}</td>
            <td>${cred.notes}</td>
            <td class="actions">
                <button class="action-btn edit" data-index="${index}" title="Editar">
                    <i class="fas fa-edit"></i>
                </button>
                <button class="action-btn delete" data-index="${index}" title="Excluir">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        credentialsList.appendChild(row);
    });
}

function handleTableActions(e) {
    const btn = e.target.closest('button');
    if (!btn) return;

    const index = parseInt(btn.getAttribute('data-index'));
    if (isNaN(index)) return;

    if (btn.classList.contains('show')) {
        togglePasswordInTable(index);
    } else if (btn.classList.contains('edit')) {
        editCredential(index);
    } else if (btn.classList.contains('delete')) {
        deleteCredential(index);
    }
}

function togglePasswordInTable(index) {
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

function editCredential(index) {
    const cred = credentials[index];
    serviceInput.value = cred.service;
    urlInput.value = cred.url;
    userInput.value = cred.user;
    passwordInput.value = cred.password;
    mfaCheckbox.checked = cred.mfa;
    notesInput.value = cred.notes;
    
    isEditingIndex = index;
    saveCredentialBtn.textContent = 'Atualizar Credencial';
    updatePasswordStrength(cred.password);
    showCredentialForm();
}

function deleteCredential(index) {
    if (confirm(`Tem certeza que deseja excluir "${credentials[index].service}"?`)) {
        credentials.splice(index, 1);
        renderCredentialsTable();
        if (isEditingIndex === index) clearCredentialForm();
        showAlert('Credencial excluída!', 'success');
        updateUI();
    }
}

// Funções do gerador de senhas
function openPasswordGenerator() {
    passwordModal.style.display = 'flex';
    generateNewPassword();
}

function closePasswordModal() {
    passwordModal.style.display = 'none';
}

function generateNewPassword() {
    const length = parseInt(document.querySelector('input[name="length"]:checked').value);
    const uppercase = document.querySelector('input[name="uppercase"]').checked;
    const lowercase = document.querySelector('input[name="lowercase"]').checked;
    const numbers = document.querySelector('input[name="numbers"]').checked;
    const symbols = document.querySelector('input[name="symbols"]').checked;
    
    generatedPasswordInput.value = generateRandomPassword(length, uppercase, lowercase, numbers, symbols);
    updatePasswordStrength(generatedPasswordInput.value, 'modal-strength-bar', 'modal-strength-text');
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
    passwordInput.value = generatedPasswordInput.value;
    updatePasswordStrength(passwordInput.value);
    closePasswordModal();
}

function copyPasswordToClipboard() {
    generatedPasswordInput.select();
    document.execCommand('copy');
    showAlert('Senha copiada!', 'success');
}

function togglePasswordVisibility() {
    const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
    passwordInput.setAttribute('type', type);
    togglePasswordBtn.innerHTML = type === 'password' ? '<i class="fas fa-eye"></i>' : '<i class="fas fa-eye-slash"></i>';
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
    
    // Atualiza UI
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
    saveBtn.disabled = !workbook || credentials.length === 0;
}

function showAlert(message, type = 'info') {
    const alertBox = document.createElement('div');
    alertBox.className = `alert alert-${type}`;
    alertBox.innerHTML = `
        <span>${message}</span>
        <button class="close-alert">&times;</button>
    `;
    
    document.body.appendChild(alertBox);
    
    // Fecha automaticamente após 3 segundos
    setTimeout(() => {
        alertBox.classList.add('fade-out');
        setTimeout(() => alertBox.remove(), 300);
    }, 3000);
    
    // Fecha manualmente
    alertBox.querySelector('.close-alert').addEventListener('click', () => {
        alertBox.remove();
    });
}