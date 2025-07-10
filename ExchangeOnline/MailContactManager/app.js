// Microsoft Graph configuration
const msalConfig = {
    auth: {
        clientId: '729bb47e-4ea6-4b86-b44e-de38234b844d', // Replace with your app registration client ID
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: window.location.origin
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
    }
};

const loginRequest = {
    scopes: [
        'User.Read',
        'Contacts.ReadWrite',
        'Directory.Read.All',
        'User.ReadBasic.All'
    ]
};

// Initialize MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Global variables
let accessToken = null;
let allContacts = [];
let currentEditingContact = null;

// DOM elements
const authButton = document.getElementById('authButton');
const connectionStatus = document.getElementById('connectionStatus');
const connectionText = document.getElementById('connectionText');
const searchInput = document.getElementById('searchInput');
const typeFilter = document.getElementById('typeFilter');
const addContactBtn = document.getElementById('addContactBtn');
const contactsGrid = document.getElementById('contactsGrid');
const loadingSpinner = document.getElementById('loadingSpinner');
const contactModal = new bootstrap.Modal(document.getElementById('contactModal'));
const contactForm = document.getElementById('contactForm');
const saveContactBtn = document.getElementById('saveContactBtn');
const deleteContactBtn = document.getElementById('deleteContactBtn');

// Event listeners
authButton.addEventListener('click', handleAuth);
searchInput.addEventListener('input', filterContacts);
typeFilter.addEventListener('change', filterContacts);
addContactBtn.addEventListener('click', () => openContactModal());
saveContactBtn.addEventListener('click', saveContact);
deleteContactBtn.addEventListener('click', deleteContact);

// Initialize app
document.addEventListener('DOMContentLoaded', async () => {
    await checkAuthState();
});

// Authentication functions
async function handleAuth() {
    if (accessToken) {
        await signOut();
    } else {
        await signIn();
    }
}

async function signIn() {
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        await getAccessToken();
        updateAuthUI(true);
        await loadContacts();
    } catch (error) {
        console.error('Login failed:', error);
        showError('Authentication failed: ' + error.message);
    }
}

async function signOut() {
    try {
        await msalInstance.logoutPopup();
        accessToken = null;
        updateAuthUI(false);
        clearContacts();
    } catch (error) {
        console.error('Logout failed:', error);
    }
}

async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        throw new Error('No accounts found');
    }

    const silentRequest = {
        ...loginRequest,
        account: accounts[0]
    };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        accessToken = response.accessToken;
        return accessToken;
    } catch (error) {
        console.error('Silent token acquisition failed:', error);
        const response = await msalInstance.acquireTokenPopup(loginRequest);
        accessToken = response.accessToken;
        return accessToken;
    }
}

async function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            await getAccessToken();
            updateAuthUI(true);
            await loadContacts();
        } catch (error) {
            console.error('Auto-login failed:', error);
            updateAuthUI(false);
        }
    }
}

function updateAuthUI(isAuthenticated) {
    if (isAuthenticated) {
        authButton.textContent = 'Sign Out';
        authButton.className = 'btn btn-outline-light';
        connectionStatus.className = 'status-indicator status-connected';
        connectionText.textContent = 'Connected';
        addContactBtn.disabled = false;
    } else {
        authButton.textContent = 'Sign In';
        authButton.className = 'btn btn-outline-light';
        connectionStatus.className = 'status-indicator status-disconnected';
        connectionText.textContent = 'Not Connected';
        addContactBtn.disabled = true;
    }
}

// Microsoft Graph API functions
async function makeGraphRequest(endpoint, method = 'GET', body = null) {
    if (!accessToken) {
        throw new Error('No access token available');
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method: method,
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: body ? JSON.stringify(body) : null
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(`Graph API error: ${error.error?.message || response.statusText}`);
    }

    return response.json();
}

async function loadContacts() {
    showLoading(true);
    try {
        // Get organizational contacts from directory
        const directoryContacts = await makeGraphRequest('/contacts');
        
        // Get mail contacts (if available through Graph)
        let mailContacts = [];
        try {
            const mailContactsResponse = await makeGraphRequest('/me/contacts');
            mailContacts = mailContactsResponse.value || [];
        } catch (error) {
            console.log('Mail contacts not accessible via Graph API');
        }

        // Combine and process contacts
        allContacts = [
            ...directoryContacts.value.map(contact => ({
                ...contact,
                type: 'directory',
                displayName: contact.displayName || contact.givenName + ' ' + contact.surname,
                emailAddress: contact.emailAddresses?.[0]?.address || '',
                jobTitle: contact.jobTitle || '',
                department: contact.department || '',
                companyName: contact.companyName || '',
                businessPhone: contact.businessPhones?.[0] || '',
                officeLocation: contact.officeLocation || ''
            })),
            ...mailContacts.map(contact => ({
                ...contact,
                type: 'mail',
                emailAddress: contact.emailAddresses?.[0]?.address || ''
            }))
        ];

        renderContacts();
    } catch (error) {
        console.error('Failed to load contacts:', error);
        showError('Failed to load contacts: ' + error.message);
    } finally {
        showLoading(false);
    }
}

async function saveContact() {
    const formData = getFormData();
    
    if (!formData.displayName || !formData.emailAddress) {
        showError('Display Name and Email Address are required');
        return;
    }

    try {
        const contactData = {
            displayName: formData.displayName,
            emailAddresses: [{
                address: formData.emailAddress,
                name: formData.displayName
            }],
            jobTitle: formData.jobTitle,
            department: formData.department,
            companyName: formData.companyName,
            businessPhones: formData.businessPhone ? [formData.businessPhone] : [],
            officeLocation: formData.officeLocation
        };

        if (currentEditingContact) {
            // Update existing contact
            await makeGraphRequest(`/me/contacts/${currentEditingContact.id}`, 'PATCH', contactData);
            showSuccess('Contact updated successfully');
        } else {
            // Create new contact
            await makeGraphRequest('/me/contacts', 'POST', contactData);
            showSuccess('Contact created successfully');
        }

        contactModal.hide();
        await loadContacts();
    } catch (error) {
        console.error('Failed to save contact:', error);
        showError('Failed to save contact: ' + error.message);
    }
}

async function deleteContact() {
    if (!currentEditingContact) return;

    if (confirm(`Are you sure you want to delete ${currentEditingContact.displayName}?`)) {
        try {
            await makeGraphRequest(`/me/contacts/${currentEditingContact.id}`, 'DELETE');
            showSuccess('Contact deleted successfully');
            contactModal.hide();
            await loadContacts();
        } catch (error) {
            console.error('Failed to delete contact:', error);
            showError('Failed to delete contact: ' + error.message);
        }
    }
}

// UI functions
function renderContacts() {
    const filteredContacts = getFilteredContacts();
    
    if (filteredContacts.length === 0) {
        contactsGrid.innerHTML = `
            <div class="col-12 text-center py-5">
                <i class="fas fa-address-book fa-3x text-muted mb-3"></i>
                <h5 class="text-muted">No contacts found</h5>
                <p class="text-muted">Try adjusting your search or filter criteria</p>
            </div>
        `;
        return;
    }

    contactsGrid.innerHTML = filteredContacts.map(contact => `
        <div class="col-md-6 col-lg-4 mb-3">
            <div class="card contact-card h-100" onclick="openContactModal('${contact.id}')">
                <div class="card-body">
                    <div class="d-flex align-items-center mb-2">
                        <div class="bg-primary text-white rounded-circle d-flex align-items-center justify-content-center me-3" 
                             style="width: 40px; height: 40px;">
                            <i class="fas fa-user"></i>
                        </div>
                        <div class="flex-grow-1">
                            <h6 class="card-title mb-0">${contact.displayName}</h6>
                            <small class="text-muted">${contact.type === 'directory' ? 'Directory Contact' : 'Mail Contact'}</small>
                        </div>
                    </div>
                    <p class="card-text">
                        <i class="fas fa-envelope text-muted me-2"></i>
                        <small>${contact.emailAddress}</small>
                    </p>
                    ${contact.jobTitle ? `
                        <p class="card-text">
                            <i class="fas fa-briefcase text-muted me-2"></i>
                            <small>${contact.jobTitle}</small>
                        </p>
                    ` : ''}
                    ${contact.department ? `
                        <p class="card-text">
                            <i class="fas fa-building text-muted me-2"></i>
                            <small>${contact.department}</small>
                        </p>
                    ` : ''}
                </div>
            </div>
        </div>
    `).join('');
}

function getFilteredContacts() {
    let filtered = allContacts;

    // Apply search filter
    const searchTerm = searchInput.value.toLowerCase();
    if (searchTerm) {
        filtered = filtered.filter(contact => 
            contact.displayName.toLowerCase().includes(searchTerm) ||
            contact.emailAddress.toLowerCase().includes(searchTerm) ||
            (contact.jobTitle && contact.jobTitle.toLowerCase().includes(searchTerm)) ||
            (contact.department && contact.department.toLowerCase().includes(searchTerm))
        );
    }

    // Apply type filter
    const typeFilterValue = typeFilter.value;
    if (typeFilterValue !== 'all') {
        filtered = filtered.filter(contact => {
            if (typeFilterValue === 'contact') return contact.type === 'mail';
            if (typeFilterValue === 'user') return contact.type === 'directory';
            return true;
        });
    }

    return filtered;
}

function filterContacts() {
    renderContacts();
}

function openContactModal(contactId = null) {
    currentEditingContact = contactId ? allContacts.find(c => c.id === contactId) : null;
    
    if (currentEditingContact) {
        // Edit mode
        document.getElementById('modalTitle').textContent = 'Edit Contact';
        populateForm(currentEditingContact);
        deleteContactBtn.style.display = 'inline-block';
    } else {
        // Create mode
        document.getElementById('modalTitle').textContent = 'Add New Contact';
        clearForm();
        deleteContactBtn.style.display = 'none';
    }
    
    contactModal.show();
}

function populateForm(contact) {
    document.getElementById('displayName').value = contact.displayName || '';
    document.getElementById('emailAddress').value = contact.emailAddress || '';
    document.getElementById('contactType').value = contact.type === 'directory' ? 'user' : 'contact';
    document.getElementById('jobTitle').value = contact.jobTitle || '';
    document.getElementById('department').value = contact.department || '';
    document.getElementById('companyName').value = contact.companyName || '';
    document.getElementById('businessPhone').value = contact.businessPhone || '';
    document.getElementById('officeLocation').value = contact.officeLocation || '';
}

function clearForm() {
    contactForm.reset();
}

function getFormData() {
    return {
        displayName: document.getElementById('displayName').value,
        emailAddress: document.getElementById('emailAddress').value,
        contactType: document.getElementById('contactType').value,
        jobTitle: document.getElementById('jobTitle').value,
        department: document.getElementById('department').value,
        companyName: document.getElementById('companyName').value,
        businessPhone: document.getElementById('businessPhone').value,
        officeLocation: document.getElementById('officeLocation').value
    };
}

function clearContacts() {
    allContacts = [];
    contactsGrid.innerHTML = `
        <div class="col-12 text-center py-5">
            <i class="fas fa-sign-in-alt fa-3x text-muted mb-3"></i>
            <h5 class="text-muted">Please sign in to view contacts</h5>
        </div>
    `;
}

function showLoading(show) {
    loadingSpinner.style.display = show ? 'block' : 'none';
    contactsGrid.style.display = show ? 'none' : 'block';
}

function showError(message) {
    // Create toast notification for errors
    const toast = document.createElement('div');
    toast.className = 'toast align-items-center text-white bg-danger border-0 position-fixed top-0 end-0 m-3';
    toast.style.zIndex = '9999';
    toast.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">
                <i class="fas fa-exclamation-triangle me-2"></i>${message}
            </div>
            <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
    `;
    document.body.appendChild(toast);
    const bsToast = new bootstrap.Toast(toast);
    bsToast.show();
    toast.addEventListener('hidden.bs.toast', () => toast.remove());
}

function showSuccess(message) {
    // Create toast notification for success
    const toast = document.createElement('div');
    toast.className = 'toast align-items-center text-white bg-success border-0 position-fixed top-0 end-0 m-3';
    toast.style.zIndex = '9999';
    toast.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">
                <i class="fas fa-check-circle me-2"></i>${message}
            </div>
            <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
    `;
    document.body.appendChild(toast);
    const bsToast = new bootstrap.Toast(toast);
    bsToast.show();
    toast.addEventListener('hidden.bs.toast', () => toast.remove());
}
