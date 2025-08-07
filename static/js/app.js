// グローバル変数
let vendors = [];
let companies = [];  // 送金会社データ
let paymentItems = [];
let currentPaymentId = null;

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    loadVendors();
    loadCompanies();  // 送金会社データを読み込み
    loadPaymentHistory();
    loadUploadedFiles();
    updateVendorStats();
    setupVendorSearch(); // 業者検索機能をセットアップ
    setupCompanySearch(); // 送金会社検索機能をセットアップ
    
    // 今日の日付をデフォルトに設定
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('payment-date').value = today;
});

// セクション表示切り替え
function showSection(sectionId) {
    // すべてのセクションを非表示
    const sections = document.querySelectorAll('.content-section');
    sections.forEach(section => section.style.display = 'none');
    
    // 指定されたセクションを表示
    document.getElementById(sectionId).style.display = 'block';
    
    // ナビゲーションのアクティブ状態を更新
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => link.classList.remove('active'));
    
    // 対応するナビゲーションリンクをアクティブにする
    const sectionToNavMap = {
        'payment-section': 0,
        'upload-section': 1,
        'vendor-section': 2,
        'history-section': 3
    };
    
    const navIndex = sectionToNavMap[sectionId];
    if (navIndex !== undefined && navLinks[navIndex]) {
        navLinks[navIndex].classList.add('active');
    }
}

// 業者データを読み込み
async function loadVendors() {
    try {
        const response = await fetch('/api/vendors');
        vendors = await response.json();
        updateVendorSelect();
        updateVendorList();
    } catch (error) {
        console.error('業者データの読み込みエラー:', error);
        showAlert('業者データの読み込みに失敗しました', 'danger');
    }
}

// 送金会社データを読み込み
async function loadCompanies() {
    try {
        const response = await fetch('/api/companies');
        companies = await response.json();
        console.log('送金会社データを読み込みました:', companies);
    } catch (error) {
        console.error('送金会社データの読み込みエラー:', error);
        showAlert('送金会社データの読み込みに失敗しました', 'danger');
    }
}

// 業者選択プルダウンを更新
function updateVendorSelect() {
    const select = document.getElementById('vendor-select');
    select.innerHTML = '<option value="">業者を選択してください</option>';
    
    vendors.forEach(vendor => {
        const option = document.createElement('option');
        option.value = vendor.id;
        option.textContent = vendor.name;
        select.appendChild(option);
    });
}



// 業者検索機能
function setupVendorSearch() {
    const searchInput = document.getElementById('vendor-search');
    const searchResults = document.getElementById('vendor-search-results');
    const vendorSelect = document.getElementById('vendor-select');
    
    if (!searchInput) return;
    
    let searchTimeout;
    
    searchInput.addEventListener('input', function() {
        const query = this.value.trim();
        
        // 検索結果をクリア
        clearTimeout(searchTimeout);
        
        if (query.length < 1) {
            searchResults.style.display = 'none';
            return;
        }
        
        // デバウンス処理（300ms後に検索実行）
        searchTimeout = setTimeout(async () => {
            try {
                const response = await fetch(`/api/vendors/search?q=${encodeURIComponent(query)}`);
                const results = await response.json();
                
                displaySearchResults(results);
            } catch (error) {
                console.error('業者検索エラー:', error);
            }
        }, 300);
    });
    
    // 検索結果表示
    function displaySearchResults(results) {
        searchResults.innerHTML = '';
        
        if (results.length === 0) {
            searchResults.innerHTML = '<div class="search-result-item text-muted">該当する業者が見つかりません</div>';
        } else {
            results.forEach(vendor => {
                const item = document.createElement('div');
                item.className = 'search-result-item';
                item.innerHTML = `
                    <div class="vendor-name">${vendor.name}</div>
                    <div class="vendor-details text-muted small">
                        ${vendor.bank_name} ${vendor.branch_name} ${vendor.account_number}
                    </div>
                `;
                
                item.addEventListener('click', () => {
                    // 業者を選択
                    vendorSelect.value = vendor.id;
                    searchInput.value = vendor.name;
                    searchResults.style.display = 'none';
                });
                
                searchResults.appendChild(item);
            });
        }
        
        searchResults.style.display = 'block';
    }
    
    // クリック外で検索結果を非表示
    document.addEventListener('click', function(event) {
        if (!searchInput.contains(event.target) && !searchResults.contains(event.target)) {
            searchResults.style.display = 'none';
        }
    });
}

// 送金会社検索機能
function setupCompanySearch() {
    const searchInput = document.getElementById('company-search');
    const searchResults = document.getElementById('company-search-results');
    const hiddenInput = document.getElementById('remittance-company');
    
    if (!searchInput) return;
    
    let searchTimeout;
    
    searchInput.addEventListener('input', function() {
        const query = this.value.trim();
        
        // 検索結果をクリア
        clearTimeout(searchTimeout);
        
        if (query.length < 1) {
            searchResults.style.display = 'none';
            hiddenInput.value = '';
            return;
        }
        
        // デバウンス処理（300ms後に検索実行）
        searchTimeout = setTimeout(() => {
            displayCompanySearchResults(query);
        }, 300);
    });
    
    // 送金会社検索結果表示
    function displayCompanySearchResults(query) {
        searchResults.innerHTML = '';
        
        console.log('検索クエリ:', query);
        console.log('送金会社データ:', companies);
        
        // ローカルの送金会社データから検索
        const filteredCompanies = companies.filter(company => 
            company.name.toLowerCase().includes(query.toLowerCase())
        );
        
        console.log('フィルタ結果:', filteredCompanies);
        
        if (filteredCompanies.length === 0) {
            searchResults.innerHTML = '<div class="search-result-item text-muted">該当する送金会社が見つかりません</div>';
        } else {
            filteredCompanies.forEach(company => {
                const item = document.createElement('div');
                item.className = 'search-result-item';
                item.innerHTML = `
                    <div class="company-name">${company.name}</div>
                    <div class="company-details text-muted small">
                        ${company.bank_name} ${company.branch_name} ${company.account_number}
                    </div>
                `;
                
                item.addEventListener('click', () => {
                    // 送金会社を選択
                    searchInput.value = company.name;
                    hiddenInput.value = company.name;
                    searchResults.style.display = 'none';
                });
                
                searchResults.appendChild(item);
            });
        }
        
        searchResults.style.display = 'block';
    }
    
    // クリック外で検索結果を非表示
    document.addEventListener('click', function(event) {
        if (!searchInput.contains(event.target) && !searchResults.contains(event.target)) {
            searchResults.style.display = 'none';
        }
    });
}

// 業者一覧テーブルを更新
function updateVendorList() {
    const tbody = document.getElementById('vendor-list');
    tbody.innerHTML = '';
    
    vendors.forEach(vendor => {
        const row = document.createElement('tr');
        const sourceIcon = vendor.source === 'upload' ? 
            '<i class="fas fa-upload text-success" title="アップロード由来"></i>' : 
            '<i class="fas fa-keyboard text-primary" title="手動登録"></i>';
        
        row.innerHTML = `
            <td>${sourceIcon} ${vendor.name}</td>
            <td>${vendor.bank_name}</td>
            <td>${vendor.branch_name}</td>
            <td>${vendor.account_number}</td>
        `;
        tbody.appendChild(row);
    });
}

// 業者を追加
async function addVendor() {
    const form = document.getElementById('vendor-form');
    const formData = new FormData(form);
    
    const vendorData = {
        name: document.getElementById('vendor-name').value,
        bank_name: document.getElementById('bank-name').value,
        branch_name: document.getElementById('branch-name').value,
        account_type: parseInt(document.getElementById('account-type').value),
        account_number: document.getElementById('account-number').value,
        account_holder: document.getElementById('account-holder').value
    };
    
    // バリデーション
    if (!vendorData.name || !vendorData.bank_name || !vendorData.branch_name || 
        !vendorData.account_number || !vendorData.account_holder) {
        showAlert('すべての項目を入力してください', 'danger');
        return;
    }
    
    try {
        const response = await fetch('/api/vendors', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(vendorData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showAlert('業者を登録しました', 'success');
            form.reset();
            loadVendors();
        } else {
            showAlert('業者の登録に失敗しました', 'danger');
        }
    } catch (error) {
        console.error('業者登録エラー:', error);
        showAlert('業者の登録に失敗しました', 'danger');
    }
}

// 支払項目を追加
function addPaymentItem() {
    const vendorId = parseInt(document.getElementById('vendor-select').value);
    const amount = parseInt(document.getElementById('amount').value);
    const description = document.getElementById('description').value;
    const remarks = document.getElementById('remarks').value;
    
    // バリデーション
    if (!vendorId || !amount || !description) {
        showAlert('業者、金額、摘要は必須項目です', 'danger');
        return;
    }
    
    const vendor = vendors.find(v => v.id === vendorId);
    if (!vendor) {
        showAlert('選択された業者が見つかりません', 'danger');
        return;
    }
    
    // 支払項目を追加
    const item = {
        id: Date.now(),
        vendor_id: vendorId,
        vendor_name: vendor.name,
        amount: amount,
        description: description,
        remarks: remarks
    };
    
    paymentItems.push(item);
    updatePaymentTable();
    
    // フォームをクリア
    document.getElementById('vendor-select').value = '';
    document.getElementById('amount').value = '';
    document.getElementById('description').value = '';
    document.getElementById('remarks').value = '';
    
    showAlert('支払項目を追加しました', 'success');
}

// 支払テーブルを更新
function updatePaymentTable() {
    const tbody = document.getElementById('payment-items');
    tbody.innerHTML = '';
    
    let total = 0;
    
    paymentItems.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.vendor_name}</td>
            <td class="amount-cell">${item.amount.toLocaleString()}</td>
            <td>${item.description}</td>
            <td>${item.remarks || '-'}</td>
            <td>
                <button class="btn btn-danger btn-sm btn-delete" onclick="removePaymentItem(${item.id})">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(row);
        total += item.amount;
    });
    
    // 合計金額を更新
    document.getElementById('total-amount').textContent = total.toLocaleString();
    
    // ボタンの有効/無効を切り替え
    const hasItems = paymentItems.length > 0;
    const hasPaymentInfo = document.getElementById('payment-date').value && 
                          document.getElementById('remittance-company').value;
    
    document.getElementById('create-btn').disabled = !(hasItems && hasPaymentInfo);
    
    // ヘッダー情報を更新
    if (hasPaymentInfo) {
        updatePaymentHeader();
    }
}

// 支払項目を削除
function removePaymentItem(itemId) {
    paymentItems = paymentItems.filter(item => item.id !== itemId);
    updatePaymentTable();
    showAlert('支払項目を削除しました', 'success');
}

// 支払ヘッダーを更新
function updatePaymentHeader() {
    const paymentDate = document.getElementById('payment-date').value;
    const remittanceCompany = document.getElementById('remittance-company').value;
    
    if (paymentDate && remittanceCompany) {
        document.getElementById('header-date').textContent = paymentDate;
        document.getElementById('header-company').textContent = remittanceCompany;
        document.getElementById('payment-header').style.display = 'block';
    } else {
        document.getElementById('payment-header').style.display = 'none';
    }
}

// 支払表を作成
async function createPaymentList() {
    if (paymentItems.length === 0) {
        showAlert('支払項目がありません', 'warning');
        return;
    }
    
    const paymentDate = document.getElementById('payment-date').value;
    const remittanceCompany = document.getElementById('remittance-company').value;
    
    if (!paymentDate || !remittanceCompany) {
        showAlert('支払日と送金会社名を入力してください', 'warning');
        return;
    }
    
    const paymentData = {
        payment_date: paymentDate,
        remittance_company: remittanceCompany,
        items: paymentItems
    };
    
    try {
        const response = await fetch('/api/payments', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(paymentData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            currentPaymentId = result.payment_id;
            
            let message = '支払表を作成しました';
            if (result.pdf_generated) {
                message += '。PDFも生成されました。';
                // PDFダウンロードボタンを表示
                showPdfDownloadButton(result.payment_id);
            } else if (result.error) {
                message += `。ただし、PDF生成に失敗しました: ${result.error}`;
            }
            
            showAlert(message, 'success');
            loadPaymentHistory();
        } else {
            showAlert('支払表の作成に失敗しました', 'danger');
        }
    } catch (error) {
        console.error('支払表作成エラー:', error);
        showAlert('支払表の作成に失敗しました', 'danger');
    }
}

// PDFダウンロードボタンを表示
function showPdfDownloadButton(paymentId) {
    const alertContainer = document.querySelector('.alert:last-child');
    if (alertContainer) {
        const downloadBtn = document.createElement('button');
        downloadBtn.className = 'btn btn-primary btn-sm ms-2';
        downloadBtn.innerHTML = '<i class="fas fa-download"></i> PDFダウンロード';
        downloadBtn.onclick = () => downloadPaymentPdf(paymentId);
        alertContainer.appendChild(downloadBtn);
    }
}

// 支払表PDFをダウンロード
function downloadPaymentPdf(paymentId) {
    const link = document.createElement('a');
    link.href = `/api/payments/${paymentId}/pdf`;
    link.download = `payment_list_${paymentId}.pdf`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 総合振込ファイルを生成
function generateTransferFile() {
    if (!currentPaymentId) {
        showAlert('先に支払表を作成してください', 'danger');
        return;
    }
    
    // ファイルダウンロード
    window.location.href = `/api/payments/${currentPaymentId}/transfer`;
    showAlert('総合振込ファイルをダウンロードしました', 'success');
}

// 支払履歴を読み込み
async function loadPaymentHistory() {
    try {
        const response = await fetch('/api/payments');
        const payments = await response.json();
        updatePaymentHistory(payments);
    } catch (error) {
        console.error('支払履歴の読み込みエラー:', error);
        showAlert('支払履歴の読み込みに失敗しました', 'danger');
    }
}

// 支払履歴テーブルを更新
function updatePaymentHistory(payments) {
    const tbody = document.getElementById('payment-history');
    tbody.innerHTML = '';
    
    payments.forEach(payment => {
        const row = document.createElement('tr');
        const totalAmount = payment.items.reduce((sum, item) => sum + parseInt(item.amount), 0);
        
        // 作成日時を日本語形式で表示
        const createdAt = new Date(payment.created_at);
        const createdAtFormatted = createdAt.toLocaleString('ja-JP', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });
        
        // 支払日を日本語形式で表示
        const paymentDate = new Date(payment.payment_date);
        const paymentDateFormatted = paymentDate.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
        
        row.innerHTML = `
            <td>${createdAtFormatted}</td>
            <td>${paymentDateFormatted}</td>
            <td>${payment.remittance_company}</td>
            <td>${payment.items.length}件</td>
            <td>${totalAmount.toLocaleString()}円</td>
            <td>
                <button class="btn btn-info btn-sm me-1" onclick="recreateFromHistory('${payment.id}')" title="この支払表をベースに新しい支払表を作成">
                    <i class="fas fa-copy"></i> 再作成
                </button>
                <button class="btn btn-success btn-sm me-1" onclick="downloadPaymentPdf('${payment.id}')">
                    <i class="fas fa-file-pdf"></i> PDF
                </button>
                <button class="btn btn-primary btn-sm me-1" onclick="downloadTransferFile('${payment.id}')">
                    <i class="fas fa-download"></i> 振込ファイル
                </button>
                <button class="btn btn-danger btn-sm" onclick="deletePayment('${payment.id}', '${paymentDateFormatted}', '${payment.remittance_company}')">
                    <i class="fas fa-trash"></i> 削除
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// 履歴から振込ファイルをダウンロード
function downloadTransferFile(paymentId) {
    window.location.href = `/api/payments/${paymentId}/transfer`;
    showAlert('総合振込ファイルをダウンロードしました', 'success');
}

// アラート表示
function showAlert(message, type = 'info') {
    // 既存のアラートを削除
    const existingAlert = document.querySelector('.alert-custom');
    if (existingAlert) {
        existingAlert.remove();
    }
    
    // 新しいアラートを作成
    const alert = document.createElement('div');
    alert.className = `alert alert-${type} alert-dismissible fade show alert-custom`;
    alert.style.position = 'fixed';
    alert.style.top = '20px';
    alert.style.right = '20px';
    alert.style.zIndex = '9999';
    alert.style.minWidth = '300px';
    
    alert.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(alert);
    
    // 3秒後に自動削除
    setTimeout(() => {
        if (alert.parentNode) {
            alert.remove();
        }
    }, 3000);
}

// マスターデータアップロード機能
async function uploadMasterFile() {
    const fileInput = document.getElementById('master-file');
    const file = fileInput.files[0];
    
    if (!file) {
        showAlert('ファイルを選択してください', 'danger');
        return;
    }
    
    const formData = new FormData();
    formData.append('file', file);
    
    // プログレス表示
    document.getElementById('upload-progress').style.display = 'block';
    
    try {
        const response = await fetch('/api/upload-file', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            showAlert(result.message, 'success');
            fileInput.value = ''; // ファイル選択をクリア
            loadVendors(); // 業者一覧を更新
            loadUploadedFiles(); // アップロードファイル一覧を更新
            updateVendorStats(); // 統計を更新
        } else {
            showAlert(result.error, 'danger');
        }
    } catch (error) {
        console.error('アップロードエラー:', error);
        showAlert('ファイルのアップロードに失敗しました', 'danger');
    } finally {
        document.getElementById('upload-progress').style.display = 'none';
    }
}

// アップロード済みファイル一覧を読み込み
async function loadUploadedFiles() {
    try {
        const response = await fetch('/api/upload-files');
        const files = await response.json();
        updateUploadedFilesList(files);
    } catch (error) {
        console.error('アップロードファイル読み込みエラー:', error);
    }
}

// アップロード済みファイル一覧を更新
function updateUploadedFilesList(files) {
    const tbody = document.getElementById('uploaded-files-list');
    tbody.innerHTML = '';
    
    if (files.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td colspan="4" class="text-center text-muted">
                アップロード済みファイルはありません
            </td>
        `;
        tbody.appendChild(row);
        return;
    }
    
    files.forEach(file => {
        const row = document.createElement('tr');
        const fileSize = formatFileSize(file.size);
        const modifiedDate = new Date(file.modified).toLocaleString('ja-JP');
        
        row.innerHTML = `
            <td>${file.filename}</td>
            <td>${fileSize}</td>
            <td>${modifiedDate}</td>
            <td>
                <button class="btn btn-danger btn-sm" onclick="deleteUploadedFile('${file.filename}')">
                    <i class="fas fa-trash"></i> 削除
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// ファイルサイズをフォーマット
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// アップロードファイルを削除
async function deleteUploadedFile(filename) {
    if (!confirm(`ファイル "${filename}" を削除しますか？\n関連する業者データも削除されます。`)) {
        return;
    }
    
    try {
        const response = await fetch(`/api/delete-file/${encodeURIComponent(filename)}`, {
            method: 'DELETE'
        });
        
        const result = await response.json();
        
        if (result.success) {
            showAlert(result.message, 'success');
            loadVendors(); // 業者一覧を更新
            loadUploadedFiles(); // アップロードファイル一覧を更新
            updateVendorStats(); // 統計を更新
        } else {
            showAlert(result.error, 'danger');
        }
    } catch (error) {
        console.error('ファイル削除エラー:', error);
        showAlert('ファイルの削除に失敗しました', 'danger');
    }
}

// 業者統計を更新
function updateVendorStats() {
    const manualVendors = vendors.filter(v => v.source !== 'upload');
    const uploadVendors = vendors.filter(v => v.source === 'upload');
    
    document.getElementById('manual-vendor-count').textContent = manualVendors.length;
    document.getElementById('upload-vendor-count').textContent = uploadVendors.length;
}

// 支払履歴を削除
async function deletePayment(paymentId, paymentDate, remittanceCompany) {
    if (confirm(`以下の支払表を削除しますか？\n\n支払日: ${paymentDate}\n送金会社: ${remittanceCompany}\n\n※この操作は取り消せません。`)) {
        fetch(`/api/payments/${paymentId}`, {
            method: 'DELETE'
        })
        .then(response => {
            if (response.ok) {
                alert('支払表を削除しました。');
                loadPaymentHistory(); // 履歴を再読み込み
            } else {
                alert('削除に失敗しました。');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('削除中にエラーが発生しました。');
        });
    }
}

// 過去履歴から再作成
function recreateFromHistory(paymentId) {
    console.log('再作成開始:', paymentId);
    
    fetch(`/api/payments/${paymentId}`)
        .then(response => {
            console.log('API応答:', response);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(payment => {
            console.log('取得した支払データ:', payment);
            
            // 支払表作成セクションを表示
            showSection('payment-section');
            
            // 基本情報を設定
            document.getElementById('payment-date').value = payment.payment_date;
            document.getElementById('remittance-company').value = payment.remittance_company;
            
            // 支払項目をクリア
            paymentItems = [];
            
            // 過去の支払項目を復元
            payment.items.forEach(item => {
                paymentItems.push({
                    id: Date.now() + Math.random(), // 新しいIDを生成
                    vendor_id: item.vendor_id,
                    vendor_name: item.vendor_name,
                    amount: item.amount,
                    description: item.description,
                    remarks: item.note || item.remarks || ''
                });
            });
            
            // 支払項目テーブルを更新
            updatePaymentTable();
            
            // メッセージ表示
            alert(`「${payment.remittance_company}」の支払表（${payment.payment_date}）を再作成用に読み込みました。\n\n内容を確認・編集して保存してください。`);
            
            // ページトップにスクロール
            window.scrollTo(0, 0);
        })
        .catch(error => {
            console.error('再作成エラー:', error);
            alert(`支払表の読み込みに失敗しました。\nエラー: ${error.message}`);
        });
}

// 支払情報の変更を監視
document.addEventListener('DOMContentLoaded', function() {
    const paymentDate = document.getElementById('payment-date');
    const remittanceCompany = document.getElementById('remittance-company');
    
    paymentDate.addEventListener('change', updatePaymentTable);
    remittanceCompany.addEventListener('input', updatePaymentTable);
});

// データバックアップ・復元機能
function createBackup() {
    if (!confirm('現在のデータのバックアップを作成しますか？')) {
        return;
    }
    
    fetch('/api/backup/create', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            alert(`バックアップを作成しました！\n支払データ: ${data.payments_count}件\n業者データ: ${data.vendors_count}件`);
        } else {
            alert('バックアップ作成に失敗しました: ' + data.error);
        }
    })
    .catch(error => {
        console.error('バックアップ作成エラー:', error);
        alert('バックアップ作成中にエラーが発生しました');
    });
}

function restoreBackup() {
    if (!confirm('バックアップからデータを復元しますか？\n現在のデータは上書きされます。')) {
        return;
    }
    
    fetch('/api/backup/restore', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        }
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            alert(`バックアップから復元しました！\n支払データ: ${data.payments_restored}件\n業者データ: ${data.vendors_restored}件`);
            // ページをリロードして最新データを表示
            location.reload();
        } else {
            alert('バックアップ復元に失敗しました: ' + data.error);
        }
    })
    .catch(error => {
        console.error('バックアップ復元エラー:', error);
        alert('バックアップ復元中にエラーが発生しました');
    });
}

function checkBackupStatus() {
    fetch('/api/backup/status')
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            let message = `バックアップ状態\n\n`;
            message += `バックアップファイル数: ${data.backup_count}件\n\n`;
            
            if (data.backup_files.length > 0) {
                message += '最新のバックアップファイル:\n';
                data.backup_files.slice(-3).forEach(file => {
                    const date = new Date(file.modified).toLocaleString('ja-JP');
                    message += `• ${file.filename} (${date})\n`;
                });
            } else {
                message += 'バックアップファイルがありません。';
            }
            
            alert(message);
        } else {
            alert('バックアップ状態確認に失敗しました: ' + data.error);
        }
    })
    .catch(error => {
        console.error('バックアップ状態確認エラー:', error);
        alert('バックアップ状態確認中にエラーが発生しました');
    });
}
