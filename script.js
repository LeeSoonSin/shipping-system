import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getDatabase, ref, set, push, onValue, remove } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js";

// ⚠️ [필독] 구글 파이어베이스 콘솔에서 발급받은 키값을 여기에 정확히 복사해 붙여넣으세요!
const firebaseConfig = {
    apiKey: "AIzaSyACQPQZwIWZS0v8dh2VcPBvp0VEXTzq0mw",
    authDomain: "my-shipping-system.firebaseapp.com",
    databaseURL: "https://my-shipping-system-default-rtdb.firebaseio.com",
    projectId: "my-shipping-system",
    storageBucket: "my-shipping-system.firebasestorage.app",
    messagingSenderId: "111295587884",
    appId: "1:111295587884:web:72f930fca8cb33d7eb5389"
};

// Firebase 초기화 실행
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// 전역 변수
let globalAllData = {}; // 서버에서 실시간 동기화되는 데이터 저장소
let currentDate = new Date();
let selectedDateStr = "";
let currentFilterCustomer = ""; 

// 초기 구동 및 이벤트 리스너 바인딩
window.onload = function() {
    startRealtimeSync(); // 서버 실시간 리스너 작동 개시
    
    // 버튼 및 이벤트 바인딩
    document.getElementById('bulkPasteInput').addEventListener('paste', handleBulkPaste);
    document.getElementById('scheduleForm').addEventListener('submit', saveRecord);
    document.getElementById('btnCloseModal').onclick = closeModal;
    document.getElementById('btnExport').onclick = exportToExcel;
    document.getElementById('pasteZone').onclick = () => document.getElementById('bulkPasteInput').focus();
    
    // 달력 화살표 이벤트 연결을 위해 전역 스코프 대신 직접 지정
    window.changeMonth = function(dir) {
        currentDate.setMonth(currentDate.getMonth() + dir);
        renderCalendar();
    };
};

// 1. 핵심: 클라우드 DB 전체 데이터 실시간 리스너 작동
function startRealtimeSync() {
    const shippingRef = ref(db, 'shippingData');
    
    // 이 함수는 서버에 값이 바뀌는 순간 '즉시' 알아서 팀원들 화면까지 실시간 재실행됩니다.
    onValue(shippingRef, (snapshot) => {
        globalAllData = snapshot.val() || {};
        renderCalendar();
        if(document.getElementById('dateModal').style.display === 'flex') {
            updateCustomerSidebar();
            updateRecordTable();
        }
    });
}

// 2. 메인 달력 드로잉
function renderCalendar() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    document.getElementById('calendarTitle').innerText = `${year}년 ${String(month + 1).padStart(2, '0')}월`;

    const firstDay = new Date(year, month, 1).getDay();
    const lastDate = new Date(year, month + 1, 0).getDate();
    const prevLastDate = new Date(year, month, 0).getDate();

    const grid = document.getElementById('calendarGrid');
    grid.innerHTML = "";

    const weekdays = ['일', '월', '화', '수', '목', '금', '토'];
    weekdays.forEach(w => {
        let div = document.createElement('div');
        div.className = 'weekday'; div.innerText = w; grid.appendChild(div);
    });

    for (let i = firstDay; i > 0; i--) {
        let dateStr = getFormattedDate(new Date(year, month - 1, prevLastDate - i + 1));
        createDayCell(prevLastDate - i + 1, true, dateStr, grid);
    }
    for (let i = 1; i <= lastDate; i++) {
        let dateStr = getFormattedDate(new Date(year, month, i));
        createDayCell(i, false, dateStr, grid);
    }
    const totalCells = grid.children.length - 7;
    const remaining = 42 - totalCells;
    for (let i = 1; i <= remaining; i++) {
        let dateStr = getFormattedDate(new Date(year, month + 1, i));
        createDayCell(i, true, dateStr, grid);
    }
}

function createDayCell(day, isOtherMonth, dateStr, grid) {
    let cell = document.createElement('div');
    cell.className = 'day-cell';
    if (isOtherMonth) cell.classList.add('other-month');

    cell.onclick = () => openModal(dateStr);

    let numDiv = document.createElement('div');
    numDiv.className = 'day-number'; numDiv.innerText = day; cell.appendChild(numDiv);

    // 실시간 감지 객체에서 개수 분석 후 뱃지 노출
    if (globalAllData[dateStr]) {
        let count = Object.keys(globalAllData[dateStr]).length;
        if(count > 0) {
            let sumDiv = document.createElement('div');
            sumDiv.className = 'day-summary'; sumDiv.innerText = `📦 출고 ${count}건`;
            cell.appendChild(sumDiv);
        }
    }
    grid.appendChild(cell);
}

// 3. 모달 팝업 액션
function openModal(dateStr) {
    selectedDateStr = dateStr;
    currentFilterCustomer = "";
    document.getElementById('selectedDateText').innerText = `📅 ${dateStr} 출고 내역 상세조회`;
    document.getElementById('dateModal').style.display = 'flex';
    
    updateCustomerSidebar();
    updateRecordTable();
    resetForm();
}

function closeModal() {
    document.getElementById('dateModal').style.display = 'none';
    renderCalendar();
}

// 4. 데이터 리스트 반환 유틸 (객체 -> 내림차순 배열 정렬 구조 변경)
function getSortedRecordsList() {
    let dayDataObj = globalAllData[selectedDateStr] || {};
    let recordsArr = [];
    
    // Firebase 특성상 고유 ID키를 가진 객체 구조이므로 가공
    for(let key in dayDataObj) {
        recordsArr.push({
            fbKey: key,
            name: dayDataObj[key].name,
            qty: dayDataObj[key].qty,
            customer: dayDataObj[key].customer
        });
    }
    // 고객사 기준 역순 정렬
    recordsArr.sort((a, b) => b.customer.localeCompare(a.customer, 'ko'));
    return recordsArr;
}

// 5. 사이드바 고객사 갱신
function updateCustomerSidebar() {
    let sidebarList = document.getElementById('customerFilterList');
    sidebarList.innerHTML = "";

    let records = getSortedRecordsList();
    
    let customers = [];
    records.forEach(r => {
        if (!customers.includes(r.customer)) customers.push(r.customer);
    });
    customers.sort((a, b) => b.localeCompare(a, 'ko'));

    let allLi = document.createElement('li');
    allLi.innerText = `🔄 전체보기 (${records.length}건)`;
    if (currentFilterCustomer === "") allLi.className = "active";
    allLi.onclick = () => {
        currentFilterCustomer = "";
        updateCustomerSidebar();
        updateRecordTable();
    };
    sidebarList.appendChild(allLi);

    customers.forEach(cust => {
        let count = records.filter(r => r.customer === cust).length;
        let li = document.createElement('li');
        li.innerText = `${cust} (${count}건)`;
        li.title = cust;
        if (currentFilterCustomer === cust) li.className = "active";
        
        li.onclick = () => {
            currentFilterCustomer = cust;
            updateCustomerSidebar();
            updateRecordTable();
        };
        sidebarList.appendChild(li);
    });
}

// 6. 테이블 화면 업데이트
function updateRecordTable() {
    let tbody = document.getElementById('recordTableBody');
    tbody.innerHTML = "";
    
    let records = getSortedRecordsList();
    let displayRecords = currentFilterCustomer ? records.filter(r => r.customer === currentFilterCustomer) : records;

    if (displayRecords.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #9ca3af; padding: 40px;">등록된 출고 내역이 없습니다. 위의 박스에 엑셀 데이터를 붙여넣어 보세요.</td></tr>`;
        return;
    }

    displayRecords.forEach((rec) => {
        let tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight: 500; color:#1e293b;">${rec.name}</td>
            <td><span style="background-color:#f1f5f9; padding: 4px 8px; border-radius:4px; font-weight:bold;">${rec.qty}</span></td>
            <td style="color: var(--primary-color); font-weight: bold;">${rec.customer}</td>
            <td>
                <button class="btn btn-sm btn-secondary" id="edit-${rec.fbKey}">수정</button>
                <button class="btn btn-sm btn-close" id="del-${rec.fbKey}">삭제</button>
            </td>
        `;
        tbody.appendChild(tr);

        // 이벤트 리스너 할당
        document.getElementById(`edit-${rec.fbKey}`).onclick = () => editRecord(rec);
        document.getElementById(`del-${rec.fbKey}`).onclick = () => deleteRecord(rec.fbKey);
    });
}

// 7. 대량 데이터 엑셀 복사-붙여넣기 파싱 및 Firebase 일괄 업로드
function handleBulkPaste(e) {
    e.preventDefault();
    let clipboardData = e.clipboardData || window.clipboardData;
    let pastedData = clipboardData.getData('Text');
    if (!pastedData) return;

    let lines = pastedData.split('\n');
    let parsedCount = 0;

    lines.forEach(line => {
        if (!line.trim()) return;
        let tokens = line.split(/\t/); // 엑셀 탭 문자 식별
        
        if (tokens.length >= 3) {
            let name = tokens[0].trim();
            let qty = parseInt(tokens[1].trim());
            let customer = tokens[2].trim();

            if (name && !isNaN(qty) && customer) {
                // 클라우드 데이터베이스에 실시간 적재 처리 (push 함수가 자동으로 고유 ID를 부여)
                const newRecordRef = push(ref(db, `shippingData/${selectedDateStr}`));
                set(newRecordRef, { name, qty, customer });
                parsedCount++;
            }
        }
    });

    if (parsedCount > 0) {
        alert(`클라우드 서버로 ${parsedCount}건의 데이터를 실시간 업로드했습니다! 팀원 화면에도 즉시 반영됩니다.`);
    } else {
        alert("데이터 파싱 실패. [품목명] [수량] [고객사] 세 열을 드래그했는지 확인하세요.");
    }
}

// 8. 수동 단건 저장 및 갱신 수정 처리
function saveRecord(e) {
    e.preventDefault();
    let name = document.getElementById('itemName').value.trim();
    let qty = parseInt(document.getElementById('quantity').value);
    let customer = document.getElementById('customer').value.trim();
    let editKey = document.getElementById('editKey').value;

    if (editKey) {
        // 기존 행 덮어쓰기 업데이트
        set(ref(db, `shippingData/${selectedDateStr}/${editKey}`), { name, qty, customer });
    } else {
        // 새로운 행 추가
        const newRecordRef = push(ref(db, `shippingData/${selectedDateStr}`));
        set(newRecordRef, { name, qty, customer });
    }
    resetForm();
}

function editRecord(rec) {
    document.getElementById('itemName').value = rec.name;
    document.getElementById('quantity').value = rec.qty;
    document.getElementById('customer').value = rec.customer;
    document.getElementById('editKey').value = rec.fbKey;
}

function deleteRecord(key) {
    if(confirm("이 출고 내역을 클라우드 서버에서 영구히 삭제할까요?")) {
        remove(ref(db, `shippingData/${selectedDateStr}/${key}`));
    }
}

// 9. 정렬 형태 그대로 엑셀 저장
function exportToExcel() {
    let records = getSortedRecordsList();
    let displayRecords = currentFilterCustomer ? records.filter(r => r.customer === currentFilterCustomer) : records;

    if(displayRecords.length === 0) return alert("데이터가 없습니다.");

    let csvContent = "\uFEFF"; 
    csvContent += `출고 일자,${selectedDateStr}\n\n`;
    csvContent += "품목명,수량,고객사\n";

    displayRecords.forEach(rec => {
        let safeName = rec.name.replace(/,/g, " ");
        csvContent += `${safeName},${rec.qty},${rec.customer}\n`;
    });

    let blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    let link = document.createElement("a");
    link.setAttribute("href", URL.createObjectURL(blob));
    link.setAttribute("download", `출고현황_${selectedDateStr}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function getFormattedDate(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function resetForm() {
    document.getElementById('scheduleForm').reset();
    document.getElementById('editKey').value = "";
}