import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import {
  getDatabase,
  ref,
  set,
  push,
  onValue,
  remove,
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js";

// 🔑 [설정] 시스템 접속 비밀번호
const ACCESS_PASSWORD = "1234";

// ⚠️ 파이어베이스 설정 정보
const firebaseConfig = {
  apiKey: "AIzaSyACQPQZwIWZS0v8dh2VcPBvp0VEXTzq0mw",
  authDomain: "my-shipping-system.firebaseapp.com",
  databaseURL: "https://my-shipping-system-default-rtdb.firebaseio.com",
  projectId: "my-shipping-system",
  storageBucket: "my-shipping-system.firebasestorage.app",
  messagingSenderId: "111295587884",
  appId: "1:111295587884:web:72f930fca8cb33d7eb5389",
};

// Firebase 초기화
const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// 전역 변수
let globalAllData = {}; // 서버에서 실시간 동기화되는 데이터 저장소
let currentDate = new Date();
let selectedDateStr = "";
let currentFilterCustomer = "";

// 초기 구동 및 이벤트 리스너 바인딩
window.onload = function () {
  // 🔒 로그인 인증 여부 판단
  if (sessionStorage.getItem("isAuthorized") === "true") {
    showMainContent();
  } else {
    showLoginOverlay();
  }

  // 로그인 관련 이벤트 연결
  document
    .getElementById("loginForm")
    .addEventListener("submit", checkPassword);
  document.getElementById("btnLogout").onclick = handleLogout;

  // 기존 버튼 및 이벤트 바인딩 유지
  document
    .getElementById("bulkPasteInput")
    .addEventListener("paste", handleBulkPaste);
  document
    .getElementById("scheduleForm")
    .addEventListener("submit", saveRecord);
  document.getElementById("btnCloseModal").onclick = closeModal;
  document.getElementById("btnExport").onclick = exportToExcel;
  document.getElementById("pasteZone").onclick = () =>
    document.getElementById("bulkPasteInput").focus();

  // 달력 화살표 이벤트 연결
  document.getElementById("btnPrevMonth").onclick = () => changeMonth(-1);
  document.getElementById("btnNextMonth").onclick = () => changeMonth(1);

  // 🟢 엑셀 동기화 버튼 및 파일 선택 이벤트 연결
  const btnUploadExcel = document.getElementById("btnUploadExcel");
  const excelFileInput = document.getElementById("excelFileInput");

  if (btnUploadExcel && excelFileInput) {
    btnUploadExcel.onclick = () => excelFileInput.click();
    excelFileInput.onchange = handleExcelSync;
  }

  // 전역 스코프 유지용
  window.changeMonth = function (dir) {
    currentDate.setMonth(currentDate.getMonth() + dir);
    renderCalendar();
  };
};

/* ==========================================================================
   🔒 접속 비밀번호 보안 제어 영역
   ========================================================================== */

function checkPassword(e) {
  e.preventDefault();
  const inputPw = document.getElementById("accessPassword").value;
  const errorMsg = document.getElementById("loginError");

  if (inputPw === ACCESS_PASSWORD) {
    sessionStorage.setItem("isAuthorized", "true");
    errorMsg.style.display = "none";
    showMainContent();
  } else {
    errorMsg.style.display = "block";
    document.getElementById("accessPassword").value = "";
    document.getElementById("accessPassword").focus();
  }
}

function showMainContent() {
  document.getElementById("loginOverlay").style.display = "none";
  document.getElementById("mainContent").style.display = "block";
  startRealtimeSync();
}

function showLoginOverlay() {
  document.getElementById("loginOverlay").style.display = "flex";
  document.getElementById("mainContent").style.display = "none";
}

function handleLogout() {
  if (confirm("출고 시스템에서 로그아웃하시겠습니까?")) {
    sessionStorage.removeItem("isAuthorized");
    location.reload();
  }
}

/* ==========================================================================
   📊 [신규] 엑셀 자동 동기화 관련 핵심 파싱 기능 ('발주현황_LIST' 전용)
   ========================================================================== */

//오늘 날짜를 YYYY-MM-DD 형식으로 구함
function getTodayDateString() {
  const today = new Date();
  return getFormattedDate(today);
}

// 엑셀 날짜 데이터를 표준 YYYY-MM-DD 문자열로 변환하는 유틸리티
function parseExcelDate(cellValue) {
  if (!cellValue) return null;

  // 1) 엑셀 날짜 일련번호(Serial Number)인 경우 (예: 45290)
  if (typeof cellValue === "number") {
    if (typeof XLSX !== "undefined" && XLSX.SSF) {
      const parsed = XLSX.SSF.parse_date_code(cellValue);
      if (parsed) {
        const y = parsed.y;
        const m = String(parsed.m).padStart(2, "0");
        const d = String(parsed.d).padStart(2, "0");
        return `${y}-${m}-${d}`;
      }
    }
    // fallback 계산
    const dateObj = new Date(Math.round((cellValue - 25569) * 86400 * 1000));
    return getFormattedDate(dateObj);
  }

  // 2) 문자열 형태인 경우 (예: "2026.08.01", "2026/08/01", "2026-08-01")
  const str = String(cellValue)
    .trim()
    .replace(/[\.\/]/g, "-");
  const match = str.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : null;
}

// 엑셀 파일 읽기 및 Firebase 전송 핸들러 (진단 로그 강화 버전)
function handleExcelSync(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
        console.error("❌ XLSX 라이브러리가 로드되지 않았습니다!");
        alert("엑셀 파싱 라이브러리(SheetJS)가 로드되지 않았습니다.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            console.log("=== 📄 1. 엑셀 파일의 모든 시트 이름 ===");
            console.log(workbook.SheetNames);

            const targetSheetName = '발주현황_LIST';
            if (!workbook.SheetNames.includes(targetSheetName)) {
                console.error(`❌ '${targetSheetName}' 시트를 찾을 수 없습니다!`);
                alert(`❌ 파일 내에 '${targetSheetName}' 시트가 존재하지 않습니다.`);
                e.target.value = "";
                return;
            }

            const worksheet = workbook.Sheets[targetSheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

            console.log("=== 📊 2. 읽어온 전체 행(Row) 개수 ===", rows.length);

            if (rows.length < 6) {
                console.warn("⚠️ 행 개수가 6개 미만입니다. 데이터가 부족합니다.");
                alert("엑셀 파일에 데이터(6번째 행 이후)가 존재하지 않습니다.");
                e.target.value = "";
                return;
            }

            // 📌 6번째 행(index 5)의 전체 데이터 확인
            const sampleRow = rows[5]; 
            console.log("=== 🔍 3. 6번째 행(데이터 첫 줄)의 전체 열 데이터 ===");
            console.log(sampleRow);

            // 열 인덱스 정의 (J=9, M=12, U=20, W=22)
            const COL_CUSTOMER = 8;  
            const COL_ITEM     = 11; 
            const COL_QTY      = 19; 
            const COL_DATE     = 21; 

            console.log("=== 🎯 4. 지정한 열에서 뽑아낸 값 ===");
            console.log("J열 (Index 9, 고객사) :", sampleRow[COL_CUSTOMER]);
            console.log("M열 (Index 12, 품목명):", sampleRow[COL_ITEM]);
            console.log("U열 (Index 20, 수량)  :", sampleRow[COL_QTY]);
            console.log("W열 (Index 22, 날짜)  :", sampleRow[COL_DATE]);

            const todayStr = getTodayDateString();
            console.log("📅 기준 오늘 날짜:", todayStr);

            let addedCount = 0;
            let skippedCount = 0;
            let failReasons = [];

            for (let i = 5; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;

                const rawDate = row[COL_DATE];
                const rawCustomer = row[COL_CUSTOMER];
                const rawItem = row[COL_ITEM];
                const rawQty = row[COL_QTY];

                const dateStr = parseExcelDate(rawDate);
                const customer = rawCustomer ? String(rawCustomer).trim() : "";
                const name = rawItem ? String(rawItem).trim() : "";
                const qty = parseInt(rawQty, 10);

                if (dateStr && name && customer && !isNaN(qty)) {
                    if (dateStr >= todayStr) {
                        const newRecordRef = push(ref(db, `shippingData/${dateStr}`));
                        set(newRecordRef, { name, qty, customer });
                        addedCount++;
                    } else {
                        skippedCount++;
                    }
                } else {
                    // 유효성 통과 실패 이유 저장 (최대 3건만)
                    if (failReasons.length < 3) {
                        failReasons.push({
                            rowNum: i + 1,
                            dateStr, name, customer, qty,
                            rawDate, rawCustomer, rawItem, rawQty
                        });
                    }
                }
            }

            console.log("=== 💡 5. 동기화 결과 요약 ===");
            console.log(`성공: ${addedCount}건, 과거데이터 스킵: ${skippedCount}건, 파싱 실패: ${failReasons.length}건 이상`);
            if (failReasons.length > 0) {
                console.log("❌ 파싱 실패 예시 (상위 3건):", failReasons);
            }

            alert(`✅ 엑셀 동기화 완료!\n- 동기화 성공: ${addedCount}건\n- 과거 데이터 제외: ${skippedCount}건`);
            
        } catch (err) {
            console.error("💥 엑셀 처리 중 오류 발생:", err);
            alert("엑셀 파일을 처리하는 중 오류가 발생했습니다.");
        } finally {
            e.target.value = "";
        }
    };

    reader.readAsArrayBuffer(file);
}

/* ==========================================================================
   📡 실시간 데이터 연동 및 업무 로직
   ========================================================================== */

function startRealtimeSync() {
  const shippingRef = ref(db, "shippingData");

  onValue(shippingRef, (snapshot) => {
    globalAllData = snapshot.val() || {};
    renderCalendar();
    if (document.getElementById("dateModal").style.display === "flex") {
      updateCustomerSidebar();
      updateRecordTable();
    }
  });
}

function renderCalendar() {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();
  document.getElementById("calendarTitle").innerText =
    `${year}년 ${String(month + 1).padStart(2, "0")}월`;

  const firstDay = new Date(year, month, 1).getDay();
  const lastDate = new Date(year, month + 1, 0).getDate();
  const prevLastDate = new Date(year, month, 0).getDate();

  const grid = document.getElementById("calendarGrid");
  grid.innerHTML = "";

  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  weekdays.forEach((w) => {
    let div = document.createElement("div");
    div.className = "weekday";
    div.innerText = w;
    grid.appendChild(div);
  });

  for (let i = firstDay; i > 0; i--) {
    let dateStr = getFormattedDate(
      new Date(year, month - 1, prevLastDate - i + 1),
    );
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
  let cell = document.createElement("div");
  cell.className = "day-cell";
  if (isOtherMonth) cell.classList.add("other-month");

  cell.onclick = () => openModal(dateStr);

  let numDiv = document.createElement("div");
  numDiv.className = "day-number";
  numDiv.innerText = day;
  cell.appendChild(numDiv);

  if (globalAllData[dateStr]) {
    let count = Object.keys(globalAllData[dateStr]).length;
    if (count > 0) {
      let sumDiv = document.createElement("div");
      sumDiv.className = "day-summary";
      sumDiv.innerText = `📦 출고 ${count}건`;
      cell.appendChild(sumDiv);
    }
  }
  grid.appendChild(cell);
}

function openModal(dateStr) {
  selectedDateStr = dateStr;
  currentFilterCustomer = "";
  document.getElementById("selectedDateText").innerText =
    `📅 ${dateStr} 출고 내역 상세조회`;
  document.getElementById("dateModal").style.display = "flex";

  updateCustomerSidebar();
  updateRecordTable();
  resetForm();
}

function closeModal() {
  document.getElementById("dateModal").style.display = "none";
  renderCalendar();
}

function getSortedRecordsList() {
  let dayDataObj = globalAllData[selectedDateStr] || {};
  let recordsArr = [];

  for (let key in dayDataObj) {
    recordsArr.push({
      fbKey: key,
      name: dayDataObj[key].name,
      qty: dayDataObj[key].qty,
      customer: dayDataObj[key].customer,
    });
  }
  recordsArr.sort((a, b) => b.customer.localeCompare(a.customer, "ko"));
  return recordsArr;
}

function updateCustomerSidebar() {
  let sidebarList = document.getElementById("customerFilterList");
  sidebarList.innerHTML = "";

  let records = getSortedRecordsList();

  let customers = [];
  records.forEach((r) => {
    if (!customers.includes(r.customer)) customers.push(r.customer);
  });
  customers.sort((a, b) => b.localeCompare(a, "ko"));

  let allLi = document.createElement("li");
  allLi.innerText = `🔄 전체보기 (${records.length}건)`;
  if (currentFilterCustomer === "") allLi.className = "active";
  allLi.onclick = () => {
    currentFilterCustomer = "";
    updateCustomerSidebar();
    updateRecordTable();
  };
  sidebarList.appendChild(allLi);

  customers.forEach((cust) => {
    let count = records.filter((r) => r.customer === cust).length;
    let li = document.createElement("li");
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

function updateRecordTable() {
  let tbody = document.getElementById("recordTableBody");
  tbody.innerHTML = "";

  let records = getSortedRecordsList();
  let displayRecords = currentFilterCustomer
    ? records.filter((r) => r.customer === currentFilterCustomer)
    : records;

  if (displayRecords.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #9ca3af; padding: 40px;">등록된 출고 내역이 없습니다. 위의 박스에 엑셀 데이터를 붙여넣어 보세요.</td></tr>`;
    return;
  }

  displayRecords.forEach((rec) => {
    let tr = document.createElement("tr");
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

    document.getElementById(`edit-${rec.fbKey}`).onclick = () =>
      editRecord(rec);
    document.getElementById(`del-${rec.fbKey}`).onclick = () =>
      deleteRecord(rec.fbKey);
  });
}

function handleBulkPaste(e) {
  e.preventDefault();
  let clipboardData = e.clipboardData || window.clipboardData;
  let pastedData = clipboardData.getData("Text");
  if (!pastedData) return;

  let lines = pastedData.split("\n");
  let parsedCount = 0;

  lines.forEach((line) => {
    if (!line.trim()) return;
    let tokens = line.split(/\t/);

    if (tokens.length >= 3) {
      let name = tokens[0].trim();
      let qty = parseInt(tokens[1].trim());
      let customer = tokens[2].trim();

      if (name && !isNaN(qty) && customer) {
        const newRecordRef = push(ref(db, `shippingData/${selectedDateStr}`));
        set(newRecordRef, { name, qty, customer });
        parsedCount++;
      }
    }
  });

  if (parsedCount > 0) {
    alert(
      `클라우드 서버로 ${parsedCount}건의 데이터를 실시간 업로드했습니다! 팀원 화면에도 즉시 반영됩니다.`,
    );
  } else {
    alert(
      "데이터 파싱 실패. [품목명] [수량] [고객사] 세 열을 드래그했는지 확인하세요.",
    );
  }
}

function saveRecord(e) {
  e.preventDefault();
  let name = document.getElementById("itemName").value.trim();
  let qty = parseInt(document.getElementById("quantity").value);
  let customer = document.getElementById("customer").value.trim();
  let editKey = document.getElementById("editKey").value;

  if (editKey) {
    set(ref(db, `shippingData/${selectedDateStr}/${editKey}`), {
      name,
      qty,
      customer,
    });
  } else {
    const newRecordRef = push(ref(db, `shippingData/${selectedDateStr}`));
    set(newRecordRef, { name, qty, customer });
  }
  resetForm();
}

function editRecord(rec) {
  document.getElementById("itemName").value = rec.name;
  document.getElementById("quantity").value = rec.qty;
  document.getElementById("customer").value = rec.customer;
  document.getElementById("editKey").value = rec.fbKey;
}

function deleteRecord(key) {
  if (confirm("이 출고 내역을 클라우드 서버에서 영구히 삭제할까요?")) {
    remove(ref(db, `shippingData/${selectedDateStr}/${key}`));
  }
}

function exportToExcel() {
  let records = getSortedRecordsList();
  let displayRecords = currentFilterCustomer
    ? records.filter((r) => r.customer === currentFilterCustomer)
    : records;

  if (displayRecords.length === 0) return alert("데이터가 없습니다.");

  let csvContent = "\uFEFF";
  csvContent += `출고 일자,${selectedDateStr}\n\n`;
  csvContent += "품목명,수량,고객사\n";

  displayRecords.forEach((rec) => {
    let safeName = rec.name.replace(/,/g, " ");
    csvContent += `${safeName},${rec.qty},${rec.customer}\n`;
  });

  let blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  let link = document.createElement("a");
  link.setAttribute("href", URL.createObjectURL(blob));
  link.setAttribute("download", `출고현황_${selectedDateStr}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function getFormattedDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function resetForm() {
  document.getElementById("scheduleForm").reset();
  document.getElementById("editKey").value = "";
}
