import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import {
  getDatabase,
  ref,
  set,
  push,
  onValue,
  remove,
  update, // 👈 Batch Update를 위해 추가
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
let currentSearchKeyword = ""; // 🔍 검색 키워드 변수 추가

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

  // 🔍 검색창 및 검색 초기화 버튼 이벤트 연동
  const searchInput = document.getElementById("searchInput");
  const btnResetSearch = document.getElementById("btnResetSearch");

  if (searchInput) {
    searchInput.addEventListener("input", (e) => {
      currentSearchKeyword = e.target.value.trim().toLowerCase();
      if (btnResetSearch) {
        btnResetSearch.style.display = currentSearchKeyword ? "inline-block" : "none";
      }
      renderCalendar();
    });
  }

  if (btnResetSearch) {
    btnResetSearch.onclick = () => {
      if (searchInput) searchInput.value = "";
      currentSearchKeyword = "";
      btnResetSearch.style.display = "none";
      renderCalendar();
    };
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
   📊 엑셀 자동 동기화 기능 ('발주현황_LIST' 전용, 단 1회 Batch Update 방식)
   ========================================================================== */

function getTodayDateString() {
  const today = new Date();
  return getFormattedDate(today);
}

function parseExcelDate(cellValue) {
  if (!cellValue) return null;

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
    const dateObj = new Date(Math.round((cellValue - 25569) * 86400 * 1000));
    return getFormattedDate(dateObj);
  }

  const str = String(cellValue)
    .trim()
    .replace(/[\.\/]/g, "-");
  const match = str.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : null;
}

async function handleExcelSync(e) {
  const file = e.target.files[0];
  if (!file) return;

  if (typeof XLSX === "undefined") {
    alert("엑셀 파싱 라이브러리(SheetJS)가 로드되지 않았습니다.");
    return;
  }

  const reader = new FileReader();

  reader.onload = async function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const targetSheetName = "발주현황_LIST";
      if (!workbook.SheetNames.includes(targetSheetName)) {
        alert(`❌ 파일 내에 '${targetSheetName}' 시트가 존재하지 않습니다.`);
        e.target.value = "";
        return;
      }

      const worksheet = workbook.Sheets[targetSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      if (rows.length <= 1) {
        alert("엑셀 파일에 데이터가 존재하지 않습니다.");
        e.target.value = "";
        return;
      }

      // 열 인덱스 지정: J열=9, M열=12, U열=20, W열=22
      const COL_CUSTOMER = 9;
      const COL_ITEM = 12;
      const COL_QTY = 20;
      const COL_DATE = 22;

      const todayStr = getTodayDateString();
      const validDataByDate = {};
      let skippedCount = 0;

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const dateStr = parseExcelDate(row[COL_DATE]);
        const customer = row[COL_CUSTOMER] ? String(row[COL_CUSTOMER]).trim() : "";
        const name = row[COL_ITEM] ? String(row[COL_ITEM]).trim() : "";
        const qty = parseInt(String(row[COL_QTY]).replace(/,/g, ""), 10);

        if (dateStr && name && customer && !isNaN(qty)) {
          if (dateStr >= todayStr) {
            if (!validDataByDate[dateStr]) {
              validDataByDate[dateStr] = [];
            }
            validDataByDate[dateStr].push({ name, qty, customer });
          } else {
            skippedCount++;
          }
        }
      }

      const targetDates = Object.keys(validDataByDate);
      if (targetDates.length === 0) {
        alert("동기화할 유효한 출고 데이터(오늘 이후)가 엑셀에 없습니다.");
        e.target.value = "";
        return;
      }

      // 🚀 반복문으로 단건 전송하지 않고, updates 객체에 모아서 단 1회 update 실행
      let updates = {};
      let totalAddedCount = 0;

      for (const dateStr of targetDates) {
        // 기존 날짜 항목 데이터 초기화
        updates[`shippingData/${dateStr}`] = null;

        const items = validDataByDate[dateStr];
        items.forEach((item) => {
          const newKey = push(ref(db, `shippingData/${dateStr}`)).key;
          updates[`shippingData/${dateStr}/${newKey}`] = item;
          totalAddedCount++;
        });
      }

      await update(ref(db), updates);

      alert(
        `✅ 엑셀 동기화 완료!\n- 최신 반영 데이터: ${totalAddedCount}건 (${targetDates.length}개 일자)\n- 과거 데이터 제외: ${skippedCount}건`
      );
    } catch (err) {
      console.error("엑셀 동기화 오류:", err);
      alert("엑셀 파일 처리 중 오류가 발생했습니다.");
    } finally {
      e.target.value = "";
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ==========================================================================
   📡 실시간 데이터 연동 및 달력 / 검색 렌더링
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
  document.getElementById("calendarTitle").innerText = `${year}년 ${String(
    month + 1
  ).padStart(2, "0")}월`;

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

  let totalMatchedDays = 0;
  let totalMatchedRecords = 0;

  const appendDay = (day, isOtherMonth, dateStr) => {
    let cell = document.createElement("div");
    cell.className = "day-cell";
    if (isOtherMonth) cell.classList.add("other-month");
    cell.onclick = () => openModal(dateStr);

    let numDiv = document.createElement("div");
    numDiv.className = "day-number";
    numDiv.innerText = day;
    cell.appendChild(numDiv);

    if (globalAllData[dateStr]) {
      const dayRecords = Object.values(globalAllData[dateStr]);
      const totalCount = dayRecords.length;

      if (totalCount > 0) {
        let matchedCount = 0;
        if (currentSearchKeyword) {
          matchedCount = dayRecords.filter(
            (r) =>
              (r.name && r.name.toLowerCase().includes(currentSearchKeyword)) ||
              (r.customer && r.customer.toLowerCase().includes(currentSearchKeyword))
          ).length;
        }

        let sumDiv = document.createElement("div");
        sumDiv.className = "day-summary";

        if (currentSearchKeyword) {
          if (matchedCount > 0) {
            cell.classList.add("search-matched");
            sumDiv.innerText = `🔍 검색 ${matchedCount}건`;

            if (!isOtherMonth) {
              totalMatchedDays++;
              totalMatchedRecords += matchedCount;
            }
          } else {
            sumDiv.innerText = `📦 출고 ${totalCount}건`;
            cell.classList.add("search-dimmed");
          }
        } else {
          sumDiv.innerText = `📦 출고 ${totalCount}건`;
        }

        cell.appendChild(sumDiv);
      }
    }
    grid.appendChild(cell);
  };

  for (let i = firstDay; i > 0; i--) {
    let dateStr = getFormattedDate(
      new Date(year, month - 1, prevLastDate - i + 1)
    );
    appendDay(prevLastDate - i + 1, true, dateStr);
  }
  for (let i = 1; i <= lastDate; i++) {
    let dateStr = getFormattedDate(new Date(year, month, i));
    appendDay(i, false, dateStr);
  }
  const remaining = 42 - (grid.children.length - 7);
  for (let i = 1; i <= remaining; i++) {
    let dateStr = getFormattedDate(new Date(year, month + 1, i));
    appendDay(i, true, dateStr);
  }

  const summaryEl = document.getElementById("searchResultSummary");
  if (summaryEl) {
    if (currentSearchKeyword) {
      summaryEl.innerText = `🔎 '${currentSearchKeyword}' 검색 결과: 총 ${totalMatchedDays}개 날짜에서 ${totalMatchedRecords}건 발견됨`;
    } else {
      summaryEl.innerText = "";
    }
  }
}

function openModal(dateStr) {
  selectedDateStr = dateStr;
  currentFilterCustomer = "";
  document.getElementById("selectedDateText").innerText = `📅 ${dateStr} 출고 내역 상세조회`;
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

  if (currentSearchKeyword) {
    displayRecords = displayRecords.filter(
      (r) =>
        (r.name && r.name.toLowerCase().includes(currentSearchKeyword)) ||
        (r.customer && r.customer.toLowerCase().includes(currentSearchKeyword))
    );
  }

  if (displayRecords.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #9ca3af; padding: 40px;">등록되었거나 일치하는 출고 내역이 없습니다.</td></tr>`;
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

    document.getElementById(`edit-${rec.fbKey}`).onclick = () => editRecord(rec);
    document.getElementById(`del-${rec.fbKey}`).onclick = () => deleteRecord(rec.fbKey);
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
    alert(`클라우드 서버로 ${parsedCount}건의 데이터를 실시간 업로드했습니다! 팀원 화면에도 즉시 반영됩니다.`);
  } else {
    alert("데이터 파싱 실패. [품목명] [수량] [고객사] 세 열을 드래그했는지 확인하세요.");
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