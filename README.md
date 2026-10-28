<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>동맹원 관리 시스템</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Malgun Gothic', sans-serif; background-color: #1e1e2f; color: #ffffff; padding: 20px; max-width: 900px; margin: auto; }
        .container, .dashboard { background-color: #2a2a40; padding: 20px; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.3); margin-bottom: 20px; }
        h2 { text-align: center; color: #4facfe; }
        .form-group { margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; }
        input, select { padding: 8px; border-radius: 5px; border: 1px solid #4facfe; background-color: #1e1e2f; color: white; width: 60%; }
        
        /* 수정 모드일 때 입력칸 배경색 처리 (잠금 표시) */
        input:read-only { background-color: #555; cursor: not-allowed; border: 1px solid #777; color: #ccc; } 
        
        button { width: 100%; padding: 15px; border: none; color: white; font-weight: bold; font-size: 16px; border-radius: 5px; cursor: pointer; margin-top: 10px; }
        .btn-submit { background: linear-gradient(to right, #4facfe 0%, #00f2fe 100%); }
        .btn-cancel { background: #9e9e9e; display: none; }
        .btn-load { background: #ff9800; margin-bottom: 10px; }
        .btn-excel { background: #4caf50; }
        button:hover { opacity: 0.9; }
        
        /* 표(테이블) 디자인 완벽 고정 (하얀 배경에 검은 글씨) */
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th { background-color: #4facfe; color: #000000; padding: 10px; border: 1px solid #444; font-weight: bold; }
        td { background-color: #ffffff; color: #000000; padding: 10px; border: 1px solid #cccccc; text-align: center; font-size: 14px; font-weight: 500; }
        
        /* 점수 뱃지 및 작은 수정 버튼 */
        .score-badge { color: #ffffff; padding: 4px 12px; border-radius: 5px; font-weight: bold; display: inline-block; }
        .btn-edit { background-color: #E91E63; color: white; border: none; padding: 5px 10px; border-radius: 3px; cursor: pointer; font-size: 12px; width: auto; margin: 0;}
        .action-buttons { display: flex; gap: 10px; }
    </style>
</head>
<body>

<div class="container">
    <h2 id="formTitle">동맹원 점수 신규 입력</h2>
    <form id="allianceForm">
        <!-- actionType을 이용해 신규(append)인지 수정(update)인지 구분 -->
        <input type="hidden" id="actionType" value="append">

        <div class="form-group"><label>아이디</label><input type="text" id="memberId" required></div>
        <div class="form-group"><label>닉네임</label><input type="text" id="nickname" required></div>
        <div class="form-group">
            <label>직책</label>
            <select id="role">
                <option value="R5">R5 </option><option value="R4">R4 </option>
                <option value="R3">R3</option><option value="R2">R2</option><option value="R1">R1 </option>
            </select>
        </div>
        <hr style="border-color:#444; margin: 20px 0;">
        <div class="form-group">
            <label>SVS (서버전)</label>
            <select id="svs"><option value="1">참여 (+1점)</option><option value="0" selected>미참여 (0점)</option><option value="-1">노쇼 (-1점)</option></select>
        </div>
        <div class="form-group">
            <label>무기공장쟁탈전</label>
            <select id="weapon"><option value="1">참여 (+1점)</option><option value="0" selected>미참여 (0점)</option><option value="-1">노쇼 (-1점)</option></select>
        </div>
        <div class="form-group">
            <label>협곡</label>
            <select id="canyon"><option value="1">참여 (+1점)</option><option value="0" selected>미참여 (0점)</option><option value="-1">노쇼 (-1점)</option></select>
        </div>
        
        <hr style="border-color:#444; margin: 20px 0;">
        <div class="form-group">
            <label>명함점수 (비교용)</label>
            <input type="number" id="namecard" value="0" placeholder="점수 입력">
        </div>
        <div class="form-group">
            <label>SVS 준비단계 점수 (비교용)</label>
            <input type="number" id="svsPrep" value="0" placeholder="점수 입력">
        </div>

        <button type="submit" class="btn-submit" id="submitBtn">구글 시트에 신규 저장</button>
        <button type="button" class="btn-cancel" id="cancelBtn" onclick="cancelEdit()">수정 취소하기</button>
    </form>
</div>

<div class="dashboard">
    <h2>전체 동맹원 현황</h2>
    
    <div class="action-buttons">
        <button class="btn-load" onclick="loadData()" id="loadBtn">데이터 불러오기 (새로고침)</button>
        <button class="btn-excel" onclick="downloadExcel()">엑셀 파일로 다운로드</button>
    </div>

    <div style="overflow-x: auto;">
        <table id="dataTable">
            <thead>
                <tr>
                    <th>아이디</th>
                    <th>닉네임</th>
                    <th>직책</th>
                    <th>SVS</th>
                    <th>무기공장</th>
                    <th>협곡</th>
                    <th>명함점수</th>
                    <th>SVS준비</th>
                    <th>총점</th>
                    <th>기록일시</th>
                    <th>관리</th> <!-- 수정 버튼을 위한 열 -->
                </tr>
            </thead>
            <tbody id="allRecordTable">
                <tr><td colspan="11">데이터를 불러와 주세요.</td></tr>
            </tbody>
        </table>
    </div>
</div>

<script>
    // ★ 알려주신 새로운 구글 앱스 스크립트 '웹 앱 URL' 적용 완료 ★
    const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbxNSGW65c-7JaqkWbCG3JakHgTkIxL3EZgOIXTete2P6QyvQswxgq1FsYmA2oYGAi8/exec";

    let globalSheetData = []; 

    document.getElementById('allianceForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const btn = document.getElementById('submitBtn');
        const isUpdate = document.getElementById('actionType').value === "update";
        btn.innerText = "데이터 전송 중...";
        btn.disabled = true;

        const svsScore = parseInt(document.getElementById('svs').value) || 0;
        const weaponScore = parseInt(document.getElementById('weapon').value) || 0;
        const canyonScore = parseInt(document.getElementById('canyon').value) || 0;
        const namecardScore = parseInt(document.getElementById('namecard').value) || 0;
        const svsPrepScore = parseInt(document.getElementById('svsPrep').value) || 0;
        
        // 총점 계산 (명함점수, SVS준비 제외)
        const total = svsScore + weaponScore + canyonScore;

        const data = {
            action: document.getElementById('actionType').value,
            id: document.getElementById('memberId').value,
            nickname: document.getElementById('nickname').value,
            role: document.getElementById('role').value,
            svs: svsScore,
            weapon: weaponScore,
            canyon: canyonScore,
            namecard: namecardScore,
            svsPrep: svsPrepScore,
            totalScore: total
        };

        fetch(GOOGLE_SHEET_URL, {
            method: 'POST',
            mode: 'no-cors', 
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        }).then(() => {
            alert(isUpdate ? '정보가 성공적으로 수정되었습니다!' : '신규 저장이 완료되었습니다!');
            cancelEdit(); 
            loadData();   
        }).catch(err => {
            alert('오류가 발생했습니다.');
            cancelEdit();
        });
    });

    function loadData() {
        const btn = document.getElementById('loadBtn');
        const tbody = document.getElementById('allRecordTable');
        btn.innerText = "불러오는 중...";
        tbody.innerHTML = '<tr><td colspan="11">구글 시트에서 데이터를 읽고 있습니다...</td></tr>';

        fetch(GOOGLE_SHEET_URL)
            .then(response => response.text())
            .then(text => {
                try {
                    const data = JSON.parse(text);
                    globalSheetData = data; 
                    tbody.innerHTML = ""; 
                    
                    if(data.length <= 1) {
                        tbody.innerHTML = '<tr><td colspan="11">기록된 데이터가 없습니다.</td></tr>';
                        btn.innerText = "데이터 불러오기 (새로고침)";
                        return;
                    }

                    for(let i = 1; i < data.length; i++) {
                        let row = data[i];
                        if(!row || row.length === 0 || !row[0]) continue; 
                        
                        let namecard = row.length > 8 && row[6] !== undefined ? row[6] : 0;
                        let svsPrep = row.length > 8 && row[7] !== undefined ? row[7] : 0;
                        let totalScore = row.length > 8 ? (row[8] !== undefined ? row[8] : 0) : (row[6] !== undefined ? row[6] : 0);
                        let dateStr = row.length > 8 ? (row[9] ? new Date(row[9]).toLocaleString('ko-KR') : "") : (row[7] ? new Date(row[7]).toLocaleString('ko-KR') : "");

                        let scoreNum = parseInt(totalScore, 10);
                        if(isNaN(scoreNum)) scoreNum = 0;

                        let badgeColor = scoreNum > 0 ? "#2196F3" : (scoreNum === 0 ? "#4CAF50" : "#F44336");

                        let tr = document.createElement('tr');
                        tr.innerHTML = `
                            <td>${row[0] !== undefined ? row[0] : ""}</td>
                            <td>${row[1] !== undefined ? row[1] : ""}</td>
                            <td>${row[2] !== undefined ? row[2] : ""}</td>
                            <td>${row[3] !== undefined ? row[3] : "0"}</td>
                            <td>${row[4] !== undefined ? row[4] : "0"}</td>
                            <td>${row[5] !== undefined ? row[5] : "0"}</td>
                            <td>${namecard}</td>
                            <td>${svsPrep}</td>
                            <td><span class="score-badge" style="background-color: ${badgeColor};">${scoreNum}</span></td>
                            <td style="font-size:11px;">${dateStr}</td>
                            <td><button class="btn-edit" onclick="editMember(${i})">수정</button></td>
                        `;
                        tbody.appendChild(tr);
                    }
                    btn.innerText = "데이터 불러오기 (새로고침)";
                } catch (e) {
                    console.error(e);
                    tbody.innerHTML = '<tr><td colspan="11" style="color:#ff5722;">데이터를 불러오지 못했습니다.</td></tr>';
                    btn.innerText = "데이터 불러오기 (새로고침)";
                }
            })
            .catch(err => {
                tbody.innerHTML = '<tr><td colspan="11">통신 오류가 발생했습니다.</td></tr>';
                btn.innerText = "데이터 불러오기 (새로고침)";
            });
    }

    function editMember(index) {
        let row = globalSheetData[index];
        
        document.getElementById('actionType').value = "update";
        document.getElementById('formTitle').innerText = "동맹원 점수 수정 모드";
        
        document.getElementById('memberId').value = row[0];
        document.getElementById('memberId').readOnly = true; 
        
        document.getElementById('nickname').value = row[1];
        document.getElementById('nickname').readOnly = true; 
        
        document.getElementById('role').value = row[2] || "R1";
        
        document.getElementById('svs').value = row[3] !== undefined ? row[3] : "0";
        document.getElementById('weapon').value = row[4] !== undefined ? row[4] : "0";
        document.getElementById('canyon').value = row[5] !== undefined ? row[5] : "0";
        
        document.getElementById('namecard').value = row.length > 8 && row[6] !== undefined ? row[6] : "0";
        document.getElementById('svsPrep').value = row.length > 8 && row[7] !== undefined ? row[7] : "0";

        let submitBtn = document.getElementById('submitBtn');
        submitBtn.innerText = "변경된 점수 저장하기 (구글 시트 덮어쓰기)";
        submitBtn.style.background = "linear-gradient(to right, #ff416c, #ff4b2b)"; 
        submitBtn.disabled = false;
        
        document.getElementById('cancelBtn').style.display = "block"; 
        window.scrollTo({top: 0, behavior: 'smooth'});
    }

    function cancelEdit() {
        document.getElementById('allianceForm').reset();
        
        document.getElementById('actionType').value = "append";
        document.getElementById('formTitle').innerText = "동맹원 점수 신규 입력";
        
        document.getElementById('memberId').readOnly = false; 
        document.getElementById('nickname').readOnly = false; 
        
        document.getElementById('namecard').value = 0;
        document.getElementById('svsPrep').value = 0;
        
        let submitBtn = document.getElementById('submitBtn');
        submitBtn.innerText = "구글 시트에 신규 저장";
        submitBtn.style.background = "linear-gradient(to right, #4facfe 0%, #00f2fe 100%)";
        submitBtn.disabled = false;
        
        document.getElementById('cancelBtn').style.display = "none";
    }

    function downloadExcel() {
        const table = document.getElementById("dataTable");
        const wb = XLSX.utils.table_to_book(table, {sheet: "동맹원기록"});
        XLSX.writeFile(wb, "동맹원_관리현황.xlsx");
    }

    window.onload = loadData;
</script>

</body>
</html>
