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
        button { width: 100%; padding: 15px; border: none; color: white; font-weight: bold; font-size: 16px; border-radius: 5px; cursor: pointer; margin-top: 10px; }
        .btn-submit { background: linear-gradient(to right, #4facfe 0%, #00f2fe 100%); }
        .btn-load { background: #ff9800; margin-bottom: 10px; }
        .btn-excel { background: #4caf50; }
        button:hover { opacity: 0.9; }
        
        /* 표(테이블) 배경색과 글자색 명확하게 고정 */
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th { background-color: #4facfe; color: #000000; padding: 10px; border: 1px solid #444; }
        td { background-color: #2a2a40; color: #ffffff; padding: 10px; border: 1px solid #555; text-align: center; font-size: 13px; }
        
        .score-badge { 
            color: #ffffff; 
            padding: 4px 12px; 
            border-radius: 5px; 
            font-weight: bold; 
            display: inline-block;
        }
        .action-buttons { display: flex; gap: 10px; }
    </style>
</head>
<body>

<div class="container">
    <h2>동맹원 점수 입력</h2>
    <form id="allianceForm">
        <div class="form-group"><label>아이디</label><input type="text" id="memberId" required></div>
        <div class="form-group"><label>닉네임</label><input type="text" id="nickname" required></div>
        <div class="form-group">
            <label>직책</label>
            <select id="role">
                <option value="R5">R5 (맹주)</option><option value="R4">R4 (간부)</option>
                <option value="R3">R3</option><option value="R2">R2</option><option value="R1">R1 (신입)</option>
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

        <button type="submit" class="btn-submit" id="submitBtn">구글 시트에 저장</button>
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
                </tr>
            </thead>
            <tbody id="allRecordTable">
                <tr><td colspan="10">데이터를 불러와 주세요.</td></tr>
            </tbody>
        </table>
    </div>
</div>

<script>
    const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbxeLvHvpmdSHcSbeGGDIaFL9Z1B3XwfZO5B7cv7F42RFJl3BjhPQMee4APQHIGtl3Oi/exec";

    document.getElementById('allianceForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const btn = document.getElementById('submitBtn');
        btn.innerText = "저장 중...";
        btn.disabled = true;

        const svsScore = parseInt(document.getElementById('svs').value) || 0;
        const weaponScore = parseInt(document.getElementById('weapon').value) || 0;
        const canyonScore = parseInt(document.getElementById('canyon').value) || 0;
        const namecardScore = parseInt(document.getElementById('namecard').value) || 0;
        const svsPrepScore = parseInt(document.getElementById('svsPrep').value) || 0;
        
        const total = svsScore + weaponScore + canyonScore;

        const data = {
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
            alert('저장 완료!');
            document.getElementById('allianceForm').reset();
            document.getElementById('namecard').value = 0;
            document.getElementById('svsPrep').value = 0;
            
            btn.innerText = "구글 시트에 저장";
            btn.disabled = false;
            loadData();
        }).catch(err => {
            alert('저장 오류가 발생했습니다.');
            btn.innerText = "구글 시트에 저장";
            btn.disabled = false;
        });
    });

    function loadData() {
        const btn = document.getElementById('loadBtn');
        const tbody = document.getElementById('allRecordTable');
        btn.innerText = "불러오는 중...";
        tbody.innerHTML = '<tr><td colspan="10">구글 시트에서 데이터를 읽고 있습니다...</td></tr>';

        fetch(GOOGLE_SHEET_URL)
            .then(response => response.text())
            .then(text => {
                try {
                    const data = JSON.parse(text);
                    tbody.innerHTML = ""; 
                    
                    if(data.length <= 1) {
                        tbody.innerHTML = '<tr><td colspan="10">기록된 데이터가 없습니다.</td></tr>';
                        btn.innerText = "데이터 불러오기 (새로고침)";
                        return;
                    }

                    for(let i = 1; i < data.length; i++) {
                        let row = data[i];
                        if(!row || row.length === 0 || !row[0]) continue; 
                        
                        let namecard = 0;
                        let svsPrep = 0;
                        let totalScore = 0;
                        let dateStr = "";

                        if (row.length <= 8) {
                            totalScore = row[6] !== undefined ? row[6] : 0;
                            dateStr = row[7] ? new Date(row[7]).toLocaleString('ko-KR') : "";
                        } else {
                            namecard = row[6] !== undefined ? row[6] : 0;
                            svsPrep = row[7] !== undefined ? row[7] : 0;
                            totalScore = row[8] !== undefined ? row[8] : 0;
                            dateStr = row[9] ? new Date(row[9]).toLocaleString('ko-KR') : "";
                        }

                        // 총점 숫자로 변환 후 색상 판별
                        let scoreNum = parseInt(totalScore, 10);
                        if(isNaN(scoreNum)) scoreNum = 0;

                        let badgeColor = "";
                        if (scoreNum > 0) {
                            badgeColor = "#2196F3"; // 양수: 파란색
                        } else if (scoreNum === 0) {
                            badgeColor = "#4CAF50"; // 0: 초록색
                        } else {
                            badgeColor = "#F44336"; // 음수: 빨간색
                        }

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
                        `;
                        tbody.appendChild(tr);
                    }
                    btn.innerText = "데이터 불러오기 (새로고침)";
                } catch (e) {
                    console.error(e);
                    tbody.innerHTML = '<tr><td colspan="10" style="color:#ff5722;">데이터를 불러오지 못했습니다.</td></tr>';
                    btn.innerText = "데이터 불러오기 (새로고침)";
                }
            })
            .catch(err => {
                tbody.innerHTML = '<tr><td colspan="10">통신 오류가 발생했습니다.</td></tr>';
                btn.innerText = "데이터 불러오기 (새로고침)";
            });
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
