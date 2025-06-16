function uploadCSV() {
    const fileInput = document.getElementById('csv_file');
    const file = fileInput.files[0];
    if (!file) return alert("請選擇檔案");

    const reader = new FileReader();
    reader.onload = function (e) {
        const csvText = e.target.result;
        const lines = csvText.split('\n').map(line => line.trim());
        const rows = lines.map(line => line.split(','));

        const data = {
            title: '',
            remarks: '',
            weeks: [],
            periods: [],
            times: [],
            breaks: [],
            content: [],
            note: ''
        };

        let current = '';
        for (const row of rows) {
            if (!row[0]) continue;
            const key = row[0].toLowerCase();
            const values = row.slice(1).map(v => v.trim());
            if (key.startsWith('[title]')) data.title = values[0];
            else if (key.startsWith('[remarks]')) data.remarks = values[0];
            else if (key.startsWith('[weeks]')) data.weeks = values;
            else if (key.startsWith('[periods]')) data.periods = values;
            else if (key.startsWith('[times]')) data.times = values;
            else if (key.startsWith('[note]')) data.note = values[0];
            else if (key.startsWith('[breaks]')) data.breaks.push(values.join(','));
            else if (key.startsWith('[content]')) {
                data.content.push(values);
                current = '[content]';
            } else if (current === '[content]') {
                data.content.push(row.map(v => v.trim()));
            }
        }

        // ✅ 把欄位填上去（現在 data 已經完成）
        document.getElementById("title").value = data.title;
        document.getElementById("remarks").value = data.remarks;
        document.getElementById("weeks").value = data.weeks.join(',');
        document.getElementById("periods").value = data.periods.join(',');
        document.getElementById("times").value = data.times.join(',');
        document.getElementById("breaks").value = data.breaks.join('\n');
        document.getElementById("footer_note").value = data.note;
        document.getElementById("content").value = data.content.map(row => row.join(',')).join('\n');
    };

    reader.readAsText(file, 'UTF-8');
    }

function getFormData() {
return {
    title: document.getElementById("title").value,
    weeks: document.getElementById("weeks").value.split(','),
    periods: document.getElementById("periods").value.split(','),
    times: document.getElementById("times").value.split(','),
    orientation: document.getElementById("orientation").value,
    remarks: document.getElementById("remarks").value,
    note: document.getElementById("footer_note").value,
    breaks: document.getElementById("breaks").value
    .split('\n')
    .filter(line => line.includes(':')),
    content: document.getElementById("content").value
    .trim()
    .split(/\r?\n/)
    .map(row => row.split(',')),
    font: document.getElementById("font").value,
};
}

function downloadFile(url, filename) {
    fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(getFormData())
    })
    .then(response => response.blob())
    .then(blob => {
        const a = document.createElement('a');
        a.href = window.URL.createObjectURL(blob);
        a.download = filename;
        a.click();
    });
}

function downloadDocx() {
    const title = document.getElementById("title").value.trim() || '課表';
    const filename = `${title}.docx`;
    downloadFile('/download-docx', filename);
}

function downloadCsv() {
    const title = document.getElementById("title").value.trim() || '備份';
    const filename = `備份_${title}.csv`;
    downloadFile('/download-csv', filename);
}

function toggleContrast() {
    document.body.classList.toggle('high-contrast');
}