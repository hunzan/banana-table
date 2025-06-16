function generateCalendarDocx() {
  // 解析固定行程
  const fixedRaw = document.getElementById("weekly_fixed").value.trim();
  let fixedEvents = [];
  if (fixedRaw) {
    fixedEvents = fixedRaw.split('\n').map(line => line.trim());
  }

  // 解析特定行程
  const specificRaw = document.getElementById("specific_events").value.trim();
  let specificEvents = [];
  if (specificRaw) {
    specificEvents = specificRaw.split('\n').map(line => {
      const [md, time, ...taskParts] = line.trim().split(/\s+/);
      const task = taskParts.join(" ");
      const [m, d] = md.split('/');
      const year = document.getElementById("year").value;
      return {
        date: `${year}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`,
        time,
        task
      };
    });
  }

  const data = {
    font: document.getElementById("font").value,
    layout: document.getElementById("layout").value,
    title: document.getElementById("title").value,
    year: document.getElementById("year").value,
    month: document.getElementById("month").value,
    fixedEvents: fixedEvents,
    specificEvents: specificEvents,
  };

  fetch('/download-month-docx', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  })
    .then(res => {
      if (!res.ok) {
        throw new Error(`HTTP error! status: ${res.status}`);
      }
      return res.blob();
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${data.title || '月行事曆'}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch(err => {
      alert("下載失敗：" + err.message);
    });
}


function downloadCalendarCsv() {
  // 同樣的解析邏輯
  const fixedRaw = document.getElementById("weekly_fixed").value.trim();
  let fixedEvents = [];
  if (fixedRaw) {
    fixedEvents = fixedRaw.split('\n').map(line => line.trim());
  }

  const specificRaw = document.getElementById("specific_events").value.trim();
  let specificEvents = [];
  if (specificRaw) {
    specificEvents = specificRaw.split('\n').map(line => {
      const [md, time, ...taskParts] = line.trim().split(/\s+/);
      const task = taskParts.join(" ");
      const [m, d] = md.split('/');
      const year = document.getElementById("year").value;
      return {
        date: `${year}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`,
        time,
        task
      };
    });
  }

  const data = {
    font: document.getElementById('font').value,
    layout: document.getElementById('layout').value,
    title: document.getElementById("title").value,
    year: document.getElementById("year").value,
    month: document.getElementById("month").value,
    fixedEvents: fixedEvents,
    specificEvents: specificEvents,
  };
  console.log("🔍 傳給後端的資料：", data);

  fetch('/download-calendar-csv', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(data)
  })
    .then(res => {
      if (!res.ok) {
        throw new Error(`HTTP error! status: ${res.status}`);
      }
      return res.blob();
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${data.title || '月行事曆'}_備份.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    })
    .catch(err => {
      alert("❌ 下載 CSV 發生錯誤！");
      console.error(err);
    });
    
}
function uploadCalendarCsv(event) {
  const input = event.target;
  if (!input.files || input.files.length === 0) {
    alert('請先選擇 CSV 檔案');
    return;
  }

  const file = input.files[0];
  const formData = new FormData();
  formData.append('csv', file);

  fetch('/upload-calendar-csv', {
    method: 'POST',
    body: formData
  })
  .then(res => {
    if (!res.ok) {
      throw new Error(`伺服器回應錯誤：${res.status}`);
    }
    return res.json();
  })
  .then(data => {
    // 🔹 回填五個欄位資料（title, year, month, layout, font）
    if (data.font !== undefined) {
      document.getElementById('font').value = data.font;
    }
    if (data.layout !== undefined) {
      document.getElementById('layout').value = data.layout;
    }
    if (data.title !== undefined) {
      document.getElementById('title').value = data.title;
    }
    if (data.year !== undefined) {
      document.getElementById('year').value = data.year;
    }
    if (data.month !== undefined) {
      document.getElementById('month').value = data.month;
    }
    
    // 🔹 固定行程回填
    if (data.fixedEvents && Array.isArray(data.fixedEvents)) {
      document.getElementById('weekly_fixed').value = data.fixedEvents.join('\n');
    } else {
      document.getElementById('weekly_fixed').value = '';
    }

    // 🔹 特定行程回填（穩定處理日期格式）
    if (data.specificEvents && Array.isArray(data.specificEvents)) {
    const lines = data.specificEvents.map(evt => {
        let dateStr = evt.date;

        // 若為 YYYY-MM-DD 格式，轉成 MM/DD
        if (typeof dateStr === 'string' && dateStr.length === 10 && dateStr.includes('-')) {
        const parts = dateStr.split('-');
        if (parts.length === 3) {
            const month = parseInt(parts[1]);
            const day = parseInt(parts[2]);
            if (!isNaN(month) && !isNaN(day)) {
            dateStr = `${month}/${day}`;
            }
        }
        }

        return `${dateStr} ${evt.time} ${evt.task}`;
    });

    document.getElementById('specific_events').value = lines.join('\n');
    } else {
    document.getElementById('specific_events').value = '';
    }

    alert('CSV 讀取並填入完成！');
  })
  .catch(err => {
    alert('❌ 上傳 CSV 失敗：' + err.message);
    console.error(err);
  });
}

function toggleContrast() {
    document.body.classList.toggle('high-contrast');
}
