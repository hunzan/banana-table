document.addEventListener("DOMContentLoaded", function () {
  const container = document.getElementById("tableContainer");
  const addBtn = document.getElementById("addItemBtn");
  const removeBtn = document.getElementById("removeItemBtn");
  const clearBtn = document.getElementById("clearBtn");
  const uploadInput = document.getElementById("uploadInput");

    // 新增一筆欄位（可帶入預設值）
    function addItem(title = "", content = "") {
    const div = document.createElement("div");
    div.className = "item-block";
    div.innerHTML = `
        <input name="titles" type="text" placeholder="項目名稱" value="${title}" required><br>
        <textarea name="contents" rows="3" placeholder="內容" required>${content}</textarea>
        <hr>
    `;
    container.appendChild(div);

    // ⭐ 抓出剛剛 div 裡的 input 並 focus
    const input = div.querySelector('input[name="titles"]');
    if (input) {
        input.focus();
    }
 }

  // 預設一筆欄位
  addItem();

  // 新增欄位按鈕
  if (addBtn) {
    addBtn.addEventListener("click", function () {
      addItem();
    });
  }

  // 移除最後一筆欄位
  if (removeBtn) {
    removeBtn.addEventListener("click", function () {
      const items = document.querySelectorAll(".item-block");
      if (items.length > 0) {
        items[items.length - 1].remove();
      }
    });
  }

  // 清除所有欄位並新增一筆空白
  if (clearBtn) {
    clearBtn.addEventListener("click", function () {
      container.innerHTML = "";
      addItem();
    });
  }

  // CSV 上傳自動填表功能
  if (uploadInput) {
    uploadInput.addEventListener("change", function (event) {
      const file = event.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
          const content = e.target.result;
          const lines = content.split(/\r?\n/).filter(line => line.trim() !== "");

          container.innerHTML = ""; // 清空欄位

          lines.forEach(line => {
            const match = line.match(/^"(.*?)","(.*?)"$/);
            if (match) {
              const title = match[1];
              const value = match[2];
              addItem(title, value);
            }
          });
        };
        reader.readAsText(file, "UTF-8");
      }
    });
  }
});

function toggleContrast() {
  document.body.classList.toggle('high-contrast');
}
function uploadCustomCsv(event) {
  const container = document.getElementById("tableContainer");
  const file = event.target.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const csvText = e.target.result;

      // 使用 PapaParse 處理
      const parsed = Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true
      });

      container.innerHTML = ""; // 清空欄位

      parsed.data.forEach(row => {
        const [title, content] = row;
        if (title && content !== undefined) {
          const div = document.createElement("div");
          div.className = "item-block";
          div.innerHTML = `
            <input name="titles" type="text" placeholder="項目名稱" value="${title}" required><br>
            <textarea name="contents" rows="3" placeholder="內容" required>${content}</textarea>
            <hr>
          `;
          container.appendChild(div);
        }
      });
    };
    reader.readAsText(file, "UTF-8");
  }
}
