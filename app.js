/**
 * Javascript xử lý tương tác cho Web Sinh Hoạt Chuyên Đề 1
 * Chi Bộ Trường THPT Phục Hòa - Tỉnh Cao Bằng
 */

// 1. Quản lý Chuyển đổi Tab
function switchTab(tabId) {
  // Ẩn tất cả panels
  const panels = document.querySelectorAll('.tab-content-panel');
  panels.forEach(p => p.classList.remove('active'));

  // Bỏ active nav-links
  const navLinks = document.querySelectorAll('.nav-link');
  navLinks.forEach(l => l.classList.remove('active'));

  // Hiện panel tương ứng
  const targetPanel = document.getElementById(`tab-${tabId}`);
  if (targetPanel) {
    targetPanel.classList.add('active');
  }

  // Active nav link
  const targetLink = document.querySelector(`.nav-link[data-tab="${tabId}"]`);
  if (targetLink) {
    targetLink.classList.add('active');
  }

  // Cuộn mượt về đầu phần nội dung
  const contentStart = document.getElementById('main-tab-viewport');
  if (contentStart) {
    contentStart.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// 2. Thanh tiến trình đọc (Reading Progress)
window.addEventListener('scroll', () => {
  const winScroll = document.body.scrollTop || document.documentElement.scrollTop;
  const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
  const scrolled = (winScroll / height) * 100;
  const progressBar = document.getElementById('reading-progress');
  if (progressBar) {
    progressBar.style.width = scrolled + '%';
  }
});

// 3. Tùy chỉnh Cỡ Chữ (Font Size Controller)
let currentFontSize = 16;
function adjustFontSize(delta) {
  currentFontSize += delta;
  if (currentFontSize < 13) currentFontSize = 13;
  if (currentFontSize > 22) currentFontSize = 22;
  
  const sheets = document.querySelectorAll('.doc-sheet, .process-stepper, .tab-content-panel');
  sheets.forEach(el => {
    el.style.fontSize = `${currentFontSize}px`;
  });
  showToast(`Đã đổi cỡ chữ: ${currentFontSize}px`);
}

function resetFontSize() {
  currentFontSize = 16;
  const sheets = document.querySelectorAll('.doc-sheet, .process-stepper, .tab-content-panel');
  sheets.forEach(el => {
    el.style.fontSize = `16px`;
  });
  showToast('Đã khôi phục cỡ chữ mặc định');
}

// 4. Sao chép nội dung & Prompt (Copy to Clipboard)
function copyText(elementId) {
  const el = document.getElementById(elementId);
  if (!el) return;
  const text = el.innerText || el.value;
  navigator.clipboard.writeText(text).then(() => {
    showToast('✓ Đã sao chép vào bộ nhớ tạm (Clipboard)!');
  }).catch(() => {
    showToast('Lỗi khi sao chép. Vui lòng thử lại!');
  });
}

function copyPromptContent(btn) {
  const promptBox = btn.closest('.prompt-box');
  if (!promptBox) return;
  // Lấy nội dung text bỏ qua nút copy
  const clone = promptBox.cloneNode(true);
  const btnClone = clone.querySelector('.prompt-copy-btn');
  if (btnClone) btnClone.remove();
  const text = clone.innerText.trim();
  
  navigator.clipboard.writeText(text).then(() => {
    const originalText = btn.innerText;
    btn.innerText = '✓ Đã sao chép';
    btn.style.backgroundColor = '#15803d';
    showToast('✓ Đã sao chép Prompt thành công! Hãy dán vào Gemini / ChatGPT.');
    setTimeout(() => {
      btn.innerText = originalText;
      btn.style.backgroundColor = '';
    }, 2000);
  });
}

// 5. Hiển thị thông báo Toast nhanh
function showToast(message) {
  let toast = document.getElementById('toast-notification');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast-notification';
    toast.style.position = 'fixed';
    toast.style.bottom = '80px';
    toast.style.right = '24px';
    toast.style.backgroundColor = '#0f172a';
    toast.style.color = '#ffffff';
    toast.style.padding = '12px 20px';
    toast.style.borderRadius = '8px';
    toast.style.boxShadow = '0 10px 25px rgba(0,0,0,0.2)';
    toast.style.fontSize = '0.9rem';
    toast.style.fontWeight = '600';
    toast.style.zIndex = '10000';
    toast.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
    document.body.appendChild(toast);
  }
  toast.innerText = message;
  toast.style.opacity = '1';
  toast.style.transform = 'translateY(0)';
  
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
  }, 2500);
}

// 6. CÔNG CỤ CHUẨN HÓA CÚ PHÁP ĐỀ THI TRÊN WORD (QUIZ FORMATTER & VALIDATOR)
function formatQuizText() {
  const input = document.getElementById('raw-quiz-input');
  const output = document.getElementById('formatted-quiz-output');
  if (!input || !output) return;

  let raw = input.value;
  if (!raw.trim()) {
    showToast('Vui lòng dán nội dung câu hỏi vào khung nhập!');
    return;
  }

  // Chuẩn hóa dòng và khoảng trắng
  let lines = raw.split('\n');
  let cleanedLines = [];
  let questionCount = 0;
  let optionCount = 0;

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    if (!line) continue;

    // Chuẩn hóa khoảng trắng kép
    line = line.replace(/\s+/g, ' ');

    // Nhận diện câu hỏi (Bắt đầu bằng Câu 1, Câu 2, hoặc 1., 2., 1/, 2/)
    let qMatch = line.match(/^(?:Câu\s*(\d+)[\.:\s]*|(\d+)[\.:\/\)]\s*)(.*)/i);
    if (qMatch) {
      questionCount++;
      let qNum = questionCount;
      let qContent = (qMatch[3] || '').trim();
      cleanedLines.push(`\nCâu ${qNum}: ${qContent}`);
      continue;
    }

    // Nhận diện đáp án A., B., C., D. hoặc a., b., c., d.
    let optMatch = line.match(/^([A-Da-d])[\.:\)\s]+(.*)/);
    if (optMatch) {
      optionCount++;
      let optLetter = optMatch[1].toUpperCase();
      let optContent = optMatch[2].trim();
      cleanedLines.push(`${optLetter}. ${optContent}`);
      continue;
    }

    // Nhận diện mệnh đề Đúng/Sai (a), b), c), d)
    let tfMatch = line.match(/^([a-d])\)\s*(.*)/i);
    if (tfMatch) {
      optionCount++;
      let letter = tfMatch[1].toLowerCase();
      let content = tfMatch[2].trim();
      cleanedLines.push(`${letter}) ${content}`);
      continue;
    }

    // Dòng văn bản bình thường (giải thích hoặc ngữ cảnh)
    cleanedLines.push(line);
  }

  const result = cleanedLines.join('\n').trim();
  output.value = result;

  // Cập nhật thống kê
  const statBox = document.getElementById('validator-stats');
  if (statBox) {
    statBox.innerHTML = `
      <span>✓ Đã nhận diện: <strong style="color:#1d4ed8">${questionCount}</strong> câu hỏi</span>
      <span>• <strong style="color:#15803d">${optionCount}</strong> phương án/mệnh đề</span>
      <span>• Định dạng: <strong style="color:#d97706">Chuẩn Microsoft Word / Azota</strong></span>
    `;
  }
  showToast(`Đã chuẩn hóa thành công ${questionCount} câu hỏi!`);
}

function loadSampleQuizToFormatter() {
  const sample = `1/ Thiết bị nào sau đây thuộc khối thiết bị ra (Output) của máy tính?
A. Chuột quang
B. Bàn phím cơ
C. Máy in laser màu
D. Máy quét hình ảnh

2. Theo nguyên tắc đạo đức số và an toàn thông tin, hành vi nào sau đây là KHÔNG được phép?
A/ Bật xác thực hai lớp (2FA) bảo vệ tài khoản
B. Tải phần mềm không rõ nguồn gốc từ các trang web lậu vào máy tính trường
C/ Thường xuyên sao lưu dữ liệu bài giảng lên Google Drive
D/ Báo cáo ban giám hiệu khi phát hiện mã độc

Câu 3: Đánh giá tính Đúng/Sai của các phát biểu về ứng dụng AI trong giảng dạy:
a) AI giúp giáo viên tiết kiệm thời gian hành chính và chuẩn bị học liệu.
b) Giáo viên có thể tin tưởng 100% dữ liệu AI cung cấp mà không cần kiểm chứng.
c) AI hỗ trợ cá nhân hóa việc học tập cho học sinh vùng cao.
d) Học sinh được phép dùng AI để sao chép bài làm kiểm tra định kỳ.`;

  const input = document.getElementById('raw-quiz-input');
  if (input) {
    input.value = sample;
    formatQuizText();
  }
}

function copyFormattedOutput() {
  const output = document.getElementById('formatted-quiz-output');
  if (!output || !output.value.trim()) {
    showToast('Chưa có nội dung chuẩn hóa để sao chép!');
    return;
  }
  navigator.clipboard.writeText(output.value).then(() => {
    showToast('✓ Đã sao chép nội dung chuẩn Word vào Clipboard! Giờ hãy dán (Ctrl+V) vào Microsoft Word.');
  });
}

function clearFormatter() {
  const input = document.getElementById('raw-quiz-input');
  const output = document.getElementById('formatted-quiz-output');
  if (input) input.value = '';
  if (output) output.value = '';
  const statBox = document.getElementById('validator-stats');
  if (statBox) statBox.innerHTML = '<span>Trạng thái: Sẵn sàng nhập liệu</span>';
}

// Khởi chạy khi tài liệu tải xong
document.addEventListener('DOMContentLoaded', () => {
  console.log("Hệ thống Web Sinh Hoạt Chuyên Đề 1 THPT Phục Hòa đã sẵn sàng.");
});
