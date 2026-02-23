// ── Navigation ──────────────────────────────────────────────
const PAGES = [
  "dashboard",
  "upload",
  "process",
  "analytics",
  "returns",
  "accounts",
  "about",
];
function showPage(name) {
  document
    .querySelectorAll(".page")
    .forEach((p) => p.classList.remove("active"));
  document
    .querySelectorAll(".nav-item")
    .forEach((n) => n.classList.remove("active"));

  const page = document.getElementById("page-" + name);
  if (page) page.classList.add("active");

  const idx = PAGES.indexOf(name);
  if (idx >= 0) {
    const navItems = document.querySelectorAll(".nav-item");
    if (navItems[idx]) navItems[idx].classList.add("active");
  }

  if (name === "analytics") {
    loadCharts();
  }
  if (name === "accounts") {
    loadAccounts();
    loadPlatformsList();
  }
}

// ── Toast ────────────────────────────────────────────────────
function showToast(msg, type = "info") {
  const icons = { success: "✅", error: "❌", info: "ℹ️" };
  const t = document.createElement("div");
  t.className = `toast toast-${type}`;
  t.innerHTML = `<span>${icons[type]}</span><span>${msg}</span>`;
  document.getElementById("toastContainer").appendChild(t);
  setTimeout(() => {
    t.style.animation = "slideOut .3s ease forwards";
    setTimeout(() => t.remove(), 300);
  }, 3500);
}

// ── Upload ───────────────────────────────────────────────────
const zone = document.getElementById("uploadZone");
if (zone) {
  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
  });
}

function handleFiles(files) {
  if (!files.length) return;
  const fd = new FormData();

  const accSelect = document.getElementById("accountSelect");
  if (accSelect) {
    fd.append("account_name", accSelect.value);
  }

  // Convert FileList to Array if needed, though FormData accepts loop
  for (let i = 0; i < files.length; i++) {
    fd.append("files", files[i]);
  }

  fetch("/upload", { method: "POST", body: fd })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message + " - توجه إلى صفحة 'معالجة البيانات' للبدء", "success");
        refreshFiles();
      } else {
        showToast(data.message, "error");
      }
    })
    .catch((err) => showToast("فشل الاتصال بالخادم", "error"));
}

function refreshFiles() {
  fetch("/api/files")
    .then((r) => r.json())
    .then((data) => {
      const list = document.getElementById("file-list");
      if (!list) return;

      if (!data.files || !data.files.length) {
        list.innerHTML =
          '<div id="empty-msg" style="text-align:center;color:var(--muted);padding:32px">لا توجد ملفات مرفوعة بعد</div>';
        const cnt = document.getElementById("file-count");
        if (cnt) cnt.textContent = "";
        return;
      }
      list.innerHTML = data.files
        .map(
          (f, i) => `
      <div class="file-item" id="fi-${i}">
        <div class="file-icon">${f.name.endsWith(".csv") ? "📄" : "📊"}</div>
        <div class="file-info">
          <div class="file-name">${f.name}</div>
          <div class="file-meta">${f.size} KB &nbsp;•&nbsp; ${f.modified}</div>
        </div>
        <button class="file-del" onclick="deleteFile('${f.name}', this)" title="حذف">✕</button>
      </div>
    `,
        )
        .join("");

      const cnt = document.getElementById("file-count");
      if (cnt) cnt.textContent = `(${data.files.length} ملف)`;
    })
    .catch(console.error);
}

function deleteFile(name, btn) {
  if (!confirm("هل أنت متأكد من حذف الملف؟")) return;

  fetch("/delete-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filename: name }),
  })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message, "success");
        refreshFiles();
      } else showToast(data.message, "error");
    });
}

// ── Reports ──────────────────────────────────────────────────
function loadReports() {
  fetch("/api/reports")
    .then((r) => r.json())
    .then((data) => {
      const list = document.getElementById("reports-list");
      if (!list) return;

      if (!data.reports || !data.reports.length) {
        list.innerHTML = '<div style="text-align:center;color:var(--muted);padding:32px">لا توجد تقارير بعد. قم بتشغيل المعالجة أولاً.</div>';
        return;
      }

      list.innerHTML = data.reports.map((r) => `
        <div class="report-item">
          <div class="report-icon">📊</div>
          <div class="report-info">
            <div class="report-name">${r.name}</div>
            <div class="report-meta">${r.size} KB &nbsp;•&nbsp; ${r.modified}</div>
          </div>
          <a href="/download/${r.name}" class="btn btn-primary">⬇️ تحميل</a>
        </div>
      `).join("");
    })
    .catch(console.error);
}

// ── Dashboard Data (NEW) ─────────────────────────────────────
async function updateDashboard() {
  try {
    const res = await fetch("/api/stats");
    const data = await res.json();

    // 1. KPI Cards
    const stats = data.stats;
    if (stats) {
      document.getElementById("kpi-orders").textContent = stats.orders;
      document.getElementById("kpi-expected").textContent =
        stats.total_expected.toLocaleString(undefined, {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      document.getElementById("kpi-collected").textContent =
        stats.total_collected.toLocaleString(undefined, {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      document.getElementById("kpi-rate").textContent =
        stats.collection_rate + "%";
      document.getElementById("kpi-collections").textContent =
        stats.collections;
    }

    // 2. Platform Table
    const tableBody = document.getElementById("platform-table");
    if (tableBody && data.platforms) {
      if (data.platforms.length === 0) {
        tableBody.innerHTML =
          '<tr><td colspan="6" style="text-align: center; color: var(--muted); padding: 24px">لا توجد بيانات بعد. قم برفع الملفات أولاً.</td></tr>';
      } else {
        tableBody.innerHTML = data.platforms
          .map((p) => {
            let badgeColor = "rose";
            if (p.platform === "Amazon") badgeColor = "amber";
            else if (p.platform === "Noon") badgeColor = "cyan";
            else if (p.platform === "Trendyol") badgeColor = "emerald";
            else if (p.platform === "Ilasouq") badgeColor = "purple";

            const collected = p.collected || 0;
            const expected = p.expected || 0;
            const cost = p.cost || 0;
            // rate logic: if expected > 0 then (collected / expected * 100) else 0
            // The backend usually provides 'rate' in stats, but here p might have it or we calc it.
            // In app.py platform_breakdown, it returns 'rate'.
            const rate =
              p.rate !== undefined
                ? p.rate
                : expected
                  ? ((collected / expected) * 100).toFixed(1)
                  : 0;

            return `
            <tr>
              <td><span class="badge badge-${badgeColor}">${p.platform}</span></td>
              <td>${p.orders}</td>
              <td>${expected.toLocaleString(undefined, { minimumFractionDigits: 2 })} ر.س</td>
              <td style="color: #94a3b8">${cost.toLocaleString(undefined, { minimumFractionDigits: 2 })} ر.س</td>
              <td>${collected.toLocaleString(undefined, { minimumFractionDigits: 2 })} ر.س</td>
              <td>
                <div>${rate}%</div>
                <div class="progress-bar">
                  <div class="progress-fill" style="width:${rate}%"></div>
                </div>
              </td>
            </tr>
          `;
          })
          .join("");
      }
    }

    // 3. Reports List (if on dashboard, or if we want to update it globally)
    // There is no explicit API for reports in updateDashboard usually, but let's fetch it if needed.
    // For now, let's stick to stats.

    // Check if we need to reload charts if on analytics
    if (
      document.getElementById("page-analytics").classList.contains("active")
    ) {
      chartsLoaded = false;
      loadCharts();
    }
  } catch (err) {
    console.error("Dashboard update failed", err);
  }
}

// ── Returns Scanner ──────────────────────────────────────────
const audioSuccess = new Audio('https://assets.mixkit.co/active_storage/sfx/936/936-preview.mp3'); // Quick beep
const audioError = new Audio('https://assets.mixkit.co/active_storage/sfx/928/928-preview.mp3');   // Error buzzer
audioSuccess.volume = 0.5;
audioError.volume = 0.5;

function playSound(type) {
  if (type === 'success') {
    audioSuccess.currentTime = 0;
    audioSuccess.play().catch(e => console.log(e));
  } else {
    audioError.currentTime = 0;
    audioError.play().catch(e => console.log(e));
  }
}

// Add event listener for Enter key on the barcode input
const barcodeInput = document.getElementById('returnBarcode');
if (barcodeInput) {
  barcodeInput.addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      submitReturn();
    }
  });
}

// Ensure focus on the input when navigating to the returns page
const navItems = document.querySelectorAll('.nav-item');
navItems.forEach(item => {
  item.addEventListener('click', () => {
    if (item.innerText.includes('استلام المرتجعات') && barcodeInput) {
      setTimeout(() => barcodeInput.focus(), 200);
      loadReturns();
    }
  });
});

function submitReturn() {
  const input = document.getElementById('returnBarcode');
  const trackingId = input.value.trim();

  if (!trackingId) {
    showToast('يرجى إدخال رقم الطلب أو التتبع', 'error');
    playSound('error');
    input.focus();
    return;
  }

  const btn = document.getElementById('btnSubmitReturn');
  const originalText = btn.innerHTML;
  btn.innerHTML = '<span class="spinner" style="display:inline-block; width:14px; height:14px; margin-left:8px;"></span> جاري...';
  btn.disabled = true;

  fetch('/api/returns/add', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tracking_id: trackingId })
  })
    .then(r => r.json())
    .then(data => {
      if (data.success) {
        playSound('success');
        showToast(data.message, 'success');
        input.value = ''; // clear input
        loadReturns();    // refresh table
      } else if (data.duplicate) {
        playSound('error');
        showToast(data.error, 'info'); // Duplicate isn't a critical failure, just a repeat
        input.value = '';
      } else {
        playSound('error');
        showToast(data.error, 'error');
      }
    })
    .catch(err => {
      console.error(err);
      playSound('error');
      showToast('فشل الاتصال بالخادم', 'error');
    })
    .finally(() => {
      btn.innerHTML = originalText;
      btn.disabled = false;
      input.focus(); // keep focus for next scan
    });
}

function loadReturns() {
  fetch('/api/returns/list?limit=50')
    .then(r => r.json())
    .then(data => {
      const tb = document.getElementById('returnsBody');
      const countSpan = document.getElementById('returnsTotalCount');

      if (!tb) return;

      if (data.success) {
        countSpan.textContent = data.total || 0;

        if (!data.data || data.data.length === 0) {
          tb.innerHTML = '<tr><td colspan="4" style="text-align:center; padding:30px; color:var(--muted);">لا توجد مرتجعات مسجلة</td></tr>';
          return;
        }

        tb.innerHTML = data.data.map((item, index) => {
          const scanTime = new Date(item.scanned_at).toLocaleString('ar-EG', { dateStyle: 'short', timeStyle: 'medium' });
          return `
                <tr style="animation: fadeRow 0.3s ease backwards; animation-delay: ${index * 0.03}s;">
                    <td style="color:var(--muted); font-size:12px;">#${item.id}</td>
                    <td style="font-weight:600; font-family: monospace; font-size:14px;">${item.tracking_id}</td>
                    <td style="color:var(--muted); font-size:12px;">${scanTime}</td>
                    <td style="text-align:center;">
                        <button class="btn btn-ghost" style="padding:4px 8px; font-size:12px; border-color:transparent; color:var(--rose);" onclick="deleteReturn(${item.id})">حذف</button>
                    </td>
                </tr>
                `;
        }).join('');
      }
    })
    .catch(err => {
      console.error("Failed to load returns", err);
    });
}

function deleteReturn(id) {
  if (!confirm('هل أنت متأكد من حذف هذا السجل؟')) return;

  fetch('/api/returns/delete/' + id, { method: 'DELETE' })
    .then(r => r.json())
    .then(data => {
      if (data.success) {
        showToast('تم الحذف بنجاح', 'success');
        loadReturns();
        document.getElementById('returnBarcode').focus();
      } else {
        showToast(data.error || 'فشل الحذف', 'error');
      }
    })
    .catch(err => {
      console.error(err);
      showToast('فشل الاتصال بالخادم', 'error');
    });
}

// ── Process Pipeline ─────────────────────────────────────────
function runPipeline() {
  const btn = document.getElementById("runBtn");
  if (!btn) return; // Might be processBtn in some versions, but sticking to runBtn based on view_file

  const spinner = document.getElementById("spinner");
  const btnText = document.getElementById("runBtnText");
  const logCard = document.getElementById("log-card");
  const logBox = document.getElementById("logBox");

  btn.disabled = true;
  if (spinner) spinner.style.display = "block";
  if (btnText) btnText.textContent = "جاري المعالجة...";
  if (logCard) logCard.style.display = "block";
  if (logBox) logBox.textContent = "بدء المعالجة...\n";

  // Animate steps
  ["step1", "step2", "step3"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.className = "step-card";
      el.querySelector(".step-num").textContent = id.replace("step", "");
    }
  });
  const s1 = document.getElementById("step1");
  if (s1) s1.classList.add("running");

  fetch("/process", { method: "POST" })
    .then((r) => r.json())
    .then(async (data) => {
      if (logBox) {
        logBox.textContent = data.log || "";
        logBox.scrollTop = logBox.scrollHeight;
      }

      // Mark steps done
      ["step1", "step2", "step3"].forEach((id) => {
        const el = document.getElementById(id);
        if (el) {
          el.className = "step-card done";
          el.querySelector(".step-num").textContent = "✓";
        }
      });

      // Update Reports List
      loadReports();

      // Force dashboard refresh
      await updateDashboard();
      chartsLoaded = false;

      showToast("تمت المعالجة بنجاح!", "success");
    })
    .catch((err) => {
      if (logBox) logBox.textContent += "\n[ERROR] " + err;
      ["step1", "step2", "step3"].forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.classList.add("error");
      });
      showToast("حدث خطأ أثناء المعالجة", "error");
    })
    .finally(() => {
      btn.disabled = false;
      if (spinner) spinner.style.display = "none";
      if (btnText) btnText.textContent = "🚀 تشغيل المعالجة الكاملة";
    });
}

// ── New Week ─────────────────────────────────────────────────
function startNewWeek() {
  if (
    !confirm(
      "🗓️ بدء أسبوع جديد\n\nسيتم حذف:\n• جميع بيانات الطلبات والتحصيل من قاعدة البيانات\n• جميع الملفات المرفوعة من مجلد samples\n\n✅ ستظل جميع التقارير السابقة محفوظة\n\nهل تريد المتابعة؟",
    )
  )
    return;

  fetch("/new-week", { method: "POST" })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message, "success");
        updateDashboard(); // Will clear stats since DB is empty
        refreshFiles();
        setTimeout(() => showPage("upload"), 800);
      } else {
        showToast(data.message, "error");
      }
    })
    .catch(() => showToast("حدث خطأ غير متوقع", "error"));
}

// ── Reset DB ────────────────────────────────────────────────
function confirmReset() {
  if (!confirm("⚠️ هل أنت متأكد من إعادة تعيين قاعدة البيانات؟")) return;
  fetch("/reset-db", { method: "POST" })
    .then((r) => r.json())
    .then((data) => {
      showToast(data.message, data.success ? "success" : "error");
      if (data.success) location.reload();
    });
}

// ── Charts ───────────────────────────────────────────────────
let chartsLoaded = false;
let chartInstances = {};

// Chart Configuration
const CHART_COLORS = {
  cyan: "rgba(34,211,238,0.85)",
  blue: "rgba(59,130,246,0.85)",
  emerald: "rgba(16,185,129,0.85)",
  amber: "rgba(245,158,11,0.85)",
  rose: "rgba(244,63,94,0.85)",
  purple: "rgba(167,139,250,0.85)",
};
const PLATFORM_COLORS = [
  CHART_COLORS.amber,
  CHART_COLORS.cyan,
  CHART_COLORS.emerald,
  CHART_COLORS.purple,
  CHART_COLORS.rose,
];
const STATUS_COLORS = [
  CHART_COLORS.emerald,
  CHART_COLORS.rose,
  CHART_COLORS.amber,
  CHART_COLORS.blue,
];

const chartDefaults = {
  plugins: {
    legend: {
      labels: { color: "#94a3b8", font: { family: "Tajawal", size: 12 } },
    },
  },
  scales: {
    x: {
      ticks: { color: "#94a3b8", font: { family: "Tajawal" } },
      grid: { color: "rgba(255,255,255,0.05)" },
    },
    y: {
      ticks: { color: "#94a3b8", font: { family: "Tajawal" } },
      grid: { color: "rgba(255,255,255,0.05)" },
    },
  },
  responsive: true,
  maintainAspectRatio: false,
  animation: { duration: 800, easing: "easeInOutQuart" },
};

function destroyChart(id) {
  if (chartInstances[id]) {
    chartInstances[id].destroy();
    delete chartInstances[id];
  }
}

function loadCharts() {
  if (chartsLoaded && Object.keys(chartInstances).length > 0) return;
  chartsLoaded = true;

  fetch("/api/charts")
    .then((r) => r.json())
    .then((data) => {
      if (data.error) {
        showToast("خطأ في تحميل البيانات: " + data.error, "error");
        return;
      }

      // 1. Platform Orders Donut
      const ctx1 = document.getElementById("chartPlatformOrders");
      if (ctx1) {
        destroyChart("chartPlatformOrders");
        chartInstances["chartPlatformOrders"] = new Chart(ctx1, {
          type: "doughnut",
          data: {
            labels: data.platforms.labels,
            datasets: [
              {
                data: data.platforms.orders,
                backgroundColor: PLATFORM_COLORS,
                borderColor: "#111827",
                borderWidth: 3,
              },
            ],
          },
          options: { ...chartDefaults, cutout: "65%", scales: {} },
        });
      }

      // 2. Payment Status Donut
      const ctx2 = document.getElementById("chartStatus");
      if (ctx2) {
        destroyChart("chartStatus");
        chartInstances["chartStatus"] = new Chart(ctx2, {
          type: "doughnut",
          data: {
            labels: data.status.labels,
            datasets: [
              {
                data: data.status.values,
                backgroundColor: STATUS_COLORS,
                borderColor: "#111827",
                borderWidth: 3,
              },
            ],
          },
          options: { ...chartDefaults, cutout: "65%", scales: {} },
        });
      }

      // 3. Expected vs Collected Bar
      const ctx3 = document.getElementById("chartPlatformAmounts");
      if (ctx3) {
        destroyChart("chartPlatformAmounts");
        chartInstances["chartPlatformAmounts"] = new Chart(ctx3, {
          type: "bar",
          data: {
            labels: data.platforms.labels,
            datasets: [
              {
                label: "المبلغ المتوقع",
                data: data.platforms.expected,
                backgroundColor: "rgba(59,130,246,0.7)",
                borderColor: CHART_COLORS.blue,
                borderWidth: 2,
                borderRadius: 8,
              },
              {
                label: "المبلغ المحصل",
                data: data.platforms.collected,
                backgroundColor: "rgba(16,185,129,0.7)",
                borderColor: CHART_COLORS.emerald,
                borderWidth: 2,
                borderRadius: 8,
              },
            ],
          },
          options: { ...chartDefaults },
        });
      }

      // 4. Weekly Trend
      const ctx4 = document.getElementById("chartWeekly");
      if (ctx4) {
        if (data.weekly.labels.length > 0) {
          destroyChart("chartWeekly");
          chartInstances["chartWeekly"] = new Chart(ctx4, {
            type: "line",
            data: {
              labels: data.weekly.labels,
              datasets: [
                {
                  label: "عدد الطلبات",
                  data: data.weekly.orders,
                  yAxisID: "y",
                  borderColor: CHART_COLORS.cyan,
                  backgroundColor: "rgba(34,211,238,0.1)",
                  fill: true,
                  tension: 0.4,
                },
                {
                  label: "إجمالي المبيعات",
                  data: data.weekly.totals,
                  yAxisID: "y1",
                  borderColor: CHART_COLORS.purple,
                  backgroundColor: "rgba(167,139,250,0.1)",
                  fill: true,
                  tension: 0.4,
                },
              ],
            },
            options: {
              ...chartDefaults,
              scales: {
                x: chartDefaults.scales.x,
                y: {
                  ...chartDefaults.scales.y,
                  type: "linear",
                  position: "right",
                },
                y1: {
                  ...chartDefaults.scales.y,
                  type: "linear",
                  position: "left",
                  grid: { drawOnChartArea: false },
                },
              },
            },
          });
        } else {
          ctx4.parentElement.innerHTML =
            '<canvas id="chartWeekly"></canvas><p style="color:var(--muted);text-align:center;padding:40px;position:absolute;top:50%;left:50%;transform:translate(-50%,-50%)">لا توجد بيانات أسبوعية بعد</p>';
        }
      }
    })
    .catch((e) => console.error("Chart load error", e));
}

// ── Accounts Management ─────────────────────────────────────
let accountsData = [];

const COUNTRY_FLAGS = {
  SA: "🇸🇦 السعودية",
  AE: "🇦🇪 الإمارات",
  EG: "🇪🇬 مصر",
  KW: "🇰🇼 الكويت",
  BH: "🇧🇭 البحرين",
  QA: "🇶🇦 قطر",
  OM: "🇴🇲 عمان",
  JO: "🇯🇴 الأردن",
  OTHER: "🌍 أخرى",
};

const PLATFORM_BADGE_COLORS = {
  Amazon: "amber",
  Noon: "cyan",
  Trendyol: "emerald",
  Ilasouq: "purple",
  Website: "blue",
  Tabby: "rose",
  Tamara: "rose",
};

function loadPlatformsList() {
  fetch("/api/platforms/list")
    .then((r) => r.json())
    .then((data) => {
      const dl = document.getElementById("platformList");
      if (dl && data.platforms) {
        dl.innerHTML = data.platforms
          .map((p) => `<option value="${p}">`)
          .join("");
      }
    })
    .catch(console.error);
}

function loadAccounts() {
  fetch("/api/accounts")
    .then((r) => r.json())
    .then((data) => {
      accountsData = data.accounts || [];
      renderAccountsTable(accountsData);
      updateAccountsKPIs(accountsData);
    })
    .catch((err) => {
      console.error("Accounts load failed", err);
      const body = document.getElementById("accountsBody");
      if (body)
        body.innerHTML =
          '<tr><td colspan="6" style="text-align:center;padding:48px;color:var(--rose);">خطأ في تحميل الحسابات</td></tr>';
    });
}

function renderAccountsTable(accounts) {
  const body = document.getElementById("accountsBody");
  if (!body) return;

  if (!accounts.length) {
    body.innerHTML = `
      <tr><td colspan="6" style="text-align:center; padding:48px; color:var(--muted);">
        <div style="font-size: 40px; margin-bottom: 12px; opacity: 0.4;">🏪</div>
        <div style="font-size: 16px; font-weight: 700; margin-bottom: 6px; color: var(--text);">لا توجد حسابات مسجلة</div>
        <div style="font-size: 13px;">أضف حسابات جديدة من النموذج أعلاه</div>
      </td></tr>`;
    document.getElementById("acc-visible-count").textContent = "0";
    return;
  }

  body.innerHTML = accounts
    .map((acc, i) => {
      const badgeColor = PLATFORM_BADGE_COLORS[acc.platform_name] || "cyan";
      const flag = COUNTRY_FLAGS[acc.country] || COUNTRY_FLAGS["OTHER"];
      const date = acc.created_at
        ? new Date(acc.created_at).toLocaleDateString("ar-EG", {
          year: "numeric",
          month: "short",
          day: "numeric",
        })
        : "—";
      const fixedShip = acc.fixed_shipping_cost || 0;
      const clientShip = acc.client_shipping_cost || 0;
      const pmCommission = acc.payment_commission_rate || 0;
      const taxRate = acc.tax_rate || 0;
      const costIncTax = acc.cost_includes_tax ? 1 : 0;

      return `
      <tr class="acc-row" data-id="${acc.id}" data-platform="${acc.platform_name}" data-name="${acc.account_name}" data-country="${acc.country}"
          style="animation: fadeRow 0.3s ease backwards; animation-delay: ${i * 0.03}s;">
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border);">
          <div style="width: 36px; height: 36px; border-radius: 10px; background: var(--surface2); display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; color: var(--muted); border: 1px solid var(--border);">${i + 1}</div>
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border);">
          <span class="badge badge-${badgeColor}">${acc.platform_name}</span>
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); font-weight: 600;">
          ${acc.account_name}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); font-size: 13px;">
          ${flag}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center; font-size: 13px; font-weight: 600;">
          ${fixedShip > 0 ? fixedShip.toFixed(2) + ' ر.س' : '<span style="color:var(--muted)">—</span>'}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center; font-size: 13px; font-weight: 600;">
          ${clientShip > 0 ? clientShip.toFixed(2) + ' ر.س' : '<span style="color:var(--muted)">—</span>'}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center; font-size: 13px; font-weight: 600;">
          ${pmCommission > 0 ? pmCommission.toFixed(2) + ' ر.س' : '<span style="color:var(--muted)">—</span>'}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center; font-size: 13px; font-weight: 600;">
          ${taxRate > 0 ? taxRate + '%' : '<span style="color:var(--muted)">—</span>'}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center;">
          ${costIncTax ? '<span style="color:var(--emerald); font-weight:700;">✅ نعم</span>' : '<span style="color:var(--muted);">❌ لا</span>'}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); font-size: 12px; color: var(--muted);">
          ${date}
        </td>
        <td style="padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: center;">
          <div style="display: flex; gap: 6px; justify-content: center;">
            <button onclick="editAccount(${acc.id})" style="
              display: inline-flex; align-items: center; gap: 4px;
              padding: 6px 14px; border-radius: 8px; font-size: 12px; font-weight: 600;
              cursor: pointer; border: 1px solid rgba(59,130,246,0.25);
              background: rgba(59,130,246,0.1); color: var(--blue);
              transition: 0.2s; font-family: 'Tajawal', sans-serif;
            " onmouseover="this.style.background='rgba(59,130,246,0.2)'" onmouseout="this.style.background='rgba(59,130,246,0.1)'">
              ✏️ تعديل
            </button>
            <button onclick="deleteAccount(${acc.id}, '${acc.account_name.replace(/'/g, "\\'")}')" style="
              display: inline-flex; align-items: center; gap: 4px;
              padding: 6px 14px; border-radius: 8px; font-size: 12px; font-weight: 600;
              cursor: pointer; border: 1px solid rgba(244,63,94,0.25);
              background: rgba(244,63,94,0.1); color: var(--rose);
              transition: 0.2s; font-family: 'Tajawal', sans-serif;
            " onmouseover="this.style.background='rgba(244,63,94,0.2)'" onmouseout="this.style.background='rgba(244,63,94,0.1)'">
              🗑️ حذف
            </button>
          </div>
        </td>
      </tr>`;
    })
    .join("");

  document.getElementById("acc-visible-count").textContent = accounts.length;
}

function updateAccountsKPIs(accounts) {
  const total = accounts.length;
  const platforms = new Set(accounts.map((a) => a.platform_name)).size;
  const sa = accounts.filter((a) => a.country === "SA").length;
  const other = total - sa;

  document.getElementById("acc-kpi-total").textContent = total;
  document.getElementById("acc-kpi-platforms").textContent = platforms;
  document.getElementById("acc-kpi-sa").textContent = sa;
  document.getElementById("acc-kpi-other").textContent = other;
}

function addNewAccount() {
  const platform = document.getElementById("accPlatform").value.trim();
  const name = document.getElementById("accName").value.trim();
  const country = document.getElementById("accCountry").value;
  const shipping = parseFloat(document.getElementById("accShipping").value) || 0;
  const clientShipping = parseFloat(document.getElementById("accClientShipping").value) || 0;
  const commission = parseFloat(document.getElementById("accCommission").value) || 0;
  const taxRate = parseFloat(document.getElementById("accTaxRate").value) || 0;
  const taxInc = document.getElementById("accTaxInc").checked;

  if (!platform) {
    showToast("يجب إدخال اسم المنصة", "error");
    document.getElementById("accPlatform").focus();
    return;
  }
  if (!name) {
    showToast("يجب إدخال اسم الحساب / الفرع", "error");
    document.getElementById("accName").focus();
    return;
  }

  fetch("/api/accounts/add", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      platform_name: platform,
      account_name: name,
      country: country,
      fixed_shipping_cost: shipping,
      client_shipping_cost: clientShipping,
      payment_commission_rate: commission,
      tax_rate: taxRate,
      cost_includes_tax: taxInc,
    }),
  })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message, "success");
        document.getElementById("accPlatform").value = "";
        document.getElementById("accName").value = "";
        document.getElementById("accCountry").value = "SA";
        document.getElementById("accShipping").value = "";
        document.getElementById("accClientShipping").value = "";
        document.getElementById("accCommission").value = "";
        document.getElementById("accTaxRate").value = "";
        document.getElementById("accTaxInc").checked = false;
        loadAccounts();
        loadPlatformsList();
      } else {
        showToast(data.message, "error");
      }
    })
    .catch(() => showToast("فشل الاتصال بالخادم", "error"));
}

function editAccount(id) {
  const acc = accountsData.find((a) => a.id === id);
  if (!acc) return;

  const countryOptions = Object.entries(COUNTRY_FLAGS)
    .map(
      ([code, label]) =>
        `<option value="${code}" ${code === acc.country ? "selected" : ""}>${label}</option>`,
    )
    .join("");

  const row = document.querySelector(`tr[data-id="${id}"]`);
  if (!row) return;

  const cells = row.querySelectorAll("td");
  // Save original content for cancel
  const originalHTML = row.innerHTML;
  const fixedShip = acc.fixed_shipping_cost || 0;
  const clientShip = acc.client_shipping_cost || 0;
  const pmCommission = acc.payment_commission_rate || 0;
  const taxRate = acc.tax_rate || 0;
  const costIncTax = acc.cost_includes_tax ? 1 : 0;

  const inputStyle = `width: 100%; padding: 8px 12px; background: var(--bg); border: 1px solid var(--cyan);
    border-radius: 8px; color: var(--text); font-size: 13px; font-family: 'Tajawal', sans-serif; outline: none;`;

  // Replace cells with edit inputs
  cells[1].innerHTML = `<input type="text" id="edit-platform-${id}" value="${acc.platform_name}" list="platformList" style="${inputStyle}">`;
  cells[2].innerHTML = `<input type="text" id="edit-name-${id}" value="${acc.account_name}" style="${inputStyle}">`;
  cells[3].innerHTML = `<select id="edit-country-${id}" style="${inputStyle} cursor: pointer;">${countryOptions}</select>`;
  cells[4].innerHTML = `<input type="number" id="edit-shipping-${id}" value="${fixedShip}" step="0.01" min="0" placeholder="0" style="${inputStyle} width:80px; text-align:center;">`;
  cells[5].innerHTML = `<input type="number" id="edit-clientShip-${id}" value="${clientShip}" step="0.01" min="0" placeholder="0" style="${inputStyle} width:80px; text-align:center;">`;
  cells[6].innerHTML = `<input type="number" id="edit-commission-${id}" value="${pmCommission}" step="0.01" min="0" placeholder="0" style="${inputStyle} width:80px; text-align:center;">`;
  cells[7].innerHTML = `<input type="number" id="edit-taxRate-${id}" value="${taxRate}" step="0.01" min="0" placeholder="0" style="${inputStyle} width:80px; text-align:center;">`;
  cells[8].innerHTML = `<label style="display:flex; align-items:center; gap:6px; cursor:pointer; justify-content:center;">
    <input type="checkbox" id="edit-tax-${id}" ${costIncTax ? 'checked' : ''} style="width:18px; height:18px; cursor:pointer;">
    <span style="font-size:12px; color:var(--text);">شاملة</span>
  </label>`;
  cells[10].innerHTML = `
    <div style="display: flex; gap: 6px; justify-content: center;">
      <button onclick="saveEditAccount(${id})" style="
        display: inline-flex; align-items: center; gap: 4px;
        padding: 6px 14px; border-radius: 8px; font-size: 12px; font-weight: 600;
        cursor: pointer; border: 1px solid rgba(16,185,129,0.25);
        background: rgba(16,185,129,0.1); color: var(--emerald);
        transition: 0.2s; font-family: 'Tajawal', sans-serif;
      " onmouseover="this.style.background='rgba(16,185,129,0.2)'" onmouseout="this.style.background='rgba(16,185,129,0.1)'">
        ✅ حفظ
      </button>
      <button onclick="cancelEditAccount(${id})" data-original='${originalHTML.replace(/'/g, "&#39;")}' style="
        display: inline-flex; align-items: center; gap: 4px;
        padding: 6px 14px; border-radius: 8px; font-size: 12px; font-weight: 600;
        cursor: pointer; border: 1px solid var(--border);
        background: var(--surface2); color: var(--muted);
        transition: 0.2s; font-family: 'Tajawal', sans-serif;
      ">
        ❌ إلغاء
      </button>
    </div>`;

  // Focus the name input
  document.getElementById(`edit-name-${id}`).focus();
}

function saveEditAccount(id) {
  const platform = document.getElementById(`edit-platform-${id}`).value.trim();
  const name = document.getElementById(`edit-name-${id}`).value.trim();
  const country = document.getElementById(`edit-country-${id}`).value;
  const fixedShipping = parseFloat(document.getElementById(`edit-shipping-${id}`).value) || 0;
  const clientShipping = parseFloat(document.getElementById(`edit-clientShip-${id}`).value) || 0;
  const pmCommission = parseFloat(document.getElementById(`edit-commission-${id}`).value) || 0;
  const taxRate = parseFloat(document.getElementById(`edit-taxRate-${id}`).value) || 0;
  const costIncTax = document.getElementById(`edit-tax-${id}`).checked;

  if (!platform || !name) {
    showToast("يجب إدخال اسم المنصة واسم الحساب", "error");
    return;
  }

  fetch("/api/accounts/update", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      id: id,
      platform_name: platform,
      account_name: name,
      country: country,
      fixed_shipping_cost: fixedShipping,
      client_shipping_cost: clientShipping,
      payment_commission_rate: pmCommission,
      tax_rate: taxRate,
      cost_includes_tax: costIncTax,
    }),
  })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message, "success");
        loadAccounts();
      } else {
        showToast(data.message, "error");
      }
    })
    .catch(() => showToast("فشل الاتصال بالخادم", "error"));
}

function cancelEditAccount(id) {
  // Simply reload the table to restore original state
  loadAccounts();
}

function deleteAccount(id, name) {
  if (!confirm(`⚠️ هل أنت متأكد من حذف الحساب "${name}"؟\n\nلن يتم حذف البيانات المرتبطة بهذا الحساب (الطلبات والتحصيل).`))
    return;

  fetch("/api/accounts/delete", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: id }),
  })
    .then((r) => r.json())
    .then((data) => {
      if (data.success) {
        showToast(data.message, "success");
        loadAccounts();
      } else {
        showToast(data.message, "error");
      }
    })
    .catch(() => showToast("فشل الاتصال بالخادم", "error"));
}

function filterAccountsList() {
  const query = document
    .getElementById("accSearchInput")
    .value.toLowerCase()
    .trim();
  const rows = document.querySelectorAll(".acc-row");
  let visible = 0;

  rows.forEach((row) => {
    const platform = (row.dataset.platform || "").toLowerCase();
    const name = (row.dataset.name || "").toLowerCase();
    const country = (row.dataset.country || "").toLowerCase();

    if (
      !query ||
      platform.includes(query) ||
      name.includes(query) ||
      country.includes(query)
    ) {
      row.style.display = "";
      visible++;
    } else {
      row.style.display = "none";
    }
  });

  document.getElementById("acc-visible-count").textContent = visible;
}

// ── Init ───────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  // 1. Refresh File List
  refreshFiles();

  // 2. Fetch Dashboard Data (Async)
  updateDashboard();

  // 3. Fetch Reports (Async)
  loadReports();
});
