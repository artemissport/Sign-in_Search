// ==========================================
// 請填入您最新的 GAS Web App 網址
// ==========================================
const GAS_URL = "https://script.google.com/macros/s/AKfycbxE5Wrqr_uPvuHZ6HhE6s3fWfh9j7QX5et4ANiI5SXXcH0_HS9ZL8OUon2B-Jrmd3eA/exec";

// 全域變數
let userData = null;
let userId = "";

google.charts.load('current', { 'packages': ['bar'] });

window.addEventListener('load', function() {
  const storedData = sessionStorage.getItem('userData');
  userId = sessionStorage.getItem('userId');

  if (!storedData || !userId) {
    alert("請先登入！");
    window.location.href = "index.html";
    return;
  }

  userData = JSON.parse(storedData);

  document.getElementById('disp_id').innerText = userData.idName || userId;
  document.getElementById('disp_dateRange').innerText = userData.dateRangeText || "";
  
  if (userData.signSheetUrl) {
    const link = document.getElementById('signSheetUrlLink');
    link.href = userData.signSheetUrl;
    link.style.display = 'inline-block';
  }

  divInit();
});

function logout() {
  sessionStorage.clear();
  window.location.href = "index.html";
}

function divInit() {
  $("#table_div").hide();
  $("#chart_sec").hide();
  $("#calendar_sec").hide();
  $("#drawContent").hide();
  $("#tips").html(""); 
  $("#loading").hide();
}

function showLoading() {
  $("#loading").show();
  $('input[type="button"]').prop('disabled', true);
}

function hideLoading() {
  $("#loading").hide();
  $('input[type="button"]').prop('disabled', false);
}

// 產生 API 網址
function getParams(op, filter = "conditional") {
  const params = new URLSearchParams({
    op: op,
    id: userId,
    sheetId: userData.sheetId,
    filter: filter
  });
  return GAS_URL + "?" + params.toString();
}

// 1. 全部心得查詢 (維持原邏輯)
function getData() {
  divInit();
  showLoading();

  fetch(getParams("getGsData", "ALL"))
    .then(res => res.json())
    .then(data => {
      // 確保 data 是陣列，因為後端有時回傳字串
      if (typeof data === 'string') data = JSON.parse(data);
      drawTable(data);
      $("#table_div").show();
    })
    .catch(err => { console.error(err); alert("讀取失敗"); })
    .finally(() => hideLoading());
}

function drawTable(obj) {
  if (!obj || obj.length === 0) {
    document.getElementById('table_div').innerHTML = "<h3 class='text-center'>查無資料</h3>";
    return;
  }
  let table = '<table class="table table-striped table-bordered"><thead><tr>';
  for (let key in obj[0]) { table += '<th class="bg-primary">' + key + '</th>'; }
  table += '</tr></thead><tbody>';
  obj.forEach(row => {
    table += '<tr>';
    for (let key in row) {
      let style = (key == "校長留言板") ? "style='color:blue; font-weight:bold;'" : "";
      table += `<td ${style}>${row[key]}</td>`;
    }
    table += '</tr>';
  });
  table += '</tbody></table>';
  document.getElementById('table_div').innerHTML = table;
}

// 2. 統計圖表 (維持原邏輯)
function drawChart() {
  divInit();
  showLoading();
  $("#chart_sec").show();

  Promise.all([
    fetch(getParams("calCoachCount")).then(r => r.json()),
    fetch(getParams("getEvents")).then(r => r.json()),
    fetch(getParams("getBadges")).then(r => r.json())
  ]).then(([coachData, eventsData, badgesData]) => {
    
    document.getElementById('tips').innerHTML = coachData.message;
    drawBarChart(eventsData);
    renderBadges(badgesData);

  }).catch(err => { console.error(err); alert("讀取圖表失敗"); })
  .finally(() => hideLoading());
}

function drawBarChart(rows) {
  let total = rows.length;
  let dataTable = new google.visualization.DataTable();
  dataTable.addColumn('string', '項目');
  dataTable.addColumn('number', '次數');
  dataTable.addRow([userData.idName || userId, 0]); 
  dataTable.addRow(['累計簽到數', total]);

  let options = {
    chart: { title: '線上減重班', subtitle: '這個月累積簽到數' },
    bars: 'horizontal',
    bar: { groupWidth: "60%" },
    height: 300,
    legend: { position: 'none' }
  };
  let chart = new google.charts.Bar(document.getElementById('chart_div'));
  chart.draw(dataTable, google.charts.Bar.convertOptions(options));
}

// 3. 勳章渲染 (將後端計算結果畫出來)
function renderBadges(badges) {
  let imgHtml = "";
  let coachHtml = "";
  badges.forEach(badge => {
    let imgTag = `<img src="${badge.src}" alt="${badge.alt}" title="${badge.alt}">`;
    if (badge.type === 'coach') coachHtml += imgTag;
    else imgHtml += imgTag;
  });
  document.getElementById('img_div').innerHTML = imgHtml;
  document.getElementById('imgCoach_div').innerHTML = coachHtml;
}

// 4. 行事曆 (維持原邏輯)
function drawCalendar() {
  divInit();
  showLoading();
  $("#calendar_sec").show();
  $("#tips").html("載入行事曆中...");

  $('#calendar_div').fullCalendar('destroy');
  $('#calendar_div').fullCalendar({
    header: { left: 'prev,next today', center: 'title', right: 'month,listWeek' },
    height: 'auto',
    events: function(start, end, timezone, callback) {
      fetch(getParams("getEvents"))
        .then(res => res.json())
        .then(events => {
          callback(events);
          hideLoading();
          $("#tips").html("");
        });
    }
  });
}

// 5. 小紙條 (維持原邏輯)
function drawSticyNote() {
  divInit();
  showLoading();
  const container = document.getElementById('drawContent');
  container.style.border = "1px solid #ccc";
  container.style.display = "block";
  
  fetch(getParams("getStickyNotes"))
    .then(res => res.json())
    .then(notes => {
      let html = "";
      if(notes.length === 0) html = "<h3 class='text-center' style='padding-top:100px;'>目前沒有小紙條喔</h3>";
      else {
        notes.forEach(note => {
          let content = note.content;
          if (content.length > 100) content = content.substring(0, 100) + "...";
          html += `<div class="sticky-note sticky-note-${note.color}">${content}<br><br><small style="float:right;">by ${note.name}</small></div>`;
        });
      }
      container.innerHTML = html;
      setTimeout(() => {
        const width = $(container).width();
        const height = $(container).height();
        $(".sticky-note").each(function() {
          const maxLeft = width - $(this).outerWidth();
          const maxTop = height - $(this).outerHeight();
          $(this).css({
            "top": Math.max(0, Math.random() * maxTop) + "px",
            "left": Math.max(0, Math.random() * maxLeft) + "px",
            "transform": "rotate(" + ((Math.random() * 20) - 10) + "deg)"
          });
        });
        hideLoading();
      }, 100);
    });
}
