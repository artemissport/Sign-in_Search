var globalSheetName = "表單回應 1";
var globalSheetConfigName = "名單";
var globalConfig = "網站基礎設定"
var currentDate = Moment.moment(new Date()).format("YYYY/MM/DD");
var dataColIndex = 5;//任務日期位置
var coachCount = 0;
var sportSignSheetId = "";
var thisClassSheetId = "";
var signSheetUrl = "";
var memoSheetId = "";
var startDate = "";
var endDate = "";
var testForm = { id: "2306019", filter: "conditional", sheetId: "16G-5oPa7q5b9bMWi5e5fDEI98NoI6gWy93oF8W-iuEk" };

function readFiles(){
  let folder = DriveApp.getFolderById("1a_2xr6HY70TEnkslY5E7stkCWkye_Fs4");
  let files = folder.getFiles();
  let file_arr = []
  while (files.hasNext()) {
    let file = files.next();
    file_arr.push([file.getName(),file.getId()]);
  };
  Logger.log(file_arr);
}

function doPost(e) {
  var id = e.parameter.id;
  return result(id);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function result(id) {
  webSiteConfig();
  var isAuth = false;
  if ((idName = findIdName(id)) != null) {
    isAuth = true;
  }

  if (!isAuth) {
    return HtmlService.createHtmlOutput("<h1 style=\"text-align:center\">電話：" + id + "<br>找不到對應的學號</h1>")
  } else {
    var webhtml = HtmlService.createTemplateFromFile("result");
    webhtml.serviceUrl = ScriptApp.getService().getUrl();
    webhtml.id = idName;
    webhtml.sheetId = thisClassSheetId;
    webhtml.signSheetUrl = signSheetUrl;
    webhtml.container_tips = startDate + "到" + endDate;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("登入記錄");
    sheet.appendRow([new Date(), idName]);

    return webhtml.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function findIdName(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("減重班ID");
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var thisSheet = getDataFromSheetById(sheet.getRange(i, 1).getValue(), globalSheetConfigName);
    var thislastRow = thisSheet.getLastRow();
    for (var j = 1; j <= thislastRow; j++) {
      if (String(thisSheet.getRange(j, 5).getValue()) == id) {
        thisClassSheetId = sheet.getRange(i, 1).getValue();
        signSheetUrl = sheet.getRange(i, 2).getValue(); //簽到ＵＲＬ
        return thisSheet.getRange(j, 1).getValue();
      }
    }
  }
}

function getDataFromSheetById(sheetId, sheetName) {
  return SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
}

function calCoachCount(form) {
  var dataLength = getEvents(form).length
  return new Number(coachCount) - dataLength;
}

function drawImg(form) {
  var count = getEvents(form).length;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("任務勳章"); 
  var lastRow = sheet.getLastRow();
  
  var html = "";
  for (var i = 2; i <= lastRow; i++) {
    if (Number(count) >= Number(sheet.getRange(i, 2).getValue()) && sheet.getRange(i, 2).getValue() != '') {
      html += '<img src="' + sheet.getRange(i, 4).getValue() + '" alt="'+sheet.getRange(i, 1).getValue()+'"/>';
    }
  }
  for (var i = 2; i <= lastRow; i++) {
    if (Number(count) < Number(sheet.getRange(i, 3).getValue()) && sheet.getRange(i, 3).getValue() != '') {
      html += '<img src="' + sheet.getRange(i, 4).getValue() + '" alt="'+sheet.getRange(i, 1).getValue()+'"/>';
      break;
    }
  }
  return html;
}

function drawCoachClass(form) {
  var count = calCoachCount(form);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("任務勳章"); 
  var lastRow = sheet.getLastRow();
  
  var html = "";
  for (var i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 2).getValue() == 'missionCoach' && count <= 0) {
      html += '<img src="' + sheet.getRange(i, 4).getValue() + '" alt="'+sheet.getRange(i, 1).getValue()+'"/>';
    }
  }
  return html;
}

function doGet() {
  var webhtml = HtmlService.createTemplateFromFile("index");
  webhtml.serviceUrl = ScriptApp.getService().getUrl();
  return webhtml.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getIdKeyName(form) {
  var a = getDataFromSheetById(form.sheetId, globalSheetName).getRange(1, 2).getValue();
  return a;
}

function getGsData(form) {
  var sheet = getDataFromSheetById(form.sheetId, globalSheetName);
  var dateKey = sheet.getRange(1, dataColIndex + 2).getValue();
  var idKey = sheet.getRange(1, 2).getValue();
  var json = sheetToJson(sheet, form);
  var formattedData = json.map(event => {
    event[dateKey] = Utilities.formatDate(event[dateKey], "GMT+8", "yyyy-MM-dd");
    return event;
  });

  formattedData = formattedData.sort(function(a, b) {
    return new Date(a[dateKey]).getTime() - new Date(b[dateKey]).getTime();
  });

  var uniqueEvents = formattedData.reduce((acc, event) => {
    var isUnique = acc.findIndex((e) => (
      e[dateKey] === event[dateKey]
      && String(e[idKey]) === String(event[idKey])
    )) === -1;
    if (isUnique) {
      acc.push(event);
    }
    return acc;
  }, []);
  return JSON.stringify(uniqueEvents);
}

function sheetToJson(sheet, form) {
  webSiteConfig();

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var data = sheet.getRange(1, 2, lastRow, lastColumn - 1).getValues();

  var headers = data.shift();
  var conditions = [];

  if (form.id) {
    conditions.push(function (row) {
      var rowDate = new Date(row[dataColIndex]);
      if (form.filter == "ALL") {
        return row[0] == form.id;
      } else {
        return row[0] == form.id && rowDate >= new Date(startDate) && rowDate <= new Date(endDate);
      }
    });
  }

  return data.filter(function (row) {
    return conditions.every(function (condition) { return condition(row); });
  }).map(function (row) {
    return headers.reduce(function (obj, key, index) {
      obj[key] = row[index];
      return obj;
    }, {});
  });
}

function getEvents(form) {
  webSiteConfig();
  var sheet = getDataFromSheetById(form.sheetId, globalSheetName);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 2, lastRow - 1, dataColIndex + 1).getValues();

  var events = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] != '' && data[i][0] == form.id
      && new Date(data[i][dataColIndex]) >= new Date(startDate)
      && new Date(data[i][dataColIndex]) <= new Date(endDate)) {
      var event = {};
      event.id = data[i][0] + " 每日簽到";
      event.title = data[i][0] + " 每日簽到";
      var formattedDate = Utilities.formatDate(data[i][dataColIndex], "GMT+8", "yyyy-MM-dd");
      event.start = formattedDate;
      events.push(event);
    }
  }

  var sportsData = getDataFromSheetById(sportSignSheetId, globalSheetName).getDataRange().getValues();
  for (var i = 1; i < sportsData.length; i++) {
    if (sportsData[i][1] != '' && sportsData[i][1] == form.id
      && new Date(sportsData[i][0]) >= new Date(startDate)
      && new Date(sportsData[i][0]) <= new Date(endDate)) {
      var event = {};
      event.id = sportsData[i][1] + " 運動簽到";
      event.title = sportsData[i][3].substring(0, sportsData[i][3].length - 1);
      var formattedDate = Utilities.formatDate(sportsData[i][0], "GMT+8", "yyyy-MM-dd");
      event.start = formattedDate;
      event.color = 'pink';
      events.push(event);
    }
  }


  var uniqueEvents = events.filter((event, index, self) =>
    index === self.findIndex((e) => (
      e.title === event.title && e.start === event.start
    ))
  );

  return uniqueEvents;
}

function getSticyNotes(form) {
  webSiteConfig();
  var sheet = getDataFromSheetById(memoSheetId, globalSheetName);
  var data = getSticyDatas(sheet, form.id);
  var lastRow = data.length;
  var sticyNoteHtml = "<div class=\"col-xs-12 mx-auto\">";
  var color = ["yellow", "pink", "green", "purple", "orange", "blue"]
  var min = 0;
  var max = 5;
  var i = 0;
  var usedNumCol = [];
  while (i < 30 && i < lastRow) {
    if (lastRow > 30) {
      var rowRandNum = getRandNum(lastRow, usedNumCol);
      usedNumCol.push(rowRandNum);
    } else {
      var rowRandNum = i;
    }
    i++;
    if (data[rowRandNum].content != '') {
      if (data[rowRandNum].content.length > 100) {
        sticyNoteHtml += "<div class=\"sticky-note sticky-note-" + color[Math.floor(Math.random() * (max - min + 1)) + min] + "\">";
        sticyNoteHtml += data[rowRandNum].content.substring(0, 100) + "...";
        sticyNoteHtml += "<br><br> by " + data[rowRandNum].name;
      } else if (data[rowRandNum].content.length < 5) {
        continue;
      } else {
        sticyNoteHtml += "<div class=\"sticky-note sticky-note-" + color[Math.floor(Math.random() * (max - min + 1)) + min] + "\">";
        sticyNoteHtml += data[rowRandNum].content;
        sticyNoteHtml += "<br><br> by " + data[rowRandNum].name;
      }
      sticyNoteHtml += "</div>";
    }
  }
  sticyNoteHtml += "</div>";
  return sticyNoteHtml;
}

function getRandNum(lastRow, usedNumCol) {
  var randomNumber = Math.floor(Math.random() * ((lastRow - 1) + 1));
  if (usedNumCol.includes(randomNumber)) {
    getRandNum(lastRow, usedNumCol);
  } else {
    return randomNumber;
  }
}

function getSticyDatas(sheet, id) {
  var postName = 1; //欄位是 你是誰
  var postId = 2; //欄位 發送人的學號
  var postC = 3 //欄位是要留言的對象是？
  var contetAll = 4; //發送的是班級
  var contetAngel = 5; //小天使
  var receiverId = 6; //收到訊息的同學
  var contetStudent = 7;//給同學的訊息
  var datas = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < datas.length; i++) {
    var result = {
      "content": "",
      "name": ""
    };
    switch (datas[i][postC]) {
      case "班上":
        if (String(datas[i][postId]).substring(0, 4) == id.substring(0, 4) ||
          datas[i][postId] == "校長" ||
          datas[i][postId] == "助教" ||
          datas[i][postId] == "班長") {
          result.content = datas[i][contetAll];
          result.name = datas[i][postName];
          results.push(result);
        }
        break;
      case "小天使們":
        if (String(datas[i][postId]).substring(0, 4) == id.substring(0, 4) ||
          datas[i][postId] == "校長" ||
          datas[i][postId] == "助教" ||
          datas[i][postId] == "班長") {
          result.content = datas[i][contetAngel];
          result.name = datas[i][postName];
          //results.push(result);
        }
        break;
      case "同學":
        if (String(datas[i][receiverId]) == id) {
          result.content = datas[i][contetStudent];
          result.name = datas[i][postName];
          results.push(result);
        }
        break;
    }

  }
  return results;
}

function webSiteConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalConfig);
  var iRow = 2;
  var lastRow = sheet.getLastRow();
  var dateRanges = [];
  while (iRow <= lastRow) {
    switch (sheet.getRange(iRow, 2).getValue()) {
      case "簽到課任務達標":
        coachCount = sheet.getRange(iRow, 1).getValue();
        break;
      case "運動簽到表":
        sportSignSheetId = sheet.getRange(iRow, 1).getValue();
        break;
      case "愛的小紙條":
        memoSheetId = sheet.getRange(iRow, 1).getValue();
        break;
      case "第一個月計算開始日":
      case "第二個月計算開始日":
      case "第三個月計算開始日":
        dateRanges.push(
          {
            start: Moment.moment(sheet.getRange(iRow, 1).getValue()).format("YYYY/MM/DD"),
            end: Moment.moment(sheet.getRange(iRow + 1, 1).getValue()).format("YYYY/MM/DD")
          });
        break;
    }
    iRow++
  }

  for (var i = 0; i < dateRanges.length; i++) {
    if (currentDate >= dateRanges[i].start
      && currentDate <= dateRanges[i].end) {
      startDate = dateRanges[i].start;
      endDate = dateRanges[i].end;
      break;
    }
  }
}
