let YandexOAuth = "OAuth y0_AgAEA7qkROu-AAh2tAAAAADQVO9rcCj-s8BUTDiA6mUThfhsv3Mt74A";
let BitrixApi = "https://truecode.bitrix24.ru/rest/623/oa5m3k5qshjzzi3r/";

async function request(restApi) {
  const response = await fetch(
    BitrixApi + restApi
  );
  return await response.json();
}
let data = {};
data.table = [];

function unique(arr) {
  let result = [];

  for (let str of arr) {
    if (!result.includes(str)) {
      result.push(str);
    }
  }
  return result;
}

function send() {
  /* generate worksheet and workbook */
  const worksheet = XLSX.utils.json_to_sheet(data.table);
  const workbook = XLSX.utils.book_new();
  worksheet["!cols"] = [{wch: 26}, {wch: 23}, {wch: 40}, {wch: 25},];
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");

  /* fix headers */
  XLSX.utils.sheet_add_aoa(
    worksheet,
    [
      [
        "Наименование задачи",
        "Наименование проекта",
        "Почта пользователя",
        "ФИО пользователя",
        "Фактическое затраченное время",
      ],
    ],
    { origin: "A1" }
  );

  /* create an XLSX file and try to save to text.xlsx */
  // XLSX.writeFile(workbook, "text.xlsx");

  var xlsbin = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "binary",
  });

  // в BLOB
  var buffer = new ArrayBuffer(xlsbin.length),
    array = new Uint8Array(buffer);
  for (var i = 0; i < xlsbin.length; i++) {
    array[i] = xlsbin.charCodeAt(i) & 0xff;
  }
  var xlsblob = new Blob([buffer], { type: "application/octet-stream" });
  delete array;
  delete buffer;
  delete xlsbin;

  var currentdate = new Date();
  var datetime =
    currentdate.getDate() +
    "." +
    (currentdate.getMonth() + 1) +
    "." +
    currentdate.getFullYear();

  const HttpUsers = new XMLHttpRequest();
  let url = `https://cloud-api.yandex.net/v1/disk/resources/upload?overwrite=true&path=/Отчёт ${datetime}.xlsx`;
  HttpUsers.open("GET", url, false);
  HttpUsers.setRequestHeader(
    "Authorization",
    YandexOAuth
  );
  HttpUsers.send();
  var obj = JSON.parse(HttpUsers.responseText);

  const upload = (file) => {
    fetch(obj.href, {
      method: "PUT",
      body: file, // объект файла
    })
      .then(        
        (response) => console.log(response) // if the response is a JSON object
      )
      .then(
        (success) => console.log(success) // Handle the success response object
      )
      .catch(
        (error) => console.log(error) // Handle the error response object
      );
  };

  upload(xlsblob);
}

function isLater(dateString1, dateString2) {
  return dateString1 <= dateString2
}

async function main() {
    data = {}
    data.table = []

    let startDate = document.getElementsByClassName('start_date')[0].value
    let closeDate = document.getElementsByClassName('end_date')[0].value

    // task list
    const tasklist = await request("tasks.task.list.json");
  
    console.log(tasklist);
  
    // list all task ElapsedTime
    console.log("ElapsedTime");
    for (const task of tasklist.result.tasks) {
      let userEmail;
      let userName;
      let elapsedTimes;
  
      // get all elapsedTime
      const elapsedTime = await request(
        "task.elapseditem.getlist?TASKID=" + task.id
      );
  
      //console.log(elapsedTime)
      console.log("name task: " + task.title);
      let taskTitle = task.title;
      
      console.log("start = " + task.createdDate)
      console.log("close = " + task.closedDate)
  
      console.log("project: " + task.group.name);
      let ProjectName = task.group.name;
      // get users on task
      for (const user of elapsedTime.result) {

        let getCreatedData
        let getClosedData

        let now = new Date()
        let nowYaer = now.getFullYear()
        let nowMonth = now.getMonth() + 1
        let nowDate = now.getDate()
        let getNowData = nowYaer + "-" + nowMonth + "-" + nowDate

        if (startDate == "") {startDate = "2015-12-24"}
        if (closeDate == "") {closeDate = getNowData}

        if (task.createdDate == null) { getCreatedData = startDate }          
        else { getCreatedData = task.createdDate.split('T')[0] }

        if (task.closedDate == null) { getClosedData = getNowData}
        else { getClosedData = task.closedDate.split('T')[0]}

        console.log(startDate + "  " + getCreatedData)
        console.log(isLater(startDate, getCreatedData))
        console.log(getClosedData + "  " + closeDate)
        console.log(isLater(getClosedData, closeDate))

        if (isLater(startDate, getCreatedData) && isLater(getClosedData, closeDate)) {
          const userInfo = await request("user.get?ID=" + user.USER_ID);
          //console.log(userInfo)
          for (const about of userInfo.result) {
            console.log(about.EMAIL);
            userEmail = about.EMAIL;
            console.log(about.LAST_NAME + " " + about.NAME);
            userName = about.LAST_NAME + " " + about.NAME;
          }
    
          if (user.MINUTES >= 60) {
            console.log(
              ((user.MINUTES / 60) | 0) + " ч " + (user.MINUTES % 60) + " мин"
            );
            elapsedTimes =
              ((user.MINUTES / 60) | 0) + " ч " + (user.MINUTES % 60) + " мин";
          } else {
            console.log(user.MINUTES + " мин");
            elapsedTimes = user.MINUTES + " мин";
          }
          console.log("/------------/");
    
          let a = {
            Task_Title: taskTitle,
            Project_Name: ProjectName,
            User_email: userEmail,
            User_FIO: userName,
            Elapsed_Time: elapsedTimes,
          };
    
          data.table.push(a);
        }
      }
  
      console.log("-----------------------------------");
    }

    data.table.push("")
  
    var emails = [];
    var names = [];
  
    for (let i = 0; i < data.table.length; i++) {
      emails[i] = data.table[i].User_email;
    }
  
    let UniqueEmails = unique(emails);
  
    for (let i = 0; i < UniqueEmails.length; i++) {
      for(let j = 0; j < data.table.length; j++) {
        if(UniqueEmails[i] == data.table[j].User_email) {
          names[i] = data.table[j].User_FIO
        }
      }
    }

    for (let i = 0; i < UniqueEmails.length; i++) {
      let TimeCounter = 0;
      let AllTime = 0;
      let AllTimeSplit = 0;
      for (let j = 0; j < data.table.length; j++) {
        if (UniqueEmails[i] == data.table[j].User_email) {
          AllTime += data.table[j].Elapsed_Time + " ";
        }
      }
      AllTime = AllTime.replace("undefined", "");
      AllTimeSplit = AllTime.split(" ");
  
      for (let j = 0; j < AllTimeSplit.length; j++) {
        if (AllTimeSplit[j] == "ч") {
          TimeCounter += Number(AllTimeSplit[j - 1] * 60);
        } else if (AllTimeSplit[j] == "мин") {
          TimeCounter += Number(AllTimeSplit[j - 1]);
        }
      }
      TimeCounter =
        ((TimeCounter / 60) | 0) + " ч " + (TimeCounter % 60) + " мин";
  
      let a = {
        Project_Name: names[i],
        User_email: UniqueEmails[i],
        User_FIO: "Итого: ",
        Elapsed_Time: TimeCounter,
      };
  
      data.table.push(a);
    }
  
    send();
  }
  