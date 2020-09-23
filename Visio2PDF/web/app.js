// Get Folders from User Input... Tkinter Callback in Python
let dir_path;
let cover_path;
let insert_version_tag;
let previewData;

// SECTION: Settings
eel.expose(setSettings);
function setSettings(settings) {
  // Dark Mode
  if (settings["defaultMode"] == "dark") {
    document.getElementById("darkSwitch").checked = true;
    document.getElementById("main-body").setAttribute("data-theme", "dark");
    eel.write_settings({ defaultMode: "dark" });
  } else {
    document.getElementById("darkSwitch").checked = false;
  }
}

// SECTION: Version Information
eel.expose(setVersion);
function setVersion(version) {
  document.getElementById("app_version").textContent = version;
}

eel.expose(setRepoVersionNotify);
function setRepoVersionNotify(upToDate) {
  if (upToDate != "True") {
    document.getElementById("version_alert").style.display = "block";
  }
}

// !SECTION: Version Information

// SECTION: Folder / Cover Selection
async function getFolder() {
  dir_path = await eel.btn_SelectDir()();
  let dir_div = document.getElementById("dir_name");
  dir_div.value = dir_path;

  // Calls attempts to find a cover sheetsetHis
  cover_path = await eel.find_cover(dir_path)();
  if (cover_path !== false) {
    let cover_div = document.getElementById("cover_name");
    cover_div.value = cover_path;
  }
}

async function getCover() {
  cover_path = await eel.btn_SelectCoverSheet()();
  let cover_div = document.getElementById("cover_name");
  cover_div.value = cover_path;
}
// !SECTION

// SECTION: Togglers
eel.expose(toggleButtons);
function toggleButtons() {
  document.getElementById("convert_btn").style.display = "block";
  document.getElementById("working_indicator").style.display = "none";
}

function toggleVersioning() {
  var current_val = document.getElementById("version_tag").disabled;

  if (current_val === true) {
    document.getElementById("version_tag").disabled = false;
    document.getElementById("engineer_name").disabled = false;
    document.getElementById("version_tag").required = true;
    document.getElementById("engineer_name").required = true;
    document.getElementById("engineer_name").placeholder =
      "Required: John C. Doe";
    document.getElementById("version_tag").placeholder = "Required: v1.01";
  } else {
    document.getElementById("version_tag").disabled = true;
    document.getElementById("engineer_name").disabled = true;
    document.getElementById("version_tag").required = false;
    document.getElementById("engineer_name").required = false;
    document.getElementById("version_tag").placeholder = "";
    document.getElementById("engineer_name").placeholder = "";
  }
}

eel.expose(setMsgVisible);
function setMsgVisible() {
  document.getElementById("msg").style.display = "block";
  document.getElementById("working_indicator").style.display = "none";
}

// SECTION: Form Validators
function validateForm() {
  var sys_name = document.getElementById("sys_name").value;
  var insert_version_tag = document.getElementById("enable_tagging").checked;
  var version_tag = document.getElementById("version_tag").value;
  var engineer_name = document.getElementById("engineer_name").valued;

  if (insert_version_tag === true) {
    var validateArray = [dir_path, sys_name, version_tag, engineer_name];
  } else {
    var validateArray = [dir_path, sys_name];
  }

  var isValid = validateEntry(validateArray);
  return isValid;
}

function validateEntry(entry_array) {
  var falseCheck;
  entry_array.forEach(function (item, _index, _array) {
    if (item === "") {
      falseCheck = false;
    } else if (item === "undefined") {
      falseCheck = false;
    }
  });
  if (falseCheck === false) {
    return false;
  } else {
    return true;
  }
}

// !SECTION: Form Validators
function getJobData() {
  if (validateForm() === false) {
    document.getElementById("required_fields_alert").style.display = "block";
    return;
  }
  document.getElementById("required_fields_alert").style.display = "none";
  document.getElementById("msg").style.display = "none";

  var jobData = {
    name: document.getElementById("sys_name").value,
    directory: dir_path,
    includeSubDir: document.getElementById("merge_subdirs").checked,
    coverSheet: cover_path,
    versionTagging: document.getElementById("enable_tagging").checked,
    versionData: {
      authorName: document.getElementById("engineer_name").value,
      versionTag: document.getElementById("version_tag").value,
    },
    fileTypes: {
      visio: document.getElementById("include_visio").checked,
      excel: document.getElementById("include_excel").checked,
      word: document.getElementById("include_word").checked,
      powerpoint: document.getElementById("include_powerpoint").checked,
      publisher: document.getElementById("include_publisher").checked,
      outlook: document.getElementById("include_outlook").checked,
      project: document.getElementById("include_project").checked,
      openOffice: document.getElementById("include_openoffice").checked,
    },
    saveJob: document.getElementById("save_job").checked,
  };

  return jobData;
}

// Preview Functions
function JSONtoTable(json, location) {
  var col = [];
  for (var i = 0; i < json.length; i++) {
    for (var key in json[i]) {
      if (col.indexOf(key) === -1) {
        col.push(key);
      }
    }
  }

  // CREATE DYNAMIC TABLE.
  var table = document.createElement("table");
  table.setAttribute("style", "width: 1000px;");

  // CREATE HTML TABLE HEADER ROW USING THE EXTRACTED HEADERS ABOVE.

  var tr = table.insertRow(-1); // TABLE ROW.

  for (var i = 0; i < col.length; i++) {
    var th = document.createElement("th"); // TABLE HEADER.
    th.setAttribute("scope", "col");
    th.innerHTML = col[i];
    tr.appendChild(th);
  }

  var th = document.createElement("th"); // TABLE HEADER.
  th.setAttribute("scope", "col");
  th.innerHTML = "Import";
  tr.appendChild(th);

  // ADD JSON DATA TO THE TABLE AS ROWS.
  for (var i = 0; i < json.length; i++) {
    tr = table.insertRow(-1);

    for (var j = 0; j < col.length; j++) {
      var tabCell = tr.insertCell(-1);

      if (json[i][col[j]] == "history") {
        // History
        tabCell.innerHTML = `<a id="history-${
          i + 1
        }" onclick="runHistory(${i})" type="button" class="btn btn btn-danger btn-sm"> Run </a>`;

        // Import History
        tabCell = tr.insertCell(-1);
        tabCell.innerHTML = `<a id="history-import${
          i + 1
        }" onclick="importHistory(${i})" type="button" class="btn btn btn-danger btn-sm"> Import </a>`;
        tabCell.setAttribute("class", "align-middle");

        // Saved
      } else if (json[i][col[j]] == "saved") {
        tabCell.innerHTML = `<a id="saved-${
          i + 1
        }" onclick="runSaved(${i})" type="button" class="btn btn btn-danger btn-sm"> Run </a>`;

        // Import Saved
        tabCell = tr.insertCell(-1);
        tabCell.innerHTML = `<a id="history-import${
          i + 1
        }" onclick="importSaved(${i})" type="button" class="btn btn btn-danger btn-sm"> Import </a>`;
        tabCell.setAttribute("class", "align-middle");
      } else {
        tabCell.innerHTML = json[i][col[j]];
      }
      tabCell.setAttribute("class", "align-middle");
    }
  }

  // FINALLY ADD THE NEWLY CREATED TABLE WITH JSON DATA TO A CONTAINER.
  var divContainer = document.getElementById(location);
  divContainer.innerHTML = "";
  divContainer.appendChild(table);
}

eel.expose(setWorking);
function setWorking() {
  document.getElementById("convert_btn").style.display = "none";
  document.getElementById("working_indicator").style.display = "block";
}

// Primary Function
async function convertVisio() {
  var jobData = getJobData();

  eel.main(jobData);
  setWorking();
}

// SECTION: Logger
function showLog() {
  var currentState = document.getElementById("output").style.display;

  if (currentState === "none") {
    document.getElementById("output").style.display = "block";
  } else {
    document.getElementById("output").style.display = "none";
  }
}

eel.expose(putMessageInOutput);
function putMessageInOutput(message) {
  const outputNode = document.querySelector("#output textarea");
  outputNode.value += message; // Add the message
  if (!message.endsWith("\n")) {
    outputNode.value += "\n"; // If there was no new line, add one
  }

  // Set the correct height to fit all the output and then scroll to the bottom
  // outputNode.style.height = 'auto';
  outputNode.style.height = outputNode.scrollHeight + 10 + "px";
  // window.scrollTo(0, document.body.scrollHeight);
}

eel.expose(signalPackagingComplete);
function signalPackagingComplete(successful) {
  setPackagingComplete(successful);

  window.scrollTo(0, document.body.scrollHeight);
}

// !SECTION: Logger

//  --------------------  //
//  SECTION: Visibility  //
// ---------------------  //

function toggleVis(elementID) {
  var eleState = document.getElementById(elementID).style.display;
  if (eleState == "block") {
    show(elementID);
  } else {
    hide(elementID);
  }
}

function show(elementID) {
  console.log(elementID);
  document.getElementById(elementID).style.display = "block";
}

function hide(elementID) {
  document.getElementById(elementID).style.display = "none";
}

//  --------------------  //
//  SECTION: Bottom Menu  //
// ---------------------  //

// Change Class of Button based onclick
function menuSelect(elementID, showID) {
  console.log(showID);
  var activeBtn = "btn btn-secondary btn-sm btn-danger";
  var inactiveBtn = "btn btn-secondary btn-sm";

  // Hide Menus
  var detailsList = document
    .getElementById("menu-details")
    .querySelectorAll(".menu-details");
  var menus = [];

  for (var i = 0, len = detailsList.length; i < len; i++) {
    detailsList[i].style.display = "none";
  }

  // Set Button Class and Show Menu
  var children = document.getElementById("bottom-menu").children;
  var ids = [];

  for (var i = 0, len = children.length; i < len; i++) {
    ids.push(children[i].id);
  }

  ids.forEach(function (item, index) {
    if (item == elementID) {
      document.getElementById(item).setAttribute("class", activeBtn);
      console.log(showID);
      show(showID);

      if (elementID == "preview-btn") {
        getPreview();
      }
    } else {
      document.getElementById(item).setAttribute("class", inactiveBtn);
    }
  });
}

function setFromImport(importData) {
  document.getElementById("sys_name").value = importData["name"];

  dir_path = importData["directory"];
  document.getElementById("dir_name").value = importData["directory"]

  document.getElementById("merge_subdirs").value = importData["includeSubDir"];

  cover_path = importData["coverSheet"];
  document.getElementById("cover_name").value = importData["coverSheet"]

  document.getElementById("enable_tagging").value = importData["versionTagging"];

  if (importData["versionTagging"] == true) {
    document.getElementById("engineer_name").value = importData["versionData"]["authorName"];
    document.getElementById("version_tag").value = importData["versionData"]["versionTag"];
  }
  document.getElementById("include_visio").checked = importData["fileTypes"]["include_visio"]
  document.getElementById("include_excel").checked = importData["fileTypes"]["include_excel"]
  document.getElementById("include_word").checked = importData["fileTypes"]["include_word"]
  document.getElementById("include_powerpoint").checked = importData["fileTypes"]["include_powerpoint"]
  document.getElementById("include_publisher").checked = importData["fileTypes"]["include_publisher"]
  document.getElementById("include_outlook").checked = importData["fileTypes"]["include_outlook"]
  document.getElementById("include_project").checked = importData["fileTypes"]["include_project"]
  document.getElementById("include_openoffice").checked = importData["fileTypes"]["include_openoffice"]

  document.getElementById("save_job").checked = importData["saveJob"]
}

// Set History Values, Called from Python
eel.expose(setHistory);
function setHistory(jobs) {
  var jobs = JSON.parse(jobs);
  console.log(jobs);

  JSONtoTable(jobs, "history");
}

function runHistory(index) {
  eel.run_history(index);
}

async function importHistory(index) {
  data = await eel.import_history(index)
  setFromImport(data)
}

// Set Saved Values, Called from Python
eel.expose(setSaved);
function setSaved(jobs) {
  var jobs = JSON.parse(jobs);
  console.log(jobs);

  JSONtoTable(jobs, "saved");
}

function runSaved(index) {
  eel.run_saved(index)
}

async function importSaved(index) {
  data = await eel.import_saved(index)()
  setFromImport(data)
}


// Calls Preview data from Python
async function getPreview() {
  var jobData = getJobData();

  previewData = await eel.getPreview(jobData)();

  previewData = JSON.parse(previewData);

  JSONtoTable(previewData, "preview");
}
