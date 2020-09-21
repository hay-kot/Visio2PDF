// Get Folders from User Input... Tkinter Callback in Python
let dir_path;
let cover_path;
let insert_version_tag;
let previewData;

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

  // Calls attempts to find a cover sheet
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

function showAdvSetting() {
  var currentState = document.getElementById("advanced_settings").style.display;

  if (currentState === "none") {
    document.getElementById("advanced_settings").style.display = "block";
  } else {
    document.getElementById("advanced_settings").style.display = "none";
  }
}

function showPreview() {
  var currentState = document.getElementById("preview").style.display;

  if (currentState === "none") {
    document.getElementById("preview").style.display = "block";
  } else {
    document.getElementById("preview").style.display = "none";
  }
}
// !SECTION: Togglers

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
  console.log(validateArray);

  var isValid = validateEntry(validateArray);
  console.log(isValid);
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
  console.log(falseCheck);
  if (falseCheck === false) {
    return false;
  } else {
    return true;
  }
}
// !SECTION: Form Validators

// Preview Functions
async function getPreview() {
  if (validateForm() === false) {
    document.getElementById("required_fields_alert").style.display = "block";
    return;
  }
  document.getElementById("required_fields_alert").style.display = "none";
  document.getElementById("msg").style.display = "none";

  let sys_name = document.getElementById("sys_name").value;
  let version_tag = document.getElementById("version_tag").value;
  let engineer_name = document.getElementById("engineer_name").value;
  let include_subdir = document.getElementById("merge_subdirs").checked;

  let insert_version_tag = document.getElementById("enable_tagging").checked;

  let filetypes = {
    visio: document.getElementById("include_visio").checked,
    excel: document.getElementById("include_excel").checked,
    word: document.getElementById("include_word").checked,
    powerpoint: document.getElementById("include_powerpoint").checked,
    publisher: document.getElementById("include_publisher").checked,
    outlook: document.getElementById("include_outlook").checked,
    project: document.getElementById("include_project").checked,
    openOffice: document.getElementById("include_openoffice").checked,
  };

  previewData = await eel.getPreview(
    dir_path,
    cover_path,
    sys_name,
    insert_version_tag,
    version_tag,
    engineer_name,
    include_subdir,
    filetypes
  )();

  previewData = JSON.parse(previewData);
  console.log(previewData);

  var col = [];
  for (var i = 0; i < previewData.length; i++) {
    for (var key in previewData[i]) {
      if (col.indexOf(key) === -1) {
        col.push(key);
        console.log(key);
      }
    }
  }

  console.log(col);

  // CREATE DYNAMIC TABLE.
  var table = document.createElement("table");

  // CREATE HTML TABLE HEADER ROW USING THE EXTRACTED HEADERS ABOVE.

  var tr = table.insertRow(-1); // TABLE ROW.

  for (var i = 0; i < col.length; i++) {
    var th = document.createElement("th"); // TABLE HEADER.
    th.setAttribute("scope", "col")
    th.innerHTML = col[i];
    tr.appendChild(th);
  }

  // ADD JSON DATA TO THE TABLE AS ROWS.
  for (var i = 0; i < previewData.length; i++) {
    tr = table.insertRow(-1);

    for (var j = 0; j < col.length; j++) {
      var tabCell = tr.insertCell(-1);
      tabCell.innerHTML = previewData[i][col[j]];
    }
  }

  // FINALLY ADD THE NEWLY CREATED TABLE WITH JSON DATA TO A CONTAINER.
  var divContainer = document.getElementById("preview");
  divContainer.innerHTML = "";
  divContainer.appendChild(table);

  showPreview();
}

// Primary Function
async function convertVisio() {
  if (validateForm() === false) {
    document.getElementById("required_fields_alert").style.display = "block";
    return;
  }
  document.getElementById("required_fields_alert").style.display = "none";
  document.getElementById("msg").style.display = "none";

  let sys_name = document.getElementById("sys_name").value;
  let version_tag = document.getElementById("version_tag").value;
  let engineer_name = document.getElementById("engineer_name").value;
  let include_subdir = document.getElementById("merge_subdirs").checked;

  let insert_version_tag = document.getElementById("enable_tagging").checked;

  let filetypes = {
    visio: document.getElementById("include_visio").checked,
    excel: document.getElementById("include_excel").checked,
    word: document.getElementById("include_word").checked,
    powerpoint: document.getElementById("include_powerpoint").checked,
    publisher: document.getElementById("include_publisher").checked,
    outlook: document.getElementById("include_outlook").checked,
    project: document.getElementById("include_project").checked,
    openOffice: document.getElementById("include_openoffice").checked,
  };

  eel.main(
    dir_path,
    cover_path,
    sys_name,
    insert_version_tag,
    version_tag,
    engineer_name,
    include_subdir,
    filetypes
  );
  document.getElementById("convert_btn").style.display = "none";
  document.getElementById("working_indicator").style.display = "block";
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

function CreateTableFromJSON(previewData) {
  console.log(`Length: (previewData.length)`);

  // EXTRACT VALUE FOR HTML HEADER.
}
